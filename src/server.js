require('dotenv').config();
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const crypto = require('crypto');

let fetchFn = globalThis.fetch;
if (!fetchFn) {
  fetchFn = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
}
const fetch = (...args) => fetchFn(...args);

const app = express();
app.use(express.json());
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // Limit uploads to 10MB
});

const PORT = process.env.PORT || 3000;
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'models/gemini-2.5-flash';
const GEMINI_API_BASE_URL = process.env.GEMINI_API_BASE_URL || 'https://generativelanguage.googleapis.com/v1beta';
const DATA_FILE = path.join(__dirname, '..', 'data', 'parsed_invoices.json');
const QUICKBOOKS_CLIENT_ID = process.env.QUICKBOOKS_CLIENT_ID;
const QUICKBOOKS_CLIENT_SECRET = process.env.QUICKBOOKS_CLIENT_SECRET;
const QUICKBOOKS_DEFAULT_CALLBACK_PATH = '/api/quickbooks/callback';
const QUICKBOOKS_DEFAULT_REDIRECT_URI = `http://localhost:${PORT}${QUICKBOOKS_DEFAULT_CALLBACK_PATH}`;
const QUICKBOOKS_REDIRECT_URI = process.env.QUICKBOOKS_REDIRECT_URI || QUICKBOOKS_DEFAULT_REDIRECT_URI;
const QUICKBOOKS_SCOPES = process.env.QUICKBOOKS_SCOPES || 'com.intuit.quickbooks.accounting';
const QUICKBOOKS_ENVIRONMENT = process.env.QUICKBOOKS_ENVIRONMENT || 'sandbox';
const QUICKBOOKS_API_BASE_URL = process.env.QUICKBOOKS_API_BASE_URL || (QUICKBOOKS_ENVIRONMENT === 'production' ? 'https://quickbooks.api.intuit.com' : 'https://sandbox-quickbooks.api.intuit.com');
const QUICKBOOKS_COMPANIES_FILE = path.join(__dirname, '..', 'data', 'quickbooks_companies.json');
const QUICKBOOKS_AUTH_URL = 'https://appcenter.intuit.com/connect/oauth2';
const QUICKBOOKS_TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
const QUICKBOOKS_STATE_TTL_MS = 10 * 60 * 1000;

const quickBooksStates = new Map();
const quickBooksCallbackPaths = getQuickBooksCallbackPaths();


function buildGeminiUrl() {
  const base = (GEMINI_API_BASE_URL || '').replace(/\/+$/, '');
  const modelPath = (GEMINI_MODEL || '').replace(/^\/+/, '');
  return `${base}/${modelPath}:generateContent`;
}

function getQuickBooksCallbackPaths() {
  const paths = new Set([QUICKBOOKS_DEFAULT_CALLBACK_PATH]);
  const derived = derivePathFromUrl(QUICKBOOKS_REDIRECT_URI);
  if (derived) {
    paths.add(derived);
  }
  return Array.from(paths);
}

function derivePathFromUrl(urlValue) {
  if (!urlValue) {
    return null;
  }
  try {
    return new URL(urlValue, `http://localhost:${PORT}`).pathname;
  } catch (error) {
    console.warn('Unable to determine QuickBooks callback path from QUICKBOOKS_REDIRECT_URI', error);
    return null;
  }
}

app.use(express.static(path.join(__dirname, '..', 'public')));

app.post('/api/parse-invoice', upload.single('invoice'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file received. Please attach an invoice.' });
    }

    if (!GEMINI_API_KEY) {
      return res.status(500).json({ error: 'Gemini API key is not configured. Set GEMINI_API_KEY in the environment.' });
    }

    const invoiceBuffer = req.file.buffer;
    const invoiceMetadata = {
      originalName: req.file.originalname,
      mimeType: req.file.mimetype,
      size: req.file.size,
      checksum: crypto.createHash('sha256').update(invoiceBuffer).digest('hex'),
    };

    const existingInvoices = await readStoredInvoices();
    const duplicateByChecksum = existingInvoices.find((entry) => entry.metadata?.checksum === invoiceMetadata.checksum);

    const parsedInvoice = await extractWithGemini(invoiceBuffer, req.file.mimetype);

    const enrichedInvoice = {
      parsedAt: new Date().toISOString(),
      metadata: invoiceMetadata,
      data: parsedInvoice,
    };

    const duplicateInvoice = findDuplicateInvoice(existingInvoices, enrichedInvoice.data);
    const duplicateMatch = duplicateInvoice || duplicateByChecksum || null;
    const duplicateReason = duplicateInvoice
      ? 'Matched existing invoice fields'
      : duplicateByChecksum
        ? 'File contents match an existing invoice (checksum)'
        : null;

    const storagePayload = {
      ...enrichedInvoice,
      duplicateOf: duplicateMatch?.metadata?.checksum || null,
    };

    await persistInvoice(existingInvoices, storagePayload);

    return res.json({
      invoice: enrichedInvoice,
      duplicate: duplicateMatch
        ? {
            reason: duplicateReason,
            match: duplicateMatch,
          }
        : null,
    });
  } catch (error) {
    console.error('Failed to process invoice', error);
    if (error.isGeminiError) {
      return res.status(502).json({ error: error.message });
    }
    return res.status(500).json({ error: 'Failed to parse invoice. Please try again.' });
  }
});

app.get('/api/quickbooks/companies', async (req, res) => {
  try {
    const companies = await readQuickBooksCompanies();
    const sanitized = companies.map(sanitizeQuickBooksCompany);
    res.json({ companies: sanitized });
  } catch (error) {
    console.error('Failed to load QuickBooks companies', error);
    res.status(500).json({ error: 'Failed to load QuickBooks connections.' });
  }
});

app.patch('/api/quickbooks/companies/:realmId', async (req, res) => {
  const realmId = req.params.realmId;
  const { companyName } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (typeof companyName !== 'string' || !companyName.trim()) {
    return res.status(400).json({ error: 'Company name must be a non-empty string.' });
  }

  try {
    const updated = await updateQuickBooksCompanyName(realmId, companyName.trim());
    res.json({ company: sanitizeQuickBooksCompany(updated) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to update QuickBooks company name', error);
    res.status(500).json({ error: 'Failed to update QuickBooks company name.' });
  }
});

app.get('/api/quickbooks/connect', (req, res) => {
  if (!QUICKBOOKS_CLIENT_ID || !QUICKBOOKS_CLIENT_SECRET) {
    return res.status(500).send('QuickBooks API credentials are not configured.');
  }

  const state = crypto.randomBytes(16).toString('hex');
  const now = Date.now();
  quickBooksStates.set(state, now);
  for (const [key, issued] of quickBooksStates.entries()) {
    if (now - issued > QUICKBOOKS_STATE_TTL_MS) {
      quickBooksStates.delete(key);
    }
  }

  const authorizeUrl = new URL(QUICKBOOKS_AUTH_URL);
  authorizeUrl.searchParams.set('client_id', QUICKBOOKS_CLIENT_ID);
  authorizeUrl.searchParams.set('response_type', 'code');
  authorizeUrl.searchParams.set('scope', QUICKBOOKS_SCOPES);
  authorizeUrl.searchParams.set('redirect_uri', QUICKBOOKS_REDIRECT_URI);
  authorizeUrl.searchParams.set('state', state);
  authorizeUrl.searchParams.set('prompt', 'consent');

  res.redirect(authorizeUrl.toString());
});

quickBooksCallbackPaths.forEach((callbackPath) => {
  app.get(callbackPath, handleQuickBooksCallback);
});

async function handleQuickBooksCallback(req, res) {
  const { code, state, realmId } = req.query;
  if (!code || !state || !realmId) {
    return res.status(400).send('Missing required parameters from QuickBooks.');
  }

  const stateIssuedAt = quickBooksStates.get(state);
  if (!stateIssuedAt || Date.now() - stateIssuedAt > QUICKBOOKS_STATE_TTL_MS) {
    return res.status(400).send('Authorization state is invalid or has expired. Please try connecting again.');
  }
  quickBooksStates.delete(state);

  try {
    const tokenSet = await exchangeQuickBooksCode(code);

    let companyName = realmId;
    try {
      const companyInfo = await fetchQuickBooksCompanyInfo(realmId, tokenSet.accessToken);
      companyName = companyInfo?.CompanyInfo?.CompanyName || companyName;
    } catch (infoError) {
      console.warn('QuickBooks company info lookup failed', infoError);
      if (infoError.isQuickBooksError && infoError.status === 403) {
        companyName = formatFallbackCompanyName(realmId);
      } else {
        throw infoError;
      }
    }

    await storeQuickBooksCompany({
      realmId,
      companyName,
      tokens: tokenSet,
    });

    return res.redirect(`/?quickbooks=connected&company=${encodeURIComponent(companyName)}`);
  } catch (error) {
    console.error('QuickBooks OAuth callback failed', error);
    return res.redirect('/?quickbooks=error');
  }
}

app.use((req, res) => {
  res.status(404).json({ error: 'Not found' });
});

app.listen(PORT, () => {
  console.log(`Invoice parser listening on http://localhost:${PORT}`);
});

async function exchangeQuickBooksCode(code) {
  if (!QUICKBOOKS_CLIENT_ID || !QUICKBOOKS_CLIENT_SECRET) {
    throw new Error('QuickBooks API credentials are not configured.');
  }

  const params = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: QUICKBOOKS_REDIRECT_URI,
  });

  const authHeader = Buffer.from(`${QUICKBOOKS_CLIENT_ID}:${QUICKBOOKS_CLIENT_SECRET}`).toString('base64');
  const response = await fetch(QUICKBOOKS_TOKEN_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      Authorization: `Basic ${authHeader}`,
      Accept: 'application/json',
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message =
      errorBody?.error_description ||
      errorBody?.error ||
      `QuickBooks token exchange failed with status ${response.status}`;
    const err = new Error(message);
    err.isQuickBooksError = true;
    throw err;
  }

  const data = await response.json();
  const now = Date.now();

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    scope: data.scope,
    tokenType: data.token_type,
    expiresAt: data.expires_in ? new Date(now + data.expires_in * 1000).toISOString() : null,
    refreshTokenExpiresAt: data.refresh_token_expires_in
      ? new Date(now + data.refresh_token_expires_in * 1000).toISOString()
      : null,
  };
}

async function fetchQuickBooksCompanyInfo(realmId, accessToken) {
  const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/companyinfo/${realmId}?minorversion=65`;
  const response = await fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      'User-Agent': 'InvoiceToQB/1.0',
    },
  });

  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message =
      errorBody?.Fault?.Error?.[0]?.Detail ||
      errorBody?.Fault?.Error?.[0]?.Message ||
      `QuickBooks company info request failed with status ${response.status}`;
    const err = new Error(message);
    err.isQuickBooksError = true;
    err.status = response.status;
    throw err;
  }

  return response.json();
}

function formatFallbackCompanyName(realmId) {
  return `QuickBooks Company #${realmId}`;
}

async function readQuickBooksCompanies() {
  try {
    const file = await fs.readFile(QUICKBOOKS_COMPANIES_FILE, 'utf-8');
    return JSON.parse(file);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return [];
    }
    throw err;
  }
}

async function storeQuickBooksCompany({ realmId, companyName, tokens }) {
  const companies = await readQuickBooksCompanies();
  const now = new Date().toISOString();
  const existingIndex = companies.findIndex((entry) => entry.realmId === realmId);

  const record = {
    realmId,
    companyName,
    environment: QUICKBOOKS_ENVIRONMENT,
    connectedAt: existingIndex >= 0 ? companies[existingIndex].connectedAt : now,
    updatedAt: now,
    tokens,
  };

  if (existingIndex >= 0) {
    companies[existingIndex] = { ...companies[existingIndex], ...record };
  } else {
    companies.push(record);
  }

  await persistQuickBooksCompanies(companies);
}

async function persistQuickBooksCompanies(companies) {
  await fs.mkdir(path.dirname(QUICKBOOKS_COMPANIES_FILE), { recursive: true });
  await fs.writeFile(QUICKBOOKS_COMPANIES_FILE, JSON.stringify(companies, null, 2));
}

async function updateQuickBooksCompanyName(realmId, companyName) {
  const companies = await readQuickBooksCompanies();
  const index = companies.findIndex((entry) => entry.realmId === realmId);

  if (index < 0) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const now = new Date().toISOString();
  const record = {
    ...companies[index],
    companyName,
    updatedAt: now,
  };

  companies[index] = record;
  await persistQuickBooksCompanies(companies);
  return record;
}

function sanitizeQuickBooksCompany(company) {
  if (!company) {
    return company;
  }

  const { tokens, ...rest } = company;
  return rest;
}

async function extractWithGemini(buffer, mimeType) {
  const url = buildGeminiUrl();
  const prompt = `You are an expert system for reading invoices. Extract the following fields and respond with **only** valid JSON matching this schema:
{
  "vendor": string | null,
  "products": [
    {
      "description": string,
      "quantity": number | null,
      "unitPrice": number | null,
      "lineTotal": number | null,
      "taxCode": string | null
    }
  ],
  "invoiceDate": string | null, // ISO 8601 (YYYY-MM-DD) when possible
  "invoiceNumber": string | null,
  "billTo": string | null,
  "subtotal": number | null,
  "vatAmount": number | null,
  "totalAmount": number | null,
  "taxCode": string | null,
  "currency": string | null // 3-letter ISO code when available
}
Rules:
- Use numbers for monetary values without currency symbols.
- If a field is missing, set it to null.
- Combine multiple bill recipients into a single string if necessary.
- For products, include at least a description. Omit other product fields if not present by setting them to null.
- Do not add commentary. Return only minified JSON with the exact fields above.`;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          { text: prompt },
          {
            inlineData: {
              data: buffer.toString('base64'),
              mimeType: mimeType || 'application/pdf',
            },
          },
        ],
      },
    ],
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': GEMINI_API_KEY,
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message = errorBody?.error?.message || `Gemini API request failed with status ${response.status}`;
    const err = new Error(message);
    err.isGeminiError = true;
    throw err;
  }

  const body = await response.json();
  const candidateText = body?.candidates?.[0]?.content?.parts?.map((part) => part.text).join('').trim();

  if (!candidateText) {
    const err = new Error('Gemini API returned no content to parse.');
    err.isGeminiError = true;
    throw err;
  }

  const jsonText = extractJson(candidateText);
  if (!jsonText) {
    const err = new Error('Unable to extract JSON from Gemini response.');
    err.isGeminiError = true;
    throw err;
  }

  return JSON.parse(jsonText);
}

function extractJson(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) {
    return null;
  }
  return text.substring(start, end + 1);
}

async function safeReadJson(response) {
  try {
    return await response.json();
  } catch (err) {
    return null;
  }
}

async function readStoredInvoices() {
  try {
    const file = await fs.readFile(DATA_FILE, 'utf-8');
    return JSON.parse(file);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return [];
    }
    throw err;
  }
}

async function persistInvoice(existingInvoices, invoice) {
  const updated = [...existingInvoices, invoice];
  await fs.mkdir(path.dirname(DATA_FILE), { recursive: true });
  await fs.writeFile(DATA_FILE, JSON.stringify(updated, null, 2));
}

function normalise(value) {
  return (value || '').toString().trim().toLowerCase();
}

function findDuplicateInvoice(existingInvoices, candidate) {
  if (!candidate) {
    return null;
  }

  const candidateInvoiceNumber = normalise(candidate.invoiceNumber);
  if (candidateInvoiceNumber) {
    const invoiceNumberMatch = existingInvoices.find((entry) => normalise(entry?.data?.invoiceNumber) === candidateInvoiceNumber);
    if (invoiceNumberMatch) {
      return invoiceNumberMatch;
    }
  }

  const candidateVendor = normalise(candidate.vendor);
  const candidateTotal = candidate.totalAmount;
  if (candidateVendor && typeof candidateTotal === 'number') {
    const vendorTotalMatch = existingInvoices.find((entry) => normalise(entry?.data?.vendor) === candidateVendor && entry?.data?.totalAmount === candidateTotal);
    if (vendorTotalMatch) {
      return vendorTotalMatch;
    }
  }

  return null;
}
