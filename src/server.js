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
const QUICKBOOKS_METADATA_DIR = path.join(__dirname, '..', 'data', 'quickbooks');
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
  const { companyName, legalName } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (companyName !== undefined && (typeof companyName !== 'string' || !companyName.trim())) {
    return res.status(400).json({ error: 'Company name must be a non-empty string when provided.' });
  }

  if (legalName !== undefined && typeof legalName !== 'string') {
    return res.status(400).json({ error: 'Legal name must be a string.' });
  }

  if (companyName === undefined && legalName === undefined) {
    return res.status(400).json({ error: 'Provide companyName and/or legalName to update.' });
  }

  const updates = {};
  if (companyName !== undefined) {
    updates.companyName = companyName.trim();
  }

  if (legalName !== undefined) {
    updates.legalName = legalName.trim() ? legalName.trim() : null;
  }

  try {
    const updated = await updateQuickBooksCompanyFields(realmId, updates);
    res.json({ company: sanitizeQuickBooksCompany(updated) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to update QuickBooks company details', error);
    res.status(500).json({ error: 'Failed to update QuickBooks company details.' });
  }
});

app.get('/api/quickbooks/companies/:realmId/metadata', async (req, res) => {
  const realmId = req.params.realmId;

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  try {
    const metadata = await readQuickBooksCompanyMetadata(realmId);
    res.json({ metadata });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to load QuickBooks metadata', error);
    res.status(500).json({ error: 'Failed to load QuickBooks metadata.' });
  }
});

app.post('/api/quickbooks/companies/:realmId/metadata/refresh', async (req, res) => {
  const realmId = req.params.realmId;

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  try {
    const metadata = await fetchAndStoreQuickBooksMetadata(realmId, { force: true });
    res.json({ metadata });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to refresh QuickBooks metadata', error);
    res.status(500).json({ error: 'Failed to refresh QuickBooks metadata.' });
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
    let legalName = null;
    try {
      const companyInfo = await fetchQuickBooksCompanyInfo(realmId, tokenSet.accessToken);
      companyName = companyInfo?.CompanyInfo?.CompanyName || companyName;
      legalName = companyInfo?.CompanyInfo?.LegalName || companyName;
    } catch (infoError) {
      console.warn('QuickBooks company info lookup failed', infoError);
      if (infoError.isQuickBooksError && infoError.status === 403) {
        companyName = formatFallbackCompanyName(realmId);
      } else {
        throw infoError;
      }
    }

    if (!legalName) {
      legalName = companyName;
    }

    await storeQuickBooksCompany({
      realmId,
      companyName,
      legalName,
      tokens: tokenSet,
    });

    fetchAndStoreQuickBooksMetadata(realmId).catch((error) => {
      console.warn(`QuickBooks metadata prefetch failed for ${realmId}`, error.message || error);
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

warmQuickBooksMetadata().catch((error) => {
  console.error('QuickBooks metadata warmup failed', error);
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

async function storeQuickBooksCompany({ realmId, companyName, legalName, tokens }) {
  const companies = await readQuickBooksCompanies();
  const now = new Date().toISOString();
  const existingIndex = companies.findIndex((entry) => entry.realmId === realmId);
  const existing = existingIndex >= 0 ? companies[existingIndex] : {};

  const record = {
    ...existing,
    realmId,
    companyName: companyName ?? existing.companyName ?? null,
    legalName: legalName ?? existing.legalName ?? null,
    environment: QUICKBOOKS_ENVIRONMENT,
    connectedAt: existing.connectedAt || now,
    updatedAt: now,
    tokens: tokens || existing.tokens || null,
  };

  if (existingIndex >= 0) {
    companies[existingIndex] = record;
  } else {
    companies.push(record);
  }

  await persistQuickBooksCompanies(companies);
  return record;
}

async function persistQuickBooksCompanies(companies) {
  await fs.mkdir(path.dirname(QUICKBOOKS_COMPANIES_FILE), { recursive: true });
  await fs.writeFile(QUICKBOOKS_COMPANIES_FILE, JSON.stringify(companies, null, 2));
}

async function updateQuickBooksCompanyFields(realmId, updates) {
  const companies = await readQuickBooksCompanies();
  const index = companies.findIndex((entry) => entry.realmId === realmId);

  if (index < 0) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const now = new Date().toISOString();
  const next = {
    ...companies[index],
    ...updates,
    updatedAt: now,
  };

  companies[index] = next;
  await persistQuickBooksCompanies(companies);
  return next;
}

async function updateQuickBooksCompanyTokens(realmId, tokens) {
  if (!tokens) {
    return null;
  }

  const companies = await readQuickBooksCompanies();
  const index = companies.findIndex((entry) => entry.realmId === realmId);

  if (index < 0) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  companies[index] = {
    ...companies[index],
    tokens,
    updatedAt: new Date().toISOString(),
  };

  await persistQuickBooksCompanies(companies);
  return companies[index];
}

async function getQuickBooksCompanyRecord(realmId) {
  const companies = await readQuickBooksCompanies();
  return companies.find((entry) => entry.realmId === realmId) || null;
}

async function fetchAndStoreQuickBooksMetadata(realmId, { force = false } = {}) {
  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const result = await performQuickBooksFetch(realmId, async (accessToken) => {
    const [vendors, accounts, taxCodes] = await Promise.all([
      fetchQuickBooksVendors(realmId, accessToken).catch((error) =>
        handleMetadataFetchError(error, 'Vendor', realmId)
      ),
      fetchQuickBooksAccounts(realmId, accessToken).catch((error) =>
        handleMetadataFetchError(error, 'Account', realmId)
      ),
      fetchQuickBooksTaxCodes(realmId, accessToken).catch((error) =>
        handleMetadataFetchError(error, 'TaxCode', realmId)
      ),
    ]);

    const timestamp = new Date().toISOString();

    await Promise.all([
      writeQuickBooksMetadataFile(realmId, 'vendors', { updatedAt: timestamp, items: vendors }),
      writeQuickBooksMetadataFile(realmId, 'accounts', { updatedAt: timestamp, items: accounts }),
      writeQuickBooksMetadataFile(realmId, 'taxCodes', { updatedAt: timestamp, items: taxCodes }),
    ]);

    await updateQuickBooksCompanyFields(realmId, {
      vendorsUpdatedAt: vendors.length ? timestamp : null,
      vendorsCount: vendors.length,
      accountsUpdatedAt: accounts.length ? timestamp : null,
      accountsCount: accounts.length,
      taxCodesUpdatedAt: taxCodes.length ? timestamp : null,
      taxCodesCount: taxCodes.length,
    });

    return {
      vendors: { updatedAt: timestamp, items: vendors },
      accounts: { updatedAt: timestamp, items: accounts },
      taxCodes: { updatedAt: timestamp, items: taxCodes },
    };
  });

  return result;
}

async function readQuickBooksCompanyMetadata(realmId) {
  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const [vendors, accounts, taxCodes] = await Promise.all([
    readQuickBooksMetadataFile(realmId, 'vendors'),
    readQuickBooksMetadataFile(realmId, 'accounts'),
    readQuickBooksMetadataFile(realmId, 'taxCodes'),
  ]);

  return {
    vendors,
    accounts,
    taxCodes,
  };
}

function handleMetadataFetchError(error, entity, realmId) {
  if (error?.status === 403) {
    console.warn(
      `QuickBooks ${entity} metadata unavailable for ${realmId}: ${error.message || 'Forbidden.'}`
    );
    return [];
  }

  throw error;
}

async function readQuickBooksMetadataFile(realmId, type) {
  try {
    const file = await fs.readFile(getQuickBooksMetadataPath(realmId, type), 'utf-8');
    const parsed = JSON.parse(file);
    return {
      updatedAt: parsed?.updatedAt || null,
      items: Array.isArray(parsed?.items) ? parsed.items : [],
    };
  } catch (error) {
    if (error.code === 'ENOENT') {
      return { updatedAt: null, items: [] };
    }
    throw error;
  }
}

async function writeQuickBooksMetadataFile(realmId, type, payload) {
  const filePath = getQuickBooksMetadataPath(realmId, type);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2));
}

function getQuickBooksMetadataPath(realmId, type) {
  return path.join(QUICKBOOKS_METADATA_DIR, realmId, `${type}.json`);
}

async function performQuickBooksFetch(realmId, requestFn) {
  let attempt = 0;

  while (attempt < 2) {
    const company = await getQuickBooksCompanyRecord(realmId);
    if (!company?.tokens?.accessToken) {
      const error = new Error('QuickBooks access token is not available.');
      error.code = 'TOKEN_MISSING';
      throw error;
    }

    try {
      return await requestFn(company.tokens.accessToken, company);
    } catch (error) {
      if (error.status === 401 && company.tokens.refreshToken && attempt === 0) {
        const refreshed = await refreshQuickBooksAccessToken(company.tokens.refreshToken);
        await updateQuickBooksCompanyTokens(realmId, refreshed);
        attempt += 1;
        continue;
      }
      throw error;
    }
  }

  const error = new Error('Failed to call QuickBooks API after refreshing token.');
  error.code = 'TOKEN_REFRESH_FAILED';
  throw error;
}

async function refreshQuickBooksAccessToken(refreshToken) {
  if (!refreshToken) {
    const error = new Error('QuickBooks refresh token is not available.');
    error.code = 'TOKEN_MISSING';
    throw error;
  }

  if (!QUICKBOOKS_CLIENT_ID || !QUICKBOOKS_CLIENT_SECRET) {
    const error = new Error('QuickBooks client credentials are not configured.');
    error.code = 'CONFIG_MISSING';
    throw error;
  }

  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
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
      `QuickBooks token refresh failed with status ${response.status}`;
    const err = new Error(message);
    err.isQuickBooksError = true;
    err.status = response.status;
    throw err;
  }

  const data = await response.json();
  const now = Date.now();

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || refreshToken,
    scope: data.scope,
    tokenType: data.token_type,
    expiresAt: data.expires_in ? new Date(now + data.expires_in * 1000).toISOString() : null,
    refreshTokenExpiresAt: data.refresh_token_expires_in
      ? new Date(now + data.refresh_token_expires_in * 1000).toISOString()
      : null,
  };
}

async function fetchQuickBooksVendors(realmId, accessToken) {
  const vendors = await fetchQuickBooksQueryList(realmId, accessToken, 'Vendor');
  return vendors
    .filter((vendor) => vendor.Active !== false)
    .map((vendor) => ({
      id: vendor.Id,
      displayName: vendor.DisplayName || vendor.CompanyName || vendor.Title || vendor.FamilyName || vendor.GivenName || `Vendor ${vendor.Id}`,
      email: vendor.PrimaryEmailAddr?.Address || null,
      phone: vendor.PrimaryPhone?.FreeFormNumber || null,
    }));
}

async function fetchQuickBooksAccounts(realmId, accessToken) {
  const accounts = await fetchQuickBooksQueryList(realmId, accessToken, 'Account');
  return accounts
    .filter((account) => account.Active !== false)
    .map((account) => ({
      id: account.Id,
      name: account.Name || account.FullyQualifiedName || `Account ${account.Id}`,
      fullyQualifiedName: account.FullyQualifiedName || null,
      accountType: account.AccountType || null,
      accountSubType: account.AccountSubType || null,
    }));
}

async function fetchQuickBooksTaxCodes(realmId, accessToken) {
  const taxCodes = await fetchQuickBooksQueryList(realmId, accessToken, 'TaxCode');
  return taxCodes
    .filter((code) => code.Active !== false)
    .map((code) => ({
      id: code.Id,
      name: code.Name || `Tax Code ${code.Id}`,
      description: code.Description || null,
      rate: deriveTaxCodeRate(code),
      agency: code.SalesTaxRateList?.TaxAgencyRef?.name || null,
      active: code.Active !== false,
    }));
}

async function fetchQuickBooksQueryList(realmId, accessToken, entity) {
  const pageSize = 200;
  let startPosition = 1;
  const items = [];

  while (true) {
    const query = `select * from ${entity} startposition ${startPosition} maxresults ${pageSize}`;
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/query?minorversion=65&query=${encodeURIComponent(query)}`;

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
        `QuickBooks query for ${entity} failed with status ${response.status}`;
      const err = new Error(message);
      err.isQuickBooksError = true;
      err.status = response.status;
      throw err;
    }

    const body = await response.json();
    const entityKey = entity;
    const rows = body?.QueryResponse?.[entityKey];
    const list = Array.isArray(rows) ? rows : rows ? [rows] : [];
    items.push(...list);

    const maxResults = body?.QueryResponse?.maxResults || list.length;
    const totalCount = body?.QueryResponse?.totalCount;

    if (list.length < pageSize || (typeof totalCount === 'number' && items.length >= totalCount)) {
      break;
    }

    startPosition += Math.max(maxResults, list.length || pageSize);
  }

  return items;
}

function deriveTaxCodeRate(taxCode) {
  const details =
    taxCode?.SalesTaxRateList?.TaxRateDetail ||
    taxCode?.PurchaseTaxRateList?.TaxRateDetail ||
    taxCode?.TaxRateDetails ||
    [];

  const list = Array.isArray(details) ? details : details ? [details] : [];

  if (!list.length) {
    return null;
  }

  const total = list.reduce((sum, detail) => {
    const rate = Number(detail.RateValue ?? detail.rateValue);
    if (Number.isFinite(rate)) {
      return sum + rate;
    }
    return sum;
  }, 0);

  return Number.isFinite(total) ? total : null;
}

async function warmQuickBooksMetadata() {
  try {
    const companies = await readQuickBooksCompanies();
    for (const company of companies) {
      try {
        await fetchAndStoreQuickBooksMetadata(company.realmId);
      } catch (error) {
        console.warn(`Failed to warm metadata for ${company.realmId}`, error.message || error);
      }
    }
  } catch (error) {
    console.error('Unable to warm QuickBooks metadata', error);
  }
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
