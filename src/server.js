require('dotenv').config();
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const crypto = require('crypto');

const { createWorker } = require('tesseract.js');

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
const INVOICE_STORAGE_DIR = path.join(__dirname, '..', 'data', 'invoice_store');
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

const OCR_IMAGE_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/heic',
  'image/heif',
  'image/tiff',
  'image/webp',
]);

const GEMINI_TEXT_LIMIT = 60000;
const VENDOR_VAT_TREATMENTS = new Set(['inclusive', 'exclusive', 'no_vat']);

let tesseractWorkerPromise = null;
let pdfjsLibPromise = null;


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

    const sourceBuffer = Buffer.isBuffer(req.file.buffer)
      ? req.file.buffer
      : Buffer.from(req.file.buffer);
    const invoiceBuffer = Buffer.allocUnsafe(sourceBuffer.length);
    sourceBuffer.copy(invoiceBuffer);
    const storageBuffer = Buffer.allocUnsafe(sourceBuffer.length);
    sourceBuffer.copy(storageBuffer);
    const invoiceMetadata = {
      originalName: req.file.originalname,
      mimeType: req.file.mimetype,
      size: req.file.size,
      checksum: crypto.createHash('sha256').update(invoiceBuffer).digest('hex'),
    };

    const existingInvoices = await readStoredInvoices();
    const duplicateByChecksum = existingInvoices.find((entry) => entry.metadata?.checksum === invoiceMetadata.checksum);

    const preprocessing = await preprocessInvoice(invoiceBuffer, req.file.mimetype);
    const parsedInvoice = await extractWithGemini(invoiceBuffer, req.file.mimetype, {
      extractedText: preprocessing.text,
    });

    const enrichedInvoice = {
      parsedAt: new Date().toISOString(),
      metadata: {
        ...invoiceMetadata,
        extraction: {
          method: preprocessing.method,
          textLength: preprocessing.text ? preprocessing.text.length : 0,
          truncated: Boolean(preprocessing.truncated),
          totalPages: preprocessing.totalPages ?? null,
          processedPages: preprocessing.processedPages ?? null,
          error: preprocessing.error || null,
        },
      },
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
      status: 'archive',
    };

    await ensureInvoiceFileStored(invoiceMetadata, storageBuffer);
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

app.patch('/api/quickbooks/companies/:realmId/vendors/:vendorId/settings', async (req, res) => {
  const realmId = req.params.realmId;
  const vendorId = req.params.vendorId;
  const { accountId, taxCodeId, vatTreatment } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (!vendorId) {
    return res.status(400).json({ error: 'Vendor ID is required.' });
  }

  if (accountId === undefined && taxCodeId === undefined && vatTreatment === undefined) {
    return res.status(400).json({ error: 'Provide accountId, taxCodeId, or vatTreatment to update.' });
  }

  try {
    const metadata = await readQuickBooksCompanyMetadata(realmId);
    const vendorExists = metadata?.vendors?.items?.some((entry) => entry.id === vendorId);
    if (!vendorExists) {
      return res.status(404).json({ error: 'Vendor not found for this company.' });
    }

    const normalised = {};

    if (accountId !== undefined) {
      if (accountId === null || accountId === '') {
        normalised.accountId = null;
      } else if (typeof accountId === 'string') {
        const accountExists = metadata?.accounts?.items?.some((entry) => entry.id === accountId);
        if (!accountExists) {
          return res.status(400).json({ error: 'Account not found for this company.' });
        }
        normalised.accountId = accountId;
      } else {
        return res.status(400).json({ error: 'accountId must be a string or null.' });
      }
    }

    if (taxCodeId !== undefined) {
      if (taxCodeId === null || taxCodeId === '') {
        normalised.taxCodeId = null;
      } else if (typeof taxCodeId === 'string') {
        const taxCodeExists = metadata?.taxCodes?.items?.some((entry) => entry.id === taxCodeId);
        if (!taxCodeExists) {
          return res.status(400).json({ error: 'Tax code not found for this company.' });
        }
        normalised.taxCodeId = taxCodeId;
      } else {
        return res.status(400).json({ error: 'taxCodeId must be a string or null.' });
      }
    }

    if (vatTreatment !== undefined) {
      if (vatTreatment === null || vatTreatment === '') {
        normalised.vatTreatment = null;
      } else if (typeof vatTreatment === 'string' && VENDOR_VAT_TREATMENTS.has(vatTreatment)) {
        normalised.vatTreatment = vatTreatment;
      } else {
        return res.status(400).json({ error: 'vatTreatment must be inclusive, exclusive, or no_vat.' });
      }
    }

    const existingSettings = await readQuickBooksVendorSettings(realmId);
    const updatedSettings = { ...existingSettings };
    const currentEntry = existingSettings[vendorId] || {};

    const mergedEntry = {
      accountId: normalised.accountId !== undefined ? normalised.accountId : currentEntry.accountId ?? null,
      taxCodeId: normalised.taxCodeId !== undefined ? normalised.taxCodeId : currentEntry.taxCodeId ?? null,
      vatTreatment:
        normalised.vatTreatment !== undefined ? normalised.vatTreatment : currentEntry.vatTreatment ?? null,
    };

    const sanitisedEntry = sanitiseVendorSettingEntry(mergedEntry);
    const hasValues = Boolean(
      sanitisedEntry.accountId || sanitisedEntry.taxCodeId || sanitisedEntry.vatTreatment
    );

    if (hasValues) {
      updatedSettings[vendorId] = sanitisedEntry;
    } else {
      delete updatedSettings[vendorId];
    }

    const persisted = await writeQuickBooksVendorSettings(realmId, updatedSettings);
    const responseEntry = persisted[vendorId] || {
      accountId: null,
      taxCodeId: null,
      vatTreatment: null,
    };

    res.json({ vendorId, settings: responseEntry });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to update vendor settings', error);
    res.status(500).json({ error: 'Failed to update vendor settings.' });
  }
});

app.get('/api/invoices', async (req, res) => {
  try {
    const invoices = await readStoredInvoices();
    res.json({ invoices });
  } catch (error) {
    console.error('Failed to load stored invoices', error);
    res.status(500).json({ error: 'Failed to load stored invoices.' });
  }
});

app.post('/api/quickbooks/companies/:realmId/vendors/import-defaults', async (req, res) => {
  const realmId = req.params.realmId;

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  try {
    const metadata = await readQuickBooksCompanyMetadata(realmId);
    const vendors = metadata?.vendors?.items || [];
    if (!vendors.length) {
      return res.status(400).json({ error: 'No vendors available for this company.' });
    }

    const accounts = metadata?.accounts?.items || [];
    const taxCodes = metadata?.taxCodes?.items || [];
    const accountIdSet = new Set(accounts.map((account) => account.id));
    const taxCodeIdSet = new Set(taxCodes.map((code) => code.id));

    const existingSettings = await readQuickBooksVendorSettings(realmId);
    const vendorsNeedingDefaults = vendors.filter((vendor) => {
      const entry = existingSettings[vendor.id];
      return !entry?.accountId || !entry?.taxCodeId;
    });

    if (!vendorsNeedingDefaults.length) {
      return res.json({ applied: [], vendorSettings: existingSettings });
    }

    const suggestions = await performQuickBooksFetch(realmId, async (accessToken) => {
      const collected = [];
      for (const vendor of vendorsNeedingDefaults) {
        try {
          // eslint-disable-next-line no-await-in-loop
          const suggestion = await fetchLastVendorDefaults(realmId, accessToken, vendor.id);
          if (suggestion?.accountId || suggestion?.taxCodeId) {
            collected.push({ vendorId: vendor.id, ...suggestion });
          }
        } catch (error) {
          console.warn(`Unable to derive defaults for vendor ${vendor.id}`, error.message || error);
        }
      }
      return collected;
    });

    if (!suggestions?.length) {
      return res.json({ applied: [], vendorSettings: existingSettings });
    }

    const updatedSettings = { ...existingSettings };
    const applied = [];

    suggestions.forEach((suggestion) => {
      const vendorId = suggestion.vendorId;
      if (!vendorId) {
        return;
      }

      const current = updatedSettings[vendorId] || {};
      let nextAccountId = current.accountId || null;
      let nextTaxCodeId = current.taxCodeId || null;
      let changed = false;

      const suggestedAccountId = sanitiseQuickBooksId(suggestion.accountId);
      const suggestedTaxCodeId = sanitiseQuickBooksId(suggestion.taxCodeId);

      if (!nextAccountId && suggestedAccountId && accountIdSet.has(suggestedAccountId)) {
        nextAccountId = suggestedAccountId;
        changed = true;
      }

      if (!nextTaxCodeId && suggestedTaxCodeId && taxCodeIdSet.has(suggestedTaxCodeId)) {
        nextTaxCodeId = suggestedTaxCodeId;
        changed = true;
      }

      if (!changed) {
        return;
      }

      updatedSettings[vendorId] = {
        accountId: nextAccountId,
        taxCodeId: nextTaxCodeId,
        vatTreatment: current.vatTreatment || null,
      };
      applied.push({
        vendorId,
        accountId: nextAccountId,
        taxCodeId: nextTaxCodeId,
        source: suggestion.source || null,
      });
    });

    if (!applied.length) {
      return res.json({ applied: [], vendorSettings: existingSettings });
    }

    const persisted = await writeQuickBooksVendorSettings(realmId, updatedSettings);

    res.json({ applied, vendorSettings: persisted });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to import vendor defaults', error);
    res.status(500).json({ error: 'Failed to import vendor defaults.' });
  }
});

app.get('/api/invoices/:checksum/file', async (req, res) => {
  const checksum = req.params.checksum;

  if (!checksum) {
    return res.status(400).json({ error: 'Checksum is required.' });
  }

  try {
    const invoice = await findStoredInvoiceByChecksum(checksum);
    if (!invoice) {
      return res.status(404).json({ error: 'Invoice not found.' });
    }

    const filePath = getStoredInvoiceFilePath(checksum, invoice.metadata);
    try {
      await fs.access(filePath);
    } catch (error) {
      if (error.code === 'ENOENT') {
        return res.status(404).json({ error: 'Invoice file not found. Please re-upload the invoice.' });
      }
      throw error;
    }

    const mimeType = invoice.metadata?.mimeType || deriveMimeTypeFromExtension(path.extname(filePath));
    const downloadName = buildDownloadFileName(checksum, invoice.metadata, filePath);

    res.setHeader('Content-Type', mimeType);
    res.setHeader('Content-Disposition', buildInlineContentDisposition(downloadName));
    return res.sendFile(filePath, (error) => {
      if (error) {
        if (error.code === 'ENOENT') {
          if (!res.headersSent) {
            res.status(404).json({ error: 'Invoice file not found. Please re-upload the invoice.' });
          }
          return;
        }
        console.error('Failed to send invoice file', error);
        if (!res.headersSent) {
          res.status(500).json({ error: 'Failed to load invoice file.' });
        }
      }
    });
  } catch (error) {
    console.error('Failed to stream invoice file', error);
    res.status(500).json({ error: 'Failed to load invoice file.' });
  }
});

app.post('/api/invoices/:checksum/status', async (req, res) => {
  const checksum = req.params.checksum;
  const { status } = req.body || {};

  if (!checksum) {
    return res.status(400).json({ error: 'Checksum is required.' });
  }

  if (status !== 'archive' && status !== 'review') {
    return res.status(400).json({ error: 'Status must be archive or review.' });
  }

  try {
    const updated = await updateStoredInvoiceStatus(checksum, status);
    if (!updated) {
      return res.status(404).json({ error: 'Invoice not found.' });
    }
    res.json({ invoice: updated });
  } catch (error) {
    console.error('Failed to update invoice status', error);
    res.status(500).json({ error: 'Failed to update invoice status.' });
  }
});

app.delete('/api/invoices/:checksum', async (req, res) => {
  const checksum = req.params.checksum;

  if (!checksum) {
    return res.status(400).json({ error: 'Checksum is required.' });
  }

  try {
    const deleted = await deleteStoredInvoice(checksum);
    if (!deleted) {
      return res.status(404).json({ error: 'Invoice not found.' });
    }
    res.json({ success: true });
  } catch (error) {
    console.error('Failed to delete invoice', error);
    res.status(500).json({ error: 'Failed to delete invoice.' });
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

    const vendorSettings = await readQuickBooksVendorSettings(realmId);

    return {
      vendors: { updatedAt: timestamp, items: vendors },
      accounts: { updatedAt: timestamp, items: accounts },
      taxCodes: { updatedAt: timestamp, items: taxCodes },
      vendorSettings,
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

  const [vendors, accounts, taxCodes, vendorSettings] = await Promise.all([
    readQuickBooksMetadataFile(realmId, 'vendors'),
    readQuickBooksMetadataFile(realmId, 'accounts'),
    readQuickBooksMetadataFile(realmId, 'taxCodes'),
    readQuickBooksVendorSettings(realmId),
  ]);

  return {
    vendors,
    accounts,
    taxCodes,
    vendorSettings,
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

async function readQuickBooksVendorSettings(realmId) {
  try {
    const contents = await fs.readFile(getQuickBooksVendorSettingsPath(realmId), 'utf-8');
    const parsed = JSON.parse(contents);
    return sanitiseVendorSettingsMap(parsed);
  } catch (error) {
    if (error.code === 'ENOENT') {
      return {};
    }
    throw error;
  }
}

async function writeQuickBooksVendorSettings(realmId, settings) {
  const filePath = getQuickBooksVendorSettingsPath(realmId);
  const sanitised = sanitiseVendorSettingsMap(settings);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(sanitised, null, 2));
  return sanitised;
}

function getQuickBooksVendorSettingsPath(realmId) {
  return path.join(QUICKBOOKS_METADATA_DIR, realmId, 'vendor-settings.json');
}

function sanitiseVendorSettingsMap(settings) {
  if (!settings || typeof settings !== 'object' || Array.isArray(settings)) {
    return {};
  }

  const result = {};
  for (const [vendorId, entry] of Object.entries(settings)) {
    const sanitised = sanitiseVendorSettingEntry(entry);
    if (sanitised.accountId || sanitised.taxCodeId || sanitised.vatTreatment) {
      result[vendorId] = sanitised;
    }
  }

  return result;
}

function sanitiseVendorSettingEntry(entry) {
  const accountId =
    typeof entry?.accountId === 'string' && entry.accountId.trim()
      ? entry.accountId.trim()
      : null;
  const taxCodeId =
    typeof entry?.taxCodeId === 'string' && entry.taxCodeId.trim()
      ? entry.taxCodeId.trim()
      : null;
  const vatValue = typeof entry?.vatTreatment === 'string' ? entry.vatTreatment.trim() : '';

  return {
    accountId,
    taxCodeId,
    vatTreatment: VENDOR_VAT_TREATMENTS.has(vatValue) ? vatValue : null,
  };
}

function sanitiseQuickBooksId(value) {
  if (value === undefined || value === null) {
    return null;
  }
  const stringValue = value.toString().trim();
  return stringValue ? stringValue : null;
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
  const [taxCodes, taxRates] = await Promise.all([
    fetchQuickBooksQueryList(realmId, accessToken, 'TaxCode'),
    fetchQuickBooksQueryList(realmId, accessToken, 'TaxRate'),
  ]);

  const taxRateLookup = buildTaxRateLookup(taxRates || []);

  return (taxCodes || [])
    .filter((code) => code?.Active !== false)
    .map((code) => ({
      id: code.Id,
      name: code.Name || `Tax Code ${code.Id}`,
      description: code.Description || null,
      rate: deriveTaxCodeRate(code, taxRateLookup),
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

function buildTaxRateLookup(taxRates) {
  const lookup = new Map();

  for (const rate of taxRates) {
    if (rate?.Active === false) {
      continue;
    }

    const rateId = rate?.Id;
    if (!rateId) {
      continue;
    }

    const numeric = normaliseRateNumber(rate?.RateValue ?? rate?.rateValue ?? rate?.Rate ?? rate?.rate);
    if (numeric !== null) {
      lookup.set(rateId, numeric);
      continue;
    }

    const parsedFromName = extractPercentageFromLabel(rate?.Name || rate?.name);
    if (parsedFromName !== null) {
      lookup.set(rateId, parsedFromName);
    }
  }

  return lookup;
}

function deriveTaxCodeRate(taxCode, taxRateLookup) {
  const details =
    taxCode?.SalesTaxRateList?.TaxRateDetail ||
    taxCode?.PurchaseTaxRateList?.TaxRateDetail ||
    taxCode?.TaxRateDetails ||
    [];

  const list = Array.isArray(details) ? details : details ? [details] : [];

  let total = 0;
  let foundRate = false;

  for (const detail of list) {
    const directRate = normaliseRateNumber(detail?.RateValue ?? detail?.rateValue);
    if (directRate !== null) {
      total += directRate;
      foundRate = true;
      continue;
    }

    const refId = detail?.TaxRateRef?.value ?? detail?.TaxRateRef?.Value;
    if (refId && taxRateLookup?.has(refId)) {
      const lookupRate = normaliseRateNumber(taxRateLookup.get(refId));
      if (lookupRate !== null) {
        total += lookupRate;
        foundRate = true;
        continue;
      }
    }

    const refName = detail?.TaxRateRef?.name ?? detail?.TaxRateRef?.Name;
    const parsedFromName = extractPercentageFromLabel(refName);
    if (parsedFromName !== null) {
      total += parsedFromName;
      foundRate = true;
    }
  }

  if (foundRate && Number.isFinite(total)) {
    return normaliseRateNumber(total);
  }

  const fallback = extractPercentageFromLabel(taxCode?.Name || taxCode?.name || taxCode?.Description || taxCode?.description);
  return fallback;
}

function normaliseRateNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return null;
  }

  return Math.round(number * 1000000) / 1000000;
}

function extractPercentageFromLabel(label) {
  if (typeof label !== 'string') {
    return null;
  }

  const match = label.match(/(-?\d+(?:\.\d+)?)\s*%/);
  if (!match) {
    return null;
  }

  const numeric = Number(match[1]);
  return Number.isFinite(numeric) ? numeric : null;
}

async function fetchLastVendorDefaults(realmId, accessToken, vendorId) {
  const vendorKey = sanitiseQuickBooksId(vendorId);
  if (!vendorKey) {
    return null;
  }

  const escapedVendorId = vendorKey.replace(/'/g, "''");
  const queries = [
    {
      entity: 'Bill',
      statement: `SELECT * FROM Bill WHERE VendorRef = '${escapedVendorId}' ORDER BY TxnDate DESC, MetaData.LastUpdatedTime DESC STARTPOSITION 1 MAXRESULTS 1`,
    },
    {
      entity: 'Purchase',
      statement: `SELECT * FROM Purchase WHERE EntityRef = '${escapedVendorId}' AND EntityRef.Type = 'Vendor' ORDER BY TxnDate DESC, MetaData.LastUpdatedTime DESC STARTPOSITION 1 MAXRESULTS 1`,
    },
  ];

  for (const { entity, statement } of queries) {
    let response;
    try {
      // eslint-disable-next-line no-await-in-loop
      response = await runQuickBooksQuery(realmId, accessToken, statement);
    } catch (error) {
      if (error.status === 400 || error.status === 404) {
        continue;
      }
      throw error;
    }
    const rows = extractEntitiesFromQueryResponse(response, entity);
    if (!rows.length) {
      continue;
    }

    for (const row of rows) {
      const extracted = extractAccountAndTaxFromTransaction(row);
      if (extracted.accountId || extracted.taxCodeId) {
        return { ...extracted, source: entity };
      }
    }
  }

  return null;
}

async function runQuickBooksQuery(realmId, accessToken, query) {
  const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/query`;
  const params = new URLSearchParams({
    minorversion: '65',
    query,
  });

  const response = await fetch(`${url}?${params.toString()}`, {
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
      `QuickBooks query failed with status ${response.status}`;
    const err = new Error(message);
    err.isQuickBooksError = true;
    err.status = response.status;
    throw err;
  }

  return response.json();
}

function extractEntitiesFromQueryResponse(response, entity) {
  const queryResponse = response?.QueryResponse;
  if (!queryResponse) {
    return [];
  }

  const rows = queryResponse[entity];
  if (Array.isArray(rows)) {
    return rows;
  }
  if (rows) {
    return [rows];
  }
  return [];
}

function extractAccountAndTaxFromTransaction(transaction) {
  const lines = Array.isArray(transaction?.Line)
    ? transaction.Line
    : transaction?.Line
      ? [transaction.Line]
      : [];

  for (const line of lines) {
    const detail = line?.AccountBasedExpenseLineDetail || line?.ItemBasedExpenseLineDetail;
    if (!detail) {
      continue;
    }

    const accountId = sanitiseQuickBooksId(detail?.AccountRef?.value || detail?.AccountRef?.Value);
    let taxCodeId = sanitiseQuickBooksId(detail?.TaxCodeRef?.value || detail?.TaxCodeRef?.Value);

    if (!taxCodeId) {
      taxCodeId = sanitiseQuickBooksId(transaction?.TxnTaxDetail?.TxnTaxCodeRef?.value || transaction?.TxnTaxDetail?.TxnTaxCodeRef?.Value);
    }

    if (accountId || taxCodeId) {
      return {
        accountId: accountId || null,
        taxCodeId: taxCodeId || null,
      };
    }
  }

  return {
    accountId: null,
    taxCodeId: sanitiseQuickBooksId(transaction?.TxnTaxDetail?.TxnTaxCodeRef?.value || transaction?.TxnTaxDetail?.TxnTaxCodeRef?.Value) || null,
  };
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

async function preprocessInvoice(buffer, mimeType) {
  if (!buffer || !mimeType) {
    return { text: null, method: 'unknown' };
  }

  try {
    if (mimeType === 'application/pdf') {
      const pdfResult = await extractTextFromPdf(buffer);
      if (pdfResult.text) {
        return {
          text: pdfResult.text,
          method: 'pdf-text',
          truncated: pdfResult.truncated,
          totalPages: pdfResult.totalPages,
          processedPages: pdfResult.processedPages,
        };
      }
      return {
        text: null,
        method: 'pdf-text-empty',
        totalPages: pdfResult.totalPages,
        processedPages: pdfResult.processedPages,
      };
    }

    if (OCR_IMAGE_MIME_TYPES.has(mimeType)) {
      const ocrText = await extractTextWithTesseract(buffer);
      if (ocrText) {
        const { text, truncated } = truncateText(ocrText, GEMINI_TEXT_LIMIT);
        return { text, method: 'image-ocr', truncated };
      }
      return { text: null, method: 'image-ocr-empty' };
    }
  } catch (error) {
    console.warn('Preprocessing failed', error);
    return { text: null, method: 'error', error: error.message };
  }

  return { text: null, method: 'unsupported' };
}

async function extractTextFromPdf(buffer) {
  const pdfData = buffer instanceof Uint8Array
    ? new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength)
    : new Uint8Array(buffer);
  const pdfjsLib = await getPdfJs();
  const loadingTask = pdfjsLib.getDocument({
    data: pdfData,
    useSystemFonts: true,
    isEvalSupported: false,
    useWorker: false,
  });

  const pdf = await loadingTask.promise;
  const totalPages = pdf.numPages;
  const processedPages = Math.min(totalPages, 10);

  let collectedText = [];

  for (let pageNumber = 1; pageNumber <= processedPages; pageNumber += 1) {
    // eslint-disable-next-line no-await-in-loop
    const page = await pdf.getPage(pageNumber);
    // eslint-disable-next-line no-await-in-loop
    const textContent = await page.getTextContent();
    const pageText = normalizeTextFromItems(textContent.items);
    if (pageText) {
      collectedText.push(`Page ${pageNumber}: ${pageText}`);
    }
  }

  await pdf.cleanup();

  const merged = collectedText.join('\n\n').trim();
  if (!merged || merged.length < 50) {
    return { text: null, totalPages, processedPages, truncated: false };
  }

  const { text, truncated } = truncateText(merged, GEMINI_TEXT_LIMIT);
  return { text, totalPages, processedPages, truncated };
}

async function extractTextWithTesseract(buffer) {
  const worker = await getTesseractWorker();
  const result = await worker.recognize(buffer);
  const rawText = result?.data?.text || '';
  const cleaned = normalizeWhitespace(rawText);
  return cleaned || null;
}

async function getTesseractWorker() {
  if (!tesseractWorkerPromise) {
    tesseractWorkerPromise = createWorker('eng', 1, {
      logger: process.env.TESSERACT_LOG ? (message) => console.log('[tesseract]', message) : undefined,
    });
  }

  return tesseractWorkerPromise;
}

async function getPdfJs() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  }

  return pdfjsLibPromise;
}

function normalizeTextFromItems(items) {
  if (!Array.isArray(items) || !items.length) {
    return '';
  }

  const raw = items
    .map((item) => (typeof item.str === 'string' ? item.str : ''))
    .join(' ');

  return normalizeWhitespace(raw);
}

function normalizeWhitespace(value) {
  if (!value) {
    return '';
  }

  return value
    .replace(/\u0000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function truncateText(text, limit) {
  if (!text) {
    return { text: '', truncated: false };
  }

  if (text.length <= limit) {
    return { text, truncated: false };
  }

  return { text: `${text.slice(0, limit)}…`, truncated: true };
}


async function extractWithGemini(buffer, mimeType, { extractedText } = {}) {
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
- If the document text is provided below, rely exclusively on that text.
- Otherwise, analyze the attached file contents.
- Use numbers for monetary values without currency symbols.
- If a field is missing, set it to null.
- Combine multiple bill recipients into a single string if necessary.
- For products, include at least a description. Omit other product fields if not present by setting them to null.
- Do not add commentary. Return only minified JSON with the exact fields above.`;

  let contents;
  if (extractedText && extractedText.trim()) {
    const documentText = extractedText.trim();
    contents = [
      {
        role: 'user',
        parts: [
          { text: `${prompt}\n\nDocument text:\n${documentText}` },
        ],
      },
    ];
  } else {
    contents = [
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
    ];
  }

  const payload = { contents };

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
    const parsed = JSON.parse(file);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.map((entry) => normalizeStoredInvoice(entry)).filter(Boolean);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return [];
    }
    throw err;
  }
}

async function persistInvoice(existingInvoices, invoice) {
  const normalizedExisting = Array.isArray(existingInvoices)
    ? existingInvoices.map((entry) => normalizeStoredInvoice(entry)).filter(Boolean)
    : [];
  const normalizedInvoice = normalizeStoredInvoice(invoice);
  const updated = normalizedInvoice ? [...normalizedExisting, normalizedInvoice] : normalizedExisting;
  await writeStoredInvoices(updated);
}

async function writeStoredInvoices(invoices) {
  await fs.mkdir(path.dirname(DATA_FILE), { recursive: true });
  await fs.writeFile(DATA_FILE, JSON.stringify(invoices, null, 2));
}

async function updateStoredInvoiceStatus(checksum, status) {
  if (!checksum) {
    return null;
  }

  const invoices = await readStoredInvoices();
  const index = invoices.findIndex((entry) => entry?.metadata?.checksum === checksum);
  if (index === -1) {
    return null;
  }

  invoices[index] = {
    ...invoices[index],
    status,
  };

  await writeStoredInvoices(invoices);
  return invoices[index];
}

async function deleteStoredInvoice(checksum) {
  if (!checksum) {
    return false;
  }

  const invoices = await readStoredInvoices();
  const target = invoices.find((entry) => entry?.metadata?.checksum === checksum);
  if (!target) {
    return false;
  }

  const filtered = invoices.filter((entry) => entry?.metadata?.checksum !== checksum);
  await writeStoredInvoices(filtered);

  const stillReferenced = filtered.some((entry) => entry?.metadata?.checksum === checksum);
  if (!stillReferenced) {
    await deleteStoredInvoiceFile(target.metadata || { checksum });
  }

  return true;
}

function normalizeStoredInvoice(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const status = entry.status === 'review' ? 'review' : 'archive';
  return {
    ...entry,
    status,
  };
}

async function ensureInvoiceFileStored(metadata, buffer) {
  const checksum = metadata?.checksum;
  if (!checksum || !buffer) {
    return null;
  }

  const filePath = getStoredInvoiceFilePath(checksum, metadata);
  try {
    await fs.access(filePath);
    return filePath;
  } catch (error) {
    if (error.code !== 'ENOENT') {
      throw error;
    }
  }

  const payload = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, payload);
  return filePath;
}

async function deleteStoredInvoiceFile(metadata) {
  const checksum = metadata?.checksum;
  if (!checksum) {
    return;
  }

  const filePath = getStoredInvoiceFilePath(checksum, metadata);
  try {
    await fs.unlink(filePath);
  } catch (error) {
    if (error.code !== 'ENOENT') {
      throw error;
    }
  }
}

function getStoredInvoiceFilePath(checksum, metadata = {}) {
  const extension = deriveInvoiceFileExtension(metadata);
  const safeChecksum = checksum.replace(/[^a-fA-F0-9]/g, '').slice(0, 64) || checksum;
  return path.join(INVOICE_STORAGE_DIR, `${safeChecksum}${extension}`);
}

function deriveInvoiceFileExtension(metadata = {}) {
  const original = typeof metadata.originalName === 'string' ? metadata.originalName.trim() : '';
  const extFromName = original ? path.extname(original).toLowerCase() : '';
  if (extFromName) {
    return extFromName;
  }

  const mimeExt = deriveDefaultExtensionFromMimeType(metadata.mimeType);
  if (mimeExt) {
    return mimeExt;
  }

  return '.bin';
}

function deriveDefaultExtensionFromMimeType(mimeType) {
  const map = {
    'application/pdf': '.pdf',
    'image/png': '.png',
    'image/jpeg': '.jpg',
    'image/jpg': '.jpg',
    'image/heic': '.heic',
    'image/heif': '.heif',
    'image/tiff': '.tiff',
    'image/webp': '.webp',
    'text/plain': '.txt',
  };
  return map[mimeType?.toLowerCase?.()] || null;
}

function deriveMimeTypeFromExtension(extension) {
  const map = {
    '.pdf': 'application/pdf',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.heic': 'image/heic',
    '.heif': 'image/heif',
    '.tif': 'image/tiff',
    '.tiff': 'image/tiff',
    '.webp': 'image/webp',
    '.txt': 'text/plain',
    '.json': 'application/json',
    '.xml': 'application/xml',
  };
  return map[extension?.toLowerCase?.()] || 'application/octet-stream';
}

async function findStoredInvoiceByChecksum(checksum) {
  if (!checksum) {
    return null;
  }

  const invoices = await readStoredInvoices();
  return invoices.find((entry) => entry?.metadata?.checksum === checksum) || null;
}

function buildDownloadFileName(checksum, metadata, filePath) {
  const original = typeof metadata?.originalName === 'string' ? metadata.originalName.trim() : '';
  if (original) {
    return original;
  }

  const ext = path.extname(filePath) || deriveInvoiceFileExtension(metadata) || '.pdf';
  return `${checksum}${ext}`;
}

function buildInlineContentDisposition(filename) {
  const safeName = sanitiseFilename(filename || 'invoice.pdf');
  const encoded = encodeRFC5987Value(filename || safeName);
  return `inline; filename="${safeName}"; filename*=UTF-8''${encoded}`;
}

function sanitiseFilename(value) {
  return (value || 'invoice.pdf')
    .toString()
    .replace(/[\r\n"\\]/g, '_')
    .slice(0, 255);
}

function encodeRFC5987Value(value) {
  return encodeURIComponent(value)
    .replace(/['()]/g, (match) => `%${match.charCodeAt(0).toString(16).toUpperCase()}`)
    .replace(/\*/g, '%2A');
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
