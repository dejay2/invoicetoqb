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

const MS_GRAPH_CLIENT_ID = process.env.MS_GRAPH_CLIENT_ID || '';
const MS_GRAPH_CLIENT_SECRET = process.env.MS_GRAPH_CLIENT_SECRET || '';
const MS_GRAPH_TENANT_ID = process.env.MS_GRAPH_TENANT_ID || '';
const MS_GRAPH_SCOPE = process.env.MS_GRAPH_SCOPE || 'https://graph.microsoft.com/.default';
const GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0';
const GRAPH_AUTH_URL = MS_GRAPH_TENANT_ID
  ? `https://login.microsoftonline.com/${MS_GRAPH_TENANT_ID}/oauth2/v2.0/token`
  : null;
const ONEDRIVE_POLL_INTERVAL_MS = Math.max(parseInt(process.env.ONEDRIVE_POLL_INTERVAL_MS || '90000', 10), 15000);
const ONEDRIVE_MAX_FILE_SIZE_BYTES = Math.max(
  parseInt(process.env.ONEDRIVE_MAX_FILE_SIZE_BYTES || `${10 * 1024 * 1024}`, 10),
  1024
);
const ONEDRIVE_MAX_DELTA_PAGES = Math.max(parseInt(process.env.ONEDRIVE_MAX_DELTA_PAGES || '5', 10), 1);

const ONEDRIVE_SYNC_ENABLED = Boolean(MS_GRAPH_CLIENT_ID && MS_GRAPH_CLIENT_SECRET && MS_GRAPH_TENANT_ID);

let graphTokenCache = { token: null, expiresAt: 0 };
const activeOneDrivePolls = new Map();
let oneDrivePollingTimer = null;

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
const AI_ACCOUNT_CONFIDENCE_VALUES = new Set(['high', 'medium', 'low', 'unknown']);
const KEYWORD_STOPWORDS = new Set([
  'the',
  'and',
  'for',
  'with',
  'from',
  'into',
  'llc',
  'ltd',
  'limited',
  'inc',
  'co',
  'company',
  'corp',
  'corporation',
  'group',
  'services',
  'service',
  'solutions',
  'holdings',
  'global',
  'international',
  'invoice',
  'number',
  'account',
  'price',
  'prices',
]);

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

async function ingestInvoiceFromSource({
  buffer,
  mimeType,
  originalName,
  fileSize,
  realmId,
  businessType,
  remoteSource,
} = {}) {
  if (!buffer) {
    throw new Error('Invoice buffer is required.');
  }

  if (!mimeType) {
    throw new Error('Invoice mimeType is required.');
  }

  if (!GEMINI_API_KEY) {
    const error = new Error('Gemini API key is not configured. Set GEMINI_API_KEY in the environment.');
    error.code = 'GEMINI_UNCONFIGURED';
    throw error;
  }

  const sourceBuffer = Buffer.isBuffer(buffer)
    ? buffer
    : Buffer.from(buffer);

  const invoiceBuffer = Buffer.allocUnsafe(sourceBuffer.length);
  sourceBuffer.copy(invoiceBuffer);
  const storageBuffer = Buffer.allocUnsafe(sourceBuffer.length);
  sourceBuffer.copy(storageBuffer);

  const checksum = crypto.createHash('sha256').update(invoiceBuffer).digest('hex');

  const metadata = {
    originalName: typeof originalName === 'string' && originalName.trim() ? originalName.trim() : 'invoice',
    mimeType,
    size: typeof fileSize === 'number' && Number.isFinite(fileSize) ? fileSize : sourceBuffer.length,
    checksum,
  };

  let resolvedRealmId = typeof realmId === 'string' ? realmId.trim() : '';
  let resolvedBusinessType = typeof businessType === 'string' ? businessType.trim().slice(0, 120) : '';

  if (resolvedRealmId) {
    try {
      const companyRecord = await getQuickBooksCompanyRecord(resolvedRealmId);
      if (companyRecord?.businessType) {
        resolvedBusinessType = companyRecord.businessType;
      }
    } catch (lookupError) {
      if (lookupError?.code !== 'NOT_FOUND') {
        console.warn(
          `Unable to load QuickBooks company profile for ${resolvedRealmId}:`,
          lookupError.message || lookupError
        );
      }
    }
  }

  const companyProfile = resolvedRealmId || resolvedBusinessType
    ? {
        realmId: resolvedRealmId || null,
        businessType: resolvedBusinessType || null,
      }
    : null;

  if (remoteSource) {
    const normalisedRemote = normalizeRemoteSourceMetadata(remoteSource);
    if (normalisedRemote) {
      metadata.remoteSource = normalisedRemote;
    }
  }

  const existingInvoices = await readStoredInvoices();

  if (metadata.remoteSource?.provider) {
    const duplicateByRemote = existingInvoices.find((entry) =>
      isSameRemoteSource(entry?.metadata?.remoteSource, metadata.remoteSource)
    );
    if (duplicateByRemote) {
      return {
        invoice: duplicateByRemote,
        duplicate: {
          reason: 'Remote file already processed.',
          match: duplicateByRemote,
        },
        skipped: true,
      };
    }
  }

  const duplicateByChecksum = existingInvoices.find((entry) => entry.metadata?.checksum === metadata.checksum) || null;

  const preprocessing = await preprocessInvoice(invoiceBuffer, mimeType);
  logPreprocessingResult(metadata.originalName, preprocessing);

  const parsedInvoice = await extractWithGemini(invoiceBuffer, mimeType, {
    extractedText: preprocessing.text,
    originalName: metadata.originalName,
    businessType: resolvedBusinessType,
  });

  const enrichedInvoice = {
    parsedAt: new Date().toISOString(),
    metadata: {
      ...metadata,
      companyProfile,
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

  await ensureInvoiceFileStored(metadata, storageBuffer);
  await persistInvoice(existingInvoices, storagePayload);

  return {
    invoice: enrichedInvoice,
    duplicate: duplicateMatch
      ? {
          reason: duplicateReason,
          match: duplicateMatch,
        }
      : null,
    stored: storagePayload,
  };
}

function isOneDriveSyncConfigured() {
  return ONEDRIVE_SYNC_ENABLED && Boolean(GRAPH_AUTH_URL);
}

async function acquireMicrosoftGraphToken() {
  if (!isOneDriveSyncConfigured()) {
    throw new Error('Microsoft Graph credentials are not configured.');
  }

  const now = Date.now();
  if (graphTokenCache.token && graphTokenCache.expiresAt - 60000 > now) {
    return graphTokenCache.token;
  }

  const params = new URLSearchParams();
  params.set('client_id', MS_GRAPH_CLIENT_ID);
  params.set('client_secret', MS_GRAPH_CLIENT_SECRET);
  params.set('grant_type', 'client_credentials');
  params.set('scope', MS_GRAPH_SCOPE);

  const response = await fetch(GRAPH_AUTH_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message =
      errorBody?.error_description ||
      errorBody?.error ||
      `Microsoft Graph token request failed with status ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    error.body = errorBody;
    throw error;
  }

  const data = await response.json();
  const expiresIn = typeof data.expires_in === 'number' ? data.expires_in : Number.parseInt(data.expires_in, 10) || 3600;
  const expiresAt = now + Math.max(expiresIn - 60, 60) * 1000;
  graphTokenCache = {
    token: data.access_token,
    expiresAt,
  };

  return graphTokenCache.token;
}

function buildGraphUrl(resource, query) {
  if (!resource) {
    throw new Error('Graph resource path is required.');
  }

  const url = /^https?:/i.test(resource)
    ? new URL(resource)
    : new URL(resource.replace(/^\/+/, ''), `${GRAPH_API_BASE_URL.replace(/\/+$/, '')}/`);

  if (query && typeof query === 'object') {
    Object.entries(query).forEach(([key, value]) => {
      if (value === null || value === undefined) {
        return;
      }
      if (Array.isArray(value)) {
        value.forEach((entry) => {
          if (entry !== null && entry !== undefined) {
            url.searchParams.append(key, entry);
          }
        });
        return;
      }
      url.searchParams.set(key, value);
    });
  }

  return url;
}

async function graphFetch(resource, { method = 'GET', headers = {}, body, query, responseType = 'json' } = {}) {
  const token = await acquireMicrosoftGraphToken();
  const url = buildGraphUrl(resource, query);

  const requestHeaders = {
    Authorization: `Bearer ${token}`,
    ...headers,
  };

  const init = {
    method,
    headers: requestHeaders,
  };

  if (body !== undefined && body !== null) {
    if (Buffer.isBuffer(body) || typeof body === 'string') {
      init.body = body;
    } else {
      init.body = JSON.stringify(body);
      init.headers['Content-Type'] = init.headers['Content-Type'] || 'application/json';
    }
  }

  const response = await fetch(url.toString(), init);

  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message =
      errorBody?.error?.message ||
      errorBody?.error_description ||
      `Microsoft Graph request failed with status ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    error.body = errorBody;
    throw error;
  }

  if (responseType === 'json') {
    return response.json();
  }
  if (responseType === 'text') {
    return response.text();
  }
  if (responseType === 'buffer') {
    const arrayBuffer = await response.arrayBuffer();
    return Buffer.from(arrayBuffer);
  }
  if (responseType === 'response') {
    return response;
  }

  return response;
}

function encodeSharingUrl(shareUrl) {
  const trimmed = sanitiseOptionalString(shareUrl);
  if (!trimmed) {
    return null;
  }
  const base64 = Buffer.from(trimmed, 'utf-8').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  return `u!${base64}`;
}

function normaliseGraphFolderPath(folderPath) {
  const trimmed = sanitiseOptionalString(folderPath) || '';
  const withoutRoot = trimmed.replace(/^drive\/?root:?/i, '').replace(/^:+/, '');
  const normalised = withoutRoot.startsWith('/') ? withoutRoot : `/${withoutRoot}`;
  return normalised.replace(/\/+$/, '');
}

async function resolveOneDriveFolderReference({ shareUrl, driveId, folderId, folderPath }) {
  if (!isOneDriveSyncConfigured()) {
    throw new Error('OneDrive integration is not configured.');
  }

  let item;
  let resolvedDriveId = sanitiseOptionalString(driveId);
  const trimmedShareUrl = sanitiseOptionalString(shareUrl);

  if (trimmedShareUrl) {
    const encoded = encodeSharingUrl(trimmedShareUrl);
    item = await graphFetch(`/shares/${encoded}/driveItem`, {
      query: {
        $select: 'id,name,parentReference,webUrl,folder',
      },
    });
    resolvedDriveId = item?.parentReference?.driveId || resolvedDriveId;
  } else if (resolvedDriveId && sanitiseOptionalString(folderId)) {
    item = await graphFetch(`/drives/${encodeURIComponent(resolvedDriveId)}/items/${encodeURIComponent(folderId)}`, {
      query: {
        $select: 'id,name,parentReference,webUrl,folder',
      },
    });
  } else if (resolvedDriveId && sanitiseOptionalString(folderPath)) {
    const normalisedPath = normaliseGraphFolderPath(folderPath);
    item = await graphFetch(`/drives/${encodeURIComponent(resolvedDriveId)}/root:${normalisedPath}:`, {
      query: {
        $select: 'id,name,parentReference,webUrl,folder',
      },
    });
  } else {
    throw new Error('Provide shareUrl or driveId with folderId/folderPath to configure OneDrive sync.');
  }

  if (!item?.folder) {
    throw new Error('The selected OneDrive resource is not a folder.');
  }

  const parentPath = item.parentReference?.path || null;
  const displayPath = parentPath ? `${parentPath}/${item.name}` : item.name || null;

  return {
    id: item.id,
    driveId: item.parentReference?.driveId || resolvedDriveId || null,
    parentId: item.parentReference?.id || null,
    name: item.name || 'Folder',
    path: displayPath,
    webUrl: item.webUrl || null,
    shareUrl: trimmedShareUrl || null,
  };
}

async function pollAllOneDriveCompanies() {
  if (!isOneDriveSyncConfigured()) {
    return;
  }

  try {
    const companies = await readQuickBooksCompanies();
    const tasks = companies
      .map((company) => {
        if (!company?.realmId) {
          return null;
        }
        const config = ensureOneDriveStateDefaults(company.oneDrive);
        if (!config || config.enabled === false) {
          return null;
        }
        return queueOneDrivePoll(company.realmId, { reason: 'interval' });
      })
      .filter(Boolean);

    await Promise.allSettled(tasks);
  } catch (error) {
    console.error('OneDrive poll enumeration failed', error);
  }
}

async function queueOneDrivePoll(realmId, { reason = 'manual' } = {}) {
  if (!isOneDriveSyncConfigured() || !realmId) {
    return null;
  }

  if (activeOneDrivePolls.has(realmId)) {
    return activeOneDrivePolls.get(realmId);
  }

  const task = (async () => {
    try {
      const company = await getQuickBooksCompanyRecord(realmId);
      if (!company) {
        return;
      }
      await pollOneDriveForCompany(company, { reason });
    } catch (error) {
      console.error(`OneDrive sync failed for realm ${realmId}`, error);
      const timestamp = new Date().toISOString();
      await updateQuickBooksCompanyOneDrive(realmId, {
        status: 'error',
        lastSyncStatus: 'error',
        lastSyncReason: reason,
        lastSyncError: {
          message: error.message || 'OneDrive sync failed.',
          at: timestamp,
        },
        updatedAt: timestamp,
      }).catch((updateError) => {
        console.warn(`Unable to persist OneDrive error state for ${realmId}`, updateError.message || updateError);
      });
    } finally {
      activeOneDrivePolls.delete(realmId);
    }
  })();

  activeOneDrivePolls.set(realmId, task);
  return task;
}

async function pollOneDriveForCompany(company, { reason = 'manual' } = {}) {
  const config = ensureOneDriveStateDefaults(company.oneDrive);
  if (!config || config.enabled === false) {
    return;
  }

  if (!config.driveId || !config.folderId) {
    await updateQuickBooksCompanyOneDrive(company.realmId, {
      status: 'error',
      lastSyncStatus: 'error',
      lastSyncReason: reason,
      lastSyncError: {
        message: 'OneDrive folder is not fully configured. Provide driveId and folderId.',
        at: new Date().toISOString(),
      },
    });
    return;
  }

  const startedAt = Date.now();
  let nextLink = config.deltaLink
    ? config.deltaLink
    : `${GRAPH_API_BASE_URL.replace(/\/+$/, '')}/drives/${encodeURIComponent(config.driveId)}/items/${encodeURIComponent(config.folderId)}/delta`;
  let latestDelta = config.deltaLink || null;
  let processedItems = 0;
  let createdCount = 0;
  let skippedCount = 0;
  let pages = 0;
  const errors = [];

  while (nextLink && pages < ONEDRIVE_MAX_DELTA_PAGES) {
    let response;
    try {
      response = await graphFetch(nextLink, { responseType: 'json' });
    } catch (error) {
      errors.push(error);
      break;
    }

    const entries = Array.isArray(response?.value) ? response.value : [];
    for (const item of entries) {
      if (!item || item.deleted || !item.file) {
        continue;
      }
      processedItems += 1;
      try {
        const result = await processOneDriveItem(company, config, item, { reason });
        if (result?.skipped) {
          skippedCount += 1;
        } else if (result?.stored) {
          createdCount += 1;
        }
      } catch (error) {
        errors.push(error);
      }
    }

    pages += 1;

    if (response['@odata.nextLink']) {
      nextLink = response['@odata.nextLink'];
    } else {
      nextLink = null;
      if (response['@odata.deltaLink']) {
        latestDelta = response['@odata.deltaLink'];
      }
    }
  }

  const completedAt = new Date().toISOString();
  const status = errors.length ? 'warning' : 'connected';
  const lastSyncStatus = errors.length ? 'partial' : 'success';
  const lastSyncError = errors.length
    ? {
        message: errors[0]?.message || 'OneDrive sync completed with errors.',
        at: completedAt,
      }
    : null;

  await updateQuickBooksCompanyOneDrive(company.realmId, {
    status,
    deltaLink: latestDelta,
    lastSyncAt: completedAt,
    lastSyncStatus,
    lastSyncReason: reason,
    lastSyncDurationMs: Date.now() - startedAt,
    lastSyncError,
    lastSyncMetrics: {
      processedItems,
      createdCount,
      skippedCount,
      pages,
      errorCount: errors.length,
    },
  });
}

async function processOneDriveItem(company, config, item, { reason = 'manual' } = {}) {
  const remoteSource = {
    provider: 'onedrive',
    driveId: config.driveId,
    itemId: item.id,
    parentId: item.parentReference?.id || config.folderId || null,
    path: buildDriveItemPath(item),
    webUrl: item.webUrl || null,
    eTag: item.eTag || null,
    lastModifiedDateTime: item.lastModifiedDateTime || null,
    syncedAt: new Date().toISOString(),
    reason,
  };

  const download = await downloadOneDriveItem(config.driveId, item.id);

  if (download.size > ONEDRIVE_MAX_FILE_SIZE_BYTES) {
    console.warn(
      `Skipping OneDrive item ${item.id} for realm ${company.realmId} because it exceeds the configured size limit (${download.size} bytes).`
    );
    return { skipped: true, reason: 'file-too-large' };
  }

  return ingestInvoiceFromSource({
    buffer: download.buffer,
    mimeType: download.mimeType,
    originalName: download.originalName,
    fileSize: download.size,
    realmId: company.realmId,
    businessType: company.businessType || null,
    remoteSource,
  });
}

async function downloadOneDriveItem(driveId, itemId) {
  const metadata = await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`, {
    query: {
      $select: 'id,name,size,file,@microsoft.graph.downloadUrl',
    },
  });

  const downloadUrl = metadata?.['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error('Download URL not available for OneDrive item.');
  }

  const response = await fetch(downloadUrl);
  if (!response.ok) {
    const message = `OneDrive file download failed with status ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }

  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  const mimeType = response.headers.get('content-type') || metadata?.file?.mimeType || 'application/octet-stream';
  const size = typeof metadata?.size === 'number' ? metadata.size : buffer.length;

  return {
    buffer,
    mimeType,
    size,
    originalName: metadata?.name || `${itemId}.bin`,
  };
}

function buildDriveItemPath(item) {
  if (!item) {
    return null;
  }
  const parentPath = item.parentReference?.path || null;
  if (parentPath) {
    return `${parentPath}/${item.name}`;
  }
  return item.name || null;
}

function startOneDriveSyncScheduler() {
  if (!isOneDriveSyncConfigured()) {
    return;
  }

  if (oneDrivePollingTimer) {
    clearInterval(oneDrivePollingTimer);
  }

  const tick = () => {
    pollAllOneDriveCompanies().catch((error) => {
      console.error('Scheduled OneDrive polling failed', error);
    });
  };

  tick();
  oneDrivePollingTimer = setInterval(tick, ONEDRIVE_POLL_INTERVAL_MS);
  console.log(
    `OneDrive folder polling enabled (interval ${Math.max(Math.round(ONEDRIVE_POLL_INTERVAL_MS / 1000), 1)}s)`
  );
}

app.post('/api/parse-invoice', upload.single('invoice'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file received. Please attach an invoice.' });
    }

    if (!GEMINI_API_KEY) {
      return res.status(500).json({ error: 'Gemini API key is not configured. Set GEMINI_API_KEY in the environment.' });
    }
    const realmId = typeof req.body?.realmId === 'string' ? req.body.realmId.trim() : '';
    const bodyBusinessType = typeof req.body?.businessType === 'string' ? req.body.businessType.trim().slice(0, 120) : '';

    const result = await ingestInvoiceFromSource({
      buffer: req.file.buffer,
      mimeType: req.file.mimetype,
      originalName: req.file.originalname,
      fileSize: req.file.size,
      realmId,
      businessType: bodyBusinessType,
    });

    return res.json({
      invoice: result.invoice,
      duplicate: result.duplicate,
    });
  } catch (error) {
    console.error('Failed to process invoice', error);
    if (error.isGeminiError) {
      return res.status(502).json({ error: error.message });
    }
    if (error.code === 'GEMINI_UNCONFIGURED') {
      return res.status(500).json({ error: error.message });
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
  const { companyName, legalName, businessType } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (companyName !== undefined && (typeof companyName !== 'string' || !companyName.trim())) {
    return res.status(400).json({ error: 'Company name must be a non-empty string when provided.' });
  }

  if (legalName !== undefined && typeof legalName !== 'string') {
    return res.status(400).json({ error: 'Legal name must be a string.' });
  }

  if (companyName === undefined && legalName === undefined && businessType === undefined) {
    return res.status(400).json({ error: 'Provide companyName, legalName, or businessType to update.' });
  }

  const updates = {};
  if (companyName !== undefined) {
    updates.companyName = companyName.trim();
  }

  if (legalName !== undefined) {
    updates.legalName = legalName.trim() ? legalName.trim() : null;
  }

  if (businessType !== undefined) {
    if (businessType === null) {
      updates.businessType = null;
    } else if (typeof businessType === 'string') {
      const trimmedType = businessType.trim();
      if (!trimmedType) {
        updates.businessType = null;
      } else {
        updates.businessType = trimmedType.slice(0, 120);
      }
    } else {
      return res.status(400).json({ error: 'Business type must be a string when provided.' });
    }
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

app.patch('/api/quickbooks/companies/:realmId/onedrive', async (req, res) => {
  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const { shareUrl, driveId, folderId, folderPath, enabled } = req.body || {};
  const enableSync = enabled !== false;

  try {
    const company = await getQuickBooksCompanyRecord(realmId);
    if (!company) {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }

    let nextState;
    if (enableSync) {
      const folder = await resolveOneDriveFolderReference({ shareUrl, driveId, folderId, folderPath });
      nextState = {
        enabled: true,
        status: 'connected',
        driveId: folder.driveId,
        folderId: folder.id,
        folderPath: folder.path,
        folderName: folder.name,
        webUrl: folder.webUrl,
        shareUrl: folder.shareUrl,
        parentId: folder.parentId,
        deltaLink: null,
        lastSyncStatus: null,
        lastSyncError: null,
        lastSyncReason: 'configuration',
        lastSyncMetrics: null,
      };
    } else {
      nextState = {
        enabled: false,
        status: 'disabled',
        deltaLink: null,
        lastSyncStatus: null,
        lastSyncError: null,
        lastSyncMetrics: null,
        lastSyncReason: 'disabled',
      };
    }

    await updateQuickBooksCompanyOneDrive(realmId, nextState, { replace: true });

    if (enableSync) {
      queueOneDrivePoll(realmId, { reason: 'configuration' }).catch((error) => {
        console.warn(`Unable to trigger OneDrive sync for ${realmId}`, error.message || error);
      });
    }

    const updated = await getQuickBooksCompanyRecord(realmId);
    return res.json({ oneDrive: sanitizeOneDriveSettings(updated?.oneDrive) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    if (error.status === 404) {
      return res.status(404).json({ error: 'Unable to locate the specified OneDrive folder.' });
    }
    if (error.status === 403) {
      return res.status(403).json({ error: 'Access to the specified OneDrive folder was denied.' });
    }
    console.error('Failed to update OneDrive settings', error);
    return res.status(500).json({ error: 'Failed to update OneDrive settings.' });
  }
});

app.delete('/api/quickbooks/companies/:realmId/onedrive', async (req, res) => {
  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  try {
    await updateQuickBooksCompanyOneDrive(realmId, null, { replace: true });
    res.json({ oneDrive: null });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to remove OneDrive settings', error);
    res.status(500).json({ error: 'Failed to remove OneDrive settings.' });
  }
});

app.post('/api/quickbooks/companies/:realmId/onedrive/sync', async (req, res) => {
  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  try {
    const company = await getQuickBooksCompanyRecord(realmId);
    if (!company) {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }

    queueOneDrivePoll(realmId, { reason: 'manual' }).catch((error) => {
      console.warn(`Unable to trigger manual OneDrive sync for ${realmId}`, error.message || error);
    });

    return res.status(202).json({ accepted: true, oneDrive: sanitizeOneDriveSettings(company.oneDrive) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to schedule OneDrive sync', error);
    return res.status(500).json({ error: 'Failed to schedule OneDrive sync.' });
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

app.post('/api/quickbooks/companies/:realmId/vendors', async (req, res) => {
  const realmId = req.params.realmId;
  const { displayName, email, phone } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    return res.status(404).json({ error: 'QuickBooks company not found.' });
  }

  const name = typeof displayName === 'string' ? displayName.trim() : '';
  if (!name) {
    return res.status(400).json({ error: 'Vendor displayName is required.' });
  }

  const normalisedEmail = typeof email === 'string' && email.trim() ? email.trim() : null;
  const normalisedPhone = typeof phone === 'string' && phone.trim() ? phone.trim() : null;

  try {
    const vendor = await createQuickBooksVendor(realmId, {
      displayName: name,
      email: normalisedEmail,
      phone: normalisedPhone,
    });

    res.status(201).json({ vendor });
  } catch (error) {
    if (error.code === 'TOKEN_MISSING') {
      return res.status(400).json({ error: 'QuickBooks access token is not available for this company.' });
    }

    if (error.isQuickBooksError) {
      const status = typeof error.status === 'number' && error.status >= 400 ? error.status : 502;
      return res.status(status).json({ error: error.message || 'QuickBooks vendor creation failed.' });
    }

    console.error('Failed to create QuickBooks vendor', error);
    res.status(500).json({ error: 'Failed to create QuickBooks vendor.' });
  }
});

app.post('/api/quickbooks/companies/:realmId/accounts', async (req, res) => {
  const realmId = req.params.realmId;
  const { name, accountType, accountSubType } = req.body || {};

  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    return res.status(404).json({ error: 'QuickBooks company not found.' });
  }

  const trimmedName = typeof name === 'string' ? name.trim() : '';
  if (!trimmedName) {
    return res.status(400).json({ error: 'Account name is required.' });
  }

  const normalisedAccountType = typeof accountType === 'string' && accountType.trim()
    ? accountType.trim()
    : null;
  const normalisedAccountSubType = typeof accountSubType === 'string' && accountSubType.trim()
    ? accountSubType.trim()
    : null;

  try {
    const account = await createQuickBooksAccount(realmId, {
      name: trimmedName,
      accountType: normalisedAccountType,
      accountSubType: normalisedAccountSubType,
    });

    res.status(201).json({ account });
  } catch (error) {
    if (error.code === 'TOKEN_MISSING') {
      return res.status(400).json({ error: 'QuickBooks access token is not available for this company.' });
    }

    if (error.isQuickBooksError) {
      const status = typeof error.status === 'number' && error.status >= 400 ? error.status : 502;
      return res.status(status).json({ error: error.message || 'QuickBooks account creation failed.' });
    }

    console.error('Failed to create QuickBooks account', error);
    res.status(500).json({ error: 'Failed to create QuickBooks account.' });
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

app.post('/api/invoices/:checksum/match', async (req, res) => {
  try {
    if (!GEMINI_API_KEY) {
      return res.status(503).json({ error: 'Gemini matching is not configured.' });
    }

    const { checksum } = req.params;
    const { realmId } = req.body || {};

    if (!checksum) {
      return res.status(400).json({ error: 'Invoice checksum is required.' });
    }

    if (!realmId) {
      return res.status(400).json({ error: 'realmId is required to match against QuickBooks data.' });
    }

    const [invoice, metadata] = await Promise.all([
      findStoredInvoiceByChecksum(checksum),
      readQuickBooksCompanyMetadata(realmId).catch((error) => {
        if (error?.code === 'NOT_FOUND') {
          return null;
        }
        throw error;
      }),
    ]);

    if (!invoice) {
      return res.status(404).json({ error: 'Invoice not found.' });
    }

    if (!metadata) {
      return res.status(404).json({ error: 'QuickBooks metadata not available for this company.' });
    }

    const vendorOptions = rankVendorCandidates(invoice, metadata?.vendors?.items || []).slice(0, 12);
    const accountOptions = rankAccountCandidates(invoice, metadata?.accounts?.items || []).slice(0, 15);

    if (!vendorOptions.length && !accountOptions.length) {
      return res.status(400).json({ error: 'No QuickBooks vendors or accounts available to match against.' });
    }

    const invoiceBusinessType = invoice?.metadata?.companyProfile?.businessType || null;

    const suggestion = await matchInvoiceWithGemini({
      invoice,
      vendorOptions,
      accountOptions,
      sourceName: invoice?.metadata?.originalName || null,
      businessType: invoiceBusinessType,
    });

    const vendorLookup = new Map(vendorOptions.map((entry) => [entry.vendor.id, entry.vendor]));
    const accountLookup = new Map(accountOptions.map((entry) => [entry.account.id, entry.account]));
    const allAccountLookup = new Map((metadata?.accounts?.items || []).map((entry) => [entry.id, entry]));

    const mappedVendor = mapAiVendorSuggestion(suggestion.vendor, vendorLookup);
    const mappedAccount = mapAiAccountSuggestion(suggestion.account, accountLookup);

    const vendorSettingsMap = metadata?.vendorSettings || {};
    const finalSuggestion = applyVendorDefaultAccount(
      { vendor: mappedVendor, account: mappedAccount },
      { vendorSettings: vendorSettingsMap, accountLookup, allAccountLookup }
    );

    return res.json(finalSuggestion);
  } catch (error) {
    console.error('AI matching failed', error);
    if (error?.isGeminiError) {
      return res.status(502).json({ error: error.message });
    }
    return res.status(500).json({ error: 'Failed to generate AI match suggestions.' });
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

if (isOneDriveSyncConfigured()) {
  startOneDriveSyncScheduler();
} else if (process.env.MS_GRAPH_CLIENT_ID || process.env.MS_GRAPH_CLIENT_SECRET || process.env.MS_GRAPH_TENANT_ID) {
  console.warn('OneDrive sync is not enabled. Provide MS_GRAPH_CLIENT_ID, MS_GRAPH_CLIENT_SECRET, and MS_GRAPH_TENANT_ID to activate folder monitoring.');
}

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
    businessType: existing.businessType ?? null,
    environment: QUICKBOOKS_ENVIRONMENT,
    connectedAt: existing.connectedAt || now,
    updatedAt: now,
    tokens: tokens || existing.tokens || null,
  };

  record.oneDrive = ensureOneDriveStateDefaults(record.oneDrive || existing.oneDrive || null);

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

  next.oneDrive = ensureOneDriveStateDefaults(next.oneDrive || companies[index].oneDrive || null);

  companies[index] = next;
  await persistQuickBooksCompanies(companies);
  return next;
}

async function updateQuickBooksCompanyOneDrive(realmId, updates, { replace = false } = {}) {
  const companies = await readQuickBooksCompanies();
  const index = companies.findIndex((entry) => entry.realmId === realmId);

  if (index < 0) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const now = new Date().toISOString();
  const currentNormalised = normalizeOneDriveState(companies[index].oneDrive || null);
  const current = ensureOneDriveStateDefaults(currentNormalised);
  let nextState = null;

  if (replace) {
    const replacement = normalizeOneDriveState(updates);
    nextState = replacement ? ensureOneDriveStateDefaults(replacement) : null;
    if (nextState) {
      nextState.updatedAt = now;
      if (!nextState.createdAt) {
        nextState.createdAt = now;
      }
    }
  } else {
    const updateFragment = normalizeOneDriveState(updates || {});
    if (current || (updateFragment && Object.keys(updateFragment).length)) {
      const merged = {
        ...(current || {}),
        ...(updateFragment || {}),
        updatedAt: now,
      };
      nextState = ensureOneDriveStateDefaults(merged);
      if (!nextState.createdAt && current?.createdAt) {
        nextState.createdAt = current.createdAt;
      }
      if (!nextState.createdAt) {
        nextState.createdAt = now;
      }
    }
  }

  companies[index] = {
    ...companies[index],
    oneDrive: nextState,
    updatedAt: now,
  };

  await persistQuickBooksCompanies(companies);
  return companies[index];
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

  companies[index].oneDrive = ensureOneDriveStateDefaults(companies[index].oneDrive || null);

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

async function upsertQuickBooksVendorMetadata(realmId, vendor) {
  if (!realmId || !vendor?.id) {
    return null;
  }

  const existing = await readQuickBooksMetadataFile(realmId, 'vendors');
  const items = Array.isArray(existing?.items) ? existing.items.slice() : [];
  const filtered = items.filter((entry) => entry.id !== vendor.id);
  filtered.push(vendor);
  filtered.sort((a, b) => {
    const aName = (a?.displayName || '').toString();
    const bName = (b?.displayName || '').toString();
    return aName.localeCompare(bName, undefined, { sensitivity: 'base' });
  });

  const timestamp = new Date().toISOString();
  await writeQuickBooksMetadataFile(realmId, 'vendors', {
    updatedAt: timestamp,
    items: filtered,
  });

  await updateQuickBooksCompanyFields(realmId, {
    vendorsUpdatedAt: timestamp,
    vendorsCount: filtered.length,
  });

  return {
    updatedAt: timestamp,
    items: filtered,
  };
}

async function upsertQuickBooksAccountMetadata(realmId, account) {
  if (!realmId || !account?.id) {
    return null;
  }

  const existing = await readQuickBooksMetadataFile(realmId, 'accounts');
  const items = Array.isArray(existing?.items) ? existing.items.slice() : [];
  const filtered = items.filter((entry) => entry.id !== account.id);
  filtered.push(account);
  filtered.sort((a, b) => {
    const aName = (a?.name || '').toString();
    const bName = (b?.name || '').toString();
    return aName.localeCompare(bName, undefined, { sensitivity: 'base' });
  });

  const timestamp = new Date().toISOString();
  await writeQuickBooksMetadataFile(realmId, 'accounts', {
    updatedAt: timestamp,
    items: filtered,
  });

  await updateQuickBooksCompanyFields(realmId, {
    accountsUpdatedAt: timestamp,
    accountsCount: filtered.length,
  });

  return {
    updatedAt: timestamp,
    items: filtered,
  };
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
    .map(transformQuickBooksVendor);
}

function transformQuickBooksVendor(vendor) {
  if (!vendor) {
    return null;
  }

  return {
    id: vendor.Id,
    displayName:
      vendor.DisplayName ||
      vendor.CompanyName ||
      vendor.Title ||
      vendor.FamilyName ||
      vendor.GivenName ||
      (vendor.Id ? `Vendor ${vendor.Id}` : null),
    email: vendor.PrimaryEmailAddr?.Address || null,
    phone: vendor.PrimaryPhone?.FreeFormNumber || null,
  };
}

function transformQuickBooksAccount(account) {
  if (!account) {
    return null;
  }

  return {
    id: account.Id,
    name: account.Name || account.FullyQualifiedName || `Account ${account.Id}`,
    fullyQualifiedName: account.FullyQualifiedName || null,
    accountType: account.AccountType || null,
    accountSubType: account.AccountSubType || null,
  };
}

async function createQuickBooksVendor(realmId, details) {
  if (!realmId) {
    const error = new Error('Realm ID is required.');
    error.code = 'BAD_REQUEST';
    throw error;
  }

  const payload = {
    DisplayName: details.displayName,
    CompanyName: details.displayName,
    PrintOnCheckName: details.displayName,
  };

  if (details.email) {
    payload.PrimaryEmailAddr = { Address: details.email };
  }

  if (details.phone) {
    payload.PrimaryPhone = { FreeFormNumber: details.phone };
  }

  const vendor = await performQuickBooksFetch(realmId, async (accessToken) => {
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/vendor?minorversion=65`;
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        'Content-Type': 'application/json',
        'User-Agent': 'InvoiceToQB/1.0',
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorBody = await safeReadJson(response);
      const message =
        errorBody?.Fault?.Error?.[0]?.Detail ||
        errorBody?.Fault?.Error?.[0]?.Message ||
        `QuickBooks vendor creation failed with status ${response.status}`;
      const err = new Error(message);
      err.isQuickBooksError = true;
      err.status = response.status;
      err.details = errorBody;
      throw err;
    }

    const data = await response.json();
    return data?.Vendor || data;
  });

  const normalised = transformQuickBooksVendor(vendor);
  if (!normalised?.id) {
    const error = new Error('QuickBooks vendor creation did not return an identifier.');
    error.isQuickBooksError = true;
    throw error;
  }

  await upsertQuickBooksVendorMetadata(realmId, normalised);
  return normalised;
}

async function createQuickBooksAccount(realmId, details) {
  if (!realmId) {
    const error = new Error('Realm ID is required.');
    error.code = 'BAD_REQUEST';
    throw error;
  }

  const payload = {
    Name: details.name,
    AccountType: details.accountType || 'Expense',
  };

  if (details.accountSubType) {
    payload.AccountSubType = details.accountSubType;
  }

  const account = await performQuickBooksFetch(realmId, async (accessToken) => {
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/account?minorversion=65`;
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        'Content-Type': 'application/json',
        'User-Agent': 'InvoiceToQB/1.0',
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorBody = await safeReadJson(response);
      const message =
        errorBody?.Fault?.Error?.[0]?.Detail ||
        errorBody?.Fault?.Error?.[0]?.Message ||
        `QuickBooks account creation failed with status ${response.status}`;
      const err = new Error(message);
      err.isQuickBooksError = true;
      err.status = response.status;
      err.details = errorBody;
      throw err;
    }

    const data = await response.json();
    return data?.Account || data;
  });

  const normalised = transformQuickBooksAccount(account);
  if (!normalised?.id) {
    const error = new Error('QuickBooks account creation did not return an identifier.');
    error.isQuickBooksError = true;
    throw error;
  }

  await upsertQuickBooksAccountMetadata(realmId, normalised);
  return normalised;
}

async function fetchQuickBooksAccounts(realmId, accessToken) {
  const accounts = await fetchQuickBooksQueryList(realmId, accessToken, 'Account');
  return accounts
    .filter((account) => account.Active !== false)
    .map(transformQuickBooksAccount)
    .filter(Boolean);
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

  const { tokens, oneDrive, ...rest } = company;
  return {
    ...rest,
    businessType: rest.businessType ?? null,
    oneDrive: sanitizeOneDriveSettings(oneDrive),
  };
}

function sanitizeOneDriveSettings(config) {
  const normalized = normalizeOneDriveState(config);
  if (!normalized) {
    return null;
  }

  const { deltaLink, clientState, ...rest } = normalized;
  return rest;
}

function ensureOneDriveStateDefaults(config) {
  const normalized = normalizeOneDriveState(config);
  if (!normalized) {
    return null;
  }

  const result = { ...normalized };
  const now = new Date().toISOString();

  if (!Object.prototype.hasOwnProperty.call(result, 'enabled')) {
    result.enabled = true;
  } else if (result.enabled !== false) {
    result.enabled = true;
  }

  const defaultStatus = result.enabled === false ? 'disabled' : 'connected';
  if (!Object.prototype.hasOwnProperty.call(result, 'status') || !result.status) {
    result.status = defaultStatus;
  }

  if (!Object.prototype.hasOwnProperty.call(result, 'createdAt') || !result.createdAt) {
    result.createdAt = now;
  }

  if (!Object.prototype.hasOwnProperty.call(result, 'updatedAt') || !result.updatedAt) {
    result.updatedAt = now;
  }

  return result;
}

function normalizeOneDriveState(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const result = {};

  if (Object.prototype.hasOwnProperty.call(config, 'enabled')) {
    result.enabled = config.enabled === false ? false : true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'status')) {
    const statusValue = sanitiseOptionalString(config.status);
    result.status = statusValue || null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'driveId')) {
    result.driveId = sanitiseOptionalString(config.driveId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'folderId')) {
    result.folderId = sanitiseOptionalString(config.folderId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'folderPath')) {
    result.folderPath = sanitiseOptionalString(config.folderPath);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'folderName')) {
    result.folderName = sanitiseOptionalString(config.folderName);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'webUrl')) {
    result.webUrl = sanitiseOptionalString(config.webUrl);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'shareUrl')) {
    result.shareUrl = sanitiseOptionalString(config.shareUrl);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'parentId')) {
    result.parentId = sanitiseOptionalString(config.parentId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'deltaLink')) {
    result.deltaLink = sanitiseOptionalString(config.deltaLink);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncAt')) {
    result.lastSyncAt = sanitiseIsoString(config.lastSyncAt);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncStatus')) {
    result.lastSyncStatus = sanitiseOptionalString(config.lastSyncStatus);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncReason')) {
    result.lastSyncReason = sanitiseOptionalString(config.lastSyncReason);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncDurationMs')) {
    result.lastSyncDurationMs = Number.isFinite(config.lastSyncDurationMs)
      ? Number(config.lastSyncDurationMs)
      : null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncError')) {
    result.lastSyncError = normalizeOneDriveSyncError(config.lastSyncError);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncMetrics')) {
    result.lastSyncMetrics = normalizeOneDriveMetrics(config.lastSyncMetrics);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'subscriptionId')) {
    result.subscriptionId = sanitiseOptionalString(config.subscriptionId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'subscriptionExpiration')) {
    result.subscriptionExpiration = sanitiseIsoString(config.subscriptionExpiration);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'clientState')) {
    result.clientState = sanitiseOptionalString(config.clientState);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'createdAt')) {
    result.createdAt = sanitiseIsoString(config.createdAt);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'updatedAt')) {
    result.updatedAt = sanitiseIsoString(config.updatedAt);
  }

  if (!Object.keys(result).length) {
    return {};
  }

  return result;
}

function normalizeOneDriveSyncError(error) {
  if (!error || typeof error !== 'object') {
    return null;
  }

  const message = sanitiseOptionalString(error.message || error.error || error.description);
  if (!message) {
    return null;
  }

  const at = sanitiseIsoString(error.at || error.timestamp) || new Date().toISOString();
  return {
    message,
    at,
  };
}

function normalizeOneDriveMetrics(metrics) {
  if (!metrics || typeof metrics !== 'object') {
    return null;
  }

  const processedItems = Number.isFinite(metrics.processedItems) ? Number(metrics.processedItems) : 0;
  const createdCount = Number.isFinite(metrics.createdCount) ? Number(metrics.createdCount) : 0;
  const skippedCount = Number.isFinite(metrics.skippedCount) ? Number(metrics.skippedCount) : 0;
  const pages = Number.isFinite(metrics.pages) ? Number(metrics.pages) : 0;
  const errorCount = Number.isFinite(metrics.errorCount) ? Number(metrics.errorCount) : 0;

  return {
    processedItems,
    createdCount,
    skippedCount,
    pages,
    errorCount,
  };
}

function sanitiseIsoString(value) {
  const text = sanitiseOptionalString(value);
  if (!text) {
    return null;
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed.toISOString();
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


function logPreprocessingResult(originalName, preprocessing) {
  const timestamp = new Date().toISOString();
  const fileLabel = originalName || 'unknown file';

  if (!preprocessing) {
    console.log(`[${timestamp}] Preprocessing produced no result for ${fileLabel}.`);
    return;
  }

  const method = preprocessing.method || 'unknown';
  const textPreview = preprocessing.text
    ? createLogPreviewText(preprocessing.text, 500)
    : '[no text extracted]';

  console.log(
    `[${timestamp}] Preprocessing (${method}) for ${fileLabel}: ${textPreview}`
  );
}

function createLogPreviewText(value, limit) {
  if (!value) {
    return '[empty]';
  }

  const { text, truncated } = truncateText(value, limit);
  if (truncated) {
    return `${text} (truncated from ${value.length} chars)`;
  }
  return text;
}

function logGeminiPrompt(stage, sourceName, contents = []) {
  const timestamp = new Date().toISOString();
  const label = sourceName || 'unknown file';

  const formatted = contents
    .map((message, index) => {
      const role = message?.role || 'unknown-role';
      const parts = Array.isArray(message?.parts) ? message.parts : [];
      const partDetails = parts
        .map((part) => {
          if (typeof part?.text === 'string') {
            return createLogPreviewText(part.text, 2000);
          }

          if (part?.inlineData) {
            const mimeType = part.inlineData.mimeType || 'application/octet-stream';
            const base64Length = part.inlineData.data ? part.inlineData.data.length : 0;
            return `[inlineData mimeType=${mimeType} size=${base64Length} base64 chars]`;
          }

          return '[unsupported part]';
        })
        .join('\n---\n');

      return `Message ${index + 1} (${role}):\n${partDetails}`;
    })
    .join('\n===\n');

  console.log(
    `[${timestamp}] Gemini ${stage} prompt for ${label}:\n${formatted || '[no prompt parts]'}`
  );
}


async function extractWithGemini(buffer, mimeType, options = {}) {
  const { extractedText, originalName, businessType } = options;
  const url = buildGeminiUrl();
  const businessContext = businessType
    ? `\nBusiness context:\n- The company operates as: ${businessType}\n`
    : '';
  const prompt = `You are an expert system for reading invoices. Extract the following fields and respond with **only** valid JSON matching this schema:
{
  "vendor": string | null,
  "products": [
    {
      "description": string,
      "quantity": number | null,
      "unitPrice": number | null,
      "lineTotal": number | null,
      "taxCode": string | null,
      "suggestedAccount": {
        "name": string | null,
        "accountType": string | null,
        "accountSubType": string | null,
        "confidence": "high" | "medium" | "low" | "unknown",
        "reason": string | null
      }
    }
  ],
  "invoiceDate": string | null, // ISO 8601 (YYYY-MM-DD) when possible
  "invoiceNumber": string | null,
  "billTo": string | null,
  "subtotal": number | null,
  "vatAmount": number | null,
  "totalAmount": number | null,
  "taxCode": string | null,
  "currency": string | null, // 3-letter ISO code when available
  "suggestedAccount": {
    "name": string | null,
    "accountType": string | null,
    "accountSubType": string | null,
    "confidence": "high" | "medium" | "low" | "unknown",
    "reason": string | null
  }
}
Rules:
- If the document text is provided below, rely exclusively on that text.
- Otherwise, analyze the attached file contents.
- Use numbers for monetary values without currency symbols.
- If a field is missing, set it to null.
- Combine multiple bill recipients into a single string if necessary.
- For products, include at least a description. Omit other product fields if not present by setting them to null.
- Do not add commentary. Return only minified JSON with the exact fields above.
Additional guidance:
- Use the business context to determine which QuickBooks Online account is most appropriate for the entire invoice and each product.
- When possible, set accountType to a valid QuickBooks Online AccountType (e.g., Expense, Cost of Goods Sold, Other Expense, Fixed Asset) and accountSubType to a valid detail type (e.g., RepairsMaintenance, SuppliesMaterialsCogs). If unsure, set them to null.
- Confidence must be "high", "medium", "low", or "unknown". Use "high" only when the match is very strong for the business context.
- Provide a brief reason that references invoice content explaining the recommendation.`;
  const effectivePrompt = businessContext ? `${prompt}${businessContext}` : `${prompt}\n`;

  let contents;
  if (extractedText && extractedText.trim()) {
    const documentText = extractedText.trim();
    contents = [
      {
        role: 'user',
        parts: [
          { text: `${effectivePrompt}\n\nDocument text:\n${documentText}` },
        ],
      },
    ];
  } else {
    contents = [
      {
        role: 'user',
        parts: [
          { text: effectivePrompt },
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
  logGeminiPrompt('extraction', originalName, payload.contents);

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

  const parsed = JSON.parse(jsonText);
  return sanitiseGeminiInvoice(parsed);
}

function sanitiseGeminiInvoice(raw) {
  if (!raw || typeof raw !== 'object') {
    return {
      vendor: null,
      products: [],
      invoiceDate: null,
      invoiceNumber: null,
      billTo: null,
      subtotal: null,
      vatAmount: null,
      totalAmount: null,
      taxCode: null,
      currency: null,
      suggestedAccount: null,
    };
  }

  const products = Array.isArray(raw.products)
    ? raw.products.map(sanitiseGeminiProduct).filter(Boolean)
    : [];

  return {
    vendor: sanitiseOptionalString(raw.vendor),
    products,
    invoiceDate: sanitiseOptionalString(raw.invoiceDate),
    invoiceNumber: sanitiseOptionalString(raw.invoiceNumber),
    billTo: sanitiseOptionalString(raw.billTo),
    subtotal: sanitiseNumericValue(raw.subtotal),
    vatAmount: sanitiseNumericValue(raw.vatAmount),
    totalAmount: sanitiseNumericValue(raw.totalAmount),
    taxCode: sanitiseOptionalString(raw.taxCode),
    currency: sanitiseCurrencyCode(raw.currency),
    suggestedAccount: sanitiseGeminiAccountSuggestion(raw.suggestedAccount),
  };
}

function sanitiseGeminiProduct(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const description = sanitiseOptionalString(entry.description);
  const quantity = sanitiseNumericValue(entry.quantity, { allowTextFraction: true });
  const unitPrice = sanitiseNumericValue(entry.unitPrice);
  const lineTotal = sanitiseNumericValue(entry.lineTotal);
  const taxCode = sanitiseOptionalString(entry.taxCode);
  const suggestedAccount = sanitiseGeminiAccountSuggestion(entry.suggestedAccount);

  if (!description && quantity === null && unitPrice === null && lineTotal === null && !taxCode && !suggestedAccount) {
    return null;
  }

  return {
    description,
    quantity,
    unitPrice,
    lineTotal,
    taxCode,
    suggestedAccount,
  };
}

function sanitiseGeminiAccountSuggestion(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const name = sanitiseOptionalString(entry.name);
  const accountType = sanitiseOptionalString(entry.accountType);
  const accountSubType = sanitiseOptionalString(entry.accountSubType);
  const confidenceRaw = typeof entry.confidence === 'string' ? entry.confidence.toLowerCase() : '';
  const confidence = AI_ACCOUNT_CONFIDENCE_VALUES.has(confidenceRaw) ? confidenceRaw : 'unknown';
  const reason = sanitiseOptionalString(entry.reason);

  if (!name && !accountType && !accountSubType && !reason && confidence === 'unknown') {
    return null;
  }

  return {
    name,
    accountType,
    accountSubType,
    confidence,
    reason,
  };
}

function sanitiseOptionalString(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const text = String(value).trim();
  return text ? text : null;
}

function normalizeRemoteSourceMetadata(remoteSource) {
  if (!remoteSource || typeof remoteSource !== 'object') {
    return null;
  }

  const provider = sanitiseOptionalString(remoteSource.provider)?.toLowerCase() || 'unknown';
  const syncedAtSource = sanitiseOptionalString(remoteSource.syncedAt);
  let syncedAt = new Date().toISOString();
  if (syncedAtSource) {
    const parsed = new Date(syncedAtSource);
    if (!Number.isNaN(parsed.getTime())) {
      syncedAt = parsed.toISOString();
    }
  }

  const entry = {
    provider,
    driveId: sanitiseOptionalString(remoteSource.driveId || remoteSource.storageId || remoteSource.parentDriveId),
    itemId: sanitiseOptionalString(remoteSource.itemId || remoteSource.resourceId),
    parentId: sanitiseOptionalString(remoteSource.parentId || remoteSource.parentReferenceId),
    path: sanitiseOptionalString(remoteSource.path || remoteSource.parentPath),
    webUrl: sanitiseOptionalString(remoteSource.webUrl || remoteSource.url),
    eTag: sanitiseOptionalString(remoteSource.eTag || remoteSource.etag),
    lastModifiedDateTime: sanitiseOptionalString(remoteSource.lastModifiedDateTime),
    syncedAt,
  };

  if (!entry.itemId && !entry.webUrl && !entry.path) {
    return null;
  }

  return entry;
}

function isSameRemoteSource(existing, candidate) {
  if (!existing || !candidate) {
    return false;
  }

  if (existing.provider !== candidate.provider) {
    return false;
  }

  if (existing.provider === 'onedrive') {
    const driveMatches = existing.driveId && candidate.driveId ? existing.driveId === candidate.driveId : true;
    if (!driveMatches) {
      return false;
    }
    if (existing.itemId && candidate.itemId && existing.itemId === candidate.itemId) {
      return true;
    }
    if (existing.webUrl && candidate.webUrl && existing.webUrl === candidate.webUrl) {
      return true;
    }
    return false;
  }

  if (existing.itemId && candidate.itemId && existing.itemId === candidate.itemId) {
    return true;
  }

  if (existing.path && candidate.path && existing.path === candidate.path) {
    return true;
  }

  if (existing.webUrl && candidate.webUrl && existing.webUrl === candidate.webUrl) {
    return true;
  }

  return false;
}

function sanitiseNumericValue(value, { allowTextFraction = false } = {}) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  if (allowTextFraction) {
    const fractionMatch = trimmed.match(/^(\d+)\s*\/\s*(\d+)$/);
    if (fractionMatch) {
      const numerator = Number(fractionMatch[1]);
      const denominator = Number(fractionMatch[2]);
      if (Number.isFinite(numerator) && Number.isFinite(denominator) && denominator !== 0) {
        return numerator / denominator;
      }
    }
  }

  const normalised = trimmed.replace(/[^0-9.+-]/g, '');
  if (!normalised || normalised === '.' || normalised === '-' || normalised === '+') {
    return null;
  }

  const parsed = Number(normalised);
  return Number.isFinite(parsed) ? parsed : null;
}

function sanitiseCurrencyCode(value) {
  const text = sanitiseOptionalString(value);
  if (!text) {
    return null;
  }

  const letters = text.replace(/[^a-zA-Z]/g, '');
  if (letters.length < 3) {
    return null;
  }

  return letters.slice(0, 3).toUpperCase();
}

async function findStoredInvoiceByChecksum(checksum) {
  if (!checksum) {
    return null;
  }

  const invoices = await readStoredInvoices().catch(() => []);
  return invoices.find((entry) => entry?.metadata?.checksum === checksum) || null;
}

function rankVendorCandidates(invoice, vendors) {
  const normalizedVendor = normaliseComparableText(invoice?.data?.vendor);
  if (!normalizedVendor || !Array.isArray(vendors)) {
    return [];
  }

  const candidates = vendors
    .map((vendor) => {
      const names = [vendor.displayName, vendor.name].filter(Boolean);
      let bestScore = 0;
      let matchedLabel = vendor.displayName || vendor.name || '';

      names.forEach((name) => {
        const normalizedCandidate = normaliseComparableText(name);
        if (!normalizedCandidate) {
          return;
        }

        const similarity = computeVendorSimilarity(normalizedVendor, normalizedCandidate);
        if (similarity.score > bestScore) {
          bestScore = similarity.score;
          matchedLabel = similarity.label;
        }
      });

      return {
        vendor,
        score: bestScore,
        matchedLabel,
      };
    })
    .filter((entry) => entry.score >= 0.4)
    .sort((a, b) => b.score - a.score);

  return candidates;
}

function rankAccountCandidates(invoice, accounts) {
  if (!Array.isArray(accounts) || !accounts.length) {
    return [];
  }

  const keywords = extractInvoiceKeywords(invoice);
  if (!keywords.length) {
    return [];
  }

  return accounts
    .map((account) => {
      const haystack = normaliseComparableText([
        account.name,
        account.fullyQualifiedName,
        account.accountType,
        account.accountSubType,
      ].filter(Boolean).join(' '));

      if (!haystack) {
        return null;
      }

      const tokens = haystack.split(' ').filter(Boolean);
      const { score, matchedKeywords } = computeKeywordScore(keywords, haystack, tokens);
      if (!score) {
        return null;
      }

      return {
        account,
        score,
        matchedKeywords,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
}

async function matchInvoiceWithGemini({ invoice, vendorOptions, accountOptions, sourceName, businessType }) {
  const url = buildGeminiUrl();
  const keywords = extractInvoiceKeywords(invoice);
  const invoiceSummary = buildInvoiceSummary(invoice, keywords);
  const vendorList = vendorOptions
    .map((entry, index) => {
      const vendor = entry.vendor;
      const name = vendor.displayName || vendor.name || `Vendor ${vendor.id}`;
      return `${index + 1}. id: ${vendor.id} • name: ${name}`;
    })
    .join('\n');

  const accountList = accountOptions
    .map((entry, index) => {
      const account = entry.account;
      const name = account.name || account.fullyQualifiedName || `Account ${account.id}`;
      const type = [account.accountType, account.accountSubType].filter(Boolean).join(' / ');
      return `${index + 1}. id: ${account.id} • name: ${name}${type ? ` • type: ${type}` : ''}`;
    })
    .join('\n');

  const businessContext = businessType ? `\nBusiness context:\n- The company operates as: ${businessType}\n` : '';
  const prompt = `You help match invoices to QuickBooks vendors and accounts. Choose the best options from the lists below.${businessContext}

Invoice summary:
${invoiceSummary}

QuickBooks vendors:
${vendorList || 'None provided'}

QuickBooks accounts:
${accountList || 'None provided'}

Respond with JSON only, matching this schema:
{
  "vendor": {
    "vendorId": string | null,
    "confidence": "high" | "medium" | "low",
    "reason": string
  },
  "account": {
    "accountId": string | null,
    "confidence": "high" | "medium" | "low",
    "reason": string
  }
}
Rules:
- Prefer vendors and accounts that clearly relate to the invoice text.
- Pick the closest option even if confidence is low; use null only when nothing fits.
- Keep reasons under 80 characters and reference evidence from the invoice.
- Do not add commentary outside the JSON.`;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }],
      },
    ],
  };

  const promptSourceName = sourceName || invoice?.metadata?.originalName || null;
  logGeminiPrompt('matching', promptSourceName, payload.contents);

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
    const err = new Error('Gemini API returned no content for matching.');
    err.isGeminiError = true;
    throw err;
  }

  const jsonText = extractJson(candidateText);
  if (!jsonText) {
    const err = new Error('Unable to extract JSON from Gemini matching response.');
    err.isGeminiError = true;
    throw err;
  }

  return JSON.parse(jsonText);
}

function buildInvoiceSummary(invoice, keywords) {
  const lines = [];
  const vendor = invoice?.data?.vendor ? `Vendor text: "${invoice.data.vendor}"` : 'Vendor text: Unknown';
  lines.push(vendor);

  if (invoice?.data?.invoiceNumber) {
    lines.push(`Invoice number: ${invoice.data.invoiceNumber}`);
  }
  if (invoice?.data?.invoiceDate) {
    lines.push(`Invoice date: ${invoice.data.invoiceDate}`);
  }
  if (invoice?.data?.totalAmount !== undefined && invoice?.data?.totalAmount !== null) {
    lines.push(`Total amount: ${invoice.data.totalAmount}`);
  }

  if (Array.isArray(invoice?.data?.products) && invoice.data.products.length) {
    const sampled = invoice.data.products.slice(0, 5).map((product, index) => {
      const description = product?.description ? product.description.trim().replace(/\s+/g, ' ') : 'No description';
      const lineTotal = product?.lineTotal !== undefined && product?.lineTotal !== null ? ` • line total ${product.lineTotal}` : '';
      return `${index + 1}. ${description}${lineTotal}`;
    });
    lines.push('Line items:');
    lines.push(sampled.join('\n'));
  }

  if (keywords.length) {
    lines.push(`Keywords: ${keywords.slice(0, 12).join(', ')}`);
  }

  return lines.join('\n');
}

function mapAiVendorSuggestion(suggestion, vendorLookup) {
  if (!suggestion) {
    return {
      vendorId: null,
      displayName: null,
      confidence: 'unknown',
      reason: null,
    };
  }

  const vendor = suggestion.vendorId ? vendorLookup.get(suggestion.vendorId) : null;
  const confidence = mapAiConfidence(suggestion.confidence);
  return {
    vendorId: vendor?.id || null,
    displayName: vendor?.displayName || vendor?.name || suggestion.displayName || null,
    confidence,
    reason: suggestion.reason || null,
  };
}

function mapAiAccountSuggestion(suggestion, accountLookup) {
  if (!suggestion) {
    return {
      accountId: null,
      name: null,
      accountType: null,
      accountSubType: null,
      confidence: 'unknown',
      reason: null,
    };
  }

  const account = suggestion.accountId ? accountLookup.get(suggestion.accountId) : null;
  const confidence = mapAiConfidence(suggestion.confidence);
  return {
    accountId: account?.id || null,
    name: account?.name || account?.fullyQualifiedName || suggestion.name || null,
    accountType: account?.accountType || null,
    accountSubType: account?.accountSubType || null,
    confidence,
    reason: suggestion.reason || null,
  };
}

function applyVendorDefaultAccount({ vendor, account }, { vendorSettings, accountLookup, allAccountLookup }) {
  const vendorId = vendor?.vendorId;
  const vendorConfidence = vendor?.confidence;
  if (!vendorId || vendorConfidence === 'unknown') {
    return { vendor, account };
  }

  const defaults = vendorSettings?.[vendorId];
  if (!defaults?.accountId) {
    return { vendor, account };
  }

  const resolved = accountLookup?.get(defaults.accountId) || allAccountLookup?.get(defaults.accountId) || null;
  if (resolved) {
    return {
      vendor,
      account: {
        accountId: resolved.id,
        name: resolved.name || resolved.fullyQualifiedName || account?.name || null,
        accountType: resolved.accountType || null,
        accountSubType: resolved.accountSubType || null,
        confidence: vendorConfidence === 'exact' ? 'exact' : 'uncertain',
        reason: 'Vendor default account.',
      },
    };
  }

  const fallbackAccount =
    (account?.accountId === defaults.accountId ? account : null) || allAccountLookup?.get(defaults.accountId) || null;

  return {
    vendor,
    account: {
      accountId: defaults.accountId,
      name: fallbackAccount?.name || fallbackAccount?.fullyQualifiedName || null,
      accountType: fallbackAccount?.accountType || null,
      accountSubType: fallbackAccount?.accountSubType || null,
      confidence: vendorConfidence === 'exact' ? 'exact' : 'uncertain',
      reason: 'Vendor default account.',
    },
  };
}

function mapAiConfidence(confidence) {
  if (!confidence) {
    return 'unknown';
  }
  const value = confidence.toString().toLowerCase();
  if (value === 'high') {
    return 'exact';
  }
  if (value === 'medium') {
    return 'uncertain';
  }
  return 'unknown';
}

function normaliseComparableText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const text = value.toString().toLowerCase();
  const normalized = typeof text.normalize === 'function' ? text.normalize('NFD') : text;

  return normalized
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function computeVendorSimilarity(normalizedVendor, normalizedCandidate) {
  let score = computeNormalisedSimilarity(normalizedVendor, normalizedCandidate);
  let label = normalizedCandidate;

  if (!normalizedVendor || !normalizedCandidate) {
    return { score, label };
  }

  if (normalizedCandidate.includes(normalizedVendor) || normalizedVendor.includes(normalizedCandidate)) {
    score = Math.max(score, 0.9);
  }

  const candidateTokens = normalizedCandidate.split(' ').filter(Boolean);
  candidateTokens.forEach((token) => {
    if (token.length < 3) {
      return;
    }
    const tokenSimilarity = computeNormalisedSimilarity(normalizedVendor, token);
    if (tokenSimilarity > score) {
      score = tokenSimilarity;
      label = token;
    }
    if (normalizedVendor.includes(token) && token.length >= 4) {
      score = Math.max(score, 0.86);
      label = token;
    }
  });

  return { score, label };
}

function computeKeywordScore(keywords, haystack, tokens) {
  const matched = [];
  const tokenList = Array.isArray(tokens) && tokens.length ? tokens : haystack.split(' ').filter(Boolean);

  const score = keywords.reduce((running, keyword) => {
    if (!keyword) {
      return running;
    }

    if (haystack.includes(keyword)) {
      matched.push(keyword);
      return running + (keyword.length >= 6 ? 2 : 1.5);
    }

    const similar = tokenList.find((token) => computeNormalisedSimilarity(keyword, token) >= 0.82);
    if (similar) {
      matched.push(keyword);
      return running + (keyword.length >= 6 ? 1.5 : 1);
    }

    return running;
  }, 0);

  return { score, matchedKeywords: matched.slice(0, 5) };
}

function extractInvoiceKeywords(invoice) {
  const keywords = new Set();
  const pushTokens = (text) => {
    if (!text) {
      return;
    }

    text
      .toString()
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .forEach((token) => {
        if (isMeaningfulKeyword(token)) {
          keywords.add(token);
        }
      });
  };

  pushTokens(invoice?.data?.vendor);
  pushTokens(invoice?.data?.invoiceNumber);
  pushTokens(invoice?.data?.taxCode);

  if (Array.isArray(invoice?.data?.products)) {
    invoice.data.products.forEach((product) => {
      pushTokens(product?.description);
    });
  }

  pushTokens(invoice?.metadata?.originalName);

  return [...keywords];
}

function isMeaningfulKeyword(token) {
  return token && token.length >= 3 && !KEYWORD_STOPWORDS.has(token);
}

function computeNormalisedSimilarity(a, b) {
  if (!a || !b) {
    return 0;
  }
  if (a === b) {
    return 1;
  }
  const distance = levenshteinDistance(a, b);
  const longest = Math.max(a.length, b.length) || 1;
  return (longest - distance) / longest;
}

function levenshteinDistance(a, b) {
  if (a === b) {
    return 0;
  }

  const aLength = a.length;
  const bLength = b.length;

  if (!aLength) {
    return bLength;
  }

  if (!bLength) {
    return aLength;
  }

  const matrix = Array.from({ length: aLength + 1 }, () => new Array(bLength + 1).fill(0));

  for (let i = 0; i <= aLength; i += 1) {
    matrix[i][0] = i;
  }
  for (let j = 0; j <= bLength; j += 1) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= aLength; i += 1) {
    for (let j = 1; j <= bLength; j += 1) {
      if (a[i - 1] === b[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j - 1] + 1,
        );
      }
    }
  }

  return matrix[aLength][bLength];
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
