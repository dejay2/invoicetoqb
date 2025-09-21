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

const PORT = process.env.PORT || 5000;
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
const QUICKBOOKS_MINOR_VERSION = process.env.QUICKBOOKS_MINOR_VERSION || '75';
const QUICKBOOKS_COMPANIES_FILE = process.env.QUICKBOOKS_COMPANIES_FILE
  ? path.resolve(process.env.QUICKBOOKS_COMPANIES_FILE)
  : path.join(__dirname, '..', 'data', 'quickbooks_companies.json');
const QUICKBOOKS_METADATA_DIR = path.join(__dirname, '..', 'data', 'quickbooks');
const QUICKBOOKS_AUTH_URL = 'https://appcenter.intuit.com/connect/oauth2';

let quickBooksCompaniesWriteMutex = Promise.resolve();
const QUICKBOOKS_HEALTH_METRICS = [];
const MAX_HEALTH_METRIC_HISTORY = 100;

function emitHealthMetric(name, dimensions = {}) {
  if (typeof name !== 'string' || !name) {
    return null;
  }

  const entry = {
    name,
    dimensions: dimensions && typeof dimensions === 'object' ? { ...dimensions } : {},
    at: new Date().toISOString(),
  };

  QUICKBOOKS_HEALTH_METRICS.push(entry);
  if (QUICKBOOKS_HEALTH_METRICS.length > MAX_HEALTH_METRIC_HISTORY) {
    QUICKBOOKS_HEALTH_METRICS.shift();
  }

  console.info(`[HealthMetric] ${name}`, entry.dimensions);
  return entry;
}

function getHealthMetricHistory() {
  return QUICKBOOKS_HEALTH_METRICS.slice();
}

async function runWithOneDriveSettingsWriteLock(task) {
  let release;
  const next = new Promise((resolve) => {
    release = resolve;
  });

  const previous = oneDriveSettingsWriteMutex;
  oneDriveSettingsWriteMutex = next;

  try {
    await previous.catch(() => {});
    return await task();
  } finally {
    release();
  }
}

function cloneGlobalOneDriveConfig(value) {
  if (value === null || value === undefined) {
    return null;
  }
  return JSON.parse(JSON.stringify(value));
}

function getDefaultGlobalOneDriveConfig() {
  const now = new Date().toISOString();
  return {
    driveId: null,
    driveName: null,
    driveType: null,
    driveWebUrl: null,
    driveOwner: null,
    shareUrl: null,
    folderId: null,
    folderPath: null,
    folderName: null,
    folderWebUrl: null,
    folderParentId: null,
    status: 'unconfigured',
    lastValidatedAt: null,
    lastValidationError: null,
    lastSyncHealth: null,
    createdAt: now,
    updatedAt: now,
  };
}

function normalizeGlobalOneDriveSyncHealth(health) {
  if (!health || typeof health !== 'object') {
    return null;
  }

  const status = sanitiseOptionalString(health.status) || null;
  const message = sanitiseOptionalString(health.message || health.reason) || null;
  const at = sanitiseIsoString(health.at) || (status || message ? new Date().toISOString() : null);

  if (!status && !message && !at) {
    return null;
  }

  return {
    status,
    message,
    at,
  };
}

function normalizeGlobalOneDriveConfig(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const result = {};

  if (Object.prototype.hasOwnProperty.call(config, 'driveId')) {
    result.driveId = sanitiseOptionalString(config.driveId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'driveName')) {
    result.driveName = sanitiseOptionalString(config.driveName);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'driveType')) {
    result.driveType = sanitiseOptionalString(config.driveType);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'driveWebUrl')) {
    result.driveWebUrl = sanitiseOptionalString(config.driveWebUrl);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'driveOwner')) {
    if (config.driveOwner && typeof config.driveOwner === 'object') {
      const ownerName = sanitiseOptionalString(
        config.driveOwner.name || config.driveOwner.displayName || config.driveOwner.email || config.driveOwner.principal
      );
      result.driveOwner = ownerName || null;
    } else {
      result.driveOwner = sanitiseOptionalString(config.driveOwner);
    }
  }

  if (Object.prototype.hasOwnProperty.call(config, 'shareUrl')) {
    result.shareUrl = sanitiseOptionalString(config.shareUrl);
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

  if (Object.prototype.hasOwnProperty.call(config, 'folderWebUrl')) {
    result.folderWebUrl = sanitiseOptionalString(config.folderWebUrl);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'folderParentId')) {
    result.folderParentId = sanitiseOptionalString(config.folderParentId);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'status')) {
    result.status = sanitiseOptionalString(config.status) || null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastValidatedAt')) {
    result.lastValidatedAt = sanitiseIsoString(config.lastValidatedAt);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastValidationError')) {
    result.lastValidationError = normalizeOneDriveSyncError(config.lastValidationError);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncHealth')) {
    result.lastSyncHealth = normalizeGlobalOneDriveSyncHealth(config.lastSyncHealth);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'createdAt')) {
    result.createdAt = sanitiseIsoString(config.createdAt);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'updatedAt')) {
    result.updatedAt = sanitiseIsoString(config.updatedAt);
  }

  return result;
}

function sanitizeGlobalOneDriveConfig(config) {
  if (!config || typeof config !== 'object') {
    return getDefaultGlobalOneDriveConfig();
  }

  const normalised = normalizeGlobalOneDriveConfig(config) || {};
  const base = {
    ...getDefaultGlobalOneDriveConfig(),
    ...normalised,
  };

  if (!base.driveId) {
    base.status = base.status || 'unconfigured';
  }

  return base;
}

function isGlobalOneDriveConfigured(config) {
  return Boolean(config?.driveId);
}

async function persistGlobalOneDriveConfig(config) {
  return runWithOneDriveSettingsWriteLock(async () => {
    const dir = path.dirname(ONEDRIVE_SETTINGS_FILE);
    const payload = `${JSON.stringify(config, null, 2)}\n`;
    const tempPath = `${ONEDRIVE_SETTINGS_FILE}.tmp-${process.pid}-${Date.now()}`;

    await fs.mkdir(dir, { recursive: true });

    let handle;
    try {
      handle = await fs.open(tempPath, 'w', 0o600);
      await handle.writeFile(payload);
      await handle.sync();
    } finally {
      if (handle) {
        await handle.close().catch(() => {});
      }
    }

    try {
      await fs.rename(tempPath, ONEDRIVE_SETTINGS_FILE);
    } catch (err) {
      await fs.unlink(tempPath).catch(() => {});
      throw err;
    }

    let dirHandle;
    try {
      dirHandle = await fs.open(dir, 'r');
      await dirHandle.sync();
    } catch (dirErr) {
      if (dirErr && dirErr.code !== 'EISDIR' && dirErr.code !== 'ENOENT') {
        console.warn(`[OneDrive] Failed to fsync directory ${dir}: ${dirErr.message}`);
      }
    } finally {
      if (dirHandle) {
        await dirHandle.close().catch(() => {});
      }
    }
  });
}

async function loadGlobalOneDriveConfig({ refresh = false } = {}) {
  if (!hasLoadedOneDriveSettings || refresh) {
    let contents = null;
    try {
      contents = await fs.readFile(ONEDRIVE_SETTINGS_FILE, 'utf-8');
    } catch (error) {
      if (error?.code !== 'ENOENT') {
        throw error;
      }
    }

    if (!contents) {
      globalOneDriveConfigCache = getDefaultGlobalOneDriveConfig();
      await persistGlobalOneDriveConfig(globalOneDriveConfigCache).catch((err) => {
        console.warn('[OneDrive] Unable to seed default OneDrive settings file', err?.message || err);
      });
    } else {
      try {
        const parsed = JSON.parse(contents);
        const normalised = sanitizeGlobalOneDriveConfig(parsed);
        const now = new Date().toISOString();
        if (!normalised.createdAt) {
          normalised.createdAt = now;
        }
        if (!normalised.updatedAt) {
          normalised.updatedAt = now;
        }
        globalOneDriveConfigCache = normalised;
      } catch (error) {
        console.error('[OneDrive] Failed to parse OneDrive settings file. Resetting to defaults.', error);
        globalOneDriveConfigCache = getDefaultGlobalOneDriveConfig();
        await persistGlobalOneDriveConfig(globalOneDriveConfigCache).catch((err) => {
          console.warn('[OneDrive] Unable to persist reset OneDrive settings file', err?.message || err);
        });
      }
    }

    hasLoadedOneDriveSettings = true;
  }

  return cloneGlobalOneDriveConfig(globalOneDriveConfigCache);
}

async function updateGlobalOneDriveConfig(updates, { replace = false } = {}) {
  const current = await loadGlobalOneDriveConfig();
  const now = new Date().toISOString();

  let next;
  if (replace) {
    next = sanitizeGlobalOneDriveConfig(updates);
    next.updatedAt = now;
    if (!next.createdAt) {
      next.createdAt = now;
    }
  } else {
    const normalisedUpdate = normalizeGlobalOneDriveConfig(updates) || {};
    next = {
      ...current,
    };

    for (const key of Object.keys(normalisedUpdate)) {
      const value = normalisedUpdate[key];
      if (value === undefined) {
        continue;
      }
      if (value === null) {
        if (next[key] !== null) {
          next[key] = null;
        }
      } else if (typeof value === 'object' && value !== null) {
        next[key] = cloneGlobalOneDriveConfig(value);
      } else {
        next[key] = value;
      }
    }

    next.updatedAt = now;
    if (!next.createdAt) {
      next.createdAt = now;
    }

    if (!next.status) {
      next.status = next.driveId ? 'ready' : 'unconfigured';
    }
  }

  globalOneDriveConfigCache = sanitizeGlobalOneDriveConfig(next);
  await persistGlobalOneDriveConfig(globalOneDriveConfigCache);
  return cloneGlobalOneDriveConfig(globalOneDriveConfigCache);
}

async function ensureGlobalOneDriveDriveContext() {
  const config = await loadGlobalOneDriveConfig();
  const driveId = sanitiseOptionalString(config?.driveId);
  if (!driveId) {
    const error = new Error('Shared OneDrive drive is not configured.');
    error.code = 'ONEDRIVE_GLOBAL_UNCONFIGURED';
    throw error;
  }
  return { config, driveId };
}

async function runWithQuickBooksCompaniesWriteLock(task) {
  let release;
  const next = new Promise((resolve) => {
    release = resolve;
  });

  const previous = quickBooksCompaniesWriteMutex;
  quickBooksCompaniesWriteMutex = next;

  try {
    await previous.catch(() => {});
    return await task();
  } finally {
    release();
  }
}
const QUICKBOOKS_TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
const QUICKBOOKS_STATE_TTL_MS = 10 * 60 * 1000;

const MS_GRAPH_CLIENT_ID = process.env.MS_GRAPH_CLIENT_ID || '';
const MS_GRAPH_CLIENT_SECRET = process.env.MS_GRAPH_CLIENT_SECRET || '';
const MS_GRAPH_TENANT_ID = process.env.MS_GRAPH_TENANT_ID || '';
const MS_GRAPH_SCOPE = process.env.MS_GRAPH_SCOPE || 'https://graph.microsoft.com/.default';
const MS_GRAPH_SERVICE_USER_ID = (process.env.MS_GRAPH_SERVICE_USER_ID || '').trim();
const MS_GRAPH_SHAREPOINT_SITE_ID = (process.env.MS_GRAPH_SHAREPOINT_SITE_ID || '').trim();
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
const ONEDRIVE_PROCESSED_FOLDER_NAME = (process.env.ONEDRIVE_PROCESSED_FOLDER_NAME || 'Synced').trim() || 'Synced';

const ONEDRIVE_SYNC_ENABLED = Boolean(MS_GRAPH_CLIENT_ID && MS_GRAPH_CLIENT_SECRET && MS_GRAPH_TENANT_ID);

const ONEDRIVE_SETTINGS_FILE = process.env.ONEDRIVE_SETTINGS_FILE
  ? path.resolve(process.env.ONEDRIVE_SETTINGS_FILE)
  : path.join(__dirname, '..', 'data', 'onedrive_settings.json');

let oneDriveSettingsWriteMutex = Promise.resolve();
let globalOneDriveConfigCache = null;
let hasLoadedOneDriveSettings = false;
let hasAppliedOneDriveMigration = false;

loadGlobalOneDriveConfig().catch((error) => {
  console.error('[OneDrive] Failed to load global OneDrive settings', error);
});

let hasLoggedOneDriveServiceIdentityGuidance = false;
let lastOneDriveServiceIdentityMessage = '';

const OCR_IMAGE_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/heic',
  'image/heif',
  'image/tiff',
  'image/webp',
]);

const GMAIL_CLIENT_ID = process.env.GMAIL_CLIENT_ID || '';
const GMAIL_CLIENT_SECRET = process.env.GMAIL_CLIENT_SECRET || '';
const GMAIL_API_BASE_URL = 'https://gmail.googleapis.com/gmail/v1';
const GMAIL_TOKEN_URL = 'https://oauth2.googleapis.com/token';
const GMAIL_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const GMAIL_SCOPES = process.env.GMAIL_SCOPES || 'https://www.googleapis.com/auth/gmail.readonly';
const GMAIL_DEFAULT_CALLBACK_PATH = '/api/gmail/callback';
const GMAIL_DEFAULT_REDIRECT_URI = `http://localhost:${PORT}${GMAIL_DEFAULT_CALLBACK_PATH}`;
const GMAIL_REDIRECT_URI = process.env.GMAIL_REDIRECT_URI || GMAIL_DEFAULT_REDIRECT_URI;
const GMAIL_STATE_TTL_MS = 10 * 60 * 1000;
const GMAIL_DEFAULT_POLL_INTERVAL_MS = Math.max(parseInt(process.env.GMAIL_POLL_INTERVAL_MS || '120000', 10), 15000);
const GMAIL_DEFAULT_MAX_RESULTS = Math.max(parseInt(process.env.GMAIL_MAX_RESULTS || '25', 10), 1);
const GMAIL_DEFAULT_SEARCH_QUERY = process.env.GMAIL_SEARCH_QUERY || 'in:inbox has:attachment';
const GMAIL_DEFAULT_LABEL_IDS = parseDelimitedList(process.env.GMAIL_LABEL_IDS);
const GMAIL_DEFAULT_MAX_ATTACHMENT_BYTES = Math.max(
  parseInt(process.env.GMAIL_MAX_ATTACHMENT_BYTES || `${10 * 1024 * 1024}`, 10),
  1024
);
const GMAIL_DEFAULT_ALLOWED_MIME_TYPES = (() => {
  const envValues = parseDelimitedList(process.env.GMAIL_ALLOWED_MIME_TYPES);
  if (envValues?.length) {
    return envValues.map((value) => value.toLowerCase());
  }
  const defaults = ['application/pdf'];
  for (const mime of OCR_IMAGE_MIME_TYPES) {
    defaults.push(mime.toLowerCase());
  }
  return defaults;
})();
const GMAIL_DEFAULT_BUSINESS_TYPE = (process.env.GMAIL_DEFAULT_BUSINESS_TYPE || '').trim();
const GMAIL_POLL_MIN_INTERVAL_MS = 15000;
const GMAIL_CLIENT_CONFIGURED = Boolean(GMAIL_CLIENT_ID && GMAIL_CLIENT_SECRET);

let graphTokenCache = { token: null, expiresAt: 0 };
const activeOneDrivePolls = new Map();
const oneDriveProcessedFolderCache = new Map();
let oneDrivePollingTimer = null;

const gmailTokenCache = new Map();
const gmailStateCache = new Map();
const activeGmailPolls = new Map();
let gmailPollingTimer = null;

const quickBooksStates = new Map();
const quickBooksCallbackPaths = getQuickBooksCallbackPaths();
const gmailStates = new Map();
const gmailCallbackPaths = getGmailCallbackPaths();

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

function getGmailCallbackPaths() {
  const paths = new Set([GMAIL_DEFAULT_CALLBACK_PATH]);
  const derived = derivePathFromUrl(GMAIL_REDIRECT_URI);
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
    console.warn('Unable to determine callback path from redirect URI', error);
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
  defaultStatus = 'archive',
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

  const resolvedDefaultStatus = defaultStatus === 'review' ? 'review' : 'archive';
  const storageStatus = duplicateMatch ? 'archive' : resolvedDefaultStatus;

  const storagePayload = {
    ...enrichedInvoice,
    duplicateOf: duplicateMatch?.metadata?.checksum || null,
    status: storageStatus,
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

function logOneDriveServiceIdentityWarning(message) {
  const text = typeof message === 'string' ? message.trim() : '';
  if (!text) {
    return;
  }

  if (hasLoggedOneDriveServiceIdentityGuidance && lastOneDriveServiceIdentityMessage === text) {
    return;
  }

  console.warn(`[OneDrive] ${text}`);
  hasLoggedOneDriveServiceIdentityGuidance = true;
  lastOneDriveServiceIdentityMessage = text;
}

function deriveGraphDriveOwner(owner) {
  if (!owner || typeof owner !== 'object') {
    return null;
  }

  const userOwner = owner.user || owner.application || owner.group || owner.site;
  if (userOwner && typeof userOwner === 'object') {
    const name = sanitiseOptionalString(userOwner.displayName);
    if (name) {
      return name;
    }
    const principal = sanitiseOptionalString(userOwner.userPrincipalName || userOwner.email);
    if (principal) {
      return principal;
    }
  }

  const groupName = sanitiseOptionalString(owner.group?.displayName);
  if (groupName) {
    return groupName;
  }

  const applicationName = sanitiseOptionalString(owner.application?.displayName);
  if (applicationName) {
    return applicationName;
  }

  const siteDisplayName = sanitiseOptionalString(owner.site?.displayName);
  if (siteDisplayName) {
    return siteDisplayName;
  }

  return null;
}

function sanitiseGraphDrive(drive) {
  if (!drive || typeof drive !== 'object') {
    return null;
  }

  const id = sanitiseOptionalString(drive.id);
  if (!id) {
    return null;
  }

  return {
    id,
    name: sanitiseOptionalString(drive.name) || 'Drive',
    driveType: sanitiseOptionalString(drive.driveType),
    webUrl: sanitiseOptionalString(drive.webUrl),
    owner: deriveGraphDriveOwner(drive.owner),
  };
}

function sanitiseGraphDriveItem(item) {
  if (!item || typeof item !== 'object') {
    return null;
  }

  const id = sanitiseOptionalString(item.id);
  if (!id) {
    return null;
  }

  const path = buildDriveItemPath(item);

  return {
    id,
    name: sanitiseOptionalString(item.name) || 'Item',
    driveId: sanitiseOptionalString(item.parentReference?.driveId),
    parentId: sanitiseOptionalString(item.parentReference?.id),
    path,
    displayPath: buildDriveItemDisplayPath(path),
    webUrl: sanitiseOptionalString(item.webUrl),
    isFolder: Boolean(item.folder),
    childCount: Number.isFinite(item.folder?.childCount) ? Number(item.folder.childCount) : null,
  };
}

async function graphListDrives({ driveId: requestedDriveId } = {}) {
  if (!isOneDriveSyncConfigured()) {
    throw new Error('OneDrive integration is not configured.');
  }

  const driveId = sanitiseOptionalString(requestedDriveId);
  const drives = [];
  const seen = new Set();
  let warning = null;

  const includeDrive = (drive) => {
    const normalised = sanitiseGraphDrive(drive);
    if (!normalised || seen.has(normalised.id)) {
      return;
    }
    seen.add(normalised.id);
    drives.push(normalised);
  };

  const serviceUserId = sanitiseOptionalString(MS_GRAPH_SERVICE_USER_ID);
  const siteId = sanitiseOptionalString(MS_GRAPH_SHAREPOINT_SITE_ID);

  if (!serviceUserId && !siteId && !driveId) {
    warning = 'Set MS_GRAPH_SERVICE_USER_ID to enable OneDrive browsing.';
    logOneDriveServiceIdentityWarning(warning);
    return { drives: [], warning };
  }

  const selectFields = 'id,name,driveType,webUrl,owner';

  if (serviceUserId) {
    const response = await graphFetch(`/users/${encodeURIComponent(serviceUserId)}/drives`, {
      query: { $select: selectFields },
    });
    const entries = Array.isArray(response?.value) ? response.value : [];
    entries.forEach(includeDrive);
  }

  if (siteId) {
    const response = await graphFetch(`/sites/${encodeURIComponent(siteId)}/drives`, {
      query: { $select: selectFields },
    });
    const entries = Array.isArray(response?.value) ? response.value : [];
    entries.forEach(includeDrive);
  }

  if (driveId && !seen.has(driveId)) {
    try {
      const drive = await graphFetch(`/drives/${encodeURIComponent(driveId)}`, {
        query: { $select: selectFields },
      });
      includeDrive(drive);
    } catch (error) {
      if (error?.status !== 404) {
        throw error;
      }
    }
  }

  return { drives, warning };
}

async function graphListDriveChildren({ driveId, itemId, path }) {
  if (!isOneDriveSyncConfigured()) {
    throw new Error('OneDrive integration is not configured.');
  }

  const resolvedDriveId = sanitiseOptionalString(driveId);
  if (!resolvedDriveId) {
    throw new Error('Drive ID is required.');
  }

  const resolvedItemId = sanitiseOptionalString(itemId);
  const resolvedPath = sanitiseOptionalString(path);

  let resource;
  if (resolvedItemId) {
    resource = `/drives/${encodeURIComponent(resolvedDriveId)}/items/${encodeURIComponent(resolvedItemId)}/children`;
  } else if (resolvedPath) {
    const normalisedPath = normaliseGraphFolderPath(resolvedPath);
    resource = `/drives/${encodeURIComponent(resolvedDriveId)}/root:${normalisedPath}:/children`;
  } else {
    resource = `/drives/${encodeURIComponent(resolvedDriveId)}/root/children`;
  }

  const response = await graphFetch(resource, {
    query: {
      $select: 'id,name,parentReference,webUrl,folder',
      $top: 200,
    },
  });

  const entries = Array.isArray(response?.value) ? response.value : [];
  return entries.map(sanitiseGraphDriveItem).filter(Boolean);
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

async function resolveOneDriveFolderReference({ shareUrl, driveId, folderId, folderPath }, { expectedDriveId = null } = {}) {
  if (!isOneDriveSyncConfigured()) {
    throw new Error('OneDrive integration is not configured.');
  }

  let item;
  let resolvedDriveId = sanitiseOptionalString(driveId);
  const trimmedShareUrl = sanitiseOptionalString(shareUrl);
  const enforcedDriveId = sanitiseOptionalString(expectedDriveId);

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
  const itemDriveId = item.parentReference?.driveId || resolvedDriveId || null;

  if (enforcedDriveId && itemDriveId && itemDriveId !== enforcedDriveId) {
    const error = new Error('Selected OneDrive folder does not belong to the shared drive.');
    error.code = 'ONEDRIVE_FOLDER_MISMATCH';
    throw error;
  }

  return {
    id: item.id,
    driveId: itemDriveId,
    parentId: item.parentReference?.id || null,
    name: item.name || 'Folder',
    path: displayPath,
    webUrl: item.webUrl || null,
    shareUrl: trimmedShareUrl || null,
    isFolder: true,
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
    if (error?.code === 'QUICKBOOKS_COMPANY_FILE_CORRUPT') {
      console.error(
        'OneDrive poll enumeration halted: QuickBooks company store is corrupt. Restore the latest backup before resuming sync.'
      );
      emitHealthMetric('quickbooks.company_file.corrupt', { source: 'onedrive_poll' });
    } else {
      console.error('OneDrive poll enumeration failed', error);
    }
  }
}

async function queueOneDrivePoll(realmId, { reason = 'manual', forceFull = false } = {}) {
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
      await pollOneDriveForCompany(company, { reason, forceFull });
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

async function pollOneDriveForCompany(company, { reason = 'manual', forceFull = false } = {}) {
  const config = ensureOneDriveStateDefaults(company.oneDrive);
  if (!config || config.enabled === false) {
    return;
  }

  let driveContext;
  try {
    driveContext = await ensureGlobalOneDriveDriveContext();
  } catch (error) {
    await updateQuickBooksCompanyOneDrive(company.realmId, {
      status: 'error',
      lastSyncStatus: 'error',
      lastSyncReason: reason,
      lastSyncError: {
        message: error.message || 'Shared OneDrive drive is not configured.',
        at: new Date().toISOString(),
      },
    });
    return;
  }

  const monitoredFolder = config.monitoredFolder || null;
  if (!monitoredFolder?.id) {
    await updateQuickBooksCompanyOneDrive(company.realmId, {
      status: 'error',
      lastSyncStatus: 'error',
      lastSyncReason: reason,
      lastSyncError: {
        message: 'OneDrive folder is not fully configured. Choose a monitored folder from the shared drive.',
        at: new Date().toISOString(),
      },
    });
    return;
  }

  let syncReason = reason;
  const startedAt = Date.now();
  const useDeltaLink = forceFull ? null : config.deltaLink;
  let nextLink = useDeltaLink
    ? useDeltaLink
    : `${GRAPH_API_BASE_URL.replace(/\/+$/, '')}/drives/${encodeURIComponent(driveContext.driveId)}/items/${encodeURIComponent(monitoredFolder.id)}/delta`;
  let latestDelta = forceFull ? null : config.deltaLink || null;
  let processedItems = 0;
  let createdCount = 0;
  let skippedCount = 0;
  let duplicateCount = 0;
  let pages = 0;
  const errors = [];
  const activityLog = [];

  while (nextLink && pages < ONEDRIVE_MAX_DELTA_PAGES) {
    let response;
    try {
      response = await graphFetch(nextLink, { responseType: 'json' });
    } catch (error) {
      if (isOneDriveDeltaResetError(error)) {
        const wasForceFull = forceFull;
        activityLog.push(
          wasForceFull
            ? 'Change feed token expired during full resync; performing snapshot crawl.'
            : 'Change feed token expired; resetting stored token and switching to a full snapshot crawl.'
        );
        latestDelta = null;
        nextLink = null;
        forceFull = true;
        if (!wasForceFull && syncReason !== 'full-resync') {
          syncReason = 'auto-resync';
        }
      } else {
        errors.push(error);
        const message = error?.message || 'Failed to fetch OneDrive change feed.';
        activityLog.push(`Error fetching OneDrive changes: ${message}`);
      }
      break;
    }

    const entries = Array.isArray(response?.value) ? response.value : [];
    for (const item of entries) {
      if (!item || item.deleted || !item.file) {
        continue;
      }
      processedItems += 1;
      try {
        const result = await processOneDriveItem(company, config, driveContext, item, { reason: syncReason });
        const displayName = result?.originalName || item.name || item.id;

        if (result?.skipped) {
          skippedCount += 1;
          const skipReason = result?.skipReason || result?.reason || 'Skipped file.';
          activityLog.push(`Skipped ${displayName}: ${skipReason}`);
        } else if (result?.duplicate) {
          duplicateCount += 1;
          const duplicateReason = result?.duplicateReason || 'Duplicate invoice detected.';
          activityLog.push(`Duplicate ${displayName}: ${duplicateReason}`);
        } else if (result?.stored) {
          createdCount += 1;
          let message = `Imported ${displayName}`;
          if (result?.checksum) {
            message += ` (${result.checksum.slice(0, 8)})`;
          }
          if (result?.movedTo) {
            message += ` → ${result.movedTo}`;
          }
          activityLog.push(message);
        }

        if (result?.moveError) {
          activityLog.push(`Move failed for ${displayName}: ${result.moveError}`);
        }
      } catch (error) {
        errors.push(error);
        const message = error?.message || 'Unknown error ingesting file.';
        activityLog.push(`Error processing ${item?.name || item?.id || 'file'}: ${message}`);
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

  if (processedItems === 0 && forceFull) {
    activityLog.push('Delta feed returned no files; performing snapshot crawl of the folder.');
    try {
      const snapshot = await ingestOneDriveFolderSnapshot(company, config, driveContext, { reason: syncReason });
      processedItems += snapshot.processed;
      createdCount += snapshot.created;
      skippedCount += snapshot.skipped;
      duplicateCount += snapshot.duplicates;
      snapshot.errors.forEach((error) => errors.push(error));
      activityLog.push(...snapshot.logEntries);
    } catch (error) {
      errors.push(error);
      activityLog.push(`Snapshot crawl failed: ${error.message || error}`);
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

  if (processedItems === 0) {
    activityLog.push(
      forceFull
        ? 'Full resync completed but OneDrive returned no files from the configured folder.'
        : 'No files reported by OneDrive during this sync cycle.'
    );
  }

  if (!activityLog.length) {
    activityLog.push('Sync completed without importing any new invoices.');
  }

  await updateQuickBooksCompanyOneDrive(company.realmId, {
    status,
    deltaLink: latestDelta,
    lastSyncAt: completedAt,
    lastSyncStatus,
    lastSyncReason: syncReason,
    lastSyncDurationMs: Date.now() - startedAt,
    lastSyncError,
    lastSyncMetrics: {
      processedItems,
      createdCount,
      skippedCount,
      duplicateCount,
      pages,
      errorCount: errors.length,
    },
    lastSyncLog: activityLog.slice(-20),
  });

  console.info(
    `OneDrive sync summary for realm ${company.realmId}: ${createdCount} created, ${processedItems} processed, ${skippedCount} skipped, ${duplicateCount} duplicates, ${errors.length} errors.`
  );
}

async function processOneDriveItem(company, config, driveContext, item, { reason = 'manual' } = {}) {
  const monitoredFolder = config.monitoredFolder || null;
  const remoteSource = {
    provider: 'onedrive',
    driveId: driveContext.driveId,
    itemId: item.id,
    parentId: item.parentReference?.id || monitoredFolder?.id || null,
    path: buildDriveItemPath(item),
    webUrl: item.webUrl || null,
    eTag: item.eTag || null,
    lastModifiedDateTime: item.lastModifiedDateTime || null,
    syncedAt: new Date().toISOString(),
    reason,
  };

  const download = await downloadOneDriveItem(driveContext.driveId, item.id);

  if (download.size > ONEDRIVE_MAX_FILE_SIZE_BYTES) {
    console.warn(
      `Skipping OneDrive item ${item.id} for realm ${company.realmId} because it exceeds the configured size limit (${download.size} bytes).`
    );
    return { skipped: true, reason: 'File exceeds configured size limit.' };
  }

  const ingestion = await ingestInvoiceFromSource({
    buffer: download.buffer,
    mimeType: download.mimeType,
    originalName: download.originalName,
    fileSize: download.size,
    realmId: company.realmId,
    businessType: company.businessType || null,
    remoteSource,
    defaultStatus: 'review',
  });

  let movedTo = null;
  let moveError = null;
  if (ingestion?.stored && !ingestion.duplicate) {
    try {
      const destination = await moveProcessedOneDriveItem(company, config, driveContext, item);
      movedTo = destination?.name || destination?.path || config.processedFolder?.name || null;
    } catch (error) {
      moveError = error.message || error;
      console.warn(
        `Unable to relocate OneDrive item ${item.id} after ingestion for realm ${company.realmId}`,
        moveError
      );
    }
  }

  return {
    ...ingestion,
    skipped: Boolean(ingestion?.skipped),
    duplicate: Boolean(ingestion?.duplicate),
    duplicateReason: ingestion?.duplicate?.reason || null,
    originalName: download.originalName || item.name || null,
    checksum:
      ingestion?.stored?.metadata?.checksum || ingestion?.duplicate?.match?.metadata?.checksum || null,
    movedTo,
    moveError,
    skipReason: ingestion?.skipped ? ingestion?.duplicate?.reason || 'Skipped by ingestion.' : null,
  };
}

async function moveProcessedOneDriveItem(company, config, driveContext, item) {
  const driveId = driveContext.driveId;
  if (!driveId || !item?.id) {
    return null;
  }

  const folder = await ensureOneDriveProcessedFolder(company, config, driveContext);
  if (!folder || !folder.id) {
    return null;
  }

  if (item.parentReference?.id === folder.id) {
    return folder;
  }

  await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}`, {
    method: 'PATCH',
    body: {
      parentReference: { id: folder.id },
    },
  });

  return folder;
}

async function ingestOneDriveFolderSnapshot(company, config, driveContext, { reason = 'manual' } = {}) {
  const monitoredFolder = config.monitoredFolder || null;
  const logEntries = [];
  const errors = [];
  let processed = 0;
  let created = 0;
  let skipped = 0;
  let duplicates = 0;
  let pages = 0;

  if (!monitoredFolder?.id) {
    return {
      processed: 0,
      created: 0,
      skipped: 0,
      duplicates: 0,
      errors: [new Error('Monitored folder is not configured.')],
      logEntries,
    };
  }

  let nextLink = `/drives/${encodeURIComponent(driveContext.driveId)}/items/${encodeURIComponent(monitoredFolder.id)}/children`;
  let initial = true;

  while (nextLink && pages < ONEDRIVE_MAX_DELTA_PAGES) {
    let response;
    try {
      if (initial) {
        response = await graphFetch(nextLink, {
          query: {
            $select: 'id,name,parentReference,webUrl,file,eTag,lastModifiedDateTime',
            $top: 200,
          },
          responseType: 'json',
        });
        initial = false;
      } else {
        response = await graphFetch(nextLink, { responseType: 'json' });
      }
    } catch (error) {
      errors.push(error);
      logEntries.push(`Snapshot fetch failed: ${error.message || error}`);
      break;
    }

    const entries = Array.isArray(response?.value) ? response.value : [];

    for (const item of entries) {
      if (!item || item.deleted || !item.file) {
        continue;
      }
      processed += 1;
      try {
        const result = await processOneDriveItem(
          company,
          config,
          driveContext,
          item,
          { reason: reason === 'full-resync' ? 'full-resync-snapshot' : reason }
        );
        const displayName = result?.originalName || item.name || item.id;

        if (result?.skipped) {
          skipped += 1;
          const skipReason = result?.skipReason || result?.reason || 'Skipped file.';
          logEntries.push(`Skipped ${displayName}: ${skipReason}`);
        } else if (result?.duplicate) {
          duplicates += 1;
          const duplicateReason = result?.duplicateReason || 'Duplicate invoice detected.';
          logEntries.push(`Duplicate ${displayName}: ${duplicateReason}`);
        } else if (result?.stored) {
          created += 1;
          let message = `Imported ${displayName}`;
          if (result?.checksum) {
            message += ` (${result.checksum.slice(0, 8)})`;
          }
          if (result?.movedTo) {
            message += ` → ${result.movedTo}`;
          }
          logEntries.push(message);
        }

        if (result?.moveError) {
          logEntries.push(`Move failed for ${displayName}: ${result.moveError}`);
        }
      } catch (error) {
        errors.push(error);
        const message = error?.message || 'Unknown error ingesting file.';
        logEntries.push(`Error processing ${item?.name || item?.id || 'file'}: ${message}`);
      }
    }

    pages += 1;

    if (response['@odata.nextLink']) {
      nextLink = response['@odata.nextLink'];
    } else {
      nextLink = null;
    }
  }

  if (processed === 0) {
    logEntries.push('Snapshot crawl found no files in the folder.');
  }

  return {
    processed,
    created,
    skipped,
    duplicates,
    errors,
    logEntries,
  };
}

function buildProcessedFolderDetails(metadata, fallbackName = ONEDRIVE_PROCESSED_FOLDER_NAME) {
  if (!metadata || typeof metadata !== 'object') {
    return null;
  }

  return normalizeOneDriveFolderConfig({
    id: metadata.id,
    name: metadata.name || fallbackName,
    path: buildDriveItemPath(metadata),
    webUrl: metadata.webUrl || null,
    parentId: metadata.parentReference?.id || null,
  });
}

async function ensureOneDriveProcessedFolder(company, config, driveContext) {
  const driveId = driveContext.driveId;
  const monitoredFolder = config.monitoredFolder || null;
  const sourceFolderId = monitoredFolder?.id;
  if (!driveId || !sourceFolderId) {
    return null;
  }

  const cacheKey = `${company.realmId || 'unknown'}:${driveId}:${sourceFolderId}`;
  const cached = oneDriveProcessedFolderCache.get(cacheKey);
  if (cached && cached.promise) {
    return cached.promise;
  }
  if (cached && !cached.promise) {
    return cached;
  }

  const resolver = (async () => {
    const preference = normalizeOneDriveFolderConfig(config.processedFolder);
    if (preference?.id) {
      return preference;
    }

    let metadata = await locateExistingProcessedFolder(driveId, sourceFolderId).catch(() => null);
    if (!metadata) {
      metadata = await createProcessedFolder(driveId, sourceFolderId);
    }

    if (!metadata) {
      return null;
    }

    const folderDetails = buildProcessedFolderDetails(metadata);
    if (folderDetails) {
      config.processedFolder = folderDetails;
      await updateQuickBooksCompanyOneDrive(company.realmId, {
        processedFolder: folderDetails,
      }).catch((error) => {
        console.warn(
          `Unable to persist processed folder details for realm ${company.realmId}`,
          error.message || error
        );
      });
    }

    return folderDetails;
  })();

  oneDriveProcessedFolderCache.set(cacheKey, { promise: resolver });
  try {
    const result = await resolver;
    oneDriveProcessedFolderCache.set(cacheKey, result);
    return result;
  } finally {
    const entry = oneDriveProcessedFolderCache.get(cacheKey);
    if (entry && entry.promise) {
      oneDriveProcessedFolderCache.delete(cacheKey);
    }
  }
}

async function fetchOneDriveFolderMetadata(driveId, folderId) {
  if (!driveId || !folderId) {
    return null;
  }
  const item = await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}`, {
    query: {
      $select: 'id,name,folder,parentReference',
    },
  });
  if (!item?.folder) {
    return null;
  }
  return item;
}

async function locateExistingProcessedFolder(driveId, folderId) {
  const response = await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}/children`, {
    query: {
      $select: 'id,name,folder,parentReference',
      $top: 200,
    },
  });

  const entries = Array.isArray(response?.value) ? response.value : [];
  const match = entries.find((entry) => entry?.folder && entry.name && entry.name.toLowerCase() === ONEDRIVE_PROCESSED_FOLDER_NAME.toLowerCase());
  if (match) {
    return match;
  }

  return null;
}

async function createProcessedFolder(driveId, parentFolderId) {
  const created = await graphFetch(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentFolderId)}/children`, {
    method: 'POST',
    body: {
      name: ONEDRIVE_PROCESSED_FOLDER_NAME,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename',
    },
  });

  if (!created?.folder) {
    return null;
  }

  return created;
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

function buildDriveItemDisplayPath(rawPath) {
  if (!rawPath) {
    return null;
  }

  let remainder = rawPath;

  const drivePrefixMatch = remainder.match(/^\/?drives\/[^/]+\/root:(.*)$/);
  if (drivePrefixMatch) {
    remainder = drivePrefixMatch[1] || '';
  } else {
    const rootPrefixMatch = remainder.match(/^\/?drive\/root:(.*)$/);
    if (rootPrefixMatch) {
      remainder = rootPrefixMatch[1] || '';
    }
  }

  if (remainder.startsWith('/')) {
    remainder = remainder.slice(1);
  }

  if (!remainder) {
    return null;
  }

  const decoded = remainder
    .split('/')
    .map((part) => {
      if (!part) {
        return part;
      }
      try {
        return decodeURIComponent(part);
      } catch (error) {
        return part;
      }
    })
    .filter(Boolean)
    .join('/');

  return decoded || null;
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


function isGmailMonitoringConfigured() {
  return GMAIL_CLIENT_CONFIGURED;
}

function getGmailStatePath(realmId) {
  return path.join(__dirname, '..', 'data', 'gmail', `${realmId}.json`);
}

async function readCompanyGmailState(realmId) {
  if (!realmId) {
    return ensureGmailRuntimeStateDefaults(null);
  }

  if (gmailStateCache.has(realmId)) {
    return gmailStateCache.get(realmId);
  }

  try {
    const contents = await fs.readFile(getGmailStatePath(realmId), 'utf-8');
    const parsed = JSON.parse(contents);
    const state = ensureGmailRuntimeStateDefaults(parsed);
    gmailStateCache.set(realmId, state);
    return state;
  } catch (error) {
    if (error.code === 'ENOENT') {
      const state = ensureGmailRuntimeStateDefaults(null);
      gmailStateCache.set(realmId, state);
      return state;
    }
    throw error;
  }
}

function ensureGmailRuntimeStateDefaults(state) {
  const base = state && typeof state === 'object' ? state : {};
  return {
    lastPollAt: typeof base.lastPollAt === 'string' ? base.lastPollAt : null,
    lastHistoryId: typeof base.lastHistoryId === 'string' ? base.lastHistoryId : null,
    processedMessages:
      base.processedMessages && typeof base.processedMessages === 'object' && !Array.isArray(base.processedMessages)
        ? { ...base.processedMessages }
        : {},
  };
}

async function persistCompanyGmailState(realmId, nextState) {
  const normalised = ensureGmailRuntimeStateDefaults(nextState);
  normalised.processedMessages = pruneGmailProcessedMessages(normalised.processedMessages, 500);
  gmailStateCache.set(realmId, normalised);

  const statePath = getGmailStatePath(realmId);
  await fs.mkdir(path.dirname(statePath), { recursive: true });
  await fs.writeFile(statePath, JSON.stringify(normalised, null, 2));
  return normalised;
}

async function deleteCompanyGmailState(realmId) {
  gmailStateCache.delete(realmId);
  if (!realmId) {
    return;
  }

  const statePath = getGmailStatePath(realmId);
  try {
    await fs.unlink(statePath);
  } catch (error) {
    if (error.code !== 'ENOENT') {
      throw error;
    }
  }
}

function getGmailPollInterval(config) {
  const interval = Number.parseInt(config?.pollIntervalMs, 10);
  if (!Number.isFinite(interval) || interval <= 0) {
    return GMAIL_DEFAULT_POLL_INTERVAL_MS;
  }
  return Math.max(interval, GMAIL_POLL_MIN_INTERVAL_MS);
}

function getGmailAllowedMimeTypes(config) {
  const values = Array.isArray(config?.allowedMimeTypes)
    ? config.allowedMimeTypes
    : parseDelimitedList(config?.allowedMimeTypes);
  const base = values.length ? values : GMAIL_DEFAULT_ALLOWED_MIME_TYPES;
  const set = new Set();
  for (const value of base) {
    if (typeof value === 'string' && value.trim()) {
      set.add(value.trim().toLowerCase());
    }
  }
  return set;
}

async function pollAllGmailCompanies() {
  if (!isGmailMonitoringConfigured()) {
    return;
  }

  let companies;
  try {
    companies = await readQuickBooksCompanies();
  } catch (error) {
    if (error?.code === 'QUICKBOOKS_COMPANY_FILE_CORRUPT') {
      console.error(
        'Gmail polling halted: QuickBooks company store is corrupt. Restore the latest backup before resuming inbox checks.'
      );
      emitHealthMetric('quickbooks.company_file.corrupt', { source: 'gmail_poll' });
      return;
    }
    throw error;
  }
  const now = Date.now();

  for (const company of companies) {
    const realmId = company?.realmId;
    if (!realmId) {
      continue;
    }

    const gmailConfig = ensureGmailConfigDefaults(company.gmail);
    if (!gmailConfig || gmailConfig.enabled === false || !gmailConfig.refreshToken) {
      continue;
    }

    let state;
    try {
      state = await readCompanyGmailState(realmId);
    } catch (error) {
      console.warn(`Unable to read Gmail state for realm ${realmId}`, error.message || error);
      state = ensureGmailRuntimeStateDefaults(null);
    }

    const lastPollMs = state.lastPollAt ? new Date(state.lastPollAt).getTime() : 0;
    const interval = getGmailPollInterval(gmailConfig);
    if (lastPollMs && now - lastPollMs < interval) {
      continue;
    }

    queueGmailPoll(realmId, { reason: 'interval' }).catch((error) => {
      console.error(`Scheduled Gmail poll failed for realm ${realmId}`, error);
    });
  }
}

async function queueGmailPoll(realmId, { reason = 'manual' } = {}) {
  if (!isGmailMonitoringConfigured() || !realmId) {
    return null;
  }

  if (activeGmailPolls.has(realmId)) {
    return activeGmailPolls.get(realmId);
  }

  const task = (async () => {
    try {
      const company = await getQuickBooksCompanyRecord(realmId);
      if (!company) {
        const error = new Error('QuickBooks company not found.');
        error.code = 'NOT_FOUND';
        throw error;
      }

      const gmailConfig = ensureGmailConfigDefaults(company.gmail);
      if (!gmailConfig?.enabled) {
        return null;
      }

      if (!gmailConfig.refreshToken) {
        const error = new Error('Gmail refresh token is not configured. Connect the mailbox again.');
        error.code = 'GMAIL_REFRESH_TOKEN_MISSING';
        throw error;
      }

      const result = await pollGmailForCompany(company, gmailConfig, { reason });

      const statusUpdate = {
        status: result.warning ? 'warning' : 'connected',
        lastSyncStatus: result.warning ? 'partial' : 'success',
        lastSyncReason: reason,
        lastSyncAt: result.lastSyncAt,
        lastSyncMetrics: result.metrics,
        lastSyncError: result.warning ? result.lastSyncError : null,
        historyId: result.historyId || null,
      };

      await updateQuickBooksCompanyGmail(realmId, statusUpdate);
      return result;
    } catch (error) {
      console.error(`Gmail sync failed for realm ${realmId}`, error);
      try {
        await updateQuickBooksCompanyGmail(realmId, {
          status: 'error',
          lastSyncStatus: 'error',
          lastSyncReason: reason,
          lastSyncAt: new Date().toISOString(),
          lastSyncError: { message: error.message || 'Gmail poll failed.' },
        });
      } catch (updateError) {
        console.warn(`Unable to persist Gmail error state for ${realmId}`, updateError.message || updateError);
      }
      gmailTokenCache.delete(realmId);
      throw error;
    } finally {
      activeGmailPolls.delete(realmId);
    }
  })();

  activeGmailPolls.set(realmId, task);
  return task;
}

async function pollGmailForCompany(company, gmailConfig, { reason = 'manual' } = {}) {
  const realmId = company.realmId;
  const startedAt = Date.now();
  const state = await readCompanyGmailState(realmId);
  const accessToken = await acquireGmailAccessToken(realmId, gmailConfig);

  const maxResults = Number.isFinite(gmailConfig.maxResults) && gmailConfig.maxResults > 0
    ? Math.min(Math.floor(gmailConfig.maxResults), 100)
    : GMAIL_DEFAULT_MAX_RESULTS;

  const query = {
    maxResults,
    q: gmailConfig.searchQuery || undefined,
  };

  if (Array.isArray(gmailConfig.labelIds) && gmailConfig.labelIds.length) {
    query.labelIds = gmailConfig.labelIds;
  }

  const listResponse = await gmailApiRequest('/users/me/messages', { query }, accessToken);
  const messages = Array.isArray(listResponse?.messages) ? listResponse.messages : [];

  if (typeof listResponse?.historyId === 'string') {
    state.lastHistoryId = listResponse.historyId;
  }

  let processedMessages = 0;
  let processedAttachments = 0;
  let skippedAttachments = 0;
  const attachmentErrors = [];

  const allowedMimeTypes = getGmailAllowedMimeTypes(gmailConfig);
  const maxAttachmentBytes = Math.max(
    Number.parseInt(gmailConfig.maxAttachmentBytes, 10) || GMAIL_DEFAULT_MAX_ATTACHMENT_BYTES,
    1024
  );

  for (const messageSummary of messages) {
    const messageId = messageSummary?.id;
    if (!messageId) {
      continue;
    }

    if (state.processedMessages[messageId]?.processedAt) {
      continue;
    }

    let message;
    try {
      message = await gmailApiRequest(`/users/me/messages/${encodeURIComponent(messageId)}`, {
        query: { format: 'full' },
      }, accessToken);
    } catch (error) {
      console.warn(`Failed to load Gmail message ${messageId} for realm ${realmId}`, error.message || error);
      attachmentErrors.push({ messageId, error: error.message || 'Unable to load message.' });
      state.processedMessages[messageId] = {
        processedAt: new Date().toISOString(),
        attachments: [],
        status: 'error',
        reason,
        snippet: messageSummary?.snippet || null,
        historyId: messageSummary?.historyId || null,
      };
      continue;
    }

    const attachments = collectGmailAttachments(message?.payload);
    if (!attachments.length) {
      state.processedMessages[messageId] = {
        processedAt: new Date().toISOString(),
        attachments: [],
        status: 'no-attachments',
        reason,
        snippet: message?.snippet || null,
        historyId: message?.historyId || null,
      };
      continue;
    }

    const processedAttachmentKeys = [];
    let messageHandled = false;

    for (const attachment of attachments) {
      const filename = attachment.filename || `${messageId}.bin`;
      const extension = path.extname(filename);
      const declaredMime = attachment.mimeType || deriveMimeTypeFromExtension(extension);
      const mimeType = (declaredMime || 'application/octet-stream').toLowerCase();

      if (allowedMimeTypes.size && !allowedMimeTypes.has(mimeType)) {
        skippedAttachments += 1;
        continue;
      }

      const declaredSize = typeof attachment.size === 'number' && Number.isFinite(attachment.size)
        ? attachment.size
        : null;

      if (declaredSize && declaredSize > maxAttachmentBytes) {
        console.warn(
          `Skipping Gmail attachment ${attachment.attachmentId || attachment.partId || 'unknown'} on ${messageId} for realm ${realmId} because it exceeds the size limit (${declaredSize} bytes).`
        );
        skippedAttachments += 1;
        continue;
      }

      let buffer;
      try {
        buffer = await fetchGmailAttachment(messageId, attachment, accessToken);
      } catch (error) {
        console.warn(
          `Failed to download Gmail attachment ${attachment.attachmentId || attachment.partId || 'unknown'} from ${messageId} (realm ${realmId})`,
          error.message || error
        );
        attachmentErrors.push({
          messageId,
          attachmentId: attachment.attachmentId || attachment.partId || null,
          error: error.message || 'Unknown error downloading attachment.',
        });
        continue;
      }

      if (buffer.length > maxAttachmentBytes) {
        console.warn(
          `Skipping Gmail attachment ${attachment.attachmentId || attachment.partId || 'unknown'} for realm ${realmId} because the downloaded size exceeds the limit (${buffer.length} bytes).`
        );
        skippedAttachments += 1;
        continue;
      }

      const businessType = gmailConfig.businessType || company.businessType || GMAIL_DEFAULT_BUSINESS_TYPE || null;

      const remoteSource = buildGmailRemoteSource({
        message,
        attachment,
        buffer,
        mimeType,
        reason,
        email: gmailConfig.email || null,
      });

      try {
        await ingestInvoiceFromSource({
          buffer,
          mimeType,
          originalName: filename,
          fileSize: buffer.length,
          realmId: company.realmId,
          businessType,
          remoteSource,
          defaultStatus: 'review',
        });
        processedAttachments += 1;
        messageHandled = true;
        processedAttachmentKeys.push(gmailAttachmentKey(messageId, attachment));
      } catch (error) {
        console.error(
          `Failed to ingest Gmail attachment ${attachment.attachmentId || attachment.partId || 'unknown'} from ${messageId} (realm ${realmId})`,
          error.message || error
        );
        attachmentErrors.push({
          messageId,
          attachmentId: attachment.attachmentId || attachment.partId || null,
          error: error.message || 'Unknown error ingesting attachment.',
        });
      }
    }

    if (messageHandled) {
      processedMessages += 1;
    }

    state.processedMessages[messageId] = {
      processedAt: new Date().toISOString(),
      attachments: processedAttachmentKeys,
      status: messageHandled ? 'processed' : 'skipped',
      reason,
      snippet: message?.snippet || null,
      historyId: message?.historyId || null,
    };
  }

  state.lastPollAt = new Date().toISOString();
  await persistCompanyGmailState(realmId, state);

  const durationMs = Date.now() - startedAt;
  const metrics = {
    processedMessages,
    processedAttachments,
    skippedAttachments,
    errorCount: attachmentErrors.length,
    durationMs,
  };

  const warning = attachmentErrors.length > 0;
  const lastSyncError = warning
    ? {
        message:
          attachmentErrors[0]?.error ||
          `${attachmentErrors.length} Gmail attachment error${attachmentErrors.length === 1 ? '' : 's'} encountered.`,
        at: state.lastPollAt,
      }
    : null;

  return {
    metrics,
    warning,
    lastSyncError,
    lastSyncAt: state.lastPollAt,
    historyId: state.lastHistoryId || null,
  };
}

async function acquireGmailAccessToken(realmId, gmailConfig) {
  if (!isGmailMonitoringConfigured()) {
    throw new Error('Gmail OAuth client is not configured.');
  }

  const now = Date.now();
  const cached = gmailTokenCache.get(realmId);
  if (cached && cached.expiresAt - 60000 > now) {
    return cached.token;
  }

  const refreshToken = gmailConfig?.refreshToken;
  if (!refreshToken) {
    const error = new Error('Gmail refresh token is not configured. Connect the mailbox again.');
    error.code = 'GMAIL_REFRESH_TOKEN_MISSING';
    throw error;
  }

  const params = new URLSearchParams();
  params.set('client_id', GMAIL_CLIENT_ID);
  params.set('client_secret', GMAIL_CLIENT_SECRET);
  params.set('grant_type', 'refresh_token');
  params.set('refresh_token', refreshToken);

  const response = await fetch(GMAIL_TOKEN_URL, {
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
      `Gmail token request failed with status ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    error.body = errorBody;
    throw error;
  }

  const data = await response.json();
  const expiresIn = typeof data.expires_in === 'number' ? data.expires_in : Number.parseInt(data.expires_in, 10) || 3600;
  const expiresAt = now + Math.max(expiresIn - 60, 60) * 1000;

  gmailTokenCache.set(realmId, {
    token: data.access_token,
    expiresAt,
  });

  if (typeof data.refresh_token === 'string' && data.refresh_token && data.refresh_token !== refreshToken) {
    try {
      await updateQuickBooksCompanyGmail(realmId, {
        refreshToken: data.refresh_token,
        lastConnectedAt: new Date().toISOString(),
      });
    } catch (error) {
      console.warn(`Failed to persist updated Gmail refresh token for realm ${realmId}`, error.message || error);
    }
  }

  return data.access_token;
}

async function gmailApiRequest(endpoint, { method = 'GET', query, headers, body } = {}, accessToken) {
  const pathWithSlash = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
  const url = new URL(`${GMAIL_API_BASE_URL}${pathWithSlash}`);

  if (query && typeof query === 'object') {
    for (const [key, value] of Object.entries(query)) {
      if (value === undefined || value === null) {
        continue;
      }

      if (Array.isArray(value)) {
        for (const item of value) {
          url.searchParams.append(key, item);
        }
      } else {
        url.searchParams.set(key, value);
      }
    }
  }

  const init = {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      ...(headers || {}),
    },
  };

  if (body !== undefined && body !== null) {
    init.body = typeof body === 'string' ? body : JSON.stringify(body);
    if (!init.headers['Content-Type']) {
      init.headers['Content-Type'] = 'application/json';
    }
  }

  const response = await fetch(url.toString(), init);
  if (!response.ok) {
    const errorBody = await safeReadJson(response);
    const message =
      errorBody?.error?.message ||
      errorBody?.error_description ||
      `Gmail API request failed (${method || 'GET'} ${endpoint}) with status ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    error.body = errorBody;
    throw error;
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

function collectGmailAttachments(payload, parentMimeType = null, pathPrefix = []) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  const results = [];
  const { filename, mimeType, body, parts, partId, headers } = payload;
  const resolvedMime = mimeType || parentMimeType || null;

  const hasAttachmentData = Boolean(body?.attachmentId || body?.data);
  const hasFilename = typeof filename === 'string' && filename.trim().length > 0;
  const disposition = extractEmailHeader(headers, 'Content-Disposition') || '';
  const isInline = disposition.toLowerCase().includes('inline');
  const size = typeof body?.size === 'number' ? body.size : null;

  if (hasFilename || hasAttachmentData) {
    results.push({
      filename: hasFilename ? filename.trim() : null,
      mimeType: resolvedMime,
      attachmentId: body?.attachmentId || null,
      data: body?.data || null,
      size,
      partId: partId || pathPrefix.join('.'),
      isInline,
    });
  }

  if (Array.isArray(parts) && parts.length) {
    const nextPrefix = partId ? [...pathPrefix, partId] : pathPrefix;
    for (const child of parts) {
      results.push(...collectGmailAttachments(child, resolvedMime, nextPrefix));
    }
  }

  return results;
}

async function fetchGmailAttachment(messageId, attachment, accessToken) {
  if (attachment?.data) {
    return decodeBase64Url(attachment.data);
  }

  if (!attachment?.attachmentId) {
    throw new Error('Gmail attachment is missing attachmentId and inline data.');
  }

  const attachmentPath = `/users/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.attachmentId)}`;
  const response = await gmailApiRequest(attachmentPath, {}, accessToken);
  if (!response?.data) {
    throw new Error('Gmail attachment download did not include data.');
  }

  return decodeBase64Url(response.data);
}

function buildGmailRemoteSource({ message, attachment, buffer, mimeType, reason, email }) {
  if (!message) {
    return null;
  }

  const messageId = message.id || null;
  const threadId = message.threadId || null;
  const historyId = message.historyId || null;
  const subject = extractEmailHeader(message.payload?.headers, 'Subject');
  const from = extractEmailHeader(message.payload?.headers, 'From');
  const internalDate = message.internalDate ? new Date(Number(message.internalDate)).toISOString() : null;
  const attachmentKey = gmailAttachmentKey(messageId, attachment);

  return {
    provider: 'gmail',
    email: email || null,
    itemId: attachmentKey,
    messageId,
    threadId,
    attachmentId: attachment?.attachmentId || null,
    attachmentSize: buffer?.length || attachment?.size || null,
    mimeType,
    filename: attachment?.filename || null,
    labelIds: Array.isArray(message.labelIds) ? message.labelIds : null,
    subject: subject || null,
    from: from || null,
    historyId: historyId || null,
    snippet: message?.snippet || null,
    receivedAt: internalDate,
    webUrl: messageId ? `https://mail.google.com/mail/u/0/#inbox/${messageId}` : null,
    syncedAt: new Date().toISOString(),
    reason,
  };
}

function extractEmailHeader(headers, name) {
  if (!Array.isArray(headers)) {
    return null;
  }
  const target = name.toLowerCase();
  const entry = headers.find((header) => typeof header?.name === 'string' && header.name.toLowerCase() === target);
  if (!entry || typeof entry.value !== 'string') {
    return null;
  }
  return entry.value.trim() || null;
}

function gmailAttachmentKey(messageId, attachment) {
  const attachmentId = attachment?.attachmentId || attachment?.partId || 'attachment';
  return [messageId || 'message', attachmentId].join('::');
}

function pruneGmailProcessedMessages(processedMap, limit = 500) {
  if (!processedMap || typeof processedMap !== 'object') {
    return {};
  }

  const entries = Object.entries(processedMap);
  if (entries.length <= limit) {
    return processedMap;
  }

  entries.sort(([, a], [, b]) => {
    const timeA = new Date(a?.processedAt || 0).getTime();
    const timeB = new Date(b?.processedAt || 0).getTime();
    return timeB - timeA;
  });

  const trimmed = entries.slice(0, limit);
  return Object.fromEntries(trimmed);
}

function startGmailMonitor() {
  if (!isGmailMonitoringConfigured()) {
    return;
  }

  if (gmailPollingTimer) {
    clearInterval(gmailPollingTimer);
  }

  const tick = () => {
    pollAllGmailCompanies().catch((error) => {
      console.error('Scheduled Gmail polling failed', error);
    });
  };

  tick();
  gmailPollingTimer = setInterval(tick, GMAIL_DEFAULT_POLL_INTERVAL_MS);
  console.log(
    `Gmail inbox monitoring enabled (base interval ${Math.max(Math.round(GMAIL_DEFAULT_POLL_INTERVAL_MS / 1000), 1)}s)`
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

app.get('/api/onedrive/drives', async (req, res) => {
  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const driveId = typeof req.query?.driveId === 'string' ? req.query.driveId : null;

  try {
    const { drives, warning } = await graphListDrives({ driveId });
    const payload = { drives };
    if (warning) {
      payload.warning = warning;
    }
    res.json(payload);
  } catch (error) {
    console.error('Failed to list OneDrive drives', error);
    const status = Number.isInteger(error?.status) ? error.status : 500;
    const message = error?.status === 403 ? 'Access to OneDrive drives was denied.' : 'Failed to list OneDrive drives.';
    res.status(status).json({ error: message });
  }
});

app.get('/api/onedrive/children', async (req, res) => {
  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const driveId = typeof req.query?.driveId === 'string' ? req.query.driveId.trim() : '';
  const itemId = typeof req.query?.itemId === 'string' ? req.query.itemId.trim() : '';
  const folderPath = typeof req.query?.path === 'string' ? req.query.path.trim() : typeof req.query?.folderPath === 'string' ? req.query.folderPath.trim() : '';

  if (!driveId) {
    return res.status(400).json({ error: 'Provide a driveId query parameter to list folders.' });
  }

  try {
    const items = await graphListDriveChildren({ driveId, itemId, path: folderPath });
    res.json({ items });
  } catch (error) {
    if (error?.status === 404) {
      return res.status(404).json({ error: 'OneDrive folder not found.' });
    }
    if (error?.status === 403) {
      return res.status(403).json({ error: 'Access to the specified OneDrive folder was denied.' });
    }
    console.error('Failed to list OneDrive folder children', error);
    res.status(500).json({ error: 'Failed to list OneDrive folder contents.' });
  }
});

app.post('/api/onedrive/resolve', async (req, res) => {
  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const shareUrl = typeof body.shareUrl === 'string' ? body.shareUrl : null;
  const driveId = typeof body.driveId === 'string' ? body.driveId : null;
  const folderId = typeof body.folderId === 'string' ? body.folderId : null;
  const folderPath = typeof body.folderPath === 'string' ? body.folderPath : null;

  if (!shareUrl && !(driveId && (folderId || folderPath))) {
    return res
      .status(400)
      .json({ error: 'Provide shareUrl or driveId with folderId or folderPath to resolve a OneDrive folder.' });
  }

  try {
    const item = await resolveOneDriveFolderReference({ shareUrl, driveId, folderId, folderPath });
    const response = item ? { ...item } : null;
    res.json({ item: response });
  } catch (error) {
    if (error?.status === 404) {
      return res.status(404).json({ error: 'Unable to locate the specified OneDrive folder.' });
    }
    if (error?.status === 403) {
      return res.status(403).json({ error: 'Access to the specified OneDrive folder was denied.' });
    }
    console.error('Failed to resolve OneDrive folder reference', error);
    res.status(500).json({ error: 'Failed to resolve OneDrive folder reference.' });
  }
});

app.get('/api/onedrive/settings', async (req, res) => {
  try {
    const config = await loadGlobalOneDriveConfig();
    res.json({ settings: sanitizeGlobalOneDriveConfig(config) });
  } catch (error) {
    console.error('Failed to load OneDrive shared settings', error);
    res.status(500).json({ error: 'Failed to load OneDrive shared settings.' });
  }
});

app.put('/api/onedrive/settings', async (req, res) => {
  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const shareUrl = sanitiseOptionalString(body.shareUrl);
  const driveIdInput = sanitiseOptionalString(body.driveId);
  const folderId = sanitiseOptionalString(body.folderId);
  const folderPath = sanitiseOptionalString(body.folderPath);
  const statusInput = sanitiseOptionalString(body.status);
  const lastSyncHealthInput = body.lastSyncHealth && typeof body.lastSyncHealth === 'object' ? body.lastSyncHealth : undefined;

  if (!shareUrl && !(driveIdInput && (folderId || folderPath)) && !statusInput && lastSyncHealthInput === undefined) {
    return res.status(400).json({
      error: 'Provide a sharing link or drive and folder reference to update the shared OneDrive settings.',
    });
  }

  try {
    let update = {};

    if (shareUrl || driveIdInput || folderId || folderPath) {
      const resolution = await resolveOneDriveFolderReference(
        {
          shareUrl,
          driveId: driveIdInput,
          folderId,
          folderPath,
        }
      );

      const driveId = resolution.driveId;
      let driveMetadata = null;
      try {
        const { drives: driveList } = await graphListDrives({ driveId });
        driveMetadata = Array.isArray(driveList) ? driveList.find((entry) => entry?.id === driveId) || null : null;
      } catch (metadataError) {
        console.warn('Unable to resolve drive metadata while updating shared OneDrive settings', metadataError.message || metadataError);
      }

      const driveOwnerMetadata = driveMetadata?.owner || null;
      const driveOwnerName = (() => {
        if (!driveOwnerMetadata) {
          return null;
        }
        if (typeof driveOwnerMetadata === 'string') {
          return driveOwnerMetadata;
        }
        if (typeof driveOwnerMetadata === 'object') {
          const nestedUser = typeof driveOwnerMetadata.user === 'object' ? driveOwnerMetadata.user : null;
          return (
            sanitiseOptionalString(driveOwnerMetadata.displayName) ||
            sanitiseOptionalString(driveOwnerMetadata.name) ||
            sanitiseOptionalString(driveOwnerMetadata.email) ||
            sanitiseOptionalString(driveOwnerMetadata.principal) ||
            sanitiseOptionalString(nestedUser?.displayName) ||
            sanitiseOptionalString(nestedUser?.email) ||
            null
          );
        }
        return null;
      })();

      update = {
        driveId,
        shareUrl: resolution.shareUrl || shareUrl || null,
        driveName: driveMetadata?.name || null,
        driveType: driveMetadata?.driveType || null,
        driveWebUrl: driveMetadata?.webUrl || null,
        driveOwner: driveOwnerName || null,
        folderId: resolution.id || null,
        folderPath: resolution.path || null,
        folderName: resolution.name || null,
        folderWebUrl: resolution.webUrl || null,
        folderParentId: resolution.parentId || null,
        status: 'ready',
        lastValidatedAt: new Date().toISOString(),
        lastValidationError: null,
      };
    }

    if (statusInput !== undefined) {
      update.status = statusInput || null;
    }

    if (lastSyncHealthInput !== undefined) {
      update.lastSyncHealth = lastSyncHealthInput;
    }

    const next = await updateGlobalOneDriveConfig(update);
    res.json({ settings: sanitizeGlobalOneDriveConfig(next) });
  } catch (error) {
    if (error?.status === 404) {
      return res.status(404).json({ error: 'Unable to locate the specified OneDrive folder.' });
    }
    if (error?.status === 403) {
      return res.status(403).json({ error: 'Access to the specified OneDrive folder was denied.' });
    }
    console.error('Failed to update OneDrive shared settings', error);
    res.status(500).json({ error: 'Failed to update OneDrive shared settings.' });
  }
});

app.get('/api/quickbooks/companies', async (req, res) => {
  let companies = [];
  try {
    companies = await readQuickBooksCompanies();
  } catch (error) {
    if (error?.code === 'QUICKBOOKS_COMPANY_FILE_CORRUPT') {
      const message =
        'QuickBooks connections are paused because the company store is corrupt. Restore the latest backup or run the repair CLI before resuming sync.';
      console.error(message, error);
      console.debug('[QuickBooks] Company store corruption details', {
        backupPath: error?.backupPath || null,
        code: error?.code || null,
      });
      res.status(503).json({
        error: message,
        code: error.code,
        backupPath: error.backupPath || null,
      });
      emitHealthMetric('quickbooks.company_file.corrupt', { source: 'http_companies' });
      return;
    }

    console.error('Failed to load QuickBooks companies', error);
    console.debug('[QuickBooks] Unexpected error while loading QuickBooks companies', {
      code: error?.code || null,
      message: error?.message || null,
    });
    res.status(500).json({
      error: 'Failed to load QuickBooks connections. See server logs for details.',
      message: error?.message || null,
    });
    return;
  }

  const sanitized = companies.map(sanitizeQuickBooksCompany);
  console.info('[QuickBooks] Responding with companies payload', {
    count: sanitized.length,
  });
  res.json({ companies: sanitized });
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

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const enableSync = body.enabled !== false;

  const normalizeIncomingFolder = (input) => {
    if (input === undefined) {
      return undefined;
    }
    if (input === null) {
      return null;
    }
    if (input && typeof input === 'object') {
      const candidate = {
        id: sanitiseOptionalString(input.id),
        path: sanitiseOptionalString(input.path),
        name: sanitiseOptionalString(input.name),
        webUrl: sanitiseOptionalString(input.webUrl),
        parentId: sanitiseOptionalString(input.parentId),
      };
      if (!candidate.id && !candidate.path) {
        return null;
      }
      return candidate;
    }
    return null;
  };

  const monitoredFolderRequest = (() => {
    const direct = normalizeIncomingFolder(body.monitoredFolder);
    if (direct !== undefined) {
      return direct;
    }
    const legacyFields = [body.folderId, body.folderPath, body.folderName, body.webUrl, body.parentId];
    if (legacyFields.some((value) => value !== undefined && value !== null && String(value).trim() !== '')) {
      return normalizeIncomingFolder({
        id: body.folderId,
        path: body.folderPath,
        name: body.folderName,
        webUrl: body.webUrl,
        parentId: body.parentId,
      });
    }
    return undefined;
  })();

  const processedFolderRequest = (() => {
    if (Object.prototype.hasOwnProperty.call(body, 'processedFolder')) {
      return normalizeIncomingFolder(body.processedFolder);
    }
    const legacyFields = [
      body.processedFolderId,
      body.processedFolderPath,
      body.processedFolderName,
      body.processedFolderWebUrl,
      body.processedFolderParentId,
    ];
    if (legacyFields.some((value) => value !== undefined && value !== null && String(value).trim() !== '')) {
      return normalizeIncomingFolder({
        id: body.processedFolderId,
        path: body.processedFolderPath,
        name: body.processedFolderName,
        webUrl: body.processedFolderWebUrl,
        parentId: body.processedFolderParentId,
      });
    }
    return undefined;
  })();

  try {
    const company = await getQuickBooksCompanyRecord(realmId);
    if (!company) {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }

    const driveContext = await ensureGlobalOneDriveDriveContext();

    if (enableSync) {
      if (!monitoredFolderRequest || (!monitoredFolderRequest.id && !monitoredFolderRequest.path)) {
        return res.status(400).json({ error: 'Select a OneDrive folder to monitor before enabling sync.' });
      }

      const monitoredResolution = await resolveOneDriveFolderReference(
        {
          shareUrl: null,
          driveId: driveContext.driveId,
          folderId: monitoredFolderRequest.id,
          folderPath: monitoredFolderRequest.path,
        },
        { expectedDriveId: driveContext.driveId }
      );

      const monitoredFolder = {
        id: monitoredResolution.id,
        name: monitoredResolution.name || monitoredFolderRequest.name || null,
        path: monitoredResolution.path || monitoredFolderRequest.path || null,
        webUrl: monitoredResolution.webUrl || monitoredFolderRequest.webUrl || null,
        parentId: monitoredResolution.parentId || monitoredFolderRequest.parentId || null,
      };

      let processedFolderUpdate = undefined;
      if (processedFolderRequest === null) {
        processedFolderUpdate = null;
      } else if (processedFolderRequest && (processedFolderRequest.id || processedFolderRequest.path)) {
        const processedResolution = await resolveOneDriveFolderReference(
          {
            shareUrl: null,
            driveId: driveContext.driveId,
            folderId: processedFolderRequest.id,
            folderPath: processedFolderRequest.path,
          },
          { expectedDriveId: driveContext.driveId }
        );

        processedFolderUpdate = {
          id: processedResolution.id,
          name: processedResolution.name || processedFolderRequest.name || null,
          path: processedResolution.path || processedFolderRequest.path || null,
          webUrl: processedResolution.webUrl || processedFolderRequest.webUrl || null,
          parentId: processedResolution.parentId || processedFolderRequest.parentId || null,
        };
      }

      const nextState = {
        enabled: true,
        status: 'connected',
        monitoredFolder,
        deltaLink: null,
        lastSyncStatus: null,
        lastSyncError: null,
        lastSyncMetrics: null,
        lastSyncReason: 'configuration',
      };

      if (processedFolderUpdate !== undefined) {
        nextState.processedFolder = processedFolderUpdate;
      }

      await updateQuickBooksCompanyOneDrive(realmId, nextState);

      queueOneDrivePoll(realmId, { reason: 'configuration' }).catch((error) => {
        console.warn(`Unable to trigger OneDrive sync for ${realmId}`, error.message || error);
      });
    } else {
      const disableState = {
        enabled: false,
        status: 'disabled',
        lastSyncStatus: null,
        lastSyncError: null,
        lastSyncMetrics: null,
        lastSyncReason: 'disabled',
      };
      await updateQuickBooksCompanyOneDrive(realmId, disableState);
    }

    const updated = await getQuickBooksCompanyRecord(realmId);
    return res.json({ oneDrive: sanitizeOneDriveSettings(updated?.oneDrive) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    if (error.code === 'ONEDRIVE_GLOBAL_UNCONFIGURED') {
      return res.status(409).json({ error: 'Configure the shared OneDrive drive before enabling company folders.' });
    }
    if (error.code === 'ONEDRIVE_FOLDER_MISMATCH') {
      return res.status(400).json({ error: 'Selected folder does not belong to the shared OneDrive drive.' });
    }
    if (error?.status === 404) {
      return res.status(404).json({ error: 'Unable to locate the specified OneDrive folder.' });
    }
    if (error?.status === 403) {
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

async function handleOneDriveSyncRequest(req, res, { forceFullDefault = false } = {}) {
  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  if (!isOneDriveSyncConfigured()) {
    return res.status(503).json({ error: 'OneDrive integration is not configured on the server.' });
  }

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const query = req.query || {};
  const forceIndicators = [body.forceFull, query.forceFull];
  const wantsFullResync =
    forceFullDefault || forceIndicators.some((value) => value === true || value === 'true' || value === '1');

  try {
    const company = await getQuickBooksCompanyRecord(realmId);
    if (!company) {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }

    const state = ensureOneDriveStateDefaults(company.oneDrive);
    let responseCompany = company;

    if (wantsFullResync) {
      if (!state || state.enabled === false || !state.monitoredFolder?.id) {
        return res
          .status(409)
          .json({ error: 'Configure an active OneDrive folder before requesting a full resync.' });
      }

      responseCompany = await updateQuickBooksCompanyOneDrive(realmId, {
        deltaLink: null,
        lastSyncStatus: null,
        lastSyncError: null,
        lastSyncMetrics: null,
        lastSyncReason: 'full-resync',
      });
    }

    const syncReason = wantsFullResync ? 'full-resync' : 'manual';
    const failureLabel = wantsFullResync ? 'full OneDrive resync' : 'manual OneDrive sync';

    queueOneDrivePoll(realmId, { reason: syncReason, forceFull: wantsFullResync }).catch((error) => {
      console.warn(`Unable to trigger ${failureLabel} for ${realmId}`, error.message || error);
    });

    return res
      .status(202)
      .json({ accepted: true, oneDrive: sanitizeOneDriveSettings(responseCompany?.oneDrive) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error(
      wantsFullResync ? 'Failed to schedule OneDrive full resync' : 'Failed to schedule OneDrive sync',
      error
    );
    return res.status(500).json({
      error: wantsFullResync ? 'Failed to schedule OneDrive full resync.' : 'Failed to schedule OneDrive sync.',
    });
  }
}

app.post('/api/quickbooks/companies/:realmId/onedrive/sync', (req, res) =>
  handleOneDriveSyncRequest(req, res)
);

app.post('/api/quickbooks/companies/:realmId/onedrive/resync', (req, res) =>
  handleOneDriveSyncRequest(req, res, { forceFullDefault: true })
);

app.post('/api/quickbooks/companies/:realmId/gmail/auth-url', async (req, res) => {
  if (!isGmailMonitoringConfigured()) {
    return res.status(503).json({ error: 'Gmail OAuth client is not configured on the server.' });
  }

  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    return res.status(404).json({ error: 'QuickBooks company not found.' });
  }

  const emailInput = typeof req.body?.email === 'string' ? req.body.email.trim() : '';
  const state = createGmailOAuthState(realmId, { email: emailInput });

  const authorizeUrl = new URL(GMAIL_AUTH_URL);
  authorizeUrl.searchParams.set('client_id', GMAIL_CLIENT_ID);
  authorizeUrl.searchParams.set('response_type', 'code');
  authorizeUrl.searchParams.set('scope', GMAIL_SCOPES);
  authorizeUrl.searchParams.set('redirect_uri', GMAIL_REDIRECT_URI);
  authorizeUrl.searchParams.set('state', state);
  authorizeUrl.searchParams.set('access_type', 'offline');
  authorizeUrl.searchParams.set('prompt', 'consent');
  authorizeUrl.searchParams.set('include_granted_scopes', 'true');
  if (emailInput) {
    authorizeUrl.searchParams.set('login_hint', emailInput);
  }

  res.json({ url: authorizeUrl.toString(), state });
});

app.patch('/api/quickbooks/companies/:realmId/gmail', async (req, res) => {
  if (!isGmailMonitoringConfigured()) {
    return res.status(503).json({ error: 'Gmail OAuth client is not configured on the server.' });
  }

  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    return res.status(404).json({ error: 'QuickBooks company not found.' });
  }

  const body = req.body || {};
  const updates = {};

  if (body.email !== undefined) {
    if (body.email === null) {
      updates.email = null;
    } else if (typeof body.email === 'string' && body.email.trim()) {
      updates.email = body.email.trim();
    } else {
      return res.status(400).json({ error: 'Email must be a non-empty string when provided.' });
    }
  }

  if (body.enabled !== undefined) {
    updates.enabled = body.enabled !== false;
  }

  if (body.searchQuery !== undefined) {
    if (body.searchQuery === null) {
      updates.searchQuery = null;
    } else if (typeof body.searchQuery === 'string') {
      updates.searchQuery = body.searchQuery.trim();
    } else {
      return res.status(400).json({ error: 'searchQuery must be a string when provided.' });
    }
  }

  if (body.labelIds !== undefined) {
    let labels;
    if (Array.isArray(body.labelIds)) {
      labels = body.labelIds;
    } else if (typeof body.labelIds === 'string') {
      labels = parseDelimitedList(body.labelIds);
    } else {
      return res.status(400).json({ error: 'labelIds must be a string or array of labels.' });
    }
    updates.labelIds = labels;
  }

  if (body.allowedMimeTypes !== undefined) {
    let allowed;
    if (Array.isArray(body.allowedMimeTypes)) {
      allowed = body.allowedMimeTypes;
    } else if (typeof body.allowedMimeTypes === 'string') {
      allowed = parseDelimitedList(body.allowedMimeTypes);
    } else {
      return res.status(400).json({ error: 'allowedMimeTypes must be a string or array of MIME types.' });
    }
    updates.allowedMimeTypes = allowed;
  }

  if (body.pollIntervalMs !== undefined) {
    const interval = Number.parseInt(body.pollIntervalMs, 10);
    if (!Number.isFinite(interval) || interval < GMAIL_POLL_MIN_INTERVAL_MS) {
      return res
        .status(400)
        .json({ error: `pollIntervalMs must be a number greater than or equal to ${GMAIL_POLL_MIN_INTERVAL_MS}.` });
    }
    updates.pollIntervalMs = interval;
  }

  if (body.maxAttachmentBytes !== undefined) {
    const maxBytes = Number.parseInt(body.maxAttachmentBytes, 10);
    if (!Number.isFinite(maxBytes) || maxBytes < 1024) {
      return res.status(400).json({ error: 'maxAttachmentBytes must be a number >= 1024.' });
    }
    updates.maxAttachmentBytes = maxBytes;
  }

  if (body.maxResults !== undefined) {
    const maxResults = Number.parseInt(body.maxResults, 10);
    if (!Number.isFinite(maxResults) || maxResults <= 0) {
      return res.status(400).json({ error: 'maxResults must be a positive number.' });
    }
    updates.maxResults = maxResults;
  }

  if (body.businessType !== undefined) {
    if (body.businessType === null) {
      updates.businessType = null;
    } else if (typeof body.businessType === 'string') {
      updates.businessType = body.businessType.trim().slice(0, 120) || null;
    } else {
      return res.status(400).json({ error: 'businessType must be a string when provided.' });
    }
  }

  const currentConfig = ensureGmailConfigDefaults(company.gmail);
  const willEnable = updates.enabled !== undefined
    ? updates.enabled
    : currentConfig
      ? currentConfig.enabled !== false
      : false;
  if (willEnable && !(currentConfig?.refreshToken || updates.refreshToken)) {
    return res.status(409).json({ error: 'Connect Gmail before enabling inbox monitoring.' });
  }

  try {
    const updated = await updateQuickBooksCompanyGmail(realmId, updates);
    const gmailConfig = ensureGmailConfigDefaults(updated.gmail);
    if (gmailConfig?.enabled && gmailConfig.refreshToken) {
      queueGmailPoll(realmId, { reason: 'configuration' }).catch((error) => {
        console.warn(`Unable to trigger Gmail sync for ${realmId}`, error.message || error);
      });
    }
    res.json({ gmail: sanitizeGmailSettings(updated.gmail) });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to update Gmail settings', error);
    res.status(500).json({ error: 'Failed to update Gmail settings.' });
  }
});

app.delete('/api/quickbooks/companies/:realmId/gmail', async (req, res) => {
  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  try {
    await updateQuickBooksCompanyGmail(realmId, null, { replace: true });
    gmailTokenCache.delete(realmId);
    try {
      await deleteCompanyGmailState(realmId);
    } catch (stateError) {
      console.warn(`Unable to remove Gmail polling state for ${realmId}`, stateError.message || stateError);
    }
    res.json({ gmail: null });
  } catch (error) {
    if (error.code === 'NOT_FOUND') {
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }
    console.error('Failed to remove Gmail settings', error);
    res.status(500).json({ error: 'Failed to remove Gmail settings.' });
  }
});

app.post('/api/quickbooks/companies/:realmId/gmail/sync', async (req, res) => {
  if (!isGmailMonitoringConfigured()) {
    return res.status(503).json({ error: 'Gmail OAuth client is not configured on the server.' });
  }

  const realmId = req.params.realmId;
  if (!realmId) {
    return res.status(400).json({ error: 'Realm ID is required.' });
  }

  const company = await getQuickBooksCompanyRecord(realmId);
  if (!company) {
    return res.status(404).json({ error: 'QuickBooks company not found.' });
  }

  const gmailConfig = ensureGmailConfigDefaults(company.gmail);
  if (!gmailConfig?.enabled) {
    return res.status(400).json({ error: 'Enable Gmail monitoring before requesting a manual sync.' });
  }

  if (!gmailConfig.refreshToken) {
    return res.status(409).json({ error: 'Connect Gmail to obtain a refresh token before syncing.' });
  }

  try {
    queueGmailPoll(realmId, { reason: 'manual' }).catch((error) => {
      console.warn(`Unable to trigger manual Gmail sync for ${realmId}`, error.message || error);
    });
    res.status(202).json({ accepted: true, gmail: sanitizeGmailSettings(gmailConfig) });
  } catch (error) {
    console.error('Failed to schedule Gmail sync', error);
    res.status(500).json({ error: 'Failed to schedule Gmail sync.' });
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

app.patch('/api/invoices/:checksum/review', async (req, res) => {
  const checksum = req.params.checksum;
  let body = req.body;

  if (typeof body === 'string') {
    try {
      body = body.trim() ? JSON.parse(body) : {};
    } catch (error) {
      console.warn('Unable to parse review update payload as JSON', { checksum });
      body = {};
    }
  } else if (!body || typeof body !== 'object') {
    body = {};
  }

  const hasVendorUpdate = Object.prototype.hasOwnProperty.call(body, 'vendorId');
  const hasAccountUpdate = Object.prototype.hasOwnProperty.call(body, 'accountId');
  const hasTaxUpdate = Object.prototype.hasOwnProperty.call(body, 'taxCodeId');
  const hasSecondaryTaxUpdate = Object.prototype.hasOwnProperty.call(body, 'secondaryTaxCodeId');

  if (!checksum) {
    return res.status(400).json({ error: 'Checksum is required.' });
  }

  if (!hasVendorUpdate && !hasAccountUpdate && !hasTaxUpdate && !hasSecondaryTaxUpdate) {
    return res.status(400).json({ error: 'No review updates provided.' });
  }

  const updates = {};
  if (hasVendorUpdate) {
    updates.vendorId = body.vendorId;
  }
  if (hasAccountUpdate) {
    updates.accountId = body.accountId;
  }
  if (hasTaxUpdate) {
    updates.taxCodeId = body.taxCodeId;
  }
  if (hasSecondaryTaxUpdate) {
    updates.secondaryTaxCodeId = body.secondaryTaxCodeId;
  }

  try {
    const updated = await updateStoredInvoiceReviewSelection(checksum, updates);

    if (!updated) {
      return res.status(404).json({ error: 'Invoice not found.' });
    }

    return res.json({ invoice: updated });
  } catch (error) {
    console.error('Failed to update invoice review selection', error);
    return res.status(500).json({ error: 'Failed to update invoice review selection.' });
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

app.get('/api/preview-quickbooks', async (req, res) => {
  const invoiceIdRaw = typeof req.query.invoiceId === 'string' ? req.query.invoiceId.trim() : '';
  const realmIdRaw = typeof req.query.realmId === 'string' ? req.query.realmId.trim() : '';

  if (!invoiceIdRaw) {
    return res.status(400).json({ error: 'invoiceId query parameter is required.' });
  }

  if (!realmIdRaw) {
    return res.status(400).json({ error: 'realmId query parameter is required.' });
  }

  try {
    const [invoice, metadata] = await Promise.all([
      findStoredInvoiceByChecksum(invoiceIdRaw),
      readQuickBooksCompanyMetadata(realmIdRaw).catch((error) => {
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
      return res.status(404).json({ error: 'QuickBooks company not found.' });
    }

    const invoiceRealmId = invoice?.metadata?.companyProfile?.realmId;
    if (invoiceRealmId && invoiceRealmId !== realmIdRaw) {
      return res.status(409).json({ error: 'Invoice belongs to a different QuickBooks company.' });
    }

    const missingMetadata = identifyMissingQuickBooksMetadata(metadata);
    if (missingMetadata.length) {
      return res.status(412).json({
        error: 'QuickBooks metadata is missing. Refresh QuickBooks data before previewing.',
        missing: missingMetadata,
      });
    }

    let payload;
    try {
      payload = buildQuickBooksInvoicePayload(invoice, { metadata, realmId: realmIdRaw });
    } catch (error) {
      if (error?.status === 422) {
        const responseBody = { error: error.message };
        if (error.details) {
          responseBody.details = error.details;
        }
        return res.status(422).json(responseBody);
      }
      throw error;
    }

    const url = buildQuickBooksInvoiceUrl(realmIdRaw);
    console.info('quickbooks-preview.generated', {
      invoiceId: invoiceIdRaw,
      realmId: realmIdRaw,
      timestamp: new Date().toISOString(),
    });

    res.json({
      method: 'POST',
      url,
      payload,
    });
  } catch (error) {
    console.error('Failed to build QuickBooks preview', error);
    res.status(500).json({ error: 'Failed to build QuickBooks preview.' });
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

gmailCallbackPaths.forEach((callbackPath) => {
  app.get(callbackPath, handleGmailCallback);
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

function createGmailOAuthState(realmId, payload = {}) {
  const state = crypto.randomBytes(16).toString('hex');
  const now = Date.now();
  gmailStates.set(state, { realmId, createdAt: now, ...payload });
  for (const [key, entry] of gmailStates.entries()) {
    if (now - entry.createdAt > GMAIL_STATE_TTL_MS) {
      gmailStates.delete(key);
    }
  }
  return state;
}

function consumeGmailOAuthState(state) {
  if (!state) {
    return null;
  }

  const entry = gmailStates.get(state);
  if (!entry) {
    return null;
  }

  gmailStates.delete(state);
  if (Date.now() - entry.createdAt > GMAIL_STATE_TTL_MS) {
    return null;
  }

  return entry;
}

async function handleGmailCallback(req, res) {
  const { code, state } = req.query;
  if (!code || !state) {
    return res.status(400).send('Missing required parameters from Gmail.');
  }

  const entry = consumeGmailOAuthState(state);
  if (!entry?.realmId) {
    return res.status(400).send('Gmail authorization state is invalid or has expired. Please try again.');
  }

  const realmId = entry.realmId;

  try {
    const tokenSet = await exchangeGmailCode(code);
    if (!tokenSet.refreshToken) {
      throw new Error('Gmail authorization did not return a refresh token. Ensure offline access is granted.');
    }

    const updatePayload = {
      refreshToken: tokenSet.refreshToken,
      enabled: true,
      status: 'connected',
      lastSyncStatus: null,
      lastSyncError: null,
      lastSyncAt: null,
      lastSyncMetrics: null,
      lastSyncReason: 'oauth',
      lastConnectedAt: new Date().toISOString(),
      email: entry.email ? entry.email.trim() : null,
      historyId: null,
    };

    const updated = await updateQuickBooksCompanyGmail(realmId, updatePayload);
    gmailTokenCache.delete(realmId);
    try {
      await deleteCompanyGmailState(realmId);
    } catch (stateError) {
      console.warn(`Unable to reset Gmail polling state for ${realmId}`, stateError.message || stateError);
    }

    const companyName = updated?.companyName || updated?.legalName || realmId;
    return res.redirect(
      `/?gmail=connected&realmId=${encodeURIComponent(realmId)}&company=${encodeURIComponent(companyName)}`
    );
  } catch (error) {
    console.error('Gmail OAuth callback failed', error);
    return res.redirect(`/?gmail=error&message=${encodeURIComponent('Failed to connect Gmail inbox.')}`);
  }
}

app.use((req, res) => {
  res.status(404).json({ error: 'Not found' });
});

if (require.main === module) {
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

  if (isGmailMonitoringConfigured()) {
    startGmailMonitor();
  } else if (GMAIL_CLIENT_ID || GMAIL_CLIENT_SECRET) {
    console.warn(
      'Gmail monitoring is not enabled. Provide both GMAIL_CLIENT_ID and GMAIL_CLIENT_SECRET to activate inbox polling.'
    );
  }
}

module.exports = {
  app,
  buildQuickBooksExpenseLines,
  resolveQuickBooksTaxCodes,
  analyseInvoiceVatSignals,
  sumLineAmounts,
  normaliseMoneyValue,
  readQuickBooksCompanies,
  persistQuickBooksCompanies,
  emitHealthMetric,
  getHealthMetricHistory,
  attemptQuickBooksCompaniesRepair,
  QUICKBOOKS_COMPANIES_FILE,
};

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

async function exchangeGmailCode(code) {
  if (!isGmailMonitoringConfigured()) {
    throw new Error('Gmail OAuth client is not configured.');
  }

  const params = new URLSearchParams({
    code,
    client_id: GMAIL_CLIENT_ID,
    client_secret: GMAIL_CLIENT_SECRET,
    redirect_uri: GMAIL_REDIRECT_URI,
    grant_type: 'authorization_code',
  });

  const response = await fetch(GMAIL_TOKEN_URL, {
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
      `Gmail token exchange failed with status ${response.status}`;
    const error = new Error(message);
    error.body = errorBody;
    error.status = response.status;
    throw error;
  }

  const data = await response.json();
  const now = Date.now();

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || null,
    scope: data.scope || null,
    tokenType: data.token_type || null,
    expiresAt: data.expires_in ? new Date(now + data.expires_in * 1000).toISOString() : null,
  };
}

async function fetchQuickBooksCompanyInfo(realmId, accessToken) {
  const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/companyinfo/${realmId}?minorversion=${encodeURIComponent(
    QUICKBOOKS_MINOR_VERSION
  )}`;
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

function findBalancedJsonArrayEndIndex(source) {
  if (!source || typeof source !== 'string') {
    return -1;
  }

  let depth = 0;
  let inString = false;
  let isEscaped = false;

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];

    if (inString) {
      if (isEscaped) {
        isEscaped = false;
        continue;
      }

      if (char === '\\') {
        isEscaped = true;
      } else if (char === '"') {
        inString = false;
      }
      continue;
    }

    if (char === '"') {
      inString = true;
      continue;
    }

    if (char === '[' || char === '{') {
      depth += 1;
      continue;
    }

    if (char === ']' || char === '}') {
      depth -= 1;
      if (depth === 0 && char === ']') {
        return index;
      }
      continue;
    }
  }

  return -1;
}

async function backupCorruptQuickBooksCompaniesFile(rawContents) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = `${QUICKBOOKS_COMPANIES_FILE}.corrupt-${timestamp}.json`;

  try {
    await fs.writeFile(backupPath, rawContents, { flag: 'wx' });
    console.error(
      `[QuickBooks] Failed to parse ${QUICKBOOKS_COMPANIES_FILE}. Backed up corrupt contents to ${backupPath}`
    );
    return backupPath;
  } catch (backupErr) {
    if (backupErr.code === 'EEXIST') {
      console.warn(`[QuickBooks] Backup file ${backupPath} already exists. Skipping overwrite.`);
    } else {
      console.error(`[QuickBooks] Failed to write corrupt backup ${backupPath}: ${backupErr.message}`);
    }
    return null;
  }
}

async function attemptQuickBooksCompaniesRepair({ raw, parseError, backupPath }) {
  const closingIndex = findBalancedJsonArrayEndIndex(raw);
  if (closingIndex === -1) {
    console.error(
      `[QuickBooks] Unable to locate balanced closing bracket while repairing ${QUICKBOOKS_COMPANIES_FILE}.`
    );
    return null;
  }

  const trimmed = `${raw.slice(0, closingIndex + 1)}\n`;
  try {
    const repaired = JSON.parse(trimmed);
    await persistQuickBooksCompanies(repaired);
    const truncatedBytes = Math.max(raw.length - (closingIndex + 1), 0);
    console.warn(
      `[QuickBooks] Repaired ${QUICKBOOKS_COMPANIES_FILE} by truncating ${truncatedBytes} trailing byte${
        truncatedBytes === 1 ? '' : 's'
      } at index ${closingIndex}`
    );
    emitHealthMetric('quickbooks.company_file.repaired', {
      truncatedBytes,
      backupPath: backupPath || null,
      parseError: parseError?.message || null,
    });
    return {
      companies: repaired,
      truncatedBytes,
      closingIndex,
    };
  } catch (repairErr) {
    console.error(
      `[QuickBooks] Attempted repair of ${QUICKBOOKS_COMPANIES_FILE} failed: ${repairErr.message}`
    );
    return null;
  }
}

async function readQuickBooksCompanies({ allowRepair = true } = {}) {
  console.info('[QuickBooks] Loading companies file', {
    file: QUICKBOOKS_COMPANIES_FILE,
  });
  let file;
  try {
    file = await fs.readFile(QUICKBOOKS_COMPANIES_FILE, 'utf-8');
  } catch (err) {
    if (err.code === 'ENOENT') {
      console.warn('[QuickBooks] Companies file missing; returning empty list', {
        file: QUICKBOOKS_COMPANIES_FILE,
      });
      return [];
    }
    throw err;
  }

  try {
    const companies = JSON.parse(file);
    console.info('[QuickBooks] Parsed companies payload', {
      count: Array.isArray(companies) ? companies.length : 0,
    });
    const migration = await applyLegacyOneDriveMigration(companies);
    if (migration?.companiesChanged) {
      await persistQuickBooksCompanies(companies);
    }
    return companies;
  } catch (parseErr) {
    const backupPath = await backupCorruptQuickBooksCompaniesFile(file);

    if (allowRepair) {
      const repaired = await attemptQuickBooksCompaniesRepair({ raw: file, parseError: parseErr, backupPath });
      if (repaired?.companies) {
        return repaired.companies;
      }
    }

    const error = new Error(
      `Failed to parse QuickBooks companies file at ${QUICKBOOKS_COMPANIES_FILE}. Restore the backup or run the repair CLI.`
    );
    error.code = 'QUICKBOOKS_COMPANY_FILE_CORRUPT';
    error.cause = parseErr;
    if (backupPath) {
      error.backupPath = backupPath;
    }
    throw error;
  }
}

async function applyLegacyOneDriveMigration(companies) {
  if (hasAppliedOneDriveMigration) {
    return { companiesChanged: false };
  }
  hasAppliedOneDriveMigration = true;

  if (!Array.isArray(companies) || !companies.length) {
    console.info('[QuickBooks] Skipping legacy OneDrive migration - no companies detected');
    return { companiesChanged: false };
  }

  console.info('[QuickBooks] Running legacy OneDrive migration', {
    companyCount: companies.length,
  });

  let companiesChanged = false;
  let discoveredDriveId = null;
  let discoveredShareUrl = null;

  const globalConfig = await loadGlobalOneDriveConfig();
  if (isGlobalOneDriveConfigured(globalConfig)) {
    console.info('[QuickBooks] Legacy OneDrive migration skipped - global config already configured');
    discoveredDriveId = sanitiseOptionalString(globalConfig.driveId) || null;
    discoveredShareUrl = sanitiseOptionalString(globalConfig.shareUrl) || null;
  }

  for (const company of companies) {
    if (!company || typeof company !== 'object') {
      continue;
    }

    const original = company.oneDrive;
    const originalSerialised = JSON.stringify(original ?? null);

    const legacyDriveId = sanitiseOptionalString(original?.driveId);
    const legacyShareUrl = sanitiseOptionalString(original?.shareUrl);

    const normalised = ensureOneDriveStateDefaults(original);
    const nextState = normalised ? { ...normalised } : null;

    if (nextState && nextState.monitoredFolder) {
      nextState.monitoredFolder = normalizeOneDriveFolderConfig(nextState.monitoredFolder);
    }
    if (nextState && Object.prototype.hasOwnProperty.call(nextState, 'processedFolder')) {
      nextState.processedFolder = nextState.processedFolder
        ? normalizeOneDriveFolderConfig(nextState.processedFolder)
        : null;
    }

    if (JSON.stringify(nextState ?? null) !== originalSerialised) {
      companiesChanged = true;
    }

    company.oneDrive = nextState;

    if (!discoveredDriveId && legacyDriveId) {
      discoveredDriveId = legacyDriveId;
      if (!discoveredShareUrl && legacyShareUrl) {
        discoveredShareUrl = legacyShareUrl;
      }
    }
  }

  let globalUpdate = null;
  if (!isGlobalOneDriveConfigured(globalConfig) && discoveredDriveId) {
    console.info('[QuickBooks] Legacy OneDrive migration detected legacy drive metadata - seeding global config');
    globalUpdate = {
      driveId: discoveredDriveId,
      shareUrl: discoveredShareUrl || null,
      status: 'ready',
      lastValidatedAt: new Date().toISOString(),
      lastValidationError: null,
    };
  }

  if (globalUpdate) {
    try {
      await updateGlobalOneDriveConfig(globalUpdate);
      console.info('[QuickBooks] Applied legacy OneDrive global migration update');
    } catch (error) {
      console.warn('[QuickBooks] Failed to persist legacy OneDrive global migration update', {
        message: error?.message || error,
      });
      globalUpdate = null;
    }
  }

  console.debug('[QuickBooks] Legacy OneDrive migration result', {
    companiesChanged,
    globalUpdateApplied: Boolean(globalUpdate),
  });

  return {
    companiesChanged,
    globalUpdate,
  };
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
  record.gmail = ensureGmailConfigDefaults(record.gmail || existing.gmail || null);

  if (existingIndex >= 0) {
    companies[existingIndex] = record;
  } else {
    companies.push(record);
  }

  await persistQuickBooksCompanies(companies);
  return record;
}

async function persistQuickBooksCompanies(companies) {
  return runWithQuickBooksCompaniesWriteLock(async () => {
    const dir = path.dirname(QUICKBOOKS_COMPANIES_FILE);
    const tempPath = `${QUICKBOOKS_COMPANIES_FILE}.tmp-${process.pid}-${Date.now()}`;
    const payload = `${JSON.stringify(companies, null, 2)}\n`;

    await fs.mkdir(dir, { recursive: true });

    let handle;
    try {
      handle = await fs.open(tempPath, 'w', 0o600);
      await handle.writeFile(payload);
      await handle.sync();
    } finally {
      if (handle) {
        await handle.close();
      }
    }

    try {
      await fs.rename(tempPath, QUICKBOOKS_COMPANIES_FILE);
    } catch (err) {
      await fs.unlink(tempPath).catch(() => {});
      throw err;
    }

    let dirHandle;
    try {
      dirHandle = await fs.open(dir, 'r');
      await dirHandle.sync();
    } catch (dirErr) {
      if (dirErr && dirErr.code !== 'EISDIR' && dirErr.code !== 'ENOENT') {
        console.warn(
          `[QuickBooks] Failed to fsync directory ${dir}: ${dirErr.message}`
        );
      }
    } finally {
      if (dirHandle) {
        await dirHandle.close().catch(() => {});
      }
    }
  });
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
  next.gmail = ensureGmailConfigDefaults(next.gmail || companies[index].gmail || null);

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

  companies[index].gmail = ensureGmailConfigDefaults(companies[index].gmail || null);

  await persistQuickBooksCompanies(companies);
  return companies[index];
}

async function updateQuickBooksCompanyGmail(realmId, updates, { replace = false } = {}) {
  const companies = await readQuickBooksCompanies();
  const index = companies.findIndex((entry) => entry.realmId === realmId);

  if (index < 0) {
    const error = new Error('QuickBooks company not found.');
    error.code = 'NOT_FOUND';
    throw error;
  }

  const now = new Date().toISOString();
  const currentNormalised = normalizeGmailConfig(companies[index].gmail || null);
  const current = ensureGmailConfigDefaults(currentNormalised);
  let nextState = null;

  if (replace) {
    const replacement = normalizeGmailConfig(updates);
    nextState = replacement ? ensureGmailConfigDefaults(replacement) : null;
    if (nextState) {
      nextState.updatedAt = now;
      if (!nextState.createdAt) {
        nextState.createdAt = now;
      }
      if (nextState.refreshToken) {
        nextState.lastConnectedAt = nextState.lastConnectedAt || now;
      }
    }
  } else {
    const updateFragment = normalizeGmailConfig(updates || {});
    if (current || (updateFragment && Object.keys(updateFragment).length)) {
      const merged = {
        ...(current || {}),
        ...(updateFragment || {}),
        updatedAt: now,
      };

      if (!updateFragment?.refreshToken && current?.refreshToken) {
        merged.refreshToken = current.refreshToken;
      }

      if (!merged.createdAt) {
        merged.createdAt = current?.createdAt || now;
      }

      if (updateFragment?.refreshToken) {
        merged.lastConnectedAt = now;
      }

      nextState = ensureGmailConfigDefaults(merged);
    }
  }

  companies[index] = {
    ...companies[index],
    gmail: nextState,
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
  companies[index].gmail = ensureGmailConfigDefaults(companies[index].gmail || null);

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

  const id = sanitiseQuickBooksId(vendor.Id);
  if (!id) {
    return null;
  }

  return {
    id,
    displayName:
      vendor.DisplayName ||
      vendor.CompanyName ||
      vendor.Title ||
      vendor.FamilyName ||
      vendor.GivenName ||
      (id ? `Vendor ${id}` : null),
    email: vendor.PrimaryEmailAddr?.Address || null,
    phone: vendor.PrimaryPhone?.FreeFormNumber || null,
  };
}

function transformQuickBooksAccount(account) {
  if (!account) {
    return null;
  }

  const id = sanitiseQuickBooksId(account.Id);
  if (!id) {
    return null;
  }

  return {
    id,
    name: account.Name || account.FullyQualifiedName || `Account ${id}`,
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
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/vendor?minorversion=${encodeURIComponent(
      QUICKBOOKS_MINOR_VERSION
    )}`;
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
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/account?minorversion=${encodeURIComponent(
      QUICKBOOKS_MINOR_VERSION
    )}`;
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
    .map((code) => {
      const id = sanitiseQuickBooksId(code.Id);
      if (!id) {
        return null;
      }

      return {
        id,
        name: code.Name || `Tax Code ${id}`,
        description: code.Description || null,
        rate: deriveTaxCodeRate(code, taxRateLookup),
        agency: code.SalesTaxRateList?.TaxAgencyRef?.name || null,
        active: code.Active !== false,
      };
    })
    .filter(Boolean);
}

async function fetchQuickBooksQueryList(realmId, accessToken, entity) {
  const pageSize = 200;
  let startPosition = 1;
  const items = [];

  while (true) {
    const query = `select * from ${entity} startposition ${startPosition} maxresults ${pageSize}`;
    const url = `${QUICKBOOKS_API_BASE_URL}/v3/company/${realmId}/query?minorversion=${encodeURIComponent(
      QUICKBOOKS_MINOR_VERSION
    )}&query=${encodeURIComponent(query)}`;

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
    minorversion: QUICKBOOKS_MINOR_VERSION,
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
    if (error?.code === 'QUICKBOOKS_COMPANY_FILE_CORRUPT') {
      console.error(
        'Unable to warm QuickBooks metadata because the company store is corrupt. Restore the backup to resume warmup.'
      );
      emitHealthMetric('quickbooks.company_file.corrupt', { source: 'metadata_warmup' });
    } else {
      console.error('Unable to warm QuickBooks metadata', error);
    }
  }
}

function sanitizeQuickBooksCompany(company) {
  if (!company) {
    return company;
  }

  const { tokens, oneDrive, gmail, ...rest } = company;
  return {
    ...rest,
    businessType: rest.businessType ?? null,
    oneDrive: sanitizeOneDriveSettings(oneDrive),
    gmail: sanitizeGmailSettings(gmail),
  };
}

function sanitizeOneDriveSettings(config) {
  const normalized = ensureOneDriveStateDefaults(config);
  if (!normalized) {
    return null;
  }

  const { deltaLink, clientState, ...rest } = normalized;
  return rest;
}

function sanitizeGmailSettings(config) {
  const normalized = ensureGmailConfigDefaults(config);
  if (!normalized) {
    return null;
  }

  const { refreshToken, ...rest } = normalized;
  return rest;
}

function normalizeOneDriveFolderConfig(folder) {
  if (!folder || typeof folder !== 'object') {
    return null;
  }

  const result = {};

  if (Object.prototype.hasOwnProperty.call(folder, 'id')) {
    result.id = sanitiseOptionalString(folder.id);
  }

  if (Object.prototype.hasOwnProperty.call(folder, 'path')) {
    result.path = sanitiseOptionalString(folder.path);
  }

  if (Object.prototype.hasOwnProperty.call(folder, 'name')) {
    result.name = sanitiseOptionalString(folder.name);
  }

  if (Object.prototype.hasOwnProperty.call(folder, 'webUrl')) {
    result.webUrl = sanitiseOptionalString(folder.webUrl);
  }

  if (Object.prototype.hasOwnProperty.call(folder, 'parentId')) {
    result.parentId = sanitiseOptionalString(folder.parentId);
  }

  if (!Object.values(result).some((value) => value)) {
    return null;
  }

  return result;
}

function extractLegacyMonitoredFolder(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const legacy = {
    id: sanitiseOptionalString(config.folderId),
    path: sanitiseOptionalString(config.folderPath),
    name: sanitiseOptionalString(config.folderName),
    webUrl: sanitiseOptionalString(config.webUrl),
    parentId: sanitiseOptionalString(config.parentId),
  };

  return normalizeOneDriveFolderConfig(legacy);
}

function extractLegacyProcessedFolder(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const legacy = {
    id: sanitiseOptionalString(config.processedFolderId),
    path: sanitiseOptionalString(config.processedFolderPath),
    name: sanitiseOptionalString(config.processedFolderName),
    webUrl: sanitiseOptionalString(config.processedFolderWebUrl),
    parentId: sanitiseOptionalString(config.processedFolderParentId),
  };

  return normalizeOneDriveFolderConfig(legacy);
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

  if (result.monitoredFolder) {
    result.monitoredFolder = normalizeOneDriveFolderConfig(result.monitoredFolder);
  }

  if (Object.prototype.hasOwnProperty.call(result, 'processedFolder')) {
    result.processedFolder = result.processedFolder
      ? normalizeOneDriveFolderConfig(result.processedFolder)
      : null;
  }

  return result;
}

function normalizeOneDriveState(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const result = {};
  let hasValue = false;

  if (Object.prototype.hasOwnProperty.call(config, 'enabled')) {
    result.enabled = config.enabled === false ? false : true;
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'status')) {
    const statusValue = sanitiseOptionalString(config.status);
    result.status = statusValue || null;
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'monitoredFolder')) {
    if (config.monitoredFolder === null) {
      result.monitoredFolder = null;
      hasValue = true;
    } else {
      const folder = normalizeOneDriveFolderConfig(config.monitoredFolder);
      if (folder) {
        result.monitoredFolder = folder;
        hasValue = true;
      }
    }
  } else {
    const legacyMonitored = extractLegacyMonitoredFolder(config);
    if (legacyMonitored) {
      result.monitoredFolder = legacyMonitored;
      hasValue = true;
    }
  }

  if (Object.prototype.hasOwnProperty.call(config, 'processedFolder')) {
    if (config.processedFolder === null) {
      result.processedFolder = null;
      hasValue = true;
    } else {
      const folder = normalizeOneDriveFolderConfig(config.processedFolder);
      if (folder) {
        result.processedFolder = folder;
        hasValue = true;
      }
    }
  } else {
    const legacyProcessed = extractLegacyProcessedFolder(config);
    if (legacyProcessed) {
      result.processedFolder = legacyProcessed;
      hasValue = true;
    }
  }

  if (Object.prototype.hasOwnProperty.call(config, 'deltaLink')) {
    result.deltaLink = sanitiseOptionalString(config.deltaLink);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncLog')) {
    result.lastSyncLog = normalizeOneDriveSyncLog(config.lastSyncLog);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncAt')) {
    result.lastSyncAt = sanitiseIsoString(config.lastSyncAt);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncStatus')) {
    result.lastSyncStatus = sanitiseOptionalString(config.lastSyncStatus);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncReason')) {
    result.lastSyncReason = sanitiseOptionalString(config.lastSyncReason);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncDurationMs')) {
    result.lastSyncDurationMs = Number.isFinite(config.lastSyncDurationMs)
      ? Number(config.lastSyncDurationMs)
      : null;
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncError')) {
    result.lastSyncError = normalizeOneDriveSyncError(config.lastSyncError);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'lastSyncMetrics')) {
    result.lastSyncMetrics = normalizeOneDriveMetrics(config.lastSyncMetrics);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'subscriptionId')) {
    result.subscriptionId = sanitiseOptionalString(config.subscriptionId);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'subscriptionExpiration')) {
    result.subscriptionExpiration = sanitiseIsoString(config.subscriptionExpiration);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'clientState')) {
    result.clientState = sanitiseOptionalString(config.clientState);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'createdAt')) {
    result.createdAt = sanitiseIsoString(config.createdAt);
    hasValue = true;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'updatedAt')) {
    result.updatedAt = sanitiseIsoString(config.updatedAt);
    hasValue = true;
  }

  if (!hasValue) {
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
  const duplicateCount = Number.isFinite(metrics.duplicateCount) ? Number(metrics.duplicateCount) : 0;
  const pages = Number.isFinite(metrics.pages) ? Number(metrics.pages) : 0;
  const errorCount = Number.isFinite(metrics.errorCount) ? Number(metrics.errorCount) : 0;

  return {
    processedItems,
    createdCount,
    skippedCount,
    duplicateCount,
    pages,
    errorCount,
  };
}

function normalizeOneDriveSyncLog(log) {
  if (!Array.isArray(log)) {
    return null;
  }

  const entries = log
    .map((entry) => sanitiseOptionalString(entry))
    .filter(Boolean);

  if (!entries.length) {
    return null;
  }

  return entries.slice(-20);
}

function isOneDriveDeltaResetError(error) {
  if (!error || typeof error !== 'object') {
    return false;
  }

  if (error.status === 410) {
    return true;
  }

  const code = sanitiseOptionalString(error?.body?.error?.code)?.toLowerCase();
  if (code && code.includes('resyncrequired')) {
    return true;
  }

  const message = sanitiseOptionalString(error.message)?.toLowerCase();
  return Boolean(message && message.includes('resync required'));
}

function ensureGmailConfigDefaults(config) {
  const normalized = normalizeGmailConfig(config);
  if (!normalized) {
    return null;
  }

  const result = { ...normalized };
  result.enabled = result.enabled !== false;
  result.searchQuery = result.searchQuery || GMAIL_DEFAULT_SEARCH_QUERY;
  result.labelIds = Array.isArray(result.labelIds) ? result.labelIds : [];
  result.pollIntervalMs = Math.max(
    Number.isFinite(result.pollIntervalMs) ? result.pollIntervalMs : GMAIL_DEFAULT_POLL_INTERVAL_MS,
    GMAIL_POLL_MIN_INTERVAL_MS
  );
  result.maxResults = Number.isFinite(result.maxResults) && result.maxResults > 0
    ? Math.floor(result.maxResults)
    : GMAIL_DEFAULT_MAX_RESULTS;
  result.maxAttachmentBytes = Math.max(
    Number.isFinite(result.maxAttachmentBytes) ? result.maxAttachmentBytes : GMAIL_DEFAULT_MAX_ATTACHMENT_BYTES,
    1024
  );

  const allowed = Array.isArray(result.allowedMimeTypes) && result.allowedMimeTypes.length
    ? result.allowedMimeTypes
    : GMAIL_DEFAULT_ALLOWED_MIME_TYPES.slice();
  result.allowedMimeTypes = Array.from(
    new Set(
      allowed
        .map((value) => (typeof value === 'string' ? value.trim().toLowerCase() : ''))
        .filter(Boolean)
    )
  );

  result.lastSyncError = normalizeGmailSyncError(result.lastSyncError);
  result.lastSyncMetrics = normalizeGmailMetrics(result.lastSyncMetrics);

  if (!result.createdAt) {
    result.createdAt = new Date().toISOString();
  }

  if (!result.updatedAt) {
    result.updatedAt = result.createdAt;
  }

  return result;
}

function normalizeGmailConfig(config) {
  if (!config || typeof config !== 'object') {
    return null;
  }

  const result = {};

  if (Object.prototype.hasOwnProperty.call(config, 'enabled')) {
    result.enabled = config.enabled === false ? false : true;
  }

  result.email = sanitiseOptionalString(config.email);

  if (Object.prototype.hasOwnProperty.call(config, 'refreshToken')) {
    const token = sanitiseOptionalString(config.refreshToken);
    result.refreshToken = token;
  }

  result.searchQuery = sanitiseOptionalString(config.searchQuery);

  if (Object.prototype.hasOwnProperty.call(config, 'labelIds')) {
    const labels = Array.isArray(config.labelIds)
      ? config.labelIds
      : parseDelimitedList(config.labelIds);
    result.labelIds = labels.map((value) => sanitiseOptionalString(value)).filter(Boolean);
  }

  if (Object.prototype.hasOwnProperty.call(config, 'pollIntervalMs')) {
    const interval = Number.parseInt(config.pollIntervalMs, 10);
    result.pollIntervalMs = Number.isFinite(interval) ? interval : null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'maxResults')) {
    const maxResults = Number.parseInt(config.maxResults, 10);
    result.maxResults = Number.isFinite(maxResults) ? maxResults : null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'maxAttachmentBytes')) {
    const maxBytes = Number.parseInt(config.maxAttachmentBytes, 10);
    result.maxAttachmentBytes = Number.isFinite(maxBytes) ? maxBytes : null;
  }

  if (Object.prototype.hasOwnProperty.call(config, 'allowedMimeTypes')) {
    const allowed = Array.isArray(config.allowedMimeTypes)
      ? config.allowedMimeTypes
      : parseDelimitedList(config.allowedMimeTypes);
    result.allowedMimeTypes = allowed.map((value) => sanitiseOptionalString(value)).filter(Boolean);
  }

  result.status = sanitiseOptionalString(config.status);
  result.lastSyncAt = sanitiseIsoString(config.lastSyncAt);
  result.lastSyncStatus = sanitiseOptionalString(config.lastSyncStatus);
  result.lastSyncReason = sanitiseOptionalString(config.lastSyncReason);
  result.lastSyncError = normalizeGmailSyncError(config.lastSyncError);
  result.lastSyncMetrics = normalizeGmailMetrics(config.lastSyncMetrics);
  result.historyId = sanitiseOptionalString(config.historyId);
  result.createdAt = sanitiseIsoString(config.createdAt);
  result.updatedAt = sanitiseIsoString(config.updatedAt);
  result.lastConnectedAt = sanitiseIsoString(config.lastConnectedAt);
  result.businessType = sanitiseOptionalString(config.businessType);

  return result;
}

function normalizeGmailSyncError(error) {
  if (!error || typeof error !== 'object') {
    return null;
  }

  const message = sanitiseOptionalString(error.message);
  if (!message) {
    return null;
  }

  const at = sanitiseIsoString(error.at || error.timestamp) || new Date().toISOString();
  const code = sanitiseOptionalString(error.code);

  return {
    message,
    at,
    code: code || null,
  };
}

function normalizeGmailMetrics(metrics) {
  if (!metrics || typeof metrics !== 'object') {
    return null;
  }

  const processedMessages = Number.isFinite(metrics.processedMessages) ? Number(metrics.processedMessages) : 0;
  const processedAttachments = Number.isFinite(metrics.processedAttachments) ? Number(metrics.processedAttachments) : 0;
  const skippedAttachments = Number.isFinite(metrics.skippedAttachments) ? Number(metrics.skippedAttachments) : 0;
  const errorCount = Number.isFinite(metrics.errorCount) ? Number(metrics.errorCount) : 0;
  const durationMs = Number.isFinite(metrics.durationMs) ? Number(metrics.durationMs) : null;

  return {
    processedMessages,
    processedAttachments,
    skippedAttachments,
    errorCount,
    durationMs,
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
    reason: sanitiseOptionalString(remoteSource.reason),
  };

  if (provider === 'gmail') {
    entry.messageId = sanitiseOptionalString(remoteSource.messageId);
    entry.threadId = sanitiseOptionalString(remoteSource.threadId);
    entry.attachmentId = sanitiseOptionalString(remoteSource.attachmentId);
    entry.filename = sanitiseOptionalString(remoteSource.filename);
    entry.mimeType = sanitiseOptionalString(remoteSource.mimeType);
    entry.historyId = sanitiseOptionalString(remoteSource.historyId);
    entry.subject = sanitiseOptionalString(remoteSource.subject);
    entry.from = sanitiseOptionalString(remoteSource.from);
    entry.snippet = sanitiseOptionalString(remoteSource.snippet);
    entry.receivedAt = sanitiseOptionalString(remoteSource.receivedAt);
    entry.email = sanitiseOptionalString(remoteSource.email);
    const attachmentSize = remoteSource.attachmentSize;
    entry.attachmentSize = typeof attachmentSize === 'number' && Number.isFinite(attachmentSize) ? attachmentSize : null;

    if (Array.isArray(remoteSource.labelIds)) {
      const labels = remoteSource.labelIds
        .map((value) => sanitiseOptionalString(value))
        .filter(Boolean);
      entry.labelIds = labels.length ? labels : null;
    } else {
      entry.labelIds = null;
    }

    if (!entry.itemId && entry.messageId && entry.attachmentId) {
      entry.itemId = `${entry.messageId}::${entry.attachmentId}`;
    }
  }

  if (!entry.itemId && !entry.webUrl && !entry.path && !entry.messageId) {
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

  if (existing.provider === 'gmail') {
    if (existing.itemId && candidate.itemId && existing.itemId === candidate.itemId) {
      return true;
    }
    if (existing.messageId && candidate.messageId && existing.messageId === candidate.messageId) {
      if (!existing.attachmentId || !candidate.attachmentId) {
        return true;
      }
      return existing.attachmentId === candidate.attachmentId;
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

function identifyMissingQuickBooksMetadata(metadata) {
  const missing = [];
  if (!metadata?.vendors || !Array.isArray(metadata.vendors.items) || !metadata.vendors.items.length) {
    missing.push('vendors');
  }
  if (!metadata?.accounts || !Array.isArray(metadata.accounts.items) || !metadata.accounts.items.length) {
    missing.push('accounts');
  }
  if (!metadata?.taxCodes || !Array.isArray(metadata.taxCodes.items) || !metadata.taxCodes.items.length) {
    missing.push('taxCodes');
  }
  return missing;
}

function buildQuickBooksInvoiceUrl(realmId) {
  const base = (QUICKBOOKS_API_BASE_URL || '').replace(/\/+$/, '');
  const encodedRealmId = encodeURIComponent(realmId);
  const url = `${base}/v3/company/${encodedRealmId}/bill`;
  return `${url}?minorversion=${encodeURIComponent(QUICKBOOKS_MINOR_VERSION)}`;
}

function buildQuickBooksInvoicePayload(invoice, { metadata, realmId }) {
  if (!invoice) {
    throw createPreviewValidationError('Invoice is not available for preview.');
  }

  if (!realmId) {
    throw createPreviewValidationError('QuickBooks realmId is required for preview.');
  }

  const vendorId = normaliseNullableId(invoice?.reviewSelection?.vendorId);
  if (!vendorId) {
    throw createPreviewValidationError('Select a QuickBooks vendor before previewing.', { missing: 'vendorId' });
  }

  const vendorLookup = buildQuickBooksLookup(metadata?.vendors?.items);
  const vendor = vendorLookup.get(vendorId);
  if (!vendor) {
    throw createPreviewValidationError(
      'Selected QuickBooks vendor is no longer available. Refresh QuickBooks data and try again.',
      { missing: 'vendor', vendorId }
    );
  }

  const vendorSettings = metadata?.vendorSettings || {};
  const vendorDefaults = vendorSettings[vendorId] || {};
  const accountId = normaliseNullableId(invoice?.reviewSelection?.accountId) || normaliseNullableId(vendorDefaults.accountId);
  if (!accountId) {
    throw createPreviewValidationError('Select a QuickBooks account before previewing.', { missing: 'accountId' });
  }

  const accountLookup = buildQuickBooksLookup(metadata?.accounts?.items);
  const account = accountLookup.get(accountId);
  if (!account) {
    throw createPreviewValidationError(
      'Selected QuickBooks account is no longer available. Refresh QuickBooks data and try again.',
      { missing: 'account', accountId }
    );
  }

  const taxLookup = buildQuickBooksLookup(metadata?.taxCodes?.items);
  const taxResolution = resolveQuickBooksTaxCodes(invoice, taxLookup, vendorDefaults);
  const primaryTaxCodeId = taxResolution.primaryTaxCodeId;

  if (!primaryTaxCodeId) {
    throw createPreviewValidationError('Add a QuickBooks tax code for this vendor before previewing.', {
      missing: 'taxCodeId',
      vendorId,
    });
  }

  if (!taxLookup.has(primaryTaxCodeId)) {
    throw createPreviewValidationError(
      'Selected QuickBooks tax code is no longer available. Refresh QuickBooks data and try again.',
      { missing: 'taxCode', taxCodeId: primaryTaxCodeId }
    );
  }

  const secondaryTaxCodeId =
    taxResolution.secondaryTaxCodeId && taxLookup.has(taxResolution.secondaryTaxCodeId)
      ? taxResolution.secondaryTaxCodeId
      : null;

  if (taxResolution.secondaryTaxCodeId && !secondaryTaxCodeId) {
    console.warn('Secondary QuickBooks tax code selection is no longer available in metadata.', {
      requestedTaxCodeId: taxResolution.secondaryTaxCodeId,
      invoiceChecksum: invoice?.metadata?.checksum || null,
    });
  }

  const expenseLinesResult = buildQuickBooksExpenseLines(invoice, {
    account,
    taxLookup,
    taxResolution: { ...taxResolution, secondaryTaxCodeId },
  });

  const { lines, vatBuckets, usedTaxCodeIds, requiresSecondaryTaxCode } = expenseLinesResult;

  if (!lines.length) {
    throw createPreviewValidationError('Invoice does not contain any line items with amounts to preview.', {
      missing: 'lines',
    });
  }

  if (requiresSecondaryTaxCode) {
    throw createPreviewValidationError('Assign zero-rated VAT code for split invoice before previewing.', {
      missing: 'secondaryTaxCodeId',
      vendorId,
    });
  }

  const totalAmount = sumLineAmounts(lines);
  if (totalAmount === null) {
    throw createPreviewValidationError('Unable to determine invoice total for preview.', {
      missing: 'totalAmount',
    });
  }

  const vendorRef = compactObject({
    value: vendor.id,
    name: sanitiseOptionalString(vendor.displayName || vendor.name),
  });

  const currency = sanitiseCurrencyCode(invoice?.data?.currency);
  const currencyRef = currency ? compactObject({ value: currency }) : null;

  const hasMultipleTaxCodes = usedTaxCodeIds && usedTaxCodeIds.size > 1;
  const invoiceLevelTaxCodeId = !hasMultipleTaxCodes
    ? vatBuckets && vatBuckets.length ? vatBuckets[0].taxCodeId || primaryTaxCodeId : primaryTaxCodeId
    : null;
  const taxDetail = invoiceLevelTaxCodeId
    ? compactObject({ TxnTaxCodeRef: compactObject({ value: invoiceLevelTaxCodeId }) })
    : null;

  const payload = compactObject({
    VendorRef: vendorRef,
    DocNumber: sanitiseDocNumberForQuickBooks(invoice?.data?.invoiceNumber),
    TxnDate: normaliseInvoiceDateForQuickBooks(invoice?.data?.invoiceDate),
    PrivateNote: buildPreviewPrivateNote(invoice),
    CurrencyRef: currencyRef,
    Line: lines,
    TxnTaxDetail: taxDetail,
    TotalAmt: totalAmount,
  });

  if (payload.CurrencyRef) {
    payload.CurrencyRef = compactObject(payload.CurrencyRef);
  }

  if (payload.TxnTaxDetail) {
    payload.TxnTaxDetail = compactObject({
      TxnTaxCodeRef: compactObject(payload.TxnTaxDetail.TxnTaxCodeRef),
    });
  }

  return payload;
}

function resolveQuickBooksTaxCodes(invoice, taxLookup, vendorDefaults) {
  const candidateFromSelection = normaliseNullableId(invoice?.reviewSelection?.taxCodeId);
  const candidateFromDefaults = normaliseNullableId(vendorDefaults?.taxCodeId);
  const candidateLabels = collectCandidateTaxCodeLabels(invoice);
  const matchFromLabels = findTaxCodeIdFromLabels(candidateLabels, taxLookup);

  const primaryCandidates = [candidateFromSelection, candidateFromDefaults, matchFromLabels].filter(Boolean);

  let primaryTaxCodeId = null;
  for (const candidate of primaryCandidates) {
    if (candidate && taxLookup?.has(candidate)) {
      primaryTaxCodeId = candidate;
      break;
    }
  }

  if (!primaryTaxCodeId) {
    primaryTaxCodeId = primaryCandidates.find(Boolean) || null;
  }

  const secondaryFromSelection = normaliseNullableId(invoice?.reviewSelection?.secondaryTaxCodeId);
  let secondaryTaxCodeId =
    secondaryFromSelection && taxLookup?.has(secondaryFromSelection) ? secondaryFromSelection : null;

  const vatAnalysis = analyseInvoiceVatSignals(invoice, taxLookup);

  if (!secondaryTaxCodeId) {
    const productTaxCodeCandidates = vatAnalysis.signals
      .map((signal) => signal.taxCodeId)
      .filter((taxCodeId) => taxCodeId && taxCodeId !== primaryTaxCodeId);

    for (const candidate of productTaxCodeCandidates) {
      if (taxLookup?.has(candidate)) {
        secondaryTaxCodeId = candidate;
        break;
      }
    }
  }

  if (!secondaryTaxCodeId) {
    const zeroBucketPresent = vatAnalysis.signals.some((signal) => signal.bucket === 'zero');
    if (zeroBucketPresent) {
      const zeroCandidates = [];
      if (taxLookup?.size) {
        for (const [id, entry] of taxLookup.entries()) {
          if (id === primaryTaxCodeId) {
            continue;
          }
          if (classifyTaxCodeEntry(entry) === 'zero') {
            zeroCandidates.push(id);
          }
        }
      }
      if (zeroCandidates.length === 1) {
        secondaryTaxCodeId = zeroCandidates[0];
      }
    }
  }

  if (secondaryTaxCodeId === primaryTaxCodeId) {
    secondaryTaxCodeId = null;
  }

  return {
    primaryTaxCodeId,
    secondaryTaxCodeId,
    detectedVatBuckets: Array.from(vatAnalysis.detectedBucketKeys || []),
    distinctVatRates: Array.from(vatAnalysis.distinctRates || []),
  };
}

function collectCandidateTaxCodeLabels(invoice) {
  const candidates = [];
  const invoiceLevel = sanitiseOptionalString(invoice?.data?.taxCode);
  if (invoiceLevel) {
    candidates.push(invoiceLevel);
  }

  if (Array.isArray(invoice?.data?.products)) {
    invoice.data.products.forEach((product) => {
      const code = sanitiseOptionalString(product?.taxCode);
      if (code) {
        candidates.push(code);
      }
    });
  }

  return candidates;
}

function findTaxCodeIdFromLabels(candidates, taxLookup) {
  if (!Array.isArray(candidates) || !candidates.length || !taxLookup?.size) {
    return null;
  }

  const entries = [];
  for (const [id, entry] of taxLookup.entries()) {
    entries.push({
      id,
      name: normaliseComparableText(entry?.name),
      description: normaliseComparableText(entry?.description),
    });
  }

  for (const label of candidates) {
    const directId = normaliseNullableId(label);
    if (directId && taxLookup.has(directId)) {
      return directId;
    }

    const normalizedLabel = normaliseComparableText(label);
    if (!normalizedLabel) {
      continue;
    }

    const matched = entries.find(
      (entry) => entry.name === normalizedLabel || entry.description === normalizedLabel
    );
    if (matched) {
      return matched.id;
    }
  }

  return null;
}

function buildQuickBooksLookup(items) {
  const lookup = new Map();
  if (!Array.isArray(items)) {
    return lookup;
  }

  items.forEach((entry) => {
    const id = normaliseNullableId(entry?.id);
    if (id) {
      lookup.set(id, entry);
    }
  });

  return lookup;
}

// Aggregates Gemini product lines into at most two VAT buckets (standard and zero-rated) so the
// QuickBooks payload mirrors invoices that span two VAT rates. Bucket selection relies on
// product-level tax hints when available (explicit QuickBooks tax code IDs, Gemini tax labels, or
// numeric tax rates) and falls back to the primary vendor tax code when no hints exist. Any third
// rate is collapsed into the closest standard bucket and logged for follow-up. QuickBooks line
// entries are emitted with per-bucket TaxCodeRef values, enabling invoices with mixed VAT.
function buildQuickBooksExpenseLines(invoice, { account, taxLookup, taxResolution }) {
  const resolution = taxResolution || {};
  const lines = [];
  const usedTaxCodeIds = new Set();
  const vatBuckets = [];

  const accountId = normaliseNullableId(account?.id);
  const accountName = sanitiseOptionalString(account?.name || account?.fullyQualifiedName);
  const accountRef = compactObject({ value: accountId, name: accountName });

  const bucketTemplates = {
    standard: {
      key: 'standard',
      description: 'Standard VAT items',
      totalCents: 0,
      taxCodeId: null,
    },
    zero: {
      key: 'zero',
      description: 'Zero-rated items',
      totalCents: 0,
      taxCodeId: null,
    },
  };

  const availableBucketTaxCodes = determineAvailableBucketTaxCodes(resolution, taxLookup);
  if (availableBucketTaxCodes.standard) {
    bucketTemplates.standard.taxCodeId = availableBucketTaxCodes.standard;
  }
  if (availableBucketTaxCodes.zero) {
    bucketTemplates.zero.taxCodeId = availableBucketTaxCodes.zero;
  }

  const vatAnalysis = analyseInvoiceVatSignals(invoice, taxLookup);

  vatAnalysis.signals.forEach((signal) => {
    const bucketKey = signal.bucket === 'zero' ? 'zero' : 'standard';
    const bucket = bucketTemplates[bucketKey];
    bucket.totalCents += Math.round(signal.amount * 100);
    if (!bucket.taxCodeId && signal.taxCodeId && taxLookup?.has(signal.taxCodeId)) {
      bucket.taxCodeId = signal.taxCodeId;
    }
  });

  ['standard', 'zero'].forEach((bucketKey) => {
    const bucket = bucketTemplates[bucketKey];
    if (!bucket) {
      return;
    }

    const normalisedAmount = normaliseMoneyValue(bucket.totalCents / 100);
    if (normalisedAmount === null || normalisedAmount === 0) {
      return;
    }

    const preferredTaxCodeId =
      bucket.taxCodeId && taxLookup?.has(bucket.taxCodeId)
        ? bucket.taxCodeId
        : bucketKey === 'standard'
        ? resolution.primaryTaxCodeId
        : resolution.secondaryTaxCodeId;
    const taxCodeId = preferredTaxCodeId && taxLookup?.has(preferredTaxCodeId) ? preferredTaxCodeId : null;

    const detail = compactObject({
      AccountRef: accountRef,
      TaxCodeRef: taxCodeId ? compactObject({ value: taxCodeId }) : null,
    });

    const line = compactObject({
      DetailType: 'AccountBasedExpenseLineDetail',
      Amount: normalisedAmount,
      Description: bucket.description,
      AccountBasedExpenseLineDetail: detail,
    });

    if (!line.Description) {
      delete line.Description;
    }

    lines.push(line);

    if (taxCodeId) {
      usedTaxCodeIds.add(taxCodeId);
    }

    vatBuckets.push({
      key: bucket.key,
      description: bucket.description,
      amount: normalisedAmount,
      taxCodeId: taxCodeId || null,
    });
  });

  let requiresSecondaryTaxCode = false;
  const zeroBucket = vatBuckets.find((entry) => entry.key === 'zero' && entry.amount !== null && entry.amount !== 0);
  if (zeroBucket && !zeroBucket.taxCodeId) {
    requiresSecondaryTaxCode = true;
  }

  if (!lines.length) {
    const totalAmount = normaliseMoneyValue(invoice?.data?.totalAmount);
    const subtotalAmount = normaliseMoneyValue(invoice?.data?.subtotal);
    const fallbackAmount = totalAmount !== null ? totalAmount : subtotalAmount;
    if (fallbackAmount !== null) {
      const fallbackTaxCodeId =
        resolution.primaryTaxCodeId && taxLookup?.has(resolution.primaryTaxCodeId)
          ? resolution.primaryTaxCodeId
          : null;
      const detail = compactObject({
        AccountRef: accountRef,
        TaxCodeRef: fallbackTaxCodeId ? compactObject({ value: fallbackTaxCodeId }) : null,
      });
      const fallbackDescription =
        sanitiseOptionalString(invoice?.data?.vendor) ||
        sanitiseOptionalString(invoice?.metadata?.originalName) ||
        'Invoice total';

      const fallbackLine = compactObject({
        DetailType: 'AccountBasedExpenseLineDetail',
        Amount: fallbackAmount,
        Description: fallbackDescription ? fallbackDescription.slice(0, 4000) : null,
        AccountBasedExpenseLineDetail: detail,
      });

      if (!fallbackLine.Description) {
        delete fallbackLine.Description;
      }

      lines.push(fallbackLine);

      if (fallbackTaxCodeId) {
        usedTaxCodeIds.add(fallbackTaxCodeId);
      }

      vatBuckets.push({
        key: 'standard',
        description: fallbackLine.Description || 'Invoice total',
        amount: fallbackAmount,
        taxCodeId: fallbackTaxCodeId,
      });

      requiresSecondaryTaxCode = false;
    }
  }

  if (vatAnalysis.requiresSplit && zeroBucket && !zeroBucket.taxCodeId) {
    requiresSecondaryTaxCode = true;
  }

  if (vatAnalysis.distinctRates && vatAnalysis.distinctRates.size > 2) {
    console.warn('Detected more than two VAT rates; collapsing to standard/zero buckets.', {
      rates: Array.from(vatAnalysis.distinctRates).sort((a, b) => a - b),
      invoiceChecksum: invoice?.metadata?.checksum || null,
      invoiceName: invoice?.metadata?.originalName || null,
    });
  }

  return {
    lines,
    vatBuckets,
    usedTaxCodeIds,
    requiresSecondaryTaxCode,
    detectedVatBuckets: Array.from(vatAnalysis.detectedBucketKeys || []),
  };
}

function determineAvailableBucketTaxCodes(taxResolution, taxLookup) {
  const mapping = {};
  if (!taxResolution) {
    return mapping;
  }

  const register = (candidate) => {
    const id = normaliseNullableId(candidate);
    if (!id || !taxLookup?.has(id)) {
      return;
    }
    const entry = taxLookup.get(id);
    const bucket = classifyTaxCodeEntry(entry);
    if (!mapping[bucket]) {
      mapping[bucket] = id;
    }
  };

  register(taxResolution.primaryTaxCodeId);
  register(taxResolution.secondaryTaxCodeId);

  return mapping;
}

function analyseInvoiceVatSignals(invoice, taxLookup) {
  const products = Array.isArray(invoice?.data?.products) ? invoice.data.products : [];
  const signals = [];
  const detectedBucketKeys = new Set();
  const distinctRates = new Set();

  products.forEach((product) => {
    const amount = computeLineAmount(product);
    if (amount === null) {
      return;
    }

    const classification = classifyProductVatSignal(product, taxLookup);
    const bucketKey = classification.bucket === 'zero' ? 'zero' : 'standard';
    detectedBucketKeys.add(bucketKey);

    const rate = typeof classification.rate === 'number' && Number.isFinite(classification.rate)
      ? Number(Math.round(classification.rate * 1000) / 1000)
      : null;
    if (rate !== null) {
      distinctRates.add(rate);
    }

    signals.push({
      bucket: bucketKey,
      taxCodeId: classification.taxCodeId || null,
      amount,
      rate,
    });
  });

  const bucketTotals = signals.reduce(
    (acc, signal) => {
      const bucketKey = signal.bucket === 'zero' ? 'zero' : 'standard';
      acc[bucketKey] = (acc[bucketKey] || 0) + Math.round(signal.amount * 100);
      return acc;
    },
    { standard: 0, zero: 0 }
  );

  const requiresSplit = bucketTotals.standard !== 0 && bucketTotals.zero !== 0;

  return {
    signals,
    detectedBucketKeys,
    distinctRates,
    requiresSplit,
  };
}

function classifyProductVatSignal(product, taxLookup) {
  if (!product || typeof product !== 'object') {
    return { bucket: 'standard', taxCodeId: null, rate: null };
  }

  const directTaxCodeId =
    normaliseNullableId(product?.quickBooksTaxCodeId) || normaliseNullableId(product?.taxCode);
  if (directTaxCodeId && taxLookup?.has(directTaxCodeId)) {
    const entry = taxLookup.get(directTaxCodeId);
    const rate = extractRateFromTaxEntry(entry);
    return {
      bucket: classifyTaxCodeEntry(entry),
      taxCodeId: directTaxCodeId,
      rate,
    };
  }

  if (typeof product?.taxCode === 'string' && product.taxCode.trim()) {
    const matchedId = findTaxCodeIdFromLabels([product.taxCode], taxLookup);
    if (matchedId && taxLookup?.has(matchedId)) {
      const entry = taxLookup.get(matchedId);
      const rate = extractRateFromTaxEntry(entry);
      return {
        bucket: classifyTaxCodeEntry(entry),
        taxCodeId: matchedId,
        rate,
      };
    }
  }

  const inferredRate = extractVatRateFromProduct(product);
  if (inferredRate !== null) {
    return {
      bucket: classifyRateToBucket(inferredRate),
      taxCodeId: null,
      rate: inferredRate,
    };
  }

  if (typeof product?.taxCode === 'string' && product.taxCode.trim()) {
    const normalizedLabel = normaliseComparableText(product.taxCode);
    if (isZeroRateLabel(normalizedLabel)) {
      return { bucket: 'zero', taxCodeId: null, rate: 0 };
    }
    if (isStandardRateLabel(normalizedLabel)) {
      return { bucket: 'standard', taxCodeId: null, rate: null };
    }
  }

  return { bucket: 'standard', taxCodeId: null, rate: null };
}

function extractVatRateFromProduct(product) {
  const rawRate = product?.taxRate;
  if (typeof rawRate === 'number' && Number.isFinite(rawRate)) {
    return Number(Math.round(rawRate * 1000) / 1000);
  }
  if (typeof rawRate === 'string' && rawRate.trim()) {
    const parsed = Number.parseFloat(rawRate);
    if (Number.isFinite(parsed)) {
      return Number(Math.round(parsed * 1000) / 1000);
    }
  }

  if (typeof product?.taxCode === 'string' && product.taxCode.trim()) {
    const match = product.taxCode.match(/(\d+(?:\.\d+)?)\s*%/);
    if (match) {
      const rate = Number.parseFloat(match[1]);
      if (Number.isFinite(rate)) {
        return Number(Math.round(rate * 1000) / 1000);
      }
    }
  }

  return null;
}

function extractRateFromTaxEntry(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const candidates = [
    entry.rate,
    entry.RateValue,
    entry.rateValue,
    entry.taxRate,
    entry.TaxRate,
    entry?.SalesTaxRateList?.TaxRateDetail?.[0]?.RateValue,
    entry?.PurchaseTaxRateList?.TaxRateDetail?.[0]?.RateValue,
  ];

  for (const candidate of candidates) {
    if (candidate === null || candidate === undefined) {
      continue;
    }
    const numeric = Number(candidate);
    if (Number.isFinite(numeric)) {
      return Number(Math.round(numeric * 1000) / 1000);
    }
  }

  return null;
}

function classifyRateToBucket(rate) {
  if (!Number.isFinite(rate)) {
    return 'standard';
  }
  return Math.abs(rate) < 0.5 ? 'zero' : 'standard';
}

function classifyTaxCodeEntry(entry) {
  const rate = extractRateFromTaxEntry(entry);
  if (rate !== null) {
    return classifyRateToBucket(rate);
  }

  const name = normaliseComparableText(entry?.name);
  const description = normaliseComparableText(entry?.description);

  if (isZeroRateLabel(name) || isZeroRateLabel(description)) {
    return 'zero';
  }

  return 'standard';
}

function isZeroRateLabel(text) {
  if (!text) {
    return false;
  }
  return (
    text.includes('zero') ||
    text.includes('no tax') ||
    text.includes('tax free') ||
    text.includes('out of scope') ||
    text.includes('non tax') ||
    text.includes('nontax') ||
    text.includes('exempt') ||
    text.includes('0%') ||
    text.includes('0 %') ||
    text.split(' ').includes('0')
  );
}

function isStandardRateLabel(text) {
  if (!text) {
    return false;
  }
  return text.includes('standard') || text.includes('vat') || text.includes('taxable') || text.includes('20');
}

function computeLineAmount(product) {
  if (!product || typeof product !== 'object') {
    return null;
  }

  const direct = normaliseMoneyValue(product.lineTotal);
  if (direct !== null) {
    return direct;
  }

  const unitPrice = normaliseMoneyValue(product.unitPrice);
  const quantityRaw = product.quantity;
  let quantity = null;
  if (typeof quantityRaw === 'number' && Number.isFinite(quantityRaw)) {
    quantity = quantityRaw;
  } else if (quantityRaw !== null && quantityRaw !== undefined) {
    const numeric = Number(quantityRaw);
    if (Number.isFinite(numeric)) {
      quantity = numeric;
    }
  }

  if (unitPrice !== null && quantity !== null) {
    return normaliseMoneyValue(unitPrice * quantity);
  }

  return null;
}

function normaliseMoneyValue(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return null;
  }

  const rounded = Math.round(numeric * 100) / 100;
  if (!Number.isFinite(rounded)) {
    return null;
  }

  return Number(rounded.toFixed(2));
}

function sumLineAmounts(lines) {
  if (!Array.isArray(lines) || !lines.length) {
    return null;
  }

  let running = 0;
  let counted = 0;

  lines.forEach((line) => {
    const amount = normaliseMoneyValue(line?.Amount);
    if (amount !== null) {
      running += amount;
      counted += 1;
    }
  });

  if (!counted) {
    return null;
  }

  return normaliseMoneyValue(running);
}

function normaliseInvoiceDateForQuickBooks(value) {
  const text = sanitiseOptionalString(value);
  if (!text) {
    return null;
  }

  const isoMatch = text.match(/\d{4}-\d{2}-\d{2}/);
  if (isoMatch) {
    return isoMatch[0];
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return null;
}

function sanitiseDocNumberForQuickBooks(value) {
  const text = sanitiseOptionalString(value);
  if (!text) {
    return null;
  }

  return text.slice(0, 21);
}

function buildPreviewPrivateNote(invoice) {
  const parts = [];
  const originalName = sanitiseOptionalString(invoice?.metadata?.originalName);
  if (originalName) {
    parts.push(`Source file: ${originalName}`);
  }

  const provider = sanitiseOptionalString(invoice?.metadata?.remoteSource?.provider);
  if (provider) {
    parts.push(`Imported via ${provider}`);
  }

  if (!parts.length) {
    return null;
  }

  return `Preview only. ${parts.join(' | ')}`.slice(0, 4000);
}

function compactObject(object) {
  if (!object || typeof object !== 'object') {
    return object;
  }

  if (Array.isArray(object)) {
    return object;
  }

  const result = {};
  Object.entries(object).forEach(([key, value]) => {
    if (value !== null && value !== undefined) {
      result[key] = value;
    }
  });
  return result;
}

function createPreviewValidationError(message, details) {
  const error = new Error(message);
  error.status = 422;
  error.code = 'PREVIEW_VALIDATION_FAILED';
  if (details) {
    error.details = details;
  }
  return error;
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

async function updateStoredInvoiceReviewSelection(checksum, updates = {}) {
  if (!checksum) {
    return null;
  }

  const invoices = await readStoredInvoices();
  const index = invoices.findIndex((entry) => entry?.metadata?.checksum === checksum);
  if (index === -1) {
    return null;
  }

  const hasVendorUpdate = Object.prototype.hasOwnProperty.call(updates, 'vendorId');
  const hasAccountUpdate = Object.prototype.hasOwnProperty.call(updates, 'accountId');
  const hasTaxUpdate = Object.prototype.hasOwnProperty.call(updates, 'taxCodeId');
  const hasSecondaryTaxUpdate = Object.prototype.hasOwnProperty.call(
    updates,
    'secondaryTaxCodeId'
  );

  if (!hasVendorUpdate && !hasAccountUpdate && !hasTaxUpdate && !hasSecondaryTaxUpdate) {
    return invoices[index];
  }

  const currentSelection =
    invoices[index] && typeof invoices[index].reviewSelection === 'object'
      ? { ...invoices[index].reviewSelection }
      : {};

  if (hasVendorUpdate) {
    currentSelection.vendorId = normaliseNullableId(updates.vendorId);
  }

  if (hasAccountUpdate) {
    currentSelection.accountId = normaliseNullableId(updates.accountId);
  }

  if (hasTaxUpdate) {
    currentSelection.taxCodeId = normaliseNullableId(updates.taxCodeId);
  }

  if (hasSecondaryTaxUpdate) {
    currentSelection.secondaryTaxCodeId = normaliseNullableId(updates.secondaryTaxCodeId);
  }

  const normalizedSelection = normalizeReviewSelection(currentSelection);

  const updatedInvoice = {
    ...invoices[index],
    reviewSelection: normalizedSelection,
  };

  invoices[index] = updatedInvoice;
  await writeStoredInvoices(invoices);
  return updatedInvoice;
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
  const reviewSelection = normalizeReviewSelection(entry.reviewSelection);
  return {
    ...entry,
    status,
    reviewSelection,
  };
}

function normalizeReviewSelection(selection) {
  if (!selection || typeof selection !== 'object') {
    return null;
  }

  const vendorId = normaliseNullableId(selection.vendorId);
  const accountId = normaliseNullableId(selection.accountId);
  const taxCodeId = normaliseNullableId(selection.taxCodeId);
  const secondaryTaxCodeId = normaliseNullableId(selection.secondaryTaxCodeId);

  if (!vendorId && !accountId && !taxCodeId && !secondaryTaxCodeId) {
    return null;
  }

  const result = {};
  if (vendorId) {
    result.vendorId = vendorId;
  }
  if (accountId) {
    result.accountId = accountId;
  }
  if (taxCodeId) {
    result.taxCodeId = taxCodeId;
  }
  if (secondaryTaxCodeId) {
    result.secondaryTaxCodeId = secondaryTaxCodeId;
  }

  return result;
}

function normaliseNullableId(value) {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed ? trimmed : null;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return String(value);
  }

  if (typeof value === 'bigint') {
    return value.toString();
  }

  if (value === null) {
    return null;
  }

  return null;
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

function decodeBase64Url(value) {
  if (typeof value !== 'string' || !value) {
    return Buffer.alloc(0);
  }

  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padLength = normalized.length % 4;
  const padded = padLength ? normalized + '='.repeat(4 - padLength) : normalized;
  return Buffer.from(padded, 'base64');
}

function parseDelimitedList(value) {
  if (value === undefined || value === null) {
    return [];
  }

  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === 'string') {
          return entry.trim();
        }
        return String(entry || '').trim();
      })
      .filter(Boolean);
  }

  const text = String(value).replace(/[;\r\n]+/g, ',');
  return text
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean);
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
