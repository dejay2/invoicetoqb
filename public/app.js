const globalStatus = document.getElementById('global-status');
const companyDashboard = document.getElementById('company-dashboard');
const companySelect = document.getElementById('company-select');
const addCompanyButton = document.getElementById('add-company');
const tabButtons = document.querySelectorAll('[data-company-tab]');
const companyPanels = document.querySelectorAll('.company-panel');
const connectionStatus = document.getElementById('connection-status');
const connectCompanyButton = document.getElementById('connect-company');
const refreshMetadataButton = document.getElementById('refresh-metadata');
const businessProfileForm = document.getElementById('business-profile-form');
const businessTypeInput = document.getElementById('company-business-type');
const businessProfileSubmit = document.getElementById('business-profile-submit');
const vendorList = document.getElementById('vendor-list');
const accountList = document.getElementById('account-list');
const importVendorDefaultsButton = document.getElementById('import-vendor-defaults');
const reviewTableBody = document.getElementById('review-table-body');
const reviewSelectAllCheckbox = document.getElementById('review-select-all');
const reviewBulkActions = document.getElementById('review-bulk-actions');
const reviewBulkPreviewButton = document.getElementById('review-bulk-preview');
const reviewBulkEditButton = document.getElementById('review-bulk-edit');
const reviewBulkSaveButton = document.getElementById('review-bulk-save');
const reviewBulkCancelButton = document.getElementById('review-bulk-cancel');
const reviewBulkArchiveButton = document.getElementById('review-bulk-archive');
const reviewBulkDeleteButton = document.getElementById('review-bulk-delete');
const reviewSelectionCount = document.getElementById('review-selection-count');
const archiveTableBody = document.getElementById('archive-table-body');
const dropZone = document.getElementById('invoice-drop-zone');
const uploadInput = document.getElementById('invoice-upload-input');
const uploadStatusList = document.getElementById('upload-status-list');
const oneDriveForm = document.getElementById('onedrive-settings-form');
const oneDriveSharedSummaryText = document.getElementById('onedrive-shared-summary-text');
const oneDriveSharedStatusBadge = document.getElementById('onedrive-shared-status');
const oneDriveSharedConnectButton = document.getElementById('onedrive-shared-connect');
const oneDriveSharedValidateButton = document.getElementById('onedrive-shared-validate');
const oneDriveSharedLastValidated = document.getElementById('onedrive-shared-last-validated');
const oneDriveSharedLastResult = document.getElementById('onedrive-shared-last-result');
const oneDriveFolderIdInput = document.getElementById('onedrive-folder-id');
const oneDriveFolderPathInput = document.getElementById('onedrive-folder-path');
const oneDriveFolderNameInput = document.getElementById('onedrive-folder-name');
const oneDriveFolderWebUrlInput = document.getElementById('onedrive-folder-weburl');
const oneDriveFolderParentIdInput = document.getElementById('onedrive-folder-parent-id');
const oneDriveEnabledInput = document.getElementById('onedrive-enabled');
const oneDriveStatusContainer = document.getElementById('onedrive-status');
const oneDriveStatusState = document.getElementById('onedrive-status-state');
const oneDriveStatusFolder = document.getElementById('onedrive-status-folder');
const oneDriveStatusLastSync = document.getElementById('onedrive-status-last-sync');
const oneDriveStatusResult = document.getElementById('onedrive-status-result');
const oneDriveStatusLog = document.getElementById('onedrive-status-log');
const oneDriveSyncButton = document.getElementById('onedrive-sync-now');
const oneDriveResyncButton = document.getElementById('onedrive-resync');
const oneDriveClearButton = document.getElementById('onedrive-clear');
const oneDriveSaveButton = document.getElementById('onedrive-save');
const oneDriveBrowseMonitoredButton = document.getElementById('onedrive-browse-monitored');
const oneDriveBrowseProcessedButton = document.getElementById('onedrive-browse-processed');
const oneDriveProcessedClearButton = document.getElementById('onedrive-processed-clear');
const oneDriveMonitoredSummary = document.getElementById('onedrive-monitored-summary');
const oneDriveProcessedSummary = document.getElementById('onedrive-processed-summary');
const oneDriveProcessedFolderIdInput = document.getElementById('onedrive-processed-folder-id');
const oneDriveProcessedFolderPathInput = document.getElementById('onedrive-processed-folder-path');
const oneDriveProcessedFolderNameInput = document.getElementById('onedrive-processed-folder-name');
const oneDriveProcessedFolderWebUrlInput = document.getElementById('onedrive-processed-folder-weburl');
const oneDriveProcessedFolderParentIdInput = document.getElementById('onedrive-processed-folder-parent-id');
const oneDriveBrowseModal = document.getElementById('onedrive-browse-modal');
const oneDriveBrowseList = document.getElementById('onedrive-browse-list');
const oneDriveBrowseBackButton = document.getElementById('onedrive-browse-back');
const oneDriveBrowseConfirmButton = document.getElementById('onedrive-browse-confirm');
const oneDriveBrowseCancelButton = document.getElementById('onedrive-browse-cancel');
const oneDriveBrowseCloseButton = document.getElementById('onedrive-browse-close');
const oneDriveBrowseWarning = document.getElementById('onedrive-browse-warning');
const oneDriveBrowseStatus = document.getElementById('onedrive-browse-status');
const oneDriveBrowsePath = document.getElementById('onedrive-browse-path');
const oneDriveBrowseTitle = document.getElementById('onedrive-browse-title');
const oneDriveBrowseModalOverlay = oneDriveBrowseModal?.querySelector('[data-modal-dismiss]') || null;
const outlookSharedCard = document.querySelector('.outlook-shared-card');
const outlookSharedEditButton = document.getElementById('outlook-shared-edit');
const outlookSharedForm = document.getElementById('outlook-shared-form');
const outlookSharedMailboxInput = document.getElementById('outlook-shared-mailbox-input');
const outlookSharedDisplayNameInput = document.getElementById('outlook-shared-display-name');
const outlookSharedBrowseBaseButton = document.getElementById('outlook-shared-browse-base');
const outlookSharedSaveButton = document.getElementById('outlook-shared-save');
const outlookSharedCancelButton = document.getElementById('outlook-shared-cancel');
const outlookSharedSummaryText = document.getElementById('outlook-shared-summary-text');
const outlookSharedStatus = document.getElementById('outlook-shared-status');
const outlookSharedMailbox = document.getElementById('outlook-shared-mailbox');
const outlookSharedBaseFolder = document.getElementById('outlook-shared-base-folder');
const outlookSharedLastValidated = document.getElementById('outlook-shared-last-validated');
const outlookSharedBaseIdInput = document.getElementById('outlook-shared-base-id');
const outlookSharedBasePathInput = document.getElementById('outlook-shared-base-path');
const outlookSharedBaseNameInput = document.getElementById('outlook-shared-base-name');
const outlookSharedBaseWebUrlInput = document.getElementById('outlook-shared-base-weburl');

const outlookForm = document.getElementById('outlook-settings-form');
const outlookPollIntervalInput = document.getElementById('outlook-poll-interval');
const outlookMaxAttachmentInput = document.getElementById('outlook-max-attachment');
const outlookMimeTypesInput = document.getElementById('outlook-mime-types');
const outlookEnabledInput = document.getElementById('outlook-enabled');
const outlookBrowseButton = document.getElementById('outlook-browse-folder');
const outlookClearFolderButton = document.getElementById('outlook-clear-folder');
const outlookSaveButton = document.getElementById('outlook-save');
const outlookSyncButton = document.getElementById('outlook-sync-now');
const outlookResyncButton = document.getElementById('outlook-resync');
const outlookDisconnectButton = document.getElementById('outlook-disconnect');
const outlookFolderSummary = document.getElementById('outlook-folder-summary');
const outlookStatusContainer = document.getElementById('outlook-status');
const outlookStatusState = document.getElementById('outlook-status-state');
const outlookStatusFolder = document.getElementById('outlook-status-folder');
const outlookStatusLastSync = document.getElementById('outlook-status-last-sync');
const outlookStatusResult = document.getElementById('outlook-status-result');
const outlookFolderIdInput = document.getElementById('outlook-folder-id');
const outlookFolderPathInput = document.getElementById('outlook-folder-path');
const outlookFolderNameInput = document.getElementById('outlook-folder-name');
const outlookFolderWebUrlInput = document.getElementById('outlook-folder-weburl');
const outlookFolderParentIdInput = document.getElementById('outlook-folder-parent-id');

const outlookBrowseModal = document.getElementById('outlook-browse-modal');
const outlookBrowseTitle = document.getElementById('outlook-browse-title');
const outlookBrowseList = document.getElementById('outlook-browse-list');
const outlookBrowsePath = document.getElementById('outlook-browse-path');
const outlookBrowseBackButton = document.getElementById('outlook-browse-back');
const outlookBrowseCancelButton = document.getElementById('outlook-browse-cancel');
const outlookBrowseConfirmButton = document.getElementById('outlook-browse-confirm');
const outlookBrowseCloseButton = document.getElementById('outlook-browse-close');
const outlookBrowseWarning = document.getElementById('outlook-browse-warning');
const outlookBrowseStatus = document.getElementById('outlook-browse-status');
const outlookBrowseModalOverlay = outlookBrowseModal?.querySelector('.modal-overlay');
const qbPreviewModal = document.getElementById('qb-preview-modal');
const qbPreviewMethod = document.getElementById('qb-preview-method');
const qbPreviewSummary = document.getElementById('qb-preview-summary');
const qbPreviewJson = document.getElementById('qb-preview-json');
const qbPreviewCopyButton = document.getElementById('qb-preview-copy');
const qbPreviewCloseButton = document.getElementById('qb-preview-close');
const qbPreviewDismissButton = document.getElementById('qb-preview-dismiss');
const qbPreviewDismissOverlay = qbPreviewModal?.querySelector('[data-modal-dismiss]') || null;

let quickBooksCompanies = [];
let selectedRealmId = '';
const companyMetadataCache = new Map();
const metadataRequests = new Map();
let storedInvoices = [];
const reviewSelectedChecksums = new Set();
const reviewEditingChecksums = new Set();
let lastQuickBooksPreviewPayload = '';
let quickBooksPreviewEscapeHandler = null;
let invoicePreviewWindow = null;
let invoicePdfWindow = null;
let invoicePdfObjectUrl = null;
const MONITORED_DEFAULT_SUMMARY = 'No folder selected.';
const PROCESSED_DEFAULT_SUMMARY =
  'Processed invoices will move into a “Synced” folder inside the monitored location.';
let oneDriveBrowseState = null;
let lastOneDriveBrowseTrigger = null;
let oneDriveBrowseEscapeHandler = null;
let sharedOneDriveSettings = null;
let sharedOutlookSettings = null;
let outlookBrowseState = null;
let lastOutlookBrowseTrigger = null;
let outlookBrowseEscapeHandler = null;

const autoMatchInFlight = new Set();
const autoMatchedChecksums = new Set();

const MATCH_CLASS_BY_TYPE = {
  exact: 'match--exact',
  uncertain: 'match--review',
  unknown: 'match--none',
};

function resolveMatchClass(matchType, hasSelection) {
  if (!hasSelection) {
    return MATCH_CLASS_BY_TYPE.unknown;
  }

  return MATCH_CLASS_BY_TYPE[matchType] || MATCH_CLASS_BY_TYPE.unknown;
}

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

const VENDOR_VAT_TREATMENT_VALUES = new Set(['inclusive', 'exclusive', 'no_vat']);
const VAT_TREATMENT_OPTIONS = [
  { value: '', label: 'No default' },
  { value: 'inclusive', label: 'Inclusive of Tax' },
  { value: 'exclusive', label: 'Exclusive of Tax' },
  { value: 'no_vat', label: 'No VAT' },
];

bootstrap();

function bootstrap() {
  attachEventListeners();
  handleQuickBooksCallback();
  refreshSharedOneDriveSettings();
  refreshSharedOutlookSettings({ silent: true });
  renderBusinessProfile();
  renderOneDriveSettings();
  renderOutlookSettings();
  refreshQuickBooksCompanies();
  loadStoredInvoices();
}

async function refreshSharedOneDriveSettings() {
  if (!oneDriveSharedSummaryText) {
    sharedOneDriveSettings = null;
    return;
  }

  try {
    const response = await fetch('/api/onedrive/settings');
    if (!response.ok) {
      throw new Error('Unable to load shared settings.');
    }
    const payload = await response.json().catch(() => ({}));
    sharedOneDriveSettings = payload?.settings || null;
  } catch (error) {
    console.error('Failed to load shared OneDrive settings', error);
    sharedOneDriveSettings = null;
  } finally {
    renderSharedOneDriveSummary();
    renderOneDriveSettings();
  }
}

function renderSharedOneDriveSummary() {
  const settings = sharedOneDriveSettings || null;
  const configured = isSharedOneDriveConfigured();

  if (oneDriveSharedStatusBadge) {
    setSharedOneDriveStatusBadge(settings?.status || (configured ? 'ready' : 'unconfigured'));
  }

  if (oneDriveSharedConnectButton) {
    oneDriveSharedConnectButton.textContent = configured
      ? 'Switch OneDrive account'
      : 'Connect OneDrive account';
  }

  if (oneDriveSharedValidateButton) {
    oneDriveSharedValidateButton.disabled = !configured;
    oneDriveSharedValidateButton.textContent = 'Run connection check';
  }

  if (oneDriveSharedSummaryText) {
    oneDriveSharedSummaryText.textContent = '';

    if (!configured) {
      oneDriveSharedSummaryText.textContent =
        'Connect the OneDrive account that stores your invoice folders. Each company can pick its own monitored and processed folders below.';
    } else {
      const label = settings.driveName || settings.driveId;
      const summaryParts = [`Connected to OneDrive account: ${label}`];
      const breadcrumb = formatOneDriveBreadcrumb(deriveOneDriveBreadcrumbFromPath(settings.folderPath || ''));
      const folderLabels = [];
      if (settings.folderName) {
        folderLabels.push(settings.folderName);
      }
      if (breadcrumb) {
        folderLabels.push(breadcrumb);
      }
      if (!folderLabels.length && settings.folderId) {
        folderLabels.push(settings.folderId);
      }

      if (folderLabels.length) {
        summaryParts.push(`Company browsing is limited to ${folderLabels.join(' • ')}`);
      } else {
        summaryParts.push('Use the selectors below to choose folders per company');
      }

      const text = `${summaryParts.join('. ')}.`;
      oneDriveSharedSummaryText.appendChild(document.createTextNode(text));

      const targetUrl = settings.folderWebUrl || settings.driveWebUrl || settings.shareUrl;
      if (targetUrl) {
        oneDriveSharedSummaryText.appendChild(document.createTextNode(' '));
        const link = document.createElement('a');
        link.href = targetUrl;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'Open in OneDrive';
        oneDriveSharedSummaryText.appendChild(link);
      }

      if (settings.status && settings.status !== 'ready') {
        oneDriveSharedSummaryText.appendChild(document.createTextNode(` Status: ${settings.status}.`));
      }
    }
  }

  if (oneDriveSharedLastValidated) {
    const { lastValidatedAt } = settings || {};
    oneDriveSharedLastValidated.textContent = lastValidatedAt ? formatTimestamp(lastValidatedAt) : 'Never';
  }

  if (oneDriveSharedLastResult) {
    oneDriveSharedLastResult.textContent = formatSharedOneDriveResult(settings);
  }
}

function isSharedOneDriveConfigured() {
  return Boolean(sharedOneDriveSettings?.driveId);
}

function setSharedOneDriveStatusBadge(status) {
  if (!oneDriveSharedStatusBadge) {
    return;
  }

  const value = typeof status === 'string' ? status.trim().toLowerCase() : '';
  const label = formatStatusLabel(value, 'Not connected');
  oneDriveSharedStatusBadge.textContent = label;
  oneDriveSharedStatusBadge.dataset.status = value || 'unconfigured';

  const variants = ['status-pill--ready', 'status-pill--error', 'status-pill--warning', 'status-pill--muted'];
  variants.forEach((variant) => oneDriveSharedStatusBadge.classList.remove(variant));

  if (value === 'ready') {
    oneDriveSharedStatusBadge.classList.add('status-pill--ready');
  } else if (value === 'error') {
    oneDriveSharedStatusBadge.classList.add('status-pill--error');
  } else if (value === 'warning' || value === 'pending' || value === 'validating') {
    oneDriveSharedStatusBadge.classList.add('status-pill--warning');
  } else {
    oneDriveSharedStatusBadge.classList.add('status-pill--muted');
  }
}

function formatSharedOneDriveResult(settings) {
  if (!settings || !settings.driveId) {
    return 'Connect the OneDrive account to enable company folder browsing.';
  }

  const error = settings.lastValidationError;
  if (error?.message) {
    return error.message;
  }

  const health = settings.lastSyncHealth;
  if (health) {
    const statusLabel = formatStatusLabel(health.status, 'OK');
    if (health.message) {
      return `${statusLabel}: ${health.message}`;
    }
    return statusLabel;
  }

  if (settings.status === 'ready') {
    return 'Ready for company folder selection.';
  }

  return formatStatusLabel(settings.status, 'Unknown');
}

function attachEventListeners() {
  if (companySelect) {
    companySelect.addEventListener('change', () => {
      applySelection(companySelect.value, { triggerMetadata: true });
    });
  }

  if (addCompanyButton) {
    addCompanyButton.addEventListener('click', () => {
      window.location.href = '/api/quickbooks/connect';
    });
  }

  if (connectCompanyButton) {
    connectCompanyButton.addEventListener('click', () => {
      if (!selectedRealmId) {
        return;
      }
      window.location.href = '/api/quickbooks/connect';
    });
  }

  if (refreshMetadataButton) {
    refreshMetadataButton.addEventListener('click', () => {
      if (!selectedRealmId) {
        return;
      }
      refreshCompanyMetadata(selectedRealmId);
    });
  }

  if (businessProfileForm) {
    businessProfileForm.addEventListener('submit', handleBusinessProfileSave);
  }

  if (oneDriveForm) {
    oneDriveForm.addEventListener('submit', handleOneDriveSettingsSave);
  }

  if (oneDriveSharedConnectButton) {
    oneDriveSharedConnectButton.addEventListener('click', handleSharedOneDriveConnect);
  }

  if (oneDriveSharedValidateButton) {
    oneDriveSharedValidateButton.addEventListener('click', handleSharedOneDriveValidate);
  }

  if (oneDriveBrowseMonitoredButton) {
    oneDriveBrowseMonitoredButton.addEventListener('click', () => {
      if (!oneDriveBrowseMonitoredButton.disabled) {
        openOneDriveBrowser('monitored');
      }
    });
  }

  if (oneDriveBrowseProcessedButton) {
    oneDriveBrowseProcessedButton.addEventListener('click', () => {
      if (!oneDriveBrowseProcessedButton.disabled) {
        openOneDriveBrowser('processed');
      }
    });
  }

  if (oneDriveProcessedClearButton) {
    oneDriveProcessedClearButton.addEventListener('click', handleOneDriveProcessedClear);
  }

  if (oneDriveBrowseConfirmButton) {
    oneDriveBrowseConfirmButton.addEventListener('click', applyOneDriveBrowserSelection);
  }

  if (oneDriveBrowseCancelButton) {
    oneDriveBrowseCancelButton.addEventListener('click', closeOneDriveBrowser);
  }

  if (oneDriveBrowseCloseButton) {
    oneDriveBrowseCloseButton.addEventListener('click', closeOneDriveBrowser);
  }

  if (oneDriveBrowseModalOverlay) {
    oneDriveBrowseModalOverlay.addEventListener('click', closeOneDriveBrowser);
  }

  if (oneDriveBrowseBackButton) {
    oneDriveBrowseBackButton.addEventListener('click', handleOneDriveBrowseBack);
  }

  if (oneDriveBrowseList) {
    oneDriveBrowseList.addEventListener('click', handleOneDriveBrowseListClick);
    oneDriveBrowseList.addEventListener('dblclick', handleOneDriveBrowseListDoubleClick);
    oneDriveBrowseList.addEventListener('keydown', handleOneDriveBrowseListKeydown);
  }

  if (oneDriveSyncButton) {
    oneDriveSyncButton.addEventListener('click', handleOneDriveSyncClick);
  }

  if (oneDriveResyncButton) {
    oneDriveResyncButton.addEventListener('click', handleOneDriveFullResync);
  }

  if (oneDriveClearButton) {
    oneDriveClearButton.addEventListener('click', handleOneDriveDisconnect);
  }

  if (outlookSharedEditButton) {
    outlookSharedEditButton.addEventListener('click', handleOutlookSharedEdit);
  }

  if (outlookSharedForm) {
    outlookSharedForm.addEventListener('submit', handleOutlookSharedSave);
  }

  if (outlookSharedCancelButton) {
    outlookSharedCancelButton.addEventListener('click', handleOutlookSharedCancel);
  }

  if (outlookSharedBrowseBaseButton) {
    outlookSharedBrowseBaseButton.addEventListener('click', () => openOutlookBrowser('shared-base'));
  }

  if (outlookForm) {
    outlookForm.addEventListener('submit', handleOutlookSettingsSave);
  }

  if (outlookBrowseButton) {
    outlookBrowseButton.addEventListener('click', () => {
      if (!outlookBrowseButton.disabled) {
        openOutlookBrowser('company');
      }
    });
  }

  if (outlookClearFolderButton) {
    outlookClearFolderButton.addEventListener('click', handleOutlookClearFolder);
  }

  if (outlookSyncButton) {
    outlookSyncButton.addEventListener('click', handleOutlookSyncClick);
  }

  if (outlookResyncButton) {
    outlookResyncButton.addEventListener('click', handleOutlookResync);
  }

  if (outlookDisconnectButton) {
    outlookDisconnectButton.addEventListener('click', handleOutlookDisconnect);
  }

  if (outlookBrowseConfirmButton) {
    outlookBrowseConfirmButton.addEventListener('click', applyOutlookBrowserSelection);
  }

  if (outlookBrowseCancelButton) {
    outlookBrowseCancelButton.addEventListener('click', closeOutlookBrowser);
  }

  if (outlookBrowseCloseButton) {
    outlookBrowseCloseButton.addEventListener('click', closeOutlookBrowser);
  }

  if (outlookBrowseModalOverlay) {
    outlookBrowseModalOverlay.addEventListener('click', closeOutlookBrowser);
  }

  if (outlookBrowseBackButton) {
    outlookBrowseBackButton.addEventListener('click', handleOutlookBrowseBack);
  }

  if (outlookBrowseList) {
    outlookBrowseList.addEventListener('click', handleOutlookBrowseListClick);
    outlookBrowseList.addEventListener('dblclick', handleOutlookBrowseListDoubleClick);
    outlookBrowseList.addEventListener('keydown', handleOutlookBrowseListKeydown);
  }

  if (importVendorDefaultsButton) {
    importVendorDefaultsButton.addEventListener('click', handleImportVendorDefaults);
  }

  if (vendorList) {
    vendorList.addEventListener('change', handleVendorSettingChange);
  }

  if (dropZone) {
    dropZone.addEventListener('click', () =>
      uploadInput?.click()
    );
    dropZone.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        uploadInput?.click();
      }
    });
    ['dragenter', 'dragover'].forEach((eventName) => {
      dropZone.addEventListener(eventName, handleDropZoneDragOver);
    });
    ['dragleave', 'dragend'].forEach((eventName) => {
      dropZone.addEventListener(eventName, handleDropZoneDragLeave);
    });
    dropZone.addEventListener('drop', handleDropZoneDrop);
  }

  if (uploadInput) {
    uploadInput.addEventListener('change', (event) => {
      processSelectedFiles(event.target.files);
    });
  }

  if (archiveTableBody) {
    archiveTableBody.addEventListener('click', handleArchiveAction);
  }

  if (reviewTableBody) {
    reviewTableBody.addEventListener('click', handleReviewAction);
    reviewTableBody.addEventListener('change', handleReviewChange);
  }

  if (reviewSelectAllCheckbox) {
    reviewSelectAllCheckbox.addEventListener('change', handleReviewSelectAllChange);
  }

  if (reviewBulkPreviewButton) {
    reviewBulkPreviewButton.addEventListener('click', handleBulkPreviewSelected);
  }

  if (reviewBulkEditButton) {
    reviewBulkEditButton.addEventListener('click', handleBulkEditSelected);
  }

  if (reviewBulkSaveButton) {
    reviewBulkSaveButton.addEventListener('click', handleBulkSaveSelected);
  }

  if (reviewBulkCancelButton) {
    reviewBulkCancelButton.addEventListener('click', handleBulkCancelSelected);
  }

  if (reviewBulkArchiveButton) {
    reviewBulkArchiveButton.addEventListener('click', handleBulkArchiveSelected);
  }

  if (reviewBulkDeleteButton) {
    reviewBulkDeleteButton.addEventListener('click', handleBulkDeleteSelected);
  }

  if (qbPreviewCloseButton) {
    qbPreviewCloseButton.addEventListener('click', hideQuickBooksPreviewModal);
  }

  if (qbPreviewDismissButton) {
    qbPreviewDismissButton.addEventListener('click', hideQuickBooksPreviewModal);
  }

  if (qbPreviewDismissOverlay) {
    qbPreviewDismissOverlay.addEventListener('click', hideQuickBooksPreviewModal);
  }

  if (qbPreviewCopyButton) {
    qbPreviewCopyButton.addEventListener('click', copyQuickBooksPreviewPayload);
  }

  tabButtons.forEach((button) => {
    button.addEventListener('click', () => {
      activateCompanyTab(button.dataset.companyTab);
    });
  });

  document.addEventListener('dragover', handleDocumentDragOver);
  document.addEventListener('drop', handleDocumentDrop);
}

function handleQuickBooksCallback() {
  const params = new URLSearchParams(window.location.search);
  let shouldClear = false;

  if (params.has('quickbooks')) {
    const status = params.get('quickbooks');
    const companyName = params.get('company');
    shouldClear = true;

    if (status === 'connected') {
      const name = companyName || 'QuickBooks company';
      showStatus(globalStatus, `Connected to ${name}.`, 'success');
    } else if (status === 'error') {
      showStatus(globalStatus, 'QuickBooks connection failed. Please try again.', 'error');
    }
  }

  if (params.has('outlook')) {
    const status = params.get('outlook');
    const companyName = params.get('company');
    const message = params.get('message');
    shouldClear = true;

    if (status === 'connected') {
      const name = companyName || 'the selected company';
      showStatus(globalStatus, `Outlook mailbox connected for ${name}.`, 'success');
    } else if (status === 'error') {
      showStatus(globalStatus, message || 'Outlook connection failed. Please try again.', 'error');
    }
  }

  if (shouldClear) {
    window.history.replaceState({}, document.title, window.location.pathname);
  }
}

function handleDropZoneDragOver(event) {
  event.preventDefault();
  event.stopPropagation();
  if (event.dataTransfer) {
    try {
      event.dataTransfer.dropEffect = 'copy';
    } catch (error) {
      // ignore browser dropEffect errors
    }
  }
  if (dropZone) {
    dropZone.classList.add('dragover');
  }
}

function handleDropZoneDragLeave(event) {
  event.preventDefault();
  if (dropZone) {
    dropZone.classList.remove('dragover');
  }
}

function handleDropZoneDrop(event) {
  event.preventDefault();
  event.stopPropagation();
  if (dropZone) {
    dropZone.classList.remove('dragover');
  }
  const files = extractFilesFromDataTransfer(event.dataTransfer);
  if (files.length) {
    processSelectedFiles(files);
  }
}

async function processSelectedFiles(fileList) {
  const files = Array.from(fileList || []).filter((file) => file && file.size > 0);
  if (!files.length) {
    return;
  }

  clearUploadStatusEmptyState();

  let movedToReview = false;
  for (const file of files) {
    const statusEntry = createUploadStatusEntry(file.name);
    // eslint-disable-next-line no-await-in-loop
    const success = await uploadAndProcessFile(file, statusEntry);
    if (success) {
      movedToReview = true;
    }
  }

  if (uploadInput) {
    uploadInput.value = '';
  }

  if (movedToReview) {
    activateCompanyTab('to-review');
  }
}

function extractFilesFromDataTransfer(dataTransfer) {
  if (!dataTransfer) {
    return [];
  }

  if (dataTransfer.files?.length) {
    return Array.from(dataTransfer.files);
  }

  if (dataTransfer.items?.length) {
    const collected = [];
    Array.from(dataTransfer.items).forEach((item) => {
      if (typeof item.getAsFile === 'function' && item.kind === 'file') {
        const file = item.getAsFile();
        if (file) {
          collected.push(file);
        }
      }
    });
    if (collected.length) {
      return collected;
    }
  }

  return [];
}

function handleDocumentDragOver(event) {
  if (!dropZone || !isAddInvoicesPanelActive()) {
    return;
  }

  const zone = findDropZoneFromEvent(event.target);
  const withinBounds = zone || isEventWithinDropZone(event);
  if (!withinBounds) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  if (event.dataTransfer) {
    try {
      event.dataTransfer.dropEffect = 'copy';
    } catch (error) {
      // ignore
    }
  }

  dropZone.classList.add('dragover');
}

function handleDocumentDrop(event) {
  if (!dropZone || !isAddInvoicesPanelActive()) {
    return;
  }

  const zone = findDropZoneFromEvent(event.target);
  const withinBounds = zone || isEventWithinDropZone(event);
  if (!withinBounds) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  dropZone.classList.remove('dragover');

  const files = extractFilesFromDataTransfer(event.dataTransfer);
  if (files.length) {
    processSelectedFiles(files);
  }
}

function findDropZoneFromEvent(target) {
  if (!target || !dropZone) {
    return null;
  }

  if (target === dropZone || dropZone.contains(target)) {
    return dropZone;
  }

  if (typeof target.closest === 'function') {
    const closest = target.closest('#invoice-drop-zone');
    if (closest === dropZone) {
      return dropZone;
    }
  }

  return null;
}

function isEventWithinDropZone(event) {
  if (!dropZone) {
    return false;
  }

  const { clientX, clientY } = event;
  if (typeof clientX !== 'number' || typeof clientY !== 'number') {
    return false;
  }

  const rect = dropZone.getBoundingClientRect();
  if (!rect) {
    return false;
  }

  const withinHorizontalBounds = clientX >= rect.left && clientX <= rect.right;
  const withinVerticalBounds = clientY >= rect.top && clientY <= rect.bottom;
  return withinHorizontalBounds && withinVerticalBounds;
}

function isAddInvoicesPanelActive() {
  const panel = document.getElementById('add-invoices');
  if (!panel) {
    return false;
  }
  return !panel.hasAttribute('hidden');
}

function clearUploadStatusEmptyState() {
  if (uploadStatusList) {
    const emptyItem = uploadStatusList.querySelector('.empty');
    if (emptyItem) {
      emptyItem.remove();
    }
  }
}

function createUploadStatusEntry(fileName) {
  if (!uploadStatusList) {
    return null;
  }

  const item = document.createElement('li');
  item.className = 'upload-status-item';

  const name = document.createElement('span');
  name.className = 'upload-status-name';
  name.textContent = fileName;

  const message = document.createElement('span');
  message.className = 'upload-status-message';
  message.textContent = 'Queued…';

  item.append(name, message);
  uploadStatusList.prepend(item);

  const entry = { element: item, messageElement: message };
  return entry;
}

function setUploadStatus(entry, message, state = 'info') {
  if (!entry || !entry.element) {
    return;
  }
  entry.element.dataset.state = state;
  if (entry.messageElement) {
    entry.messageElement.textContent = message;
  }
}

async function uploadAndProcessFile(file, statusEntry) {
  if (!file) {
    return false;
  }

  if (!selectedRealmId) {
    setUploadStatus(statusEntry, 'Select a company before uploading invoices.', 'error');
    return false;
  }

  setUploadStatus(statusEntry, 'Uploading to Gemini…');

  const formData = new FormData();
  formData.append('invoice', file);
  formData.append('realmId', selectedRealmId);
  const company = getSelectedCompany();
  if (company?.businessType) {
    formData.append('businessType', company.businessType);
  }

  try {
    const response = await fetch('/api/parse-invoice', {
      method: 'POST',
      body: formData,
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to parse invoice.';
      throw new Error(message);
    }

    const checksum = payload?.invoice?.metadata?.checksum;
    if (checksum) {
      try {
        const statusResponse = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/status`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ status: 'review' }),
        });

        if (!statusResponse.ok) {
          const statusPayload = await statusResponse.json().catch(() => null);
          const message = statusPayload?.error || 'Failed to route invoice to review.';
          throw new Error(message);
        }
      } catch (error) {
        throw new Error(error.message || 'Failed to route invoice to review.');
      }
    }

    await loadStoredInvoices();
    setUploadStatus(statusEntry, 'Parsed and routed to review.', 'success');
    return true;
  } catch (error) {
    console.error(error);
    setUploadStatus(statusEntry, error.message || 'Upload failed.', 'error');
    return false;
  }
}

async function refreshQuickBooksCompanies(preferredRealmId = selectedRealmId) {
  const previousSelection = preferredRealmId || selectedRealmId || '';

  try {
    const response = await fetch('/api/quickbooks/companies');
    let payload = null;
    const responseText = await response.text();
    console.debug('[QuickBooks] /api/quickbooks/companies raw response', responseText);

    if (responseText && responseText.trim()) {
      try {
        payload = JSON.parse(responseText);
      } catch (parseError) {
        console.warn('Failed to parse response from /api/quickbooks/companies', parseError);
        const message = 'The QuickBooks companies response could not be parsed.';
        const error = new Error(message);
        error.cause = parseError;
        throw error;
      }
    }

    if (!response.ok) {
      const message = (payload && typeof payload.error === 'string' && payload.error.trim())
        ? payload.error.trim()
        : responseText && responseText.trim()
          ? responseText.trim()
          : 'Unable to load QuickBooks companies.';
      const error = new Error(message);
      if (payload && typeof payload.code === 'string') {
        error.code = payload.code;
      }
      if (payload && typeof payload.backupPath === 'string') {
        error.backupPath = payload.backupPath;
      }
      throw error;
    }

    if (!payload || typeof payload !== 'object') {
      throw new Error('QuickBooks companies response was empty.');
    }
    if (!Array.isArray(payload?.companies)) {
      throw new Error('QuickBooks companies response did not include any connections.');
    }
    quickBooksCompanies = payload.companies.map((company) => ({
      ...company,
      outlook: company.outlook || null,
    }));

    const activeRealms = new Set(quickBooksCompanies.map((company) => company.realmId));
    Array.from(companyMetadataCache.keys()).forEach((realmId) => {
      if (!activeRealms.has(realmId)) {
        companyMetadataCache.delete(realmId);
      }
    });

    populateCompanySelect(previousSelection, { triggerMetadataWhenSame: true });
  } catch (error) {
    console.error(error);
    quickBooksCompanies = [];
    populateCompanySelect('', { triggerMetadataWhenSame: false });
    showStatus(globalStatus, error.message || 'Failed to load QuickBooks companies.', 'error');
  }
}

function populateCompanySelect(preferredRealmId, { triggerMetadataWhenSame = false } = {}) {
  if (!companySelect) {
    return;
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Choose a company';

  companySelect.innerHTML = '';
  companySelect.appendChild(placeholder);

  quickBooksCompanies.forEach((company) => {
    const option = document.createElement('option');
    option.value = company.realmId;
    option.textContent = company.companyName || company.legalName || `Realm ${company.realmId}`;
    companySelect.appendChild(option);
  });

  if (companySelect) {
    companySelect.disabled = quickBooksCompanies.length === 0;
  }

  const hasPreferred = quickBooksCompanies.some((company) => company.realmId === preferredRealmId);
  const fallbackRealmId = quickBooksCompanies.length ? quickBooksCompanies[0].realmId : '';
  const targetRealmId = hasPreferred ? preferredRealmId : fallbackRealmId;
  const triggerMetadata = hasPreferred ? triggerMetadataWhenSame : Boolean(targetRealmId);
  applySelection(targetRealmId, { triggerMetadata });
}

function applySelection(realmId, { triggerMetadata = true } = {}) {
  selectedRealmId = realmId || '';

  if (companySelect && companySelect.value !== selectedRealmId) {
    companySelect.value = selectedRealmId;
  }

  updateActionStates();

  if (!selectedRealmId) {
    hide(companyDashboard);
    resetCompanyPanels();
    return;
  }

  show(companyDashboard);
  activateCompanyTab('company-settings');
  renderCompanySettings();

  if (!triggerMetadata) {
    const cached = companyMetadataCache.get(selectedRealmId);
    if (cached) {
      renderVendorList(cached, 'No vendors available for this company.');
      renderAccountList(cached, 'No accounts available for this company.');
      return;
    }
  }

  loadCompanyMetadata(selectedRealmId);
}

function updateActionStates() {
  const hasSelection = Boolean(selectedRealmId);
  if (refreshMetadataButton) {
    refreshMetadataButton.disabled = !hasSelection;
  }
  if (connectCompanyButton) {
    connectCompanyButton.disabled = !hasSelection;
  }
}

function resetCompanyPanels() {
  if (connectionStatus) {
    connectionStatus.textContent = 'Select a company to view connection details.';
  }
  renderVendorList(null, 'Select a company to load vendor details.');
  renderAccountList(null, 'Select a company to load account details.');
  renderBusinessProfile();
  renderOneDriveSettings();
  renderOutlookSettings();
}

function activateCompanyTab(targetId) {
  if (!targetId) {
    return;
  }

  tabButtons.forEach((button) => {
    const isActive = button.dataset.companyTab === targetId;
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  companyPanels.forEach((panel) => {
    const isActive = panel.id === targetId;
    panel.classList.toggle('active', isActive);
    if (isActive) {
      panel.removeAttribute('hidden');
    } else {
      panel.setAttribute('hidden', '');
    }
  });
}

function getSelectedCompany() {
  if (!selectedRealmId) {
    return null;
  }
  return quickBooksCompanies.find((company) => company.realmId === selectedRealmId) || null;
}

function renderCompanySettings() {
  const company = getSelectedCompany();
  if (!connectionStatus) {
    return;
  }

  if (!company) {
    connectionStatus.textContent = 'Select a company to view connection details.';
    renderBusinessProfile();
    renderOneDriveSettings();
    renderOutlookSettings();
    return;
  }

  const lines = [];
  const name = company.companyName || company.legalName || `Realm ${company.realmId}`;
  lines.push(name);

  if (company.environment) {
    lines.push(company.environment === 'production' ? 'Production environment' : 'Sandbox environment');
  }

  if (company.connectedAt) {
    lines.push(`Connected ${formatTimestamp(company.connectedAt)}`);
  }

  if (company.updatedAt && company.updatedAt !== company.connectedAt) {
    lines.push(`Updated ${formatTimestamp(company.updatedAt)}`);
  }

  if (company.businessType) {
    lines.push(`Business type: ${company.businessType}`);
  }

  const counts = [];
  if (typeof company.vendorsCount === 'number') {
    counts.push(`${company.vendorsCount} vendors`);
  }
  if (typeof company.accountsCount === 'number') {
    counts.push(`${company.accountsCount} accounts`);
  }
  if (typeof company.taxCodesCount === 'number') {
    counts.push(`${company.taxCodesCount} tax codes`);
  }
  if (counts.length) {
    lines.push(counts.join(' • '));
  }

  connectionStatus.textContent = lines.join('\n');
  renderBusinessProfile();
  renderOneDriveSettings();
  renderOutlookSettings();
}

function renderBusinessProfile() {
  if (!businessTypeInput || !businessProfileSubmit) {
    return;
  }

  const company = getSelectedCompany();
  if (!company) {
    businessTypeInput.value = '';
    businessTypeInput.disabled = true;
    businessProfileSubmit.disabled = true;
    return;
  }

  businessTypeInput.disabled = false;
  businessTypeInput.value = company.businessType || '';
  businessProfileSubmit.disabled = false;
}

async function handleBusinessProfileSave(event) {
  event.preventDefault();

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before updating the business profile.', 'error');
    return;
  }

  const currentCompany = getSelectedCompany();
  const rawValue = businessTypeInput ? businessTypeInput.value : '';
  const trimmedValue = typeof rawValue === 'string' ? rawValue.trim() : '';
  const nextBusinessType = trimmedValue || null;
  const currentBusinessType = currentCompany?.businessType || null;

  if ((currentBusinessType || null) === nextBusinessType) {
    showStatus(globalStatus, 'Business profile is already up to date.', 'info');
    return;
  }

  const endpoint = `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}`;
  const originalLabel = businessProfileSubmit ? businessProfileSubmit.textContent : '';

  if (businessProfileSubmit) {
    businessProfileSubmit.disabled = true;
    businessProfileSubmit.textContent = 'Saving…';
  }
  if (businessTypeInput) {
    businessTypeInput.disabled = true;
  }

  try {
    const response = await fetch(endpoint, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ businessType: nextBusinessType }),
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (parseError) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to update business profile.';
      throw new Error(message);
    }

    const updatedCompany = payload?.company;
    if (updatedCompany) {
      const index = quickBooksCompanies.findIndex((entry) => entry.realmId === updatedCompany.realmId);
      if (index >= 0) {
        quickBooksCompanies[index] = {
          ...quickBooksCompanies[index],
          ...updatedCompany,
        };
      } else {
        quickBooksCompanies.push(updatedCompany);
      }
      renderCompanySettings();
    } else {
      await refreshQuickBooksCompanies(selectedRealmId);
    }

    showStatus(globalStatus, 'Business profile saved.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update business profile.', 'error');
  } finally {
    if (businessTypeInput) {
      businessTypeInput.disabled = !selectedRealmId;
    }
    if (businessProfileSubmit) {
      businessProfileSubmit.disabled = !selectedRealmId;
      businessProfileSubmit.textContent = originalLabel || 'Save profile';
    }
    renderBusinessProfile();
  }
}

function renderOneDriveSettings() {
  renderSharedOneDriveSummary();
  if (!oneDriveEnabledInput || !oneDriveForm) {
    return;
  }

  const company = getSelectedCompany();
  const hasCompany = Boolean(company);
  const sharedConfigured = isSharedOneDriveConfigured();

  oneDriveEnabledInput.disabled = !hasCompany;

  if (oneDriveBrowseMonitoredButton) {
    oneDriveBrowseMonitoredButton.disabled = !hasCompany || !sharedConfigured;
  }
  if (oneDriveBrowseProcessedButton) {
    oneDriveBrowseProcessedButton.disabled = !hasCompany || !sharedConfigured;
  }
  if (oneDriveProcessedClearButton) {
    oneDriveProcessedClearButton.disabled = !hasCompany;
  }
  if (oneDriveSaveButton) {
    oneDriveSaveButton.disabled = !hasCompany;
  }

  if (!hasCompany) {
    if (oneDriveForm) {
      oneDriveForm.reset();
    }
    [
      oneDriveFolderIdInput,
      oneDriveFolderPathInput,
      oneDriveFolderNameInput,
      oneDriveFolderWebUrlInput,
      oneDriveFolderParentIdInput,
      oneDriveProcessedFolderIdInput,
      oneDriveProcessedFolderPathInput,
      oneDriveProcessedFolderNameInput,
      oneDriveProcessedFolderWebUrlInput,
      oneDriveProcessedFolderParentIdInput,
    ].forEach((input) => {
      if (input) {
        input.value = '';
        if (input.dataset) {
          delete input.dataset.cleared;
        }
      }
    });
    setOneDriveInputBreadcrumb(oneDriveFolderPathInput, '');
    setOneDriveInputBreadcrumb(oneDriveProcessedFolderPathInput, '');
    updateOneDriveSelectionPreviewFromInputs();
    updateOneDriveProcessedSummaryFromInputs();
    if (oneDriveStatusContainer) {
      oneDriveStatusContainer.hidden = true;
    }
    if (oneDriveStatusLog) {
      oneDriveStatusLog.textContent = '—';
    }
    if (oneDriveSyncButton) {
      oneDriveSyncButton.disabled = true;
    }
    if (oneDriveResyncButton) {
      oneDriveResyncButton.disabled = true;
    }
    if (oneDriveClearButton) {
      oneDriveClearButton.disabled = true;
    }
    return;
  }

  const config = company.oneDrive || null;
  const monitored = config?.monitoredFolder || null;
  const processed = config?.processedFolder || null;

  const assignValue = (input, value) => {
    if (input) {
      input.value = value || '';
    }
  };

  assignValue(oneDriveFolderIdInput, monitored?.id);
  assignValue(oneDriveFolderPathInput, monitored?.path);
  assignValue(oneDriveFolderNameInput, monitored?.name);
  assignValue(oneDriveFolderWebUrlInput, monitored?.webUrl);
  assignValue(oneDriveFolderParentIdInput, monitored?.parentId);
  setOneDriveInputBreadcrumb(
    oneDriveFolderPathInput,
    deriveOneDriveBreadcrumbFromPath(monitored?.path || '')
  );

  assignValue(oneDriveProcessedFolderIdInput, processed?.id);
  assignValue(oneDriveProcessedFolderPathInput, processed?.path);
  assignValue(oneDriveProcessedFolderNameInput, processed?.name);
  assignValue(oneDriveProcessedFolderWebUrlInput, processed?.webUrl);
  assignValue(oneDriveProcessedFolderParentIdInput, processed?.parentId);
  setOneDriveInputBreadcrumb(
    oneDriveProcessedFolderPathInput,
    deriveOneDriveBreadcrumbFromPath(processed?.path || '')
  );
  if (oneDriveProcessedFolderIdInput?.dataset) {
    delete oneDriveProcessedFolderIdInput.dataset.cleared;
  }

  const isEnabled = config ? config.enabled !== false : false;
  oneDriveEnabledInput.checked = isEnabled;

  if (oneDriveSyncButton) {
    oneDriveSyncButton.disabled = !config || !isEnabled || !monitored?.id;
  }
  if (oneDriveResyncButton) {
    oneDriveResyncButton.disabled = !config || !isEnabled || !monitored?.id;
  }
  if (oneDriveClearButton) {
    oneDriveClearButton.disabled = !config;
  }

  if (!oneDriveStatusContainer) {
    updateOneDriveSelectionPreviewFromInputs();
    updateOneDriveProcessedSummaryFromInputs();
    return;
  }

  if (!config) {
    oneDriveStatusContainer.hidden = true;
    if (oneDriveStatusLog) {
      oneDriveStatusLog.textContent = '—';
    }
    updateOneDriveSelectionPreviewFromInputs();
    updateOneDriveProcessedSummaryFromInputs();
    return;
  }

  oneDriveStatusContainer.hidden = false;

  if (oneDriveStatusState) {
    oneDriveStatusState.textContent = formatStatusLabel(config.status, isEnabled ? 'Connected' : 'Disabled');
  }

  if (oneDriveStatusFolder) {
    const folderParts = [];
    if (monitored?.name) {
      folderParts.push(monitored.name);
    }
    if (monitored?.path) {
      const friendlyPath = formatOneDriveBreadcrumb(
        deriveOneDriveBreadcrumbFromPath(monitored.path)
      );
      folderParts.push(friendlyPath || monitored.path);
    }
    if (!folderParts.length && monitored?.webUrl) {
      folderParts.push(monitored.webUrl);
    }
    oneDriveStatusFolder.textContent = folderParts.length ? folderParts.join(' • ') : 'Not specified';
  }

  if (oneDriveStatusLastSync) {
    if (config.lastSyncAt) {
      oneDriveStatusLastSync.textContent = formatTimestamp(config.lastSyncAt);
    } else if (isEnabled) {
      oneDriveStatusLastSync.textContent = 'Not run yet';
    } else {
      oneDriveStatusLastSync.textContent = 'Disabled';
    }
  }

  if (oneDriveStatusResult) {
    oneDriveStatusResult.textContent = buildOneDriveResultText(config, { enabled: isEnabled });
  }

  if (oneDriveStatusLog) {
    const entries = Array.isArray(config.lastSyncLog)
      ? config.lastSyncLog
          .map((entry) => (typeof entry === 'string' ? entry.trim() : ''))
          .filter(Boolean)
      : [];

    if (!entries.length) {
      oneDriveStatusLog.textContent = isEnabled
        ? 'No recent sync activity details.'
        : 'Not available while disabled.';
    } else {
      oneDriveStatusLog.textContent = entries.join('\n');
    }
  }

  updateOneDriveSelectionPreviewFromInputs();
  updateOneDriveProcessedSummaryFromInputs();
}
function updateOneDriveSelectionPreviewFromInputs() {
  if (!oneDriveMonitoredSummary) {
    return;
  }

  if (!isSharedOneDriveConfigured()) {
    oneDriveMonitoredSummary.textContent = 'Connect the OneDrive account before choosing a folder for this company.';
    return;
  }

  const folder = {
    id: normaliseTextInput(oneDriveFolderIdInput?.value),
    name: normaliseTextInput(oneDriveFolderNameInput?.value),
    path: normaliseTextInput(oneDriveFolderPathInput?.value),
    webUrl: normaliseTextInput(oneDriveFolderWebUrlInput?.value),
  };
  const breadcrumb = getOneDriveInputBreadcrumb(oneDriveFolderPathInput);

  let summary = MONITORED_DEFAULT_SUMMARY;
  if (folder.name || folder.path || breadcrumb) {
    summary = buildOneDriveFolderSummary(folder.name, folder.path, folder.webUrl, MONITORED_DEFAULT_SUMMARY, breadcrumb);
  } else if (folder.id) {
    summary = `Folder ${folder.id}`;
  }

  oneDriveMonitoredSummary.textContent = summary;
}

function updateOneDriveProcessedSummaryFromInputs() {
  if (!oneDriveProcessedSummary) {
    return;
  }

  const folder = {
    id: normaliseTextInput(oneDriveProcessedFolderIdInput?.value),
    name: normaliseTextInput(oneDriveProcessedFolderNameInput?.value),
    path: normaliseTextInput(oneDriveProcessedFolderPathInput?.value),
    webUrl: normaliseTextInput(oneDriveProcessedFolderWebUrlInput?.value),
  };
  const breadcrumb = getOneDriveInputBreadcrumb(oneDriveProcessedFolderPathInput);

  let summary = buildOneDriveFolderSummary(folder.name, folder.path, folder.webUrl, PROCESSED_DEFAULT_SUMMARY, breadcrumb);

  if (!summary && folder.id) {
    summary = `Folder ${folder.id}`;
  }

  oneDriveProcessedSummary.textContent = summary || PROCESSED_DEFAULT_SUMMARY;
}

function collectMonitoredFolderFromInputs() {
  const id = normaliseTextInput(oneDriveFolderIdInput?.value);
  const path = normaliseTextInput(oneDriveFolderPathInput?.value);
  const name = normaliseTextInput(oneDriveFolderNameInput?.value);
  const webUrl = normaliseTextInput(oneDriveFolderWebUrlInput?.value);
  const parentId = normaliseTextInput(oneDriveFolderParentIdInput?.value);

  if (!id && !path) {
    return null;
  }

  const folder = {};
  if (id) {
    folder.id = id;
  }
  if (path) {
    folder.path = path;
  }
  if (name) {
    folder.name = name;
  }
  if (webUrl) {
    folder.webUrl = webUrl;
  }
  if (parentId) {
    folder.parentId = parentId;
  }
  return folder;
}

function collectProcessedFolderFromInputs() {
  const id = normaliseTextInput(oneDriveProcessedFolderIdInput?.value);
  const path = normaliseTextInput(oneDriveProcessedFolderPathInput?.value);
  const name = normaliseTextInput(oneDriveProcessedFolderNameInput?.value);
  const webUrl = normaliseTextInput(oneDriveProcessedFolderWebUrlInput?.value);
  const parentId = normaliseTextInput(oneDriveProcessedFolderParentIdInput?.value);

  if (!id && !path) {
    return null;
  }

  const folder = {};
  if (id) {
    folder.id = id;
  }
  if (path) {
    folder.path = path;
  }
  if (name) {
    folder.name = name;
  }
  if (webUrl) {
    folder.webUrl = webUrl;
  }
  if (parentId) {
    folder.parentId = parentId;
  }
  return folder;
}

async function handleOneDriveSettingsSave(event) {
  event.preventDefault();

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before updating OneDrive settings.', 'error');
    return;
  }

  const enableSync = Boolean(oneDriveEnabledInput?.checked);
  const monitoredFolder = collectMonitoredFolderFromInputs();
  const processedFolder = collectProcessedFolderFromInputs();
  const processedCleared = Boolean(oneDriveProcessedFolderIdInput?.dataset?.cleared);

  if (enableSync && !isSharedOneDriveConfigured()) {
    showStatus(globalStatus, 'Connect the OneDrive account before enabling sync.', 'error');
    return;
  }

  if (enableSync && !monitoredFolder) {
    showStatus(globalStatus, 'Select a OneDrive folder to monitor before enabling sync.', 'error');
    return;
  }

  const payload = {
    enabled: enableSync,
  };

  if (monitoredFolder) {
    payload.monitoredFolder = monitoredFolder;
  }

  if (processedFolder) {
    payload.processedFolder = processedFolder;
  } else if (processedCleared) {
    payload.processedFolder = null;
  }

  const originalLabel = oneDriveSaveButton ? oneDriveSaveButton.textContent : '';

  if (oneDriveSaveButton) {
    oneDriveSaveButton.disabled = true;
    oneDriveSaveButton.textContent = 'Saving…';
  }

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      }
    );

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to update OneDrive settings.';
      throw new Error(message);
    }

    let stateForMessage = null;
    if (body?.oneDrive !== undefined) {
      stateForMessage = body.oneDrive;
      updateLocalCompanyOneDrive(selectedRealmId, body.oneDrive);
    } else {
      await refreshQuickBooksCompanies(selectedRealmId);
      stateForMessage = getSelectedCompany()?.oneDrive || null;
    }

    const isEnabled = stateForMessage ? stateForMessage.enabled !== false : enableSync;
    showStatus(
      globalStatus,
      isEnabled
        ? 'OneDrive settings saved. Sync will run shortly.'
        : 'OneDrive sync disabled for this company.',
      'success'
    );
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update OneDrive settings.', 'error');
  } finally {
    if (oneDriveSaveButton) {
      oneDriveSaveButton.disabled = false;
      oneDriveSaveButton.textContent = originalLabel || 'Save OneDrive settings';
    }
  }
}

async function handleOneDriveSyncClick() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before triggering a OneDrive sync.', 'error');
    return;
  }

  if (!oneDriveSyncButton || oneDriveSyncButton.disabled) {
    return;
  }

  const originalLabel = oneDriveSyncButton.textContent;
  oneDriveSyncButton.disabled = true;
  oneDriveSyncButton.textContent = 'Syncing…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive/sync`,
      { method: 'POST' }
    );

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to start OneDrive sync.';
      throw new Error(message);
    }

    if (body?.oneDrive !== undefined) {
      updateLocalCompanyOneDrive(selectedRealmId, body.oneDrive);
    }

    showStatus(globalStatus, 'OneDrive sync queued. Check the status card for updates.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to start OneDrive sync.', 'error');
  } finally {
    if (oneDriveSyncButton) {
      oneDriveSyncButton.disabled = false;
      oneDriveSyncButton.textContent = originalLabel || 'Sync now';
    }
    renderOneDriveSettings();
  }
}

async function handleOneDriveFullResync() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before requesting a OneDrive resync.', 'error');
    return;
  }

  if (!oneDriveResyncButton || oneDriveResyncButton.disabled) {
    return;
  }

  const originalLabel = oneDriveResyncButton.textContent;
  oneDriveResyncButton.disabled = true;
  oneDriveResyncButton.textContent = 'Queuing…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive/resync`,
      { method: 'POST' }
    );

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to start OneDrive resync.';
      throw new Error(message);
    }

    if (body?.oneDrive !== undefined) {
      updateLocalCompanyOneDrive(selectedRealmId, body.oneDrive);
    }

    showStatus(globalStatus, 'OneDrive full resync queued. This may take a few minutes.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to start OneDrive resync.', 'error');
  } finally {
    if (oneDriveResyncButton) {
      oneDriveResyncButton.disabled = false;
      oneDriveResyncButton.textContent = originalLabel || 'Request full resync';
    }
    renderOneDriveSettings();
  }
}

async function handleOneDriveDisconnect() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before disconnecting OneDrive.', 'error');
    return;
  }

  if (!oneDriveClearButton || oneDriveClearButton.disabled) {
    return;
  }

  const confirmed = window.confirm('Disconnect OneDrive for this company?');
  if (!confirmed) {
    return;
  }

  const originalLabel = oneDriveClearButton.textContent;
  oneDriveClearButton.disabled = true;
  oneDriveClearButton.textContent = 'Disconnecting…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive`,
      { method: 'DELETE' }
    );

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to disconnect OneDrive.';
      throw new Error(message);
    }

    updateLocalCompanyOneDrive(selectedRealmId, body?.oneDrive || null);
    showStatus(globalStatus, 'OneDrive settings cleared for this company.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to disconnect OneDrive.', 'error');
  } finally {
    if (oneDriveClearButton) {
      oneDriveClearButton.disabled = false;
      oneDriveClearButton.textContent = originalLabel || 'Disconnect OneDrive';
    }
    renderOneDriveSettings();
  }
}

function formatStatusLabel(status, fallback = 'Unknown') {
  const value = typeof status === 'string' ? status.trim() : '';
  if (!value) {
    return fallback;
  }

  return value
    .split(/[_\s-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(' ');
}

function buildOneDriveResultText(config, { enabled = true } = {}) {
  if (!config) {
    return enabled ? 'Not connected.' : 'Disabled.';
  }

  if (config.lastSyncError?.message) {
    const label = config.lastSyncStatus ? formatStatusLabel(config.lastSyncStatus, 'Error') : 'Error';
    return `${label} • ${config.lastSyncError.message}`;
  }

  if (!config.lastSyncStatus) {
    return enabled ? 'No syncs yet.' : 'Disabled.';
  }

  const metrics = config.lastSyncMetrics || {};
  const parts = [];

  if (typeof metrics.createdCount === 'number' && metrics.createdCount > 0) {
    parts.push(`${metrics.createdCount} new ${metrics.createdCount === 1 ? 'invoice' : 'invoices'}`);
  }

  if (typeof metrics.processedItems === 'number' && metrics.processedItems > 0) {
    if (!parts.length || metrics.processedItems !== metrics.createdCount) {
      parts.push(`${metrics.processedItems} file${metrics.processedItems === 1 ? '' : 's'} processed`);
    }
  }

  if (typeof metrics.duplicateCount === 'number' && metrics.duplicateCount > 0) {
    parts.push(`${metrics.duplicateCount} duplicate${metrics.duplicateCount === 1 ? '' : 's'} skipped`);
  }

  const processed = config.processedFolder || null;
  const moveTarget = processed?.path || processed?.name;
  if (moveTarget) {
    let label = processed?.name || moveTarget;
    if (processed?.path) {
      const friendly = formatOneDriveBreadcrumb(deriveOneDriveBreadcrumbFromPath(processed.path));
      label = friendly || processed.path;
    }
    parts.push(`Moved to ${label}`);
  }

  if (!parts.length && config.lastSyncStatus === 'success') {
    parts.push('Completed without errors');
  }

  const label = formatStatusLabel(config.lastSyncStatus, enabled ? 'Success' : 'Disabled');
  return parts.length ? `${label} • ${parts.join(', ')}` : label;
}

function normaliseTextInput(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  return String(value).trim();
}

function renderEntityEmptyState(target, message) {
  if (!target) {
    return;
  }

  target.innerHTML = '';
  const entry = document.createElement('li');
  entry.className = 'empty';
  entry.textContent = message;
  target.appendChild(entry);
}

function normaliseCompanyMetadata(raw = {}) {
  const vendors = prepareMetadataSection(raw.vendors);
  const accounts = prepareMetadataSection(raw.accounts);
  const taxCodes = prepareMetadataSection(raw.taxCodes);
  const vendorSettings = prepareVendorSettings(raw.vendorSettings, { vendors, accounts, taxCodes });

  return {
    vendors,
    accounts,
    taxCodes,
    vendorSettings,
    refreshedAt: raw.refreshedAt || raw.updatedAt || raw.syncedAt || null,
  };
}

async function loadCompanyMetadata(realmId, { force = false } = {}) {
  if (!realmId) {
    return null;
  }

  const cached = companyMetadataCache.get(realmId) || null;
  if (cached && !force) {
    if (selectedRealmId === realmId) {
      renderVendorList(cached, 'No vendors available for this company.');
      renderAccountList(cached, 'No accounts available for this company.');
    }
    return cached;
  }

  if (metadataRequests.has(realmId)) {
    return metadataRequests.get(realmId);
  }

  const request = (async () => {
    try {
      const response = await fetch(
        `/api/quickbooks/companies/${encodeURIComponent(realmId)}/metadata`
      );
      const payload = await response.json().catch(() => ({}));

      if (!response.ok) {
        const message = payload?.error || 'Failed to load QuickBooks metadata.';
        throw new Error(message);
      }

      const metadata = normaliseCompanyMetadata(payload?.metadata || {});
      companyMetadataCache.set(realmId, metadata);

      if (selectedRealmId === realmId) {
        renderVendorList(metadata, 'No vendors available for this company.');
        renderAccountList(metadata, 'No accounts available for this company.');
        renderInvoices();
      }

      autoMatchInvoices(realmId).catch((error) => console.warn('Auto-match scheduling failed', error));

      return metadata;
    } catch (error) {
      if (selectedRealmId === realmId) {
        console.error(error);
        showStatus(globalStatus, error.message || 'Failed to load QuickBooks metadata.', 'error');
        renderVendorList(null, 'Unable to load vendor list. Try refreshing metadata.');
        renderAccountList(null, 'Unable to load account list. Try refreshing metadata.');
      }
      throw error;
    } finally {
      metadataRequests.delete(realmId);
    }
  })();

  metadataRequests.set(realmId, request);
  return request;
}

async function refreshCompanyMetadata(realmId) {
  if (!realmId) {
    showStatus(globalStatus, 'Select a company before refreshing metadata.', 'error');
    return;
  }

  const originalLabel = refreshMetadataButton ? refreshMetadataButton.textContent : '';
  if (refreshMetadataButton) {
    refreshMetadataButton.disabled = true;
    refreshMetadataButton.textContent = 'Refreshing…';
  }

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(realmId)}/metadata/refresh`,
      { method: 'POST' }
    );
    const body = await response.json().catch(() => ({}));

    if (!response.ok) {
      const message = body?.error || 'Failed to refresh QuickBooks metadata.';
      throw new Error(message);
    }

    const metadata = normaliseCompanyMetadata(body?.metadata || {});
    companyMetadataCache.set(realmId, metadata);

    if (selectedRealmId === realmId) {
      renderVendorList(metadata, 'No vendors available for this company.');
      renderAccountList(metadata, 'No accounts available for this company.');
      renderInvoices();
    }

    autoMatchInvoices(realmId).catch((error) => console.warn('Auto-match scheduling failed', error));

    showStatus(globalStatus, 'QuickBooks metadata refreshed.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to refresh QuickBooks metadata.', 'error');
  } finally {
    if (refreshMetadataButton) {
      refreshMetadataButton.disabled = !selectedRealmId;
      refreshMetadataButton.textContent = originalLabel || 'Refresh metadata';
    }
  }
}

function renderVendorList(metadata, emptyMessage) {
  if (!vendorList) {
    return;
  }

  vendorList.innerHTML = '';

  const vendors = metadata?.vendors?.items || [];
  if (!vendors.length) {
    renderEntityEmptyState(vendorList, emptyMessage || 'No vendors available for this company.');
    return;
  }

  const vendorSettingsEntries = metadata?.vendorSettings?.entries || {};
  const accountOptions = buildAccountOptions(metadata?.accounts?.items || []);
  const taxCodeOptions = buildTaxCodeOptions(metadata?.taxCodes?.items || []);

  const sorted = [...vendors].sort((a, b) => {
    const left = (a.displayName || a.companyName || a.fullyQualifiedName || a.id || '').toLowerCase();
    const right = (b.displayName || b.companyName || b.fullyQualifiedName || b.id || '').toLowerCase();
    return left.localeCompare(right, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((vendor) => {
    const item = document.createElement('li');
    item.className = 'entity-item';

    const vendorName =
      vendor.displayName || vendor.companyName || vendor.fullyQualifiedName || `Vendor ${vendor.id}`;

    const name = document.createElement('div');
    name.className = 'entity-name';
    name.textContent = vendorName;
    item.appendChild(name);

    const metaParts = [];
    if (vendor.companyName && vendor.companyName !== vendor.displayName) {
      metaParts.push(vendor.companyName);
    }
    if (vendor.email) {
      metaParts.push(vendor.email);
    }
    if (vendor.phone) {
      metaParts.push(vendor.phone);
    }
    if (metaParts.length) {
      const meta = document.createElement('div');
      meta.className = 'entity-meta';
      meta.textContent = metaParts.join(' • ');
      item.appendChild(meta);
    }

    const controls = createVendorSettingsControls({
      vendorId: vendor.id,
      vendorName,
      defaults: vendorSettingsEntries[vendor.id] || {},
      accountOptions,
      taxCodeOptions,
    });
    item.appendChild(controls);

    vendorList.appendChild(item);
  });
}

function buildAccountOptions(accounts) {
  const options = [{ value: '', label: 'No default' }];
  const sorted = [...accounts].sort((a, b) => {
    const left = (a.fullyQualifiedName || a.name || a.id || '').toLowerCase();
    const right = (b.fullyQualifiedName || b.name || b.id || '').toLowerCase();
    return left.localeCompare(right, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((account) => {
    if (!account?.id) {
      return;
    }
    const baseLabel = account.fullyQualifiedName || account.name || `Account ${account.id}`;
    const parts = [baseLabel];
    if (account.accountType) {
      parts.push(account.accountType);
    }
    if (account.accountSubType && account.accountSubType !== account.accountType) {
      parts.push(account.accountSubType);
    }
    options.push({ value: account.id, label: parts.join(' • ') });
  });

  return options;
}

function buildTaxCodeOptions(taxCodes) {
  const options = [{ value: '', label: 'No default' }];
  const sorted = [...taxCodes].sort((a, b) => {
    const left = (a.name || a.id || '').toLowerCase();
    const right = (b.name || b.id || '').toLowerCase();
    return left.localeCompare(right, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((code) => {
    if (!code?.id) {
      return;
    }

    const baseLabel = code.name || `Tax Code ${code.id}`;
    const numericRate = typeof code.rate === 'number' ? code.rate : null;
    const hasEmbeddedRate =
      numericRate !== null && ratesApproximatelyEqual(numericRate, parsePercentageFromLabel(baseLabel));
    const suffix = numericRate !== null && !hasEmbeddedRate ? `${formatRateValue(numericRate)}%` : null;
    const label = suffix ? `${baseLabel} • ${suffix}` : baseLabel;

    options.push({ value: code.id, label });
  });

  return options;
}

function createVendorSettingsControls({ vendorId, vendorName, defaults, accountOptions, taxCodeOptions }) {
  const container = document.createElement('div');
  container.className = 'vendor-settings-grid';

  const categorySelect = createVendorSettingSelect({
    options: accountOptions,
    value: defaults.accountId || '',
    vendorId,
    vendorName,
    field: 'accountId',
    disabled: accountOptions.length <= 1,
    disabledHint: 'No QuickBooks accounts available. Refresh metadata to sync accounts.',
  });

  const vatBasisSelect = createVendorSettingSelect({
    options: VAT_TREATMENT_OPTIONS,
    value: defaults.vatTreatment || '',
    vendorId,
    vendorName,
    field: 'vatTreatment',
  });

  const taxCodeSelect = createVendorSettingSelect({
    options: taxCodeOptions,
    value: defaults.taxCodeId || '',
    vendorId,
    vendorName,
    field: 'taxCodeId',
    disabled: taxCodeOptions.length <= 1,
    disabledHint: 'No QuickBooks tax codes available. Refresh metadata to sync tax codes.',
  });

  container.appendChild(createVendorSettingGroup('Category', categorySelect));
  container.appendChild(createVendorSettingGroup('Amounts are', vatBasisSelect));
  container.appendChild(createVendorSettingGroup('VAT code', taxCodeSelect));

  return container;
}

function createVendorSettingGroup(labelText, control) {
  const wrapper = document.createElement('label');
  wrapper.className = 'vendor-setting-group';

  const heading = document.createElement('span');
  heading.className = 'vendor-setting-heading';
  heading.textContent = labelText;

  wrapper.appendChild(heading);
  wrapper.appendChild(control);

  return wrapper;
}

function createVendorSettingSelect({
  options,
  value,
  vendorId,
  vendorName,
  field,
  disabled = false,
  disabledHint = '',
}) {
  const select = document.createElement('select');
  select.className = 'vendor-setting-select';
  select.dataset.vendorId = vendorId;
  select.dataset.settingField = field;
  select.dataset.vendorName = vendorName;

  options.forEach((option) => {
    const optionElement = document.createElement('option');
    optionElement.value = option.value;
    optionElement.textContent = option.label;
    if (option.disabled) {
      optionElement.disabled = true;
    }
    select.appendChild(optionElement);
  });

  const availableValues = options.map((option) => option.value);
  const initialValue = availableValues.includes(value) ? value : '';
  select.value = initialValue;
  select.dataset.previousValue = initialValue;

  if (disabled) {
    select.disabled = true;
    if (disabledHint) {
      select.title = disabledHint;
    }
  }

  return select;
}

function resolveVendorSettingFieldValue(field, settings) {
  if (field === 'accountId' || field === 'taxCodeId') {
    return sanitizeReviewSelectionId(settings?.[field]) || '';
  }
  if (field === 'vatTreatment') {
    return VENDOR_VAT_TREATMENT_VALUES.has(settings?.vatTreatment) ? settings.vatTreatment : '';
  }
  return '';
}

function applyVendorSettingUpdate(realmId, vendorId, settings) {
  const metadata = companyMetadataCache.get(realmId);
  if (!metadata) {
    return;
  }

  const vendorSettingsEntries = { ...(metadata.vendorSettings?.entries || {}) };
  const canonicalVendorId = sanitizeReviewSelectionId(vendorId);
  if (!canonicalVendorId) {
    return;
  }

  const accountLookup = metadata.accounts?.lookup || new Map();
  const taxLookup = metadata.taxCodes?.lookup || new Map();

  const accountId = sanitizeReviewSelectionId(settings?.accountId);
  const validAccountId = accountId && accountLookup.has(accountId) ? accountId : null;

  const taxCodeId = sanitizeReviewSelectionId(settings?.taxCodeId);
  const validTaxCodeId = taxCodeId && taxLookup.has(taxCodeId) ? taxCodeId : null;

  const vatTreatment = VENDOR_VAT_TREATMENT_VALUES.has(settings?.vatTreatment)
    ? settings.vatTreatment
    : null;

  if (validAccountId || validTaxCodeId || vatTreatment) {
    vendorSettingsEntries[canonicalVendorId] = {
      accountId: validAccountId,
      taxCodeId: validTaxCodeId,
      vatTreatment,
    };
  } else {
    delete vendorSettingsEntries[canonicalVendorId];
  }

  applyVendorSettingsStructure(metadata, vendorSettingsEntries);
}

function parsePercentageFromLabel(label) {
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

function formatRateValue(rate) {
  if (!Number.isFinite(rate)) {
    return String(rate);
  }

  return rate % 1 === 0 ? rate.toString() : rate.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
}

function ratesApproximatelyEqual(a, b) {
  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    return false;
  }

  return Math.abs(a - b) < 0.0001;
}

function renderAccountList(metadata, emptyMessage) {
  if (!accountList) {
    return;
  }

  accountList.innerHTML = '';

  const accounts = metadata?.accounts?.items || [];
  if (!accounts.length) {
    renderEntityEmptyState(accountList, emptyMessage || 'No accounts available for this company.');
    return;
  }

  const sorted = [...accounts].sort((a, b) => {
    const left = (a.fullyQualifiedName || a.name || a.id || '').toLowerCase();
    const right = (b.fullyQualifiedName || b.name || b.id || '').toLowerCase();
    return left.localeCompare(right, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((account) => {
    const item = document.createElement('li');
    item.className = 'entity-item';

    const name = document.createElement('div');
    name.className = 'entity-name';
    name.textContent = account.fullyQualifiedName || account.name || `Account ${account.id}`;
    item.appendChild(name);

    const metaParts = [];
    if (account.accountType) {
      metaParts.push(account.accountType);
    }
    if (account.accountSubType) {
      metaParts.push(account.accountSubType);
    }
    if (typeof account.currentBalance === 'number') {
      metaParts.push(`Balance: ${formatAmount(account.currentBalance)}`);
    }

    if (metaParts.length) {
      const meta = document.createElement('div');
      meta.className = 'entity-meta';
      meta.textContent = metaParts.join(' • ');
      item.appendChild(meta);
    }

    accountList.appendChild(item);
  });
}

async function loadStoredInvoices() {
  try {
    const response = await fetch('/api/invoices');
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload?.error || 'Failed to load stored invoices.');
    }
    storedInvoices = Array.isArray(payload?.invoices) ? payload.invoices : [];
    renderInvoices();
    autoMatchInvoices().catch((error) => console.warn('Auto-match scheduling failed', error));
  } catch (error) {
    storedInvoices = [];
    console.warn('Unable to load stored invoices', error);
    renderInvoices();
    autoMatchInvoices().catch((error) => console.warn('Auto-match scheduling failed', error));
  }
}

function renderInvoices() {
  // Clear existing content
  if (reviewTableBody) {
    reviewTableBody.innerHTML = '';
  }
  if (archiveTableBody) {
    archiveTableBody.innerHTML = '';
  }

  if (!storedInvoices || storedInvoices.length === 0) {
    if (reviewTableBody) {
      reviewTableBody.innerHTML = '<tr><td colspan="10" class="text-center">No invoices in review</td></tr>';
    }
    if (archiveTableBody) {
      archiveTableBody.innerHTML = '<tr><td colspan="5" class="text-center">No archived invoices</td></tr>';
    }
    updateBulkActionsState();
    return;
  }

  // Separate invoices by status
  const reviewInvoices = storedInvoices.filter(
    (invoice) => resolveStoredInvoiceStatus(invoice) === 'review'
  );
  const archiveInvoices = storedInvoices.filter(
    (invoice) => resolveStoredInvoiceStatus(invoice) === 'archive'
  );

  if (reviewEditingChecksums.size) {
    const reviewChecksums = new Set(
      reviewInvoices
        .map((invoice) => invoice?.metadata?.checksum)
        .filter((value) => typeof value === 'string' && value)
    );
    for (const checksum of Array.from(reviewEditingChecksums)) {
      if (!reviewChecksums.has(checksum)) {
        reviewEditingChecksums.delete(checksum);
      }
    }
  }

  // Render review invoices
  if (reviewTableBody) {
    if (reviewInvoices.length === 0) {
      reviewTableBody.innerHTML = '<tr><td colspan="10" class="text-center">No invoices in review</td></tr>';
    } else {
      reviewInvoices.forEach(invoice => {
        const row = createInvoiceRow(invoice, true);
        reviewTableBody.appendChild(row);
      });
    }
  }

  // Render archive invoices
  if (archiveTableBody) {
    if (archiveInvoices.length === 0) {
      archiveTableBody.innerHTML = '<tr><td colspan="5" class="text-center">No archived invoices</td></tr>';
    } else {
      archiveInvoices.forEach(invoice => {
        const row = createInvoiceRow(invoice, false);
        archiveTableBody.appendChild(row);
      });
    }
  }

  // Update bulk actions state
  updateBulkActionsState();
}

function resolveStoredInvoiceStatus(invoice) {
  const status = typeof invoice?.status === 'string' ? invoice.status : invoice?.metadata?.status;
  return status === 'review' ? 'review' : 'archive';
}

function createInvoiceRow(invoice, isReview) {
  const row = document.createElement('tr');

  const metadata = invoice.metadata || {};
  const extracted = invoice?.extracted && typeof invoice.extracted === 'object' ? invoice.extracted : {};
  const parsed = invoice?.data && typeof invoice.data === 'object' ? invoice.data : {};
  const selection =
    invoice?.reviewSelection && typeof invoice.reviewSelection === 'object'
      ? { ...invoice.reviewSelection }
      : {};

  const matchHints = extractInvoiceMatchHints(invoice);
  const vendorMatchHint = matchHints.vendor;
  const accountMatchHint = matchHints.account;
  const taxMatchHint = matchHints.taxCode;

  const checksum = metadata.checksum || '';
  row.dataset.checksum = checksum;

  const realmId = (metadata?.companyProfile && metadata.companyProfile.realmId) || selectedRealmId || '';
  const companyMetadata = realmId ? companyMetadataCache.get(realmId) || null : null;
  const isEditing = Boolean(isReview && checksum && reviewEditingChecksums.has(checksum));

  const resolveField = (field) => {
    const parsedValue = parsed[field];
    if (parsedValue !== undefined && parsedValue !== null && parsedValue !== '') {
      return parsedValue;
    }

    const extractedValue = extracted[field];
    if (extractedValue !== undefined && extractedValue !== null && extractedValue !== '') {
      return extractedValue;
    }

    return undefined;
  };

  const resolveNumericField = (field) => {
    const value = resolveField(field);
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value;
    }

    if (typeof value === 'string') {
      const parsedNumber = Number.parseFloat(value.replace(/[^\d.+-]/g, ''));
      if (!Number.isNaN(parsedNumber) && Number.isFinite(parsedNumber)) {
        return parsedNumber;
      }
    }

    return null;
  };

  const invoiceNumberRaw = resolveField('invoiceNumber');
  const invoiceNumber = invoiceNumberRaw !== undefined && invoiceNumberRaw !== null && invoiceNumberRaw !== ''
    ? String(invoiceNumberRaw)
    : metadata.invoiceFilename || metadata.originalName || '—';

  const invoiceDateRaw = resolveField('invoiceDate');
  const invoiceDate = formatDate(invoiceDateRaw);

  const invoiceVendorRaw =
    resolveField('vendor') || metadata.invoiceFilename || metadata.originalName || 'Unknown';
  const invoiceVendor = typeof invoiceVendorRaw === 'string' ? invoiceVendorRaw : String(invoiceVendorRaw);

  const vendorSettingsEntries =
    companyMetadata?.vendorSettings?.entries && typeof companyMetadata.vendorSettings.entries === 'object'
      ? companyMetadata.vendorSettings.entries
      : {};

  const selectedVendorId = sanitizeReviewSelectionId(selection.vendorId);
  let vendorId = selectedVendorId;
  let vendorMatchType = vendorId ? 'exact' : 'unknown';
  let vendorLabel = invoiceVendor;
  let vendorReason = '';
  let vendorMatch = null;

  if (!vendorId && vendorMatchHint?.status === 'exact' && vendorMatchHint.id) {
    vendorId = vendorMatchHint.id;
    vendorMatchType = 'exact';
  } else if (vendorMatchHint?.status && vendorMatchType === 'unknown') {
    vendorMatchType =
      vendorMatchHint.status === 'exact' && !vendorMatchHint.id
        ? 'uncertain'
        : vendorMatchHint.status;
  }

  if (vendorMatchHint?.reason) {
    vendorReason = vendorMatchHint.reason;
  }

  if (vendorMatchHint?.label) {
    vendorLabel = vendorMatchHint.label;
  }

  if (companyMetadata) {
    if (!vendorId) {
      vendorMatch = findBestVendorMatch(invoiceVendor, companyMetadata?.vendors?.items || []);
      if (vendorMatch && vendorMatch.score >= 0.9) {
        vendorId = vendorMatch.vendor.id;
        vendorMatchType = 'exact';
        vendorReason = 'Matched vendor name';
      } else if (vendorMatch && vendorMatch.score >= 0.7) {
        vendorMatchType = 'uncertain';
        vendorReason = `Possible match (${vendorMatch.matchedLabel})`;
      }
    }

    const vendorEntry =
      vendorId && companyMetadata?.vendors?.lookup
        ? companyMetadata.vendors.lookup.get(vendorId)
        : null;
    if (vendorEntry) {
      vendorLabel =
        vendorEntry.displayName ||
        vendorEntry.companyName ||
        vendorEntry.name ||
        vendorEntry.fullyQualifiedName ||
        vendorLabel;
    }
  }

  const vendorDefaults =
    vendorId && vendorSettingsEntries
      ? vendorSettingsEntries[vendorId] || null
      : null;
  const selectedVendorDefaults =
    selectedVendorId && vendorSettingsEntries ? vendorSettingsEntries[selectedVendorId] || null : null;

  const selectedAccountId = sanitizeReviewSelectionId(selection.accountId);
  let accountId = selectedAccountId;
  let accountMatchType = accountId ? 'exact' : 'unknown';
  let accountLabel = '—';
  let accountReason = '';

  if (!accountId && accountMatchHint?.status === 'exact' && accountMatchHint.id) {
    accountId = accountMatchHint.id;
    accountMatchType = 'exact';
  } else if (accountMatchHint?.status && accountMatchType === 'unknown') {
    accountMatchType =
      accountMatchHint.status === 'exact' && !accountMatchHint.id
        ? 'uncertain'
        : accountMatchHint.status;
  }

  if (accountMatchHint?.reason) {
    accountReason = accountMatchHint.reason;
  }

  if (!accountId && vendorDefaults?.accountId) {
    accountId = vendorDefaults.accountId;
    accountMatchType = 'exact';
    accountReason = 'Vendor default';
  }

  let accountEntry =
    accountId && companyMetadata?.accounts?.lookup
      ? companyMetadata.accounts.lookup.get(accountId)
      : null;

  if (accountEntry) {
    accountLabel =
      accountEntry.fullyQualifiedName || accountEntry.name || `Account ${accountEntry.id}`;
  } else if (accountMatchHint?.label) {
    accountLabel = accountMatchHint.label;
  } else if (parsed?.suggestedAccount?.name) {
    accountLabel = parsed.suggestedAccount.name;
    const aiConfidence = (parsed.suggestedAccount.confidence || '').toString().toLowerCase();
    if (aiConfidence === 'high') {
      accountMatchType = 'exact';
    } else if (aiConfidence === 'medium') {
      accountMatchType = 'uncertain';
    } else {
      accountMatchType = 'unknown';
    }
    if (!accountReason && parsed.suggestedAccount.reason) {
      accountReason = parsed.suggestedAccount.reason;
    }
  }

  const resolvedAccountSelection = sanitizeReviewSelectionId(selection.accountId);
  let accountSelectValue = resolvedAccountSelection;
  if (accountSelectValue) {
    accountId = accountSelectValue;
  } else if (accountId) {
    accountSelectValue = accountId;
  }

  let accountAutofilled = false;
  if (!accountSelectValue && accountMatchType === 'exact') {
    const matchedAccount = findAccountMatchByLabel(accountLabel, companyMetadata?.accounts?.items || []);
    if (matchedAccount && matchedAccount.account) {
      accountSelectValue = matchedAccount.account.id;
      accountId = matchedAccount.account.id;
      accountEntry = matchedAccount.account;
      accountLabel =
        matchedAccount.account.fullyQualifiedName ||
        matchedAccount.account.name ||
        accountLabel;
      accountAutofilled = true;
    }
  }

  if (accountAutofilled) {
    row.dataset.autofilledAccount = 'true';
    if (invoice && typeof invoice === 'object') {
      const currentSelection =
        invoice.reviewSelection && typeof invoice.reviewSelection === 'object'
          ? invoice.reviewSelection
          : (invoice.reviewSelection = {});
      currentSelection.accountId = accountSelectValue;
    }
  }

  let taxCodeId = sanitizeReviewSelectionId(selection.taxCodeId);
  let taxMatchType = taxCodeId ? 'exact' : 'unknown';

  if (!taxCodeId && taxMatchHint?.status === 'exact' && taxMatchHint.id) {
    taxCodeId = taxMatchHint.id;
    taxMatchType = 'exact';
  } else if (taxMatchHint?.status && taxMatchType === 'unknown') {
    taxMatchType =
      taxMatchHint.status === 'exact' && !taxMatchHint.id
        ? 'uncertain'
        : taxMatchHint.status;
  }

  if (!taxCodeId && vendorDefaults?.taxCodeId) {
    taxCodeId = vendorDefaults.taxCodeId;
    taxMatchType = 'exact';
  }

  const taxEntry =
    taxCodeId && companyMetadata?.taxCodes?.lookup
      ? companyMetadata.taxCodes.lookup.get(taxCodeId)
      : null;

  let taxCodeLabel = '—';
  if (taxEntry) {
    taxCodeLabel = formatTaxCodeLabel(taxEntry.displayName || taxEntry.name || `Tax Code ${taxEntry.id}`);
  } else if (taxMatchHint?.label) {
    taxCodeLabel = formatTaxCodeLabel(taxMatchHint.label);
  } else {
    const parsedTaxCodeRaw = parsed.taxCodeLabel || parsed.taxCode;
    const parsedTaxCodeText = parsedTaxCodeRaw !== undefined && parsedTaxCodeRaw !== null
      ? parsedTaxCodeRaw.toString()
      : '';
    if (/[A-Za-z%]/.test(parsedTaxCodeText)) {
      taxCodeLabel = formatTaxCodeLabel(parsedTaxCodeText);
      if (taxMatchType === 'unknown') {
        taxMatchType = 'uncertain';
      }
    }
  }

  const subtotal = resolveNumericField('subtotal');
  const vatAmount = resolveNumericField('vatAmount');
  const totalAmount = resolveNumericField('totalAmount');

  const netAmount = Number.isFinite(subtotal)
    ? subtotal
    : Number.isFinite(totalAmount) && Number.isFinite(vatAmount)
      ? Math.max(totalAmount - vatAmount, 0)
      : Number.isFinite(totalAmount)
        ? totalAmount
        : null;

  const vatResolved = Number.isFinite(vatAmount)
    ? vatAmount
    : Number.isFinite(totalAmount) && Number.isFinite(netAmount)
      ? Math.max(totalAmount - netAmount, 0)
      : null;

  const grossAmount = Number.isFinite(totalAmount)
    ? totalAmount
    : Number.isFinite(netAmount) && Number.isFinite(vatResolved)
      ? netAmount + vatResolved
      : netAmount;

  const currencyCode = parsed.currency || extracted.currency || '';
  const netDisplay = formatCurrencyAmount(netAmount, currencyCode);
  const vatDisplay = formatCurrencyAmount(vatResolved, currencyCode);
  const grossDisplay = formatCurrencyAmount(grossAmount, currencyCode);

  const rowMatchClass =
    vendorMatchType === 'exact' && accountMatchType === 'exact'
      ? 'match-exact'
      : vendorMatchType !== 'unknown' || accountMatchType !== 'unknown'
        ? 'match-uncertain'
        : 'match-unknown';
  row.classList.add(rowMatchClass);

  const invoiceSecondary = metadata.invoiceFilename || metadata.originalName || '';
  const vendorSecondary = invoiceVendor && invoiceVendor !== vendorLabel ? invoiceVendor : '';

  const vendorSelectValue = sanitizeReviewSelectionId(selection.vendorId) || vendorId || '';
  const taxSelectValue = sanitizeReviewSelectionId(selection.taxCodeId) || taxCodeId || '';

  const vendorAutoSelected =
    vendorMatchHint?.status === 'exact' && vendorMatchHint?.id && vendorMatchHint.id === vendorSelectValue;
  const accountAutoSelected =
    accountMatchHint?.status === 'exact' && accountMatchHint?.id && accountMatchHint.id === accountSelectValue;
  const taxAutoSelected =
    taxMatchHint?.status === 'exact' && taxMatchHint?.id && taxMatchHint.id === taxSelectValue;

  const vendorOptionsHtml = buildVendorSelectOptions(companyMetadata, vendorSelectValue);
  const accountOptionsHtml = buildAccountSelectOptions(companyMetadata, accountSelectValue);
  const taxOptionsHtml = buildTaxSelectOptions(companyMetadata, taxSelectValue);

  const vendorSelectDisabled = !companyMetadata || !Array.isArray(companyMetadata?.vendors?.items) || companyMetadata.vendors.items.length === 0;
  const accountSelectDisabled = !companyMetadata || !Array.isArray(companyMetadata?.accounts?.items) || companyMetadata.accounts.items.length === 0;
  const taxSelectDisabled = !companyMetadata || !Array.isArray(companyMetadata?.taxCodes?.items) || companyMetadata.taxCodes.items.length === 0;
  const hasVendorSelection = Boolean(vendorSelectValue);
  const accountHasSelection = Boolean(accountSelectValue);
  const taxHasSelection = Boolean(taxSelectValue);

  const vendorMatchClass = resolveMatchClass(vendorMatchType, hasVendorSelection);
  const accountMatchClass = resolveMatchClass(accountMatchType, accountHasSelection);
  const taxMatchClass = resolveMatchClass(taxMatchType, taxHasSelection);

  const netInputValue = formatAmountInputValue(Number.isFinite(netAmount) ? netAmount : null);
  const vatInputValue = formatAmountInputValue(Number.isFinite(vatResolved) ? vatResolved : null);
  const grossInputValue = formatAmountInputValue(Number.isFinite(grossAmount) ? grossAmount : null);

  const previewVendorId = selectedVendorId || (vendorAutoSelected ? vendorSelectValue : null);
  const previewAccountId =
    selectedAccountId ||
    (accountAutoSelected ? accountSelectValue : null) ||
    sanitizeReviewSelectionId(selectedVendorDefaults?.accountId) ||
    (accountSelectValue || null);
  const previewDisabled = !previewVendorId || !previewAccountId;
  let previewDisabledReason = '';
  if (!previewVendorId) {
    previewDisabledReason = 'Match this invoice to a QuickBooks vendor before previewing.';
  } else if (!previewAccountId) {
    const vendorDescriptor = vendorLabel || 'this vendor';
    previewDisabledReason = `Assign a QuickBooks account to ${vendorDescriptor} in Account / category settings before previewing.`;
  }

  const previewButtonAttributes = [
    'type="button"',
    'class="review-action-button review-action-button--primary"',
    'data-action="preview"',
    `data-checksum="${escapeHtml(checksum)}"`,
  ];
  if (previewDisabled) {
    previewButtonAttributes.push('disabled', 'aria-disabled="true"');
    if (previewDisabledReason) {
      previewButtonAttributes.push(`title="${escapeHtml(previewDisabledReason)}"`);
    }
  }
  const previewButtonAttributeString = previewButtonAttributes.join(' ');
  const viewPdfButtonHtml = `
          <button type="button" class="review-action-button review-action-button--secondary"
                  data-action="view-pdf" data-checksum="${escapeHtml(checksum)}"
                  title="Open source PDF in a new window" aria-label="Open source PDF in a new window">
            View PDF
          </button>`;

  const vendorSelectClasses = ['review-field', 'review-field--inline', 'select-field', vendorMatchClass];
  const filteredVendorSelectClasses = vendorSelectClasses.filter(Boolean);
  const vendorSelectAttributes = [
    `id="review-vendor-${escapeHtml(checksum)}"`,
    'data-field="vendorId"',
    'aria-label="Vendor"',
  ];

  if (!isEditing || vendorSelectDisabled) {
    vendorSelectAttributes.push('disabled');
  }

  if (!isEditing) {
    vendorSelectAttributes.push('tabindex="-1"', 'data-readonly="true"', 'aria-disabled="true"');
  } else if (vendorSelectDisabled) {
    vendorSelectAttributes.push('aria-disabled="true"');
  }

  if (vendorAutoSelected && vendorSelectValue) {
    vendorSelectAttributes.push('data-auto-selected="true"');
    vendorSelectAttributes.push(`data-auto-value="${escapeHtml(vendorSelectValue)}"`);
  }

  const vendorControl = `
        <select class="${filteredVendorSelectClasses.join(' ')}" ${vendorSelectAttributes.join(' ')}>
          ${vendorOptionsHtml}
        </select>
      `;

  const accountSelectClasses = ['review-field', 'review-field--inline', 'select-field', accountMatchClass];
  const filteredAccountSelectClasses = accountSelectClasses.filter(Boolean);
  const accountSelectAttributes = [
    `id="review-account-${escapeHtml(checksum)}"`,
    'data-field="accountId"',
    'aria-label="Account"',
  ];

  if (!isEditing || accountSelectDisabled) {
    accountSelectAttributes.push('disabled');
  }

  if (!isEditing) {
    accountSelectAttributes.push('tabindex="-1"', 'data-readonly="true"', 'aria-disabled="true"');
  } else if (accountSelectDisabled) {
    accountSelectAttributes.push('aria-disabled="true"');
  }

  if (accountAutofilled) {
    accountSelectAttributes.push('data-autofilled="true"');
  }

  if (accountAutoSelected && accountSelectValue) {
    accountSelectAttributes.push('data-auto-selected="true"');
    accountSelectAttributes.push(`data-auto-value="${escapeHtml(accountSelectValue)}"`);
  }

  const accountControl = `
        <select class="${filteredAccountSelectClasses.join(' ')}" ${accountSelectAttributes.join(' ')}>
          ${accountOptionsHtml}
        </select>
      `;

  const taxSelectClasses = ['review-field', 'review-field--inline', 'select-field', taxMatchClass];
  const filteredTaxSelectClasses = taxSelectClasses.filter(Boolean);
  const taxSelectAttributes = [
    `id="review-tax-${escapeHtml(checksum)}"`,
    'data-field="taxCodeId"',
    'aria-label="Tax code"',
  ];

  if (!isEditing || taxSelectDisabled) {
    taxSelectAttributes.push('disabled');
  }

  if (!isEditing) {
    taxSelectAttributes.push('tabindex="-1"', 'data-readonly="true"', 'aria-disabled="true"');
  } else if (taxSelectDisabled) {
    taxSelectAttributes.push('aria-disabled="true"');
  }

  if (taxAutoSelected && taxSelectValue) {
    taxSelectAttributes.push('data-auto-selected="true"');
    taxSelectAttributes.push(`data-auto-value="${escapeHtml(taxSelectValue)}"`);
  }

  const taxControl = `
        <select class="${filteredTaxSelectClasses.join(' ')}" ${taxSelectAttributes.join(' ')}>
          ${taxOptionsHtml}
        </select>
      `;

  const netControl = isEditing
    ? `
        <input id="review-net-${escapeHtml(checksum)}" class="review-field review-field--inline review-field--amount"
               data-field="netAmount" type="number" inputmode="decimal" step="0.01"
               value="${escapeHtml(netInputValue)}" aria-label="Net amount">
      `
    : '';

  const vatControl = isEditing
    ? `
        <input id="review-vat-${escapeHtml(checksum)}" class="review-field review-field--inline review-field--amount"
               data-field="vatAmount" type="number" inputmode="decimal" step="0.01"
               value="${escapeHtml(vatInputValue)}" aria-label="VAT amount">
      `
    : '';

  const grossControl = isEditing
    ? `
        <input id="review-total-${escapeHtml(checksum)}" class="review-field review-field--inline review-field--amount"
               data-field="totalAmount" type="number" inputmode="decimal" step="0.01"
               value="${escapeHtml(grossInputValue)}" aria-label="Total amount">
      `
    : '';

  const invoiceTitleAttr = invoiceSecondary
    ? ` title="${escapeHtml(invoiceSecondary)}"`
    : '';

  const vendorTitleParts = [];
  if (vendorSecondary) {
    vendorTitleParts.push(vendorSecondary);
  }
  if (vendorReason) {
    vendorTitleParts.push(vendorReason);
  }
  if (vendorAutoSelected) {
    vendorTitleParts.push('Auto-selected from QuickBooks metadata');
  }
  const vendorTitleAttr = vendorTitleParts.length
    ? ` title="${escapeHtml(vendorTitleParts.join(' • '))}"`
    : '';

  const accountTitleParts = [];
  if (accountReason) {
    accountTitleParts.push(accountReason);
  }
  if (accountAutofilled && !accountReason) {
    accountTitleParts.push('Auto-filled from QuickBooks metadata');
  }
  const accountTitleAttr = accountTitleParts.length
    ? ` title="${escapeHtml(accountTitleParts.join(' • '))}"`
    : '';

  const taxTitleParts = [];
  if (taxCodeLabel && taxCodeLabel !== '—') {
    taxTitleParts.push(taxCodeLabel);
  }
  if (taxMatchHint?.label && taxMatchHint.label !== taxCodeLabel) {
    taxTitleParts.push(taxMatchHint.label);
  }
  if (taxMatchHint?.reason) {
    taxTitleParts.push(taxMatchHint.reason);
  }
  if (taxAutoSelected) {
    taxTitleParts.push('Auto-selected from QuickBooks metadata');
  }
  const taxTitleAttr = taxTitleParts.length
    ? ` title="${escapeHtml(taxTitleParts.join(' • '))}"`
    : '';

  const vendorCellContent = `
        <div class="cell-stack"${vendorTitleAttr}>
          ${vendorControl}
        </div>
      `;

  const accountCellContent = `
        <div class="cell-stack"${accountTitleAttr}>
          ${accountControl}
        </div>
      `;

  const taxCellContent = `
        <div class="cell-stack"${taxTitleAttr}>
          ${taxControl}
        </div>
      `;

  const netCellContent = isEditing
    ? netControl
    : `<span class="cell-inline-label">${escapeHtml(netDisplay)}</span>`;

  const vatCellContent = isEditing
    ? vatControl
    : `<span class="cell-inline-label">${escapeHtml(vatDisplay)}</span>`;

  const grossCellContent = isEditing
    ? grossControl
    : `<span class="cell-inline-label">${escapeHtml(grossDisplay)}</span>`;

  if (isReview) {
    row.innerHTML = `
      <td class="cell-select">
        <input type="checkbox" class="form-check-input review-checkbox"
               data-checksum="${escapeHtml(checksum)}"
               ${reviewSelectedChecksums.has(checksum) ? 'checked' : ''}>
      </td>
      <td class="cell-invoice"${invoiceTitleAttr}>
        <span class="cell-inline-label">${escapeHtml(invoiceNumber || '—')}</span>
      </td>
      <td class="cell-date">${escapeHtml(invoiceDate)}</td>
      <td class="cell-vendor">
        ${vendorCellContent}
      </td>
      <td class="cell-account">
        ${accountCellContent}
      </td>
      <td class="cell-tax">
        ${taxCellContent}
      </td>
      <td class="numeric cell-amount">
        <div class="cell-inline" title="${escapeHtml(netDisplay)}">
          ${netCellContent}
        </div>
      </td>
      <td class="numeric cell-amount">
        <div class="cell-inline" title="${escapeHtml(vatDisplay)}">
          ${vatCellContent}
        </div>
      </td>
      <td class="numeric cell-amount">
        <div class="cell-inline" title="${escapeHtml(grossDisplay)}">
          ${grossCellContent}
        </div>
      </td>
      <td class="cell-actions${isEditing ? ' is-editing' : ''}">
        <div class="review-row-actions">
          ${isEditing
            ? `
          <button type="button" class="review-action-button review-action-button--primary" disabled aria-disabled="true"
                  title="Save changes before previewing.">
            Preview
          </button>
${viewPdfButtonHtml}
          <button type="button" class="review-action-button review-action-button--primary"
                  data-action="save" data-checksum="${escapeHtml(checksum)}">
            Save
          </button>
          <button type="button" class="review-action-button review-action-button--secondary"
                  data-action="cancel" data-checksum="${escapeHtml(checksum)}">
            Cancel
          </button>
          <button type="button" class="review-action-button review-action-button--danger"
                  data-action="delete" data-checksum="${escapeHtml(checksum)}">
            Delete
          </button>
          `
            : `
          <button ${previewButtonAttributeString}>
            Preview
          </button>
${viewPdfButtonHtml}
          <button type="button" class="review-action-button review-action-button--secondary"
                  data-action="edit" data-checksum="${escapeHtml(checksum)}">
            Edit
          </button>
          <button type="button" class="review-action-button review-action-button--danger"
                  data-action="delete" data-checksum="${escapeHtml(checksum)}">
            Delete
          </button>
          `}
        </div>
      </td>
    `;
  } else {
    row.innerHTML = `
      <td class="cell-invoice"${invoiceTitleAttr}>
        <span class="cell-inline-label">${escapeHtml(invoiceNumber || '—')}</span>
      </td>
      <td>${escapeHtml(vendorLabel)}</td>
      <td class="cell-date">${escapeHtml(invoiceDate)}</td>
      <td class="numeric">${escapeHtml(grossDisplay)}</td>
      <td class="cell-actions">
        <div class="review-row-actions">
          <button ${previewButtonAttributeString}>
            Preview
          </button>
${viewPdfButtonHtml}
          <button type="button" class="review-action-button review-action-button--danger"
                  data-action="delete" data-checksum="${escapeHtml(checksum)}">
            Delete
          </button>
        </div>
      </td>
    `;
  }

  return row;
}

function buildVendorSelectOptions(metadata, selectedValue) {
  const vendors = Array.isArray(metadata?.vendors?.items) ? metadata.vendors.items : [];
  const options = ['<option value="">Select vendor</option>'];

  if (!vendors.length) {
    return options.join('');
  }

  const sorted = [...vendors].sort((left, right) => {
    const leftLabel = (left?.displayName || left?.companyName || left?.fullyQualifiedName || left?.id || '').toLowerCase();
    const rightLabel = (right?.displayName || right?.companyName || right?.fullyQualifiedName || right?.id || '').toLowerCase();
    return leftLabel.localeCompare(rightLabel, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((vendor) => {
    const value = sanitizeReviewSelectionId(vendor?.id);
    if (!value) {
      return;
    }

    const baseLabel = vendor?.displayName || vendor?.companyName || vendor?.fullyQualifiedName || `Vendor ${vendor.id}`;
    const details = [];
    if (vendor?.companyName && vendor.companyName !== vendor.displayName) {
      details.push(vendor.companyName);
    }
    if (vendor?.email) {
      details.push(vendor.email);
    }

    const label = details.length ? `${baseLabel} • ${details.join(' • ')}` : baseLabel;
    const isSelected = value === selectedValue;
    options.push(
      `<option value="${escapeHtml(value)}"${isSelected ? ' selected' : ''}>${escapeHtml(label)}</option>`
    );
  });

  return options.join('');
}

function buildAccountSelectOptions(metadata, selectedValue) {
  const accounts = Array.isArray(metadata?.accounts?.items) ? metadata.accounts.items : [];
  const options = ['<option value="">Select account</option>'];

  if (!accounts.length) {
    return options.join('');
  }

  const sorted = [...accounts].sort((left, right) => {
    const leftLabel = (left?.fullyQualifiedName || left?.name || left?.id || '').toLowerCase();
    const rightLabel = (right?.fullyQualifiedName || right?.name || right?.id || '').toLowerCase();
    return leftLabel.localeCompare(rightLabel, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((account) => {
    const value = sanitizeReviewSelectionId(account?.id);
    if (!value) {
      return;
    }

    const baseLabel = account?.fullyQualifiedName || account?.name || `Account ${account.id}`;
    const parts = [baseLabel];
    if (account?.accountType) {
      parts.push(account.accountType);
    }
    if (account?.accountSubType && account.accountSubType !== account.accountType) {
      parts.push(account.accountSubType);
    }
    const label = parts.join(' • ');
    const isSelected = value === selectedValue;
    options.push(
      `<option value="${escapeHtml(value)}"${isSelected ? ' selected' : ''}>${escapeHtml(label)}</option>`
    );
  });

  return options.join('');
}

function buildTaxSelectOptions(metadata, selectedValue) {
  const taxCodes = Array.isArray(metadata?.taxCodes?.items) ? metadata.taxCodes.items : [];
  const options = ['<option value="">Select tax code</option>'];

  if (!taxCodes.length) {
    return options.join('');
  }

  const sorted = [...taxCodes].sort((left, right) => {
    const leftLabel = (left?.name || left?.id || '').toLowerCase();
    const rightLabel = (right?.name || right?.id || '').toLowerCase();
    return leftLabel.localeCompare(rightLabel, undefined, { sensitivity: 'base' });
  });

  sorted.forEach((code) => {
    const value = sanitizeReviewSelectionId(code?.id);
    if (!value) {
      return;
    }

    const baseLabel = formatTaxCodeLabel(code?.displayName || code?.name || `Tax Code ${code.id}`);
    const numericRate = typeof code?.rate === 'number' && Number.isFinite(code.rate) ? code.rate : null;
    const label = numericRate !== null ? `${baseLabel} • ${Number(numericRate).toFixed(2)}%` : baseLabel;
    const isSelected = value === selectedValue;
    options.push(
      `<option value="${escapeHtml(value)}"${isSelected ? ' selected' : ''}>${escapeHtml(label)}</option>`
    );
  });

  return options.join('');
}

function formatTaxCodeLabel(label) {
  if (label === null || label === undefined) {
    return '';
  }

  return label
    .toString()
    .trim()
    .replace(/(\d+)(\.0+)%/g, '$1%');
}

function formatAmountInputValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return '';
  }

  return numeric.toFixed(2);
}

function syncReviewSelectAllState() {
  if (!reviewSelectAllCheckbox || !reviewTableBody) {
    return;
  }

  const checkboxNodes = reviewTableBody.querySelectorAll('input.review-checkbox');
  if (!checkboxNodes.length) {
    reviewSelectAllCheckbox.checked = false;
    reviewSelectAllCheckbox.indeterminate = false;
    return;
  }

  let checkedCount = 0;
  checkboxNodes.forEach((checkbox) => {
    if (checkbox.checked) {
      checkedCount += 1;
    }
  });

  if (checkedCount === checkboxNodes.length) {
    reviewSelectAllCheckbox.checked = true;
    reviewSelectAllCheckbox.indeterminate = false;
    return;
  }

  if (checkedCount === 0) {
    reviewSelectAllCheckbox.checked = false;
    reviewSelectAllCheckbox.indeterminate = false;
    return;
  }

  reviewSelectAllCheckbox.checked = false;
  reviewSelectAllCheckbox.indeterminate = true;
}

function updateBulkActionsState() {
  const selections = Array.from(reviewSelectedChecksums);
  const selectionCount = selections.length;
  const hasSelections = selectionCount > 0;
  const editingSelections = selections.filter((checksum) => reviewEditingChecksums.has(checksum));
  const hasEditingSelections = editingSelections.length > 0;
  const allEditing = hasSelections && editingSelections.length === selectionCount;

  if (reviewSelectionCount) {
    reviewSelectionCount.textContent = hasSelections
      ? `${selectionCount} selected`
      : 'No invoices selected.';
  }

  setBulkButtonState(reviewBulkPreviewButton, selectionCount === 1);
  setBulkButtonState(reviewBulkEditButton, hasSelections);
  setBulkButtonState(reviewBulkSaveButton, allEditing);
  setBulkButtonState(reviewBulkCancelButton, hasEditingSelections);
  setBulkButtonState(reviewBulkArchiveButton, hasSelections);
  setBulkButtonState(reviewBulkDeleteButton, hasSelections);

  syncReviewSelectAllState();
}

function setBulkButtonState(button, enabled) {
  if (!button) {
    return;
  }

  button.disabled = !enabled;
}

async function handleImportVendorDefaults() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before importing vendor defaults.', 'error');
    return;
  }

  if (!importVendorDefaultsButton) {
    return;
  }

  const originalLabel = importVendorDefaultsButton.textContent;
  importVendorDefaultsButton.disabled = true;
  importVendorDefaultsButton.textContent = 'Importing…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/vendors/import-defaults`,
      { method: 'POST' }
    );
    const body = await response.json().catch(() => ({}));

    if (!response.ok) {
      const message = body?.error || 'Failed to import vendor defaults.';
      throw new Error(message);
    }

    const metadata = companyMetadataCache.get(selectedRealmId) || null;
    if (metadata) {
      const vendorSettingsMap =
        body?.vendorSettings && typeof body.vendorSettings === 'object'
          ? body.vendorSettings
          : metadata.vendorSettings?.entries || {};
      applyVendorSettingsStructure(metadata, vendorSettingsMap);
      companyMetadataCache.set(selectedRealmId, metadata);
      renderVendorList(metadata, 'No vendors available for this company.');
    }

    const appliedCount = Array.isArray(body?.applied) ? body.applied.length : 0;
    if (appliedCount > 0) {
      showStatus(
        globalStatus,
        `Imported defaults for ${appliedCount} vendor${appliedCount === 1 ? '' : 's'}.`,
        'success'
      );
    } else {
      showStatus(globalStatus, 'No new defaults were found for your vendors.', 'info');
    }
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to import vendor defaults.', 'error');
  } finally {
    importVendorDefaultsButton.disabled = false;
    importVendorDefaultsButton.textContent = originalLabel || 'Import suggested defaults';
  }
}

async function handleVendorSettingChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) {
    return;
  }

  if (!target.classList.contains('vendor-setting-select')) {
    return;
  }

  const vendorId = target.dataset.vendorId;
  const field = target.dataset.settingField;
  if (!vendorId || !field) {
    return;
  }

  const previousValue = target.dataset.previousValue ?? '';
  const nextValue = target.value ?? '';

  if (nextValue === previousValue) {
    return;
  }

  if (!selectedRealmId) {
    target.value = previousValue;
    return;
  }

  if (target.disabled) {
    return;
  }

  const payload = {};
  if (field === 'accountId' || field === 'taxCodeId') {
    payload[field] = nextValue || null;
  } else if (field === 'vatTreatment') {
    payload.vatTreatment = nextValue || null;
  } else {
    return;
  }

  const vendorName = target.dataset.vendorName || 'Vendor';

  target.disabled = true;
  target.classList.add('is-pending');
  target.setAttribute('aria-busy', 'true');

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/vendors/${encodeURIComponent(vendorId)}/settings`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      }
    );

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to update vendor defaults.';
      throw new Error(message);
    }

    const settings = body?.settings || {};
    const resolvedValue = resolveVendorSettingFieldValue(field, settings);
    target.value = resolvedValue;
    target.dataset.previousValue = resolvedValue;

    applyVendorSettingUpdate(selectedRealmId, vendorId, settings);
    showStatus(globalStatus, `Saved defaults for ${vendorName}.`, 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update vendor defaults.', 'error');
    target.value = previousValue;
    target.dataset.previousValue = previousValue;
  } finally {
    target.classList.remove('is-pending');
    target.removeAttribute('aria-busy');
    target.disabled = false;
  }
}

async function handleArchiveAction(event) {
  event.preventDefault();
  const target = event.target;
  const button = target.closest('button');

  if (!button || !button.dataset.action || !button.dataset.checksum) {
    return;
  }

  const action = button.dataset.action;
  const checksum = button.dataset.checksum;

  if (action === 'delete') {
    await handleIndividualDelete(checksum);
  } else if (action === 'preview') {
    await handleInvoicePreview(checksum);
  } else if (action === 'view-pdf') {
    await handleInvoicePdfView(checksum);
  }
}

async function handleReviewAction(event) {
  const target = event.target;
  if (target && target.closest && target.closest('input[type="checkbox"]')) {
    return;
  }

  const button = target && target.closest ? target.closest('button') : null;

  if (!button || !button.dataset.action || !button.dataset.checksum) {
    return;
  }

  event.preventDefault();

  const action = button.dataset.action;
  const checksum = button.dataset.checksum;

  if (action === 'delete') {
    await handleIndividualDelete(checksum);
  } else if (action === 'preview') {
    await handleInvoicePreview(checksum);
  } else if (action === 'view-pdf') {
    await handleInvoicePdfView(checksum);
  } else if (action === 'edit') {
    enterReviewEditMode(checksum);
  } else if (action === 'cancel') {
    exitReviewEditMode(checksum);
  } else if (action === 'save') {
    await handleReviewSave(checksum, button);
  }
}

function enterReviewEditMode(checksum) {
  if (!checksum) {
    return;
  }

  reviewEditingChecksums.add(checksum);
  renderInvoices();

  const scheduleFocus = typeof queueMicrotask === 'function'
    ? queueMicrotask
    : (cb) => Promise.resolve().then(cb);

  scheduleFocus(() => {
    const row = findReviewRowElement(checksum);
    if (!row) {
      return;
    }

    const focusTarget =
      row.querySelector('[data-field="vendorId"]:not([disabled])') ||
      row.querySelector('[data-field="accountId"]:not([disabled])') ||
      row.querySelector('[data-field="taxCodeId"]:not([disabled])') ||
      row.querySelector('[data-field="netAmount"]');

    if (focusTarget && typeof focusTarget.focus === 'function') {
      focusTarget.focus({ preventScroll: true });
    }
  });
}

function exitReviewEditMode(checksum) {
  if (!checksum) {
    return;
  }

  reviewEditingChecksums.delete(checksum);
  renderInvoices();
}

async function handleReviewSave(checksum, triggerButton) {
  if (!checksum) {
    return;
  }

  const invoice = findStoredInvoice(checksum);
  if (!invoice) {
    showStatus(globalStatus, 'Unable to locate invoice for editing.', 'error');
    return;
  }

  const row = findReviewRowElement(checksum);
  if (!row) {
    showStatus(globalStatus, 'Unable to locate invoice row for editing.', 'error');
    return;
  }

  const selectionUpdates = {};
  const amountUpdates = {};
  let hasChanges = false;

  const vendorSelect = row.querySelector('[data-field="vendorId"]');
  if (vendorSelect) {
    const newVendorId = sanitizeReviewSelectionId(vendorSelect.value);
    const currentVendorId = sanitizeReviewSelectionId(invoice?.reviewSelection?.vendorId);
    if (newVendorId !== currentVendorId) {
      selectionUpdates.vendorId = newVendorId;
      hasChanges = true;
    }
  }

  const accountSelect = row.querySelector('[data-field="accountId"]');
  if (accountSelect) {
    const newAccountId = sanitizeReviewSelectionId(accountSelect.value);
    const currentAccountId = sanitizeReviewSelectionId(invoice?.reviewSelection?.accountId);
    if (newAccountId !== currentAccountId) {
      selectionUpdates.accountId = newAccountId;
      hasChanges = true;
    }
  }

  const taxSelect = row.querySelector('[data-field="taxCodeId"]');
  if (taxSelect) {
    const newTaxCodeId = sanitizeReviewSelectionId(taxSelect.value);
    const currentTaxCodeId = sanitizeReviewSelectionId(invoice?.reviewSelection?.taxCodeId);
    if (newTaxCodeId !== currentTaxCodeId) {
      selectionUpdates.taxCodeId = newTaxCodeId;
      hasChanges = true;
    }
  }

  const netInput = row.querySelector('[data-field="netAmount"]');
  const vatInput = row.querySelector('[data-field="vatAmount"]');
  const totalInput = row.querySelector('[data-field="totalAmount"]');

  const netParsed = parseAmountInputValue(netInput);
  if (netParsed.provided) {
    if (!netParsed.valid) {
      showStatus(globalStatus, 'Enter a valid net amount (for example 71.22).', 'error');
      if (netInput && typeof netInput.focus === 'function') {
        netInput.focus({ preventScroll: true });
      }
      return;
    }
    const currentNet = normaliseAmountValue(invoice?.data?.subtotal);
    if (amountValuesDiffer(netParsed.value, currentNet)) {
      amountUpdates.netAmount = netParsed.value;
      hasChanges = true;
    }
  }

  const vatParsed = parseAmountInputValue(vatInput);
  if (vatParsed.provided) {
    if (!vatParsed.valid) {
      showStatus(globalStatus, 'Enter a valid VAT amount (for example 11.87).', 'error');
      if (vatInput && typeof vatInput.focus === 'function') {
        vatInput.focus({ preventScroll: true });
      }
      return;
    }
    const currentVat = normaliseAmountValue(invoice?.data?.vatAmount);
    if (amountValuesDiffer(vatParsed.value, currentVat)) {
      amountUpdates.vatAmount = vatParsed.value;
      hasChanges = true;
    }
  }

  const totalParsed = parseAmountInputValue(totalInput);
  if (totalParsed.provided) {
    if (!totalParsed.valid) {
      showStatus(globalStatus, 'Enter a valid total amount (for example 82.45).', 'error');
      if (totalInput && typeof totalInput.focus === 'function') {
        totalInput.focus({ preventScroll: true });
      }
      return;
    }
    const currentTotal = normaliseAmountValue(invoice?.data?.totalAmount);
    if (amountValuesDiffer(totalParsed.value, currentTotal)) {
      amountUpdates.totalAmount = totalParsed.value;
      hasChanges = true;
    }
  }

  if (!hasChanges) {
    exitReviewEditMode(checksum);
    return;
  }

  const payload = { ...selectionUpdates };
  if (Object.keys(amountUpdates).length) {
    payload.amounts = amountUpdates;
  }

  if (!Object.keys(payload).length) {
    exitReviewEditMode(checksum);
    return;
  }

  if (triggerButton) {
    triggerButton.disabled = true;
  }

  showStatus(globalStatus, 'Saving invoice updates…', 'info');

  try {
    const updatedInvoice = await patchInvoiceReviewSelection(checksum, payload);

    if (updatedInvoice && typeof updatedInvoice === 'object') {
      const index = storedInvoices.findIndex((entry) => entry?.metadata?.checksum === checksum);
      if (index !== -1) {
        storedInvoices[index] = updatedInvoice;
      }
    }

    reviewEditingChecksums.delete(checksum);
    renderInvoices();
    showStatus(globalStatus, 'Invoice updates saved.', 'success');
  } catch (error) {
    console.error('Failed to save review updates', error);
    if (triggerButton) {
      triggerButton.disabled = false;
    }
    showStatus(globalStatus, error.message || 'Failed to save invoice updates.', 'error');
  }
}

function findReviewRowElement(checksum) {
  if (!reviewTableBody || !checksum) {
    return null;
  }

  const selectorValue = escapeSelectorValue(checksum);
  return reviewTableBody.querySelector(`tr[data-checksum="${selectorValue}"]`);
}

function escapeSelectorValue(value) {
  if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') {
    return CSS.escape(value);
  }

  return value.replace(/['"\\\[\]]/g, '\\$&');
}

function parseAmountInputValue(input) {
  if (!input) {
    return { provided: false, valid: true, value: null };
  }

  const raw = typeof input.value === 'string' ? input.value.trim() : '';
  if (!raw) {
    return { provided: true, valid: true, value: null };
  }

  const normalised = raw.replace(/,/g, '');
  const numeric = Number(normalised);
  if (!Number.isFinite(numeric)) {
    return { provided: true, valid: false, value: null };
  }

  const rounded = Math.round(numeric * 100) / 100;
  return { provided: true, valid: true, value: rounded };
}

function normaliseAmountValue(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return null;
  }

  return Math.round(numeric * 100) / 100;
}

function amountValuesDiffer(left, right) {
  if (left === null && right === null) {
    return false;
  }

  if (left === null || right === null) {
    return true;
  }

  return Math.abs(Number(left) - Number(right)) > 0.005;
}

async function handleIndividualDelete(checksum) {
  const confirmed = await showDeleteConfirmation('invoice');

  if (!confirmed) {
    return;
  }

  showStatus(globalStatus, 'Deleting invoice...', 'info');

  try {
    await deleteInvoice(checksum);
    showStatus(globalStatus, 'Invoice deleted successfully.', 'success');
    await loadStoredInvoices();
  } catch (error) {
    console.error('Individual delete error:', error);
    showStatus(globalStatus, 'Failed to delete invoice. Please try again.', 'error');
  }
}

function handleReviewChange(event) {
  const target = event.target && event.target.matches ? event.target : null;
  if (!target) {
    return;
  }

  if (target.type === 'checkbox' && target.dataset.checksum) {
    const checksum = target.dataset.checksum;
    if (target.checked) {
      reviewSelectedChecksums.add(checksum);
    } else {
      reviewSelectedChecksums.delete(checksum);
    }

    updateBulkActionsState();
    return;
  }

  if (target.dataset && target.dataset.field) {
    handleReviewFieldOverride(target);
  }
}

function getReviewFieldInlineLabelKey(fieldElement) {
  if (!fieldElement || !fieldElement.dataset) {
    return '';
  }

  switch (fieldElement.dataset.field) {
    case 'vendorId':
      return 'vendor';
    case 'accountId':
      return 'account';
    case 'taxCodeId':
      return 'tax';
    default:
      return '';
  }
}

function getReviewFieldInlineLabelText(fieldElement) {
  if (!fieldElement) {
    return '';
  }

  if (fieldElement.tagName === 'SELECT') {
    const { selectedIndex, options } = fieldElement;
    if (selectedIndex >= 0 && options && options[selectedIndex]) {
      return (options[selectedIndex].textContent || '').trim();
    }

    return '';
  }

  if (typeof fieldElement.value === 'string') {
    return fieldElement.value.trim();
  }

  return '';
}

function updateInlineLabelForField(fieldElement) {
  if (!fieldElement) {
    return;
  }

  const statusContainer =
    fieldElement.closest('.cell-stack') || fieldElement.closest('.cell-inline');

  if (!statusContainer) {
    return;
  }

  const statusRow = statusContainer.querySelector('.cell-status-row') || statusContainer;
  if (!statusRow) {
    return;
  }

  const labelKey = getReviewFieldInlineLabelKey(fieldElement);
  const labelSelector = labelKey
    ? `.cell-inline-label[data-inline-field="${labelKey}"]`
    : '.cell-inline-label';

  const isAutoSelected = fieldElement.dataset.autoSelected === 'true';
  if (isAutoSelected) {
    const autoValue = sanitizeReviewSelectionId(fieldElement.dataset.autoValue);
    const currentValue = sanitizeReviewSelectionId(fieldElement.value);
    if (autoValue === currentValue) {
      const existingLabel = statusRow.querySelector(labelSelector);
      if (existingLabel) {
        existingLabel.remove();
      }
      return;
    }
  }

  let inlineLabel = statusRow.querySelector(labelSelector);
  if (!inlineLabel) {
    inlineLabel = document.createElement('span');
    inlineLabel.classList.add('cell-inline-label');
    if (labelKey) {
      inlineLabel.dataset.inlineField = labelKey;
    }
    statusRow.appendChild(inlineLabel);
  }

  const labelText = getReviewFieldInlineLabelText(fieldElement);
  inlineLabel.textContent = labelText;

  inlineLabel.hidden = false;
  inlineLabel.removeAttribute('hidden');
  if (inlineLabel.dataset && inlineLabel.dataset.inlineAutoHidden !== undefined) {
    delete inlineLabel.dataset.inlineAutoHidden;
  }

}

function handleReviewFieldOverride(fieldElement) {
  if (!fieldElement || typeof fieldElement !== 'object') {
    return;
  }

  if (fieldElement.dataset.autoSelected !== 'true') {
    return;
  }

  const autoValue = sanitizeReviewSelectionId(fieldElement.dataset.autoValue);
  const currentValue = sanitizeReviewSelectionId(fieldElement.value);
  if (autoValue === currentValue) {
    return;
  }

  fieldElement.dataset.autoSelected = 'false';
  delete fieldElement.dataset.autoValue;

  updateInlineLabelForField(fieldElement);

  const row = fieldElement.closest('tr[data-checksum]');
  if (!row) {
    return;
  }

  let badgeKey = null;
  switch (fieldElement.dataset.field) {
    case 'vendorId':
      badgeKey = 'vendor';
      break;
    case 'accountId':
      badgeKey = 'account';
      break;
    case 'taxCodeId':
      badgeKey = 'tax';
      break;
    default:
      break;
  }

  if (!badgeKey) {
    return;
  }

  const badge = row.querySelector(`[data-match-badge="${badgeKey}"]`);
  if (badge && badge.parentNode) {
    badge.parentNode.removeChild(badge);
  }
}

function handleReviewSelectAllChange(event) {
  if (!reviewTableBody || !reviewSelectAllCheckbox) {
    return;
  }

  const checkboxNodes = reviewTableBody.querySelectorAll('input.review-checkbox');
  if (!checkboxNodes.length) {
    reviewSelectedChecksums.clear();
    updateBulkActionsState();
    return;
  }

  const shouldSelectAll = Boolean(event?.target?.checked);

  reviewSelectedChecksums.clear();

  checkboxNodes.forEach((checkbox) => {
    checkbox.checked = shouldSelectAll;
    const checksum = checkbox.dataset.checksum;
    if (shouldSelectAll && checksum) {
      reviewSelectedChecksums.add(checksum);
    }
  });

  updateBulkActionsState();
}

function getSelectedReviewChecksums() {
  return Array.from(reviewSelectedChecksums);
}

async function handleBulkPreviewSelected() {
  const selections = getSelectedReviewChecksums();

  if (selections.length !== 1) {
    showStatus(globalStatus, 'Select a single invoice to preview.', 'info');
    return;
  }

  await handleInvoicePreview(selections[0]);
}

function handleBulkEditSelected() {
  const selections = getSelectedReviewChecksums();

  if (!selections.length) {
    showStatus(globalStatus, 'Select invoices before entering edit mode.', 'info');
    return;
  }

  selections.forEach((checksum) => enterReviewEditMode(checksum));
}

async function handleBulkSaveSelected() {
  const selections = getSelectedReviewChecksums().filter((checksum) => reviewEditingChecksums.has(checksum));

  if (!selections.length) {
    showStatus(globalStatus, 'Edit an invoice before saving changes.', 'info');
    return;
  }

  for (const checksum of selections) {
    // eslint-disable-next-line no-await-in-loop
    await handleReviewSave(checksum, null);
  }
}

function handleBulkCancelSelected() {
  const selections = getSelectedReviewChecksums().filter((checksum) => reviewEditingChecksums.has(checksum));

  if (!selections.length) {
    showStatus(globalStatus, 'No edited invoices to cancel.', 'info');
    return;
  }

  selections.forEach((checksum) => exitReviewEditMode(checksum));
}

async function handleBulkArchiveSelected() {
  if (!reviewSelectedChecksums.size) {
    showStatus(globalStatus, 'Select invoices before running a bulk action.', 'info');
    return;
  }

  const checksums = Array.from(reviewSelectedChecksums);
  const countLabel = `${checksums.length} invoice${checksums.length === 1 ? '' : 's'}`;
  showStatus(globalStatus, `Archiving ${countLabel}...`, 'info');

  try {
    const response = await fetch('/api/invoices/bulk-status', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ checksums, status: 'archive' })
    });

    if (response.ok) {
      const result = await response.json().catch(() => ({}));
      const rawSuccessful = Number(result?.successful);
      const rawFailed = Number(result?.failed);
      const successful = Number.isFinite(rawSuccessful) && rawSuccessful >= 0 ? rawSuccessful : 0;
      const failed = Number.isFinite(rawFailed) && rawFailed >= 0 ? rawFailed : 0;
      let archivedCount = successful;
      if (!archivedCount && failed === 0) {
        archivedCount = checksums.length;
      }

      const archivedLabel = `${archivedCount} invoice${archivedCount === 1 ? '' : 's'}`;
      const firstError = Array.isArray(result?.errors) && result.errors.length ? result.errors[0] : '';

      if (failed === 0) {
        showStatus(globalStatus, `Archived ${archivedLabel}.`, 'success');
      } else if (archivedCount > 0) {
        const detail = firstError ? ` ${firstError}` : '';
        showStatus(globalStatus, `Archived ${archivedLabel}. ${failed} failed.${detail}`, 'warning');
      } else {
        const detail = firstError ? ` ${firstError}` : '';
        showStatus(globalStatus, `Failed to archive selected invoices.${detail}`, 'error');
      }

      checksums.forEach((checksum) => reviewEditingChecksums.delete(checksum));
      reviewSelectedChecksums.clear();
      updateBulkActionsState();
      await loadStoredInvoices();
      return;
    }

    const responseBody = await response.json().catch(() => ({}));
    const errorMessage = responseBody?.error || `Bulk archive failed with status ${response.status}.`;
    console.warn('Falling back to individual archive requests:', errorMessage);
    await performBulkStatusIndividual(checksums, 'archive', errorMessage);
  } catch (error) {
    console.error('Bulk archive error:', error);
    await performBulkStatusIndividual(checksums, 'archive', error?.message || 'Bulk archive failed.');
  }
}

async function handleBulkDeleteSelected() {
  if (!reviewSelectedChecksums.size) {
    showStatus(globalStatus, 'Select invoices before running a bulk action.', 'info');
    return;
  }

  const checksums = Array.from(reviewSelectedChecksums);
  const confirmed = await showDeleteConfirmation(`${checksums.length} invoice${checksums.length === 1 ? '' : 's'}`);

  if (!confirmed) {
    return;
  }

  // Show progress
  showStatus(globalStatus, `Deleting ${checksums.length} invoice${checksums.length === 1 ? '' : 's'}...`, 'info');

  try {
    // Try bulk delete endpoint first
    const response = await fetch('/api/invoices/bulk-delete', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ checksums })
    });

    if (response.ok) {
      const result = await response.json();
      const successful = result.successful || 0;
      const failed = result.failed || 0;

      if (failed === 0) {
        showStatus(globalStatus, `Successfully deleted ${successful} invoice${successful === 1 ? '' : 's'}.`, 'success');
      } else {
        showStatus(globalStatus, `Deleted ${successful} invoice${successful === 1 ? '' : 's'}. ${failed} failed.`, 'warning');
      }

      // Clear selections and refresh
      reviewSelectedChecksums.clear();
      await loadStoredInvoices();
    } else {
      // Fallback to individual deletions
      await performBulkDeleteIndividual(checksums);
    }
  } catch (error) {
    console.error('Bulk delete error:', error);
    // Fallback to individual deletions
    await performBulkDeleteIndividual(checksums);
  }
}

async function performBulkStatusIndividual(checksums, status, failureReason) {
  if (!Array.isArray(checksums) || !checksums.length) {
    return;
  }

  const results = await Promise.allSettled(
    checksums.map((checksum) => updateInvoiceStatus(checksum, status))
  );

  const successfulChecksums = [];
  const failedResults = [];

  results.forEach((result, index) => {
    const checksum = checksums[index];
    if (result.status === 'fulfilled') {
      successfulChecksums.push(checksum);
    } else {
      const errorMessage = result.reason && result.reason.message
        ? result.reason.message
        : failureReason || 'Update failed.';
      failedResults.push({ checksum, errorMessage });
    }
  });

  const successfulCount = successfulChecksums.length;
  const failedCount = failedResults.length;
  const verb = status === 'archive' ? 'Archived' : 'Updated';
  const suffix = status === 'review' ? ' to review' : '';

  let message = '';
  let state = 'success';

  if (failedCount === 0) {
    const count = successfulCount || checksums.length;
    message = `${verb} ${count} invoice${count === 1 ? '' : 's'}${suffix}.`;
    state = 'success';
  } else if (successfulCount === 0) {
    const detail = failedResults[0]?.errorMessage || failureReason || 'Request failed.';
    message = `Failed to update selected invoices. ${detail}`;
    state = 'error';
  } else {
    const count = successfulCount;
    const detail = failedResults[0]?.errorMessage || failureReason || '';
    const detailSuffix = detail ? ` ${detail}` : '';
    message = `${verb} ${count} invoice${count === 1 ? '' : 's'}${suffix}. ${failedCount} failed.${detailSuffix}`;
    state = 'warning';
  }

  successfulChecksums.forEach((checksum) => reviewEditingChecksums.delete(checksum));
  reviewSelectedChecksums.clear();
  updateBulkActionsState();
  showStatus(globalStatus, message.trim(), state);
  await loadStoredInvoices();
}

async function performBulkDeleteIndividual(checksums) {
  const results = await Promise.allSettled(
    checksums.map(checksum => deleteInvoice(checksum))
  );

  const successful = results.filter(r => r.status === 'fulfilled').length;
  const failed = results.filter(r => r.status === 'rejected').length;

  if (failed === 0) {
    showStatus(globalStatus, `Successfully deleted ${successful} invoice${successful === 1 ? '' : 's'}.`, 'success');
  } else {
    showStatus(globalStatus, `Deleted ${successful} invoice${successful === 1 ? '' : 's'}. ${failed} failed.`, 'warning');
  }

  // Clear selections and refresh
  reviewSelectedChecksums.clear();
  await loadStoredInvoices();
}

async function deleteInvoice(checksum) {
  try {
    const response = await fetch(`/api/invoices/${checksum}`, {
      method: 'DELETE'
    });

    if (!response.ok) {
      throw new Error(`Failed to delete invoice: ${response.status}`);
    }

    return true;
  } catch (error) {
    console.error(`Error deleting invoice ${checksum}:`, error);
    throw error;
  }
}

async function updateInvoiceStatus(checksum, status) {
  try {
    const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/status`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ status })
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || `Failed to update invoice status (${response.status}).`;
      throw new Error(`Invoice ${checksum}: ${message}`);
    }

    return payload?.invoice || null;
  } catch (error) {
    const message = error?.message || 'Failed to update invoice status.';
    if (message.startsWith('Invoice ')) {
      throw error;
    }
    throw new Error(`Invoice ${checksum}: ${message}`);
  }
}

function showDeleteConfirmation(itemDescription) {
  return new Promise((resolve) => {
    // Create modal if it doesn't exist
    let modal = document.getElementById('delete-confirmation-modal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'delete-confirmation-modal';
      modal.className = 'modal';
      modal.style.display = 'none';
      modal.style.position = 'fixed';
      modal.style.zIndex = '1050';
      modal.style.left = '0';
      modal.style.top = '0';
      modal.style.width = '100%';
      modal.style.height = '100%';
      modal.style.backgroundColor = 'rgba(0,0,0,0.5)';
      modal.innerHTML = `
        <div style="position: relative; width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;">
          <div style="background: white; padding: 20px; border-radius: 8px; max-width: 400px; width: 90%;">
            <div style="margin-bottom: 15px;">
              <h5 style="margin: 0;">Confirm Delete</h5>
            </div>
            <div style="margin-bottom: 15px;">
              <p style="margin: 0;">Are you sure you want to delete <strong class="item-description"></strong>?</p>
              <p style="margin: 5px 0 0 0; color: #856404;">This action cannot be undone.</p>
            </div>
            <div style="display: flex; gap: 10px; justify-content: flex-end;">
              <button type="button" class="cancel-delete" style="padding: 8px 16px; border: 1px solid #ccc; background: #6c757d; color: white; border-radius: 4px; cursor: pointer;">Cancel</button>
              <button type="button" class="confirm-delete" style="padding: 8px 16px; border: none; background: #dc3545; color: white; border-radius: 4px; cursor: pointer;">Delete</button>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(modal);

      // Handle confirm button
      const confirmBtn = modal.querySelector('.confirm-delete');
      confirmBtn.addEventListener('click', () => {
        modal.style.display = 'none';
        resolve(true);
      });

      // Handle cancel button
      const cancelBtn = modal.querySelector('.cancel-delete');
      cancelBtn.addEventListener('click', () => {
        modal.style.display = 'none';
        resolve(false);
      });

      // Handle backdrop click
      modal.addEventListener('click', (e) => {
        if (e.target === modal) {
          modal.style.display = 'none';
          resolve(false);
        }
      });
    }

    // Update description
    const descriptionEl = modal.querySelector('.item-description');
    descriptionEl.textContent = itemDescription;

    // Show modal
    modal.style.display = 'block';
  });
}

function findStoredInvoice(checksum) {
  if (!checksum || !Array.isArray(storedInvoices)) {
    return null;
  }

  return storedInvoices.find((entry) => entry?.metadata?.checksum === checksum) || null;
}

function showQuickBooksPreviewModal() {
  if (!qbPreviewModal) {
    return;
  }

  qbPreviewModal.hidden = false;
  qbPreviewModal.setAttribute('aria-hidden', 'false');

  if (qbPreviewJson) {
    qbPreviewJson.scrollTop = 0;
  }

  if (!quickBooksPreviewEscapeHandler) {
    quickBooksPreviewEscapeHandler = (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        hideQuickBooksPreviewModal();
      }
    };
    document.addEventListener('keydown', quickBooksPreviewEscapeHandler);
  }

  if (qbPreviewCloseButton) {
    qbPreviewCloseButton.focus({ preventScroll: true });
  }
}

function buildQuickBooksPreviewSummary(invoice, metadata, context) {
  if (!invoice || !metadata || !context) {
    return '';
  }

  const { realmId, url, payload } = context;
  const companyEntry = quickBooksCompanies.find((company) => company.realmId === realmId) || null;
  const companyName = companyEntry?.companyName || companyEntry?.legalName || realmId;

  const selection = invoice?.reviewSelection && typeof invoice.reviewSelection === 'object'
    ? invoice.reviewSelection
    : {};

  const vendorId = sanitizeReviewSelectionId(selection.vendorId);
  const accountId = sanitizeReviewSelectionId(selection.accountId);
  const taxCodeId = sanitizeReviewSelectionId(selection.taxCodeId);

  const vendorEntry = vendorId && metadata?.vendors?.lookup instanceof Map
    ? metadata.vendors.lookup.get(vendorId)
    : null;
  const accountEntry = accountId && metadata?.accounts?.lookup instanceof Map
    ? metadata.accounts.lookup.get(accountId)
    : null;
  const taxEntry = taxCodeId && metadata?.taxCodes?.lookup instanceof Map
    ? metadata.taxCodes.lookup.get(taxCodeId)
    : null;

  const vendorName = vendorEntry?.displayName || vendorEntry?.companyName || vendorEntry?.name || 'Not selected';
  const accountName = accountEntry?.fullyQualifiedName || accountEntry?.name || 'Not selected';
  const taxCodeName = taxEntry?.name || (taxCodeId ? `Tax Code ${taxCodeId}` : 'Not selected');

  const payloadTotal = payload?.TotalAmt ?? invoice?.data?.totalAmount ?? null;
  const payloadCurrency = payload?.CurrencyRef?.value || invoice?.data?.currency || null;
  const totalDisplay = formatCurrencyAmount(payloadTotal, payloadCurrency);

  const lines = [
    `Company: ${companyName} (Realm ${realmId})`,
    url ? `Endpoint: ${url}` : 'Endpoint: —',
    `Vendor: ${vendorName}`,
    `Account: ${accountName}`,
    `Tax code: ${taxCodeName}`,
    `Invoice total: ${totalDisplay}`,
  ];

  return lines.join('\n');
}


async function copyQuickBooksPreviewPayload() {
  if (!lastQuickBooksPreviewPayload) {
    showStatus(globalStatus, 'Nothing to copy yet.', 'info');
    return;
  }

  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(lastQuickBooksPreviewPayload);
    } else {
      const textarea = document.createElement('textarea');
      textarea.value = lastQuickBooksPreviewPayload;
      textarea.setAttribute('readonly', 'true');
      textarea.style.position = 'absolute';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
    }
    showStatus(globalStatus, 'QuickBooks payload copied to clipboard.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, 'Failed to copy QuickBooks payload.', 'error');
  }
}

function revokeInvoicePdfObjectUrl() {
  if (!invoicePdfObjectUrl) {
    return;
  }

  try {
    URL.revokeObjectURL(invoicePdfObjectUrl);
  } catch (error) {
    console.warn('Unable to revoke invoice PDF object URL', error);
  }

  invoicePdfObjectUrl = null;
}

function renderInvoicePdfWindowContent(targetWindow, title, body) {
  if (!targetWindow || targetWindow.closed) {
    return;
  }

  try {
    const doc = targetWindow.document;
    doc.open();
    doc.write(`<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title>
      <style>
        :root { color-scheme: light dark; }
        body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
        .shell {
          min-height: 100vh;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 1rem;
          padding: 2rem;
          background: #f4f4f6;
          color: #1f2933;
        }
        .shell--error { background: #fff1f2; color: #9f1239; }
        .spinner {
          width: 2.5rem;
          height: 2.5rem;
          border-radius: 50%;
          border: 0.4rem solid rgba(0, 0, 0, 0.1);
          border-top-color: rgba(59, 130, 246, 0.9);
          animation: spin 0.85s linear infinite;
        }
        @keyframes spin {
          100% { transform: rotate(360deg); }
        }
        embed, iframe {
          border: 0;
          width: 100%;
          height: 100%;
          flex: 1 1 auto;
        }
        .pdf-container {
          position: fixed;
          inset: 0;
          display: flex;
          flex-direction: column;
          background: #111827;
        }
        .pdf-container > embed,
        .pdf-container > iframe {
          flex: 1;
        }
      </style>
    </head><body>${body}</body></html>`);
    doc.close();
  } catch (error) {
    console.warn('Unable to render invoice PDF content', error);
  }
}

function renderInvoicePdfLoading(targetWindow) {
  const body = `
    <main class="shell">
      <div class="spinner" role="presentation" aria-hidden="true"></div>
      <p>Loading invoice PDF…</p>
    </main>
  `;
  renderInvoicePdfWindowContent(targetWindow, 'Loading invoice PDF…', body);
}

function renderInvoicePdfError(targetWindow, message) {
  const body = `
    <main class="shell shell--error">
      <strong>Could not open the invoice PDF.</strong>
      <p>${escapeHtml(message)}</p>
    </main>
  `;
  renderInvoicePdfWindowContent(targetWindow, 'Invoice PDF unavailable', body);
}

function renderInvoicePdfEmbed(targetWindow, objectUrl) {
  const body = `
    <div class="pdf-container">
      <embed src="${escapeHtml(objectUrl)}" type="application/pdf" />
    </div>
  `;
  renderInvoicePdfWindowContent(targetWindow, 'Invoice PDF', body);
}

async function extractApiErrorMessage(response, defaultMessage) {
  const fallback = defaultMessage || 'Request failed.';
  if (!response || typeof response.clone !== 'function') {
    return fallback;
  }

  try {
    const cloned = response.clone();
    const contentType = cloned.headers?.get('content-type') || '';
    if (contentType.includes('application/json')) {
      const body = await cloned.json().catch(() => ({}));
      const message = body?.error || body?.message || null;
      if (message) {
        return String(message);
      }
    } else {
      const text = await cloned.text().catch(() => '');
      if (text) {
        try {
          const parsed = JSON.parse(text);
          if (parsed && typeof parsed === 'object') {
            const message = parsed.error || parsed.message || null;
            if (message) {
              return String(message);
            }
          }
        } catch (_error) {
          if (text.trim()) {
            return text.trim();
          }
        }
      }
    }
  } catch (error) {
    console.warn('Failed to extract API error message', error);
  }

  return response.statusText || fallback;
}

async function handleInvoicePdfView(checksum) {
  if (!checksum) {
    showStatus(globalStatus, 'Invoice reference is missing for PDF view.', 'error');
    return;
  }

  const targetUrl = `/api/invoices/${encodeURIComponent(checksum)}/file`;

  if (!invoicePdfWindow || invoicePdfWindow.closed) {
    invoicePdfWindow = window.open('', 'invoice-pdf', 'width=1024,height=768');

    if (invoicePdfWindow) {
      try {
        invoicePdfWindow.opener = null;
      } catch (_) {
        // Some browsers prevent modifying opener; ignore to preserve isolation.
      }
    }
  } else {
    invoicePdfWindow.focus();
  }

  if (!invoicePdfWindow) {
    showStatus(globalStatus, 'Allow pop-ups in your browser to view the invoice PDF.', 'error');
    return;
  }

  renderInvoicePdfLoading(invoicePdfWindow);
  invoicePdfWindow.focus();

  invoicePdfWindow.onbeforeunload = () => {
    revokeInvoicePdfObjectUrl();
    invoicePdfWindow = null;
  };

  try {
    const response = await fetch(targetUrl);

    if (!response.ok) {
      const message = await extractApiErrorMessage(response, 'Failed to load invoice PDF.');
      renderInvoicePdfError(invoicePdfWindow, message);
      showStatus(globalStatus, message, 'error');
      return;
    }

    const contentType = (response.headers?.get('content-type') || '').toLowerCase();
    if (!contentType.includes('application/pdf')) {
      const message = await extractApiErrorMessage(response, 'The invoice file is unavailable.');
      renderInvoicePdfError(invoicePdfWindow, message);
      showStatus(globalStatus, message, 'error');
      return;
    }

    const blob = await response.blob();
    revokeInvoicePdfObjectUrl();
    invoicePdfObjectUrl = URL.createObjectURL(blob);
    let navigatedToBlob = false;
    try {
      invoicePdfWindow.location.replace(invoicePdfObjectUrl);
      navigatedToBlob = true;
    } catch (navigationError) {
      console.warn('Failed to navigate popup directly to PDF blob', navigationError);
    }

    if (!navigatedToBlob) {
      renderInvoicePdfEmbed(invoicePdfWindow, invoicePdfObjectUrl);
    }
    invoicePdfWindow.focus();
  } catch (error) {
    console.error('Unable to load invoice PDF', error);
    revokeInvoicePdfObjectUrl();
    renderInvoicePdfError(invoicePdfWindow, 'Unable to load invoice PDF. Please try again.');
    showStatus(globalStatus, 'Unable to load invoice PDF. Please try again.', 'error');
  }
}

async function handleInvoicePreview(checksum) {
  if (!checksum) {
    showStatus(globalStatus, 'Invoice reference is missing for preview.', 'error');
    return;
  }

  const invoice = findStoredInvoice(checksum);
  if (!invoice) {
    showStatus(globalStatus, 'Unable to locate invoice for preview.', 'error');
    return;
  }

  const invoiceRealmId = sanitizeReviewSelectionId(invoice?.metadata?.companyProfile?.realmId);
  const realmId = invoiceRealmId || selectedRealmId || '';

  if (!realmId) {
    showStatus(globalStatus, 'Select a QuickBooks company before previewing.', 'error');
    return;
  }

  showStatus(globalStatus, 'Generating QuickBooks preview…', 'info');

  try {
    const metadata = await loadCompanyMetadata(realmId, { force: false }).catch((error) => {
      console.warn('Unable to load metadata for preview', error);
      return null;
    });

    if (!metadata) {
      throw new Error('QuickBooks metadata is unavailable. Refresh data and try again.');
    }

    const selection =
      invoice?.reviewSelection && typeof invoice.reviewSelection === 'object'
        ? invoice.reviewSelection
        : {};
    const selectedVendorId = sanitizeReviewSelectionId(selection.vendorId);
    if (!selectedVendorId) {
      throw new Error('Select a QuickBooks vendor before previewing.');
    }

    const vendorLookup =
      metadata?.vendors?.lookup instanceof Map ? metadata.vendors.lookup : null;
    const vendorSettingsEntries =
      metadata?.vendorSettings?.entries && typeof metadata.vendorSettings.entries === 'object'
        ? metadata.vendorSettings.entries
        : {};
    const selectedVendorDefaults = vendorSettingsEntries[selectedVendorId] || null;
    const accountIdFromSelection = sanitizeReviewSelectionId(selection.accountId);
    const effectiveAccountId =
      accountIdFromSelection ||
      (selectedVendorDefaults
        ? sanitizeReviewSelectionId(selectedVendorDefaults.accountId)
        : null);

    if (!effectiveAccountId) {
      const vendorEntry = vendorLookup ? vendorLookup.get(selectedVendorId) : null;
      const vendorDescriptor =
        vendorEntry?.displayName ||
        vendorEntry?.companyName ||
        vendorEntry?.name ||
        vendorEntry?.fullyQualifiedName ||
        'this vendor';
      throw new Error(
        `Assign a QuickBooks account to ${vendorDescriptor} in Account / category settings before previewing.`
      );
    }

    const params = new URLSearchParams({ invoiceId: checksum, realmId });
    const response = await fetch(`/api/preview-quickbooks?${params.toString()}`);
    const bodyText = await response.text();
    let body = {};
    if (bodyText) {
      try {
        body = JSON.parse(bodyText);
      } catch (error) {
        console.warn('Unable to parse QuickBooks preview body', error);
      }
    }

    if (!response.ok) {
      let message = body?.error || 'Failed to build QuickBooks preview.';
      const detailParts = [];
      if (Array.isArray(body?.missing) && body.missing.length) {
        detailParts.push(`Missing: ${body.missing.join(', ')}`);
      }
      if (body?.details && typeof body.details === 'object') {
        Object.entries(body.details).forEach(([key, value]) => {
          if (value === undefined || value === null || value === '') {
            return;
          }
          detailParts.push(`${key}: ${value}`);
        });
      }
      if (detailParts.length) {
        message = `${message} (${detailParts.join('; ')})`;
      }
      throw new Error(message);
    }

    const method = body?.method || 'POST';
    const url = body?.url || '';
    const payload = body?.payload || {};

    lastQuickBooksPreviewPayload = JSON.stringify(payload, null, 2);

    if (qbPreviewMethod) {
      qbPreviewMethod.textContent = method;
    }

    if (qbPreviewJson) {
      qbPreviewJson.textContent = lastQuickBooksPreviewPayload;
    }

    if (qbPreviewSummary) {
      const summary = buildQuickBooksPreviewSummary(invoice, metadata, {
        realmId,
        url,
        payload,
      });
      if (summary) {
        qbPreviewSummary.textContent = summary;
        qbPreviewSummary.hidden = false;
      } else {
        qbPreviewSummary.textContent = '';
        qbPreviewSummary.hidden = true;
      }
    }

    showQuickBooksPreviewModal();
    showStatus(globalStatus, 'QuickBooks preview generated.', 'success');
  } catch (error) {
    console.error('QuickBooks preview error', error);
    showStatus(globalStatus, error.message || 'Failed to generate QuickBooks preview.', 'error');
  }
}

function hideQuickBooksPreviewModal() {
  if (!qbPreviewModal) {
    return;
  }

  qbPreviewModal.hidden = true;
  qbPreviewModal.setAttribute('aria-hidden', 'true');

  if (quickBooksPreviewEscapeHandler) {
    document.removeEventListener('keydown', quickBooksPreviewEscapeHandler);
    quickBooksPreviewEscapeHandler = null;
  }
}

function deriveOneDriveBreadcrumbFromPath(rawPath) {
  const value = typeof rawPath === 'string' ? rawPath.trim() : '';
  if (!value) {
    return '';
  }

  let remainder = value;
  const drivePrefix = remainder.match(/^\/?drives\/[^/]+\/root:(.*)$/i);
  if (drivePrefix) {
    remainder = drivePrefix[1] || '';
  } else {
    const rootPrefix = remainder.match(/^\/?drive\/root:(.*)$/i);
    if (rootPrefix) {
      remainder = rootPrefix[1] || '';
    }
  }

  if (remainder.startsWith('/')) {
    remainder = remainder.slice(1);
  }

  if (!remainder) {
    return '';
  }

  const parts = remainder.split('/').map((part) => {
    if (!part) {
      return part;
    }
    try {
      return decodeURIComponent(part);
    } catch (error) {
      return part;
    }
  });

  return parts.filter(Boolean).join('/');
}

function deriveOneDriveItemBreadcrumb(entry) {
  if (!entry || typeof entry !== 'object') {
    return '';
  }
  const fromDisplay = typeof entry.displayPath === 'string' ? entry.displayPath.trim() : '';
  if (fromDisplay) {
    return fromDisplay;
  }
  return deriveOneDriveBreadcrumbFromPath(entry.path || '');
}

function formatOneDriveBreadcrumb(breadcrumb) {
  const value = typeof breadcrumb === 'string' ? breadcrumb.trim() : '';
  if (!value) {
    return '';
  }

  return value
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean)
    .join(' › ');
}

function setOneDriveInputBreadcrumb(input, breadcrumb) {
  if (!input) {
    return;
  }

  const value = typeof breadcrumb === 'string' ? breadcrumb.trim() : '';
  if (value) {
    input.dataset.breadcrumb = value;
  } else if (input.dataset && Object.prototype.hasOwnProperty.call(input.dataset, 'breadcrumb')) {
    delete input.dataset.breadcrumb;
  }
}

function getOneDriveInputBreadcrumb(input) {
  if (!input) {
    return '';
  }

  const stored = typeof input.dataset?.breadcrumb === 'string' ? input.dataset.breadcrumb.trim() : '';
  if (stored) {
    return stored;
  }

  return deriveOneDriveBreadcrumbFromPath(input.value);
}

function buildOneDriveFolderSummary(name, path, webUrl, fallback = '', breadcrumb = '') {
  const parts = [];
  if (name) {
    parts.push(name);
  }

  const friendlyPath = formatOneDriveBreadcrumb(breadcrumb || deriveOneDriveBreadcrumbFromPath(path));
  if (friendlyPath) {
    parts.push(friendlyPath);
  } else if (path) {
    parts.push(path);
  }

  if (!parts.length && webUrl) {
    parts.push(webUrl);
  }

  if (!parts.length) {
    return fallback || '';
  }

  return parts.join(' • ');
}

function handleOneDriveProcessedClear() {
  if (oneDriveProcessedClearButton?.disabled) {
    return;
  }

  if (oneDriveProcessedFolderIdInput) {
    oneDriveProcessedFolderIdInput.value = '';
    if (oneDriveProcessedFolderIdInput.dataset) {
      oneDriveProcessedFolderIdInput.dataset.cleared = 'true';
    }
  }
  if (oneDriveProcessedFolderPathInput) {
    oneDriveProcessedFolderPathInput.value = '';
    setOneDriveInputBreadcrumb(oneDriveProcessedFolderPathInput, '');
  }
  if (oneDriveProcessedFolderNameInput) {
    oneDriveProcessedFolderNameInput.value = '';
  }
  if (oneDriveProcessedFolderWebUrlInput) {
    oneDriveProcessedFolderWebUrlInput.value = '';
  }
  if (oneDriveProcessedFolderParentIdInput) {
    oneDriveProcessedFolderParentIdInput.value = '';
  }

  updateOneDriveProcessedSummaryFromInputs();
}

function openOneDriveBrowser(target) {
  if (!oneDriveBrowseModal) {
    return;
  }

  if ((target === 'monitored' || target === 'processed') && !isSharedOneDriveConfigured()) {
    showStatus(globalStatus, 'Connect the OneDrive account before browsing folders.', 'error');
    return;
  }

  const driveFilter = target === 'shared' ? null : sharedOneDriveSettings?.driveId || null;

  oneDriveBrowseState = {
    target,
    stack: [],
    items: [],
    selectedIndex: -1,
    currentDriveId: null,
    loading: false,
    requestToken: null,
    driveFilter,
  };

  if (driveFilter && target !== 'shared') {
    oneDriveBrowseState.stack = [
      {
        kind: 'drive',
        id: driveFilter,
        name: sharedOneDriveSettings?.driveName || driveFilter,
        driveId: driveFilter,
      },
    ];
  }

  lastOneDriveBrowseTrigger =
    target === 'shared'
      ? oneDriveSharedConnectButton || null
      : target === 'processed'
      ? oneDriveBrowseProcessedButton || null
      : oneDriveBrowseMonitoredButton || null;

  if (oneDriveBrowseConfirmButton) {
    oneDriveBrowseConfirmButton.disabled = true;
    if (target === 'processed') {
      oneDriveBrowseConfirmButton.textContent = 'Select processed folder';
    } else if (target === 'shared') {
      oneDriveBrowseConfirmButton.textContent = 'Use this base folder';
    } else {
      oneDriveBrowseConfirmButton.textContent = 'Select folder';
    }
  }

  if (oneDriveBrowseTitle) {
    if (target === 'processed') {
      oneDriveBrowseTitle.textContent = 'Choose processed OneDrive folder';
    } else if (target === 'shared') {
      oneDriveBrowseTitle.textContent = 'Choose OneDrive base folder (optional)';
    } else {
      oneDriveBrowseTitle.textContent = 'Choose OneDrive folder';
    }
  }

  setOneDriveBrowseWarning('');
  setOneDriveBrowseStatus('');
  renderOneDriveBrowserBreadcrumb();
  renderOneDriveBrowserList([]);

  oneDriveBrowseModal.hidden = false;
  oneDriveBrowseModal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  setOneDriveBrowseEscapeHandler(true);

  if (oneDriveBrowseBackButton) {
    oneDriveBrowseBackButton.disabled = !oneDriveBrowseState.stack.length;
  }

  if (driveFilter && target !== 'shared') {
    loadOneDriveChildren(driveFilter);
  } else {
    loadOneDriveDrives();
  }
}

function closeOneDriveBrowser() {
  if (!oneDriveBrowseModal || oneDriveBrowseModal.hidden) {
    return;
  }

  oneDriveBrowseModal.hidden = true;
  oneDriveBrowseModal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
  setOneDriveBrowseEscapeHandler(false);
  setOneDriveBrowseWarning('');
  setOneDriveBrowseStatus('');
  if (oneDriveBrowseList) {
    oneDriveBrowseList.innerHTML = '';
  }
  oneDriveBrowseState = null;
  if (lastOneDriveBrowseTrigger) {
    lastOneDriveBrowseTrigger.focus();
  }
  lastOneDriveBrowseTrigger = null;
}

function setOneDriveBrowseEscapeHandler(active) {
  if (active) {
    if (oneDriveBrowseEscapeHandler) {
      return;
    }
    oneDriveBrowseEscapeHandler = (event) => {
      if (event.key === 'Escape') {
        closeOneDriveBrowser();
      }
    };
    window.addEventListener('keydown', oneDriveBrowseEscapeHandler);
  } else if (oneDriveBrowseEscapeHandler) {
    window.removeEventListener('keydown', oneDriveBrowseEscapeHandler);
    oneDriveBrowseEscapeHandler = null;
  }
}

function setOneDriveBrowseWarning(message) {
  if (!oneDriveBrowseWarning) {
    return;
  }

  const text = typeof message === 'string' ? message.trim() : '';
  if (!text) {
    oneDriveBrowseWarning.hidden = true;
    oneDriveBrowseWarning.textContent = '';
  } else {
    oneDriveBrowseWarning.hidden = false;
    oneDriveBrowseWarning.textContent = text;
  }
}

function setOneDriveBrowseStatus(message, { tone = 'info' } = {}) {
  if (!oneDriveBrowseStatus) {
    return;
  }

  const text = typeof message === 'string' ? message.trim() : '';
  if (!text) {
    oneDriveBrowseStatus.hidden = true;
    oneDriveBrowseStatus.textContent = '';
    oneDriveBrowseStatus.classList.remove('is-error');
    return;
  }

  oneDriveBrowseStatus.hidden = false;
  oneDriveBrowseStatus.textContent = text;
  oneDriveBrowseStatus.classList.toggle('is-error', tone === 'error');
}

function setOneDriveBrowseLoading(isLoading, message) {
  if (!oneDriveBrowseState) {
    return;
  }

  oneDriveBrowseState.loading = Boolean(isLoading);

  if (isLoading) {
    setOneDriveBrowseStatus(message || 'Loading…');
    if (oneDriveBrowseConfirmButton) {
      oneDriveBrowseConfirmButton.disabled = true;
    }
  } else if (typeof message === 'string') {
    setOneDriveBrowseStatus(message);
  }
}

function renderOneDriveBrowserBreadcrumb() {
  if (!oneDriveBrowsePath) {
    return;
  }

  if (!oneDriveBrowseState || !Array.isArray(oneDriveBrowseState.stack) || !oneDriveBrowseState.stack.length) {
    const fallbackLabel =
      oneDriveBrowseState?.driveFilter && oneDriveBrowseState.target !== 'shared'
        ? sharedOneDriveSettings?.driveName || 'OneDrive account'
        : 'All drives';
    oneDriveBrowsePath.textContent = fallbackLabel;
    return;
  }

  const parts = oneDriveBrowseState.stack
    .map((entry) => entry?.name || entry?.id)
    .filter((part) => typeof part === 'string' && part.trim());

  oneDriveBrowsePath.textContent = parts.length ? parts.join(' / ') : 'All drives';
}

function renderOneDriveBrowserList(items) {
  if (!oneDriveBrowseList) {
    return;
  }

  oneDriveBrowseList.innerHTML = '';

  if (!Array.isArray(items) || !items.length) {
    const empty = document.createElement('li');
    empty.className = 'onedrive-browser-empty';
    const message = oneDriveBrowseState?.stack?.length
      ? 'No subfolders found.'
      : 'No drives available for browsing.';
    empty.textContent = message;
    oneDriveBrowseList.appendChild(empty);
    if (oneDriveBrowseConfirmButton) {
      oneDriveBrowseConfirmButton.disabled = true;
    }
    return;
  }

  items.forEach((item, index) => {
    const li = document.createElement('li');
    li.className = 'onedrive-browser-item';

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'onedrive-browser-entry';
    button.dataset.index = String(index);
    button.dataset.kind = item.kind;

    const name = document.createElement('span');
    name.className = 'onedrive-entry-name';
    name.textContent = item.name || item.id;
    button.appendChild(name);

    const meta = document.createElement('span');
    meta.className = 'onedrive-entry-meta';

    if (item.kind === 'drive') {
      const details = [item.owner, item.driveType].filter(Boolean);
      meta.textContent = details.join(' • ');
    } else if (item.kind === 'folder') {
      const friendlyBreadcrumb = formatOneDriveBreadcrumb(item.breadcrumb);
      meta.textContent =
        friendlyBreadcrumb || (typeof item.childCount === 'number' ? `${item.childCount} items` : 'Folder');
    }

    if (meta.textContent) {
      button.appendChild(meta);
    }

    if (oneDriveBrowseState?.selectedIndex === index) {
      button.classList.add('is-selected');
    }

    li.appendChild(button);
    oneDriveBrowseList.appendChild(li);
  });

  if (oneDriveBrowseConfirmButton) {
    const selected =
      typeof oneDriveBrowseState?.selectedIndex === 'number'
        ? items[oneDriveBrowseState.selectedIndex] || null
        : null;
    oneDriveBrowseConfirmButton.disabled = !selected || selected.kind !== 'folder';
  }
}

function selectOneDriveBrowserItem(index) {
  if (!oneDriveBrowseState || !Array.isArray(oneDriveBrowseState.items)) {
    return;
  }

  if (index < 0 || index >= oneDriveBrowseState.items.length) {
    return;
  }

  const item = oneDriveBrowseState.items[index];
  if (!item) {
    return;
  }

  oneDriveBrowseState.selectedIndex = index;

  if (oneDriveBrowseList) {
    oneDriveBrowseList.querySelectorAll('.onedrive-browser-entry').forEach((entry, entryIndex) => {
      if (entryIndex === index) {
        entry.classList.add('is-selected');
      } else {
        entry.classList.remove('is-selected');
      }
    });
  }

  if (oneDriveBrowseConfirmButton) {
    oneDriveBrowseConfirmButton.disabled = item.kind !== 'folder';
  }
}

function handleOneDriveBrowseListClick(event) {
  const button = event.target.closest('button.onedrive-browser-entry');
  if (!button || !oneDriveBrowseState) {
    return;
  }

  const index = Number.parseInt(button.dataset.index || '', 10);
  if (!Number.isFinite(index)) {
    return;
  }

  const item = oneDriveBrowseState.items[index];
  if (!item) {
    return;
  }

  if (item.kind === 'drive') {
    enterOneDriveDrive(item);
    return;
  }

  const wasSelected = oneDriveBrowseState.selectedIndex === index;
  selectOneDriveBrowserItem(index);

  if (item.kind === 'folder' && wasSelected) {
    enterOneDriveFolder(item);
  }
}

function handleOneDriveBrowseListDoubleClick(event) {
  const button = event.target.closest('button.onedrive-browser-entry');
  if (!button || !oneDriveBrowseState) {
    return;
  }

  const index = Number.parseInt(button.dataset.index || '', 10);
  if (!Number.isFinite(index)) {
    return;
  }

  const item = oneDriveBrowseState.items[index];
  if (!item) {
    return;
  }

  if (item.kind === 'drive') {
    enterOneDriveDrive(item);
  } else if (item.kind === 'folder') {
    enterOneDriveFolder(item);
  }
}

function handleOneDriveBrowseListKeydown(event) {
  if (!oneDriveBrowseState) {
    return;
  }

  if (event.key !== 'Enter' && event.key !== 'ArrowRight') {
    return;
  }

  const button = event.target.closest('button.onedrive-browser-entry');
  if (!button) {
    return;
  }

  const index = Number.parseInt(button.dataset.index || '', 10);
  if (!Number.isFinite(index)) {
    return;
  }

  const item = oneDriveBrowseState.items[index];
  if (!item) {
    return;
  }

  event.preventDefault();

  if (item.kind === 'drive') {
    enterOneDriveDrive(item);
    return;
  }

  if (oneDriveBrowseState.selectedIndex !== index) {
    selectOneDriveBrowserItem(index);
    return;
  }

  if (item.kind === 'folder') {
    enterOneDriveFolder(item);
  }
}

function enterOneDriveDrive(item) {
  if (!oneDriveBrowseState || !item?.id) {
    return;
  }

  oneDriveBrowseState.stack = [
    {
      kind: 'drive',
      id: item.id,
      name: item.name || item.id,
      driveId: item.id,
    },
  ];

  oneDriveBrowseState.selectedIndex = -1;
  renderOneDriveBrowserBreadcrumb();
  if (oneDriveBrowseBackButton) {
    oneDriveBrowseBackButton.disabled = false;
  }
  loadOneDriveChildren(item.id);
}

function enterOneDriveFolder(item) {
  if (!oneDriveBrowseState || !item?.id) {
    return;
  }

  const driveId = item.driveId || oneDriveBrowseState.currentDriveId;
  if (!driveId) {
    return;
  }

  oneDriveBrowseState.stack.push({
    kind: 'folder',
    id: item.id,
    name: item.name || item.id,
    driveId,
    path: item.path || null,
    parentId: item.parentId || null,
    breadcrumb: item.breadcrumb || null,
  });
  oneDriveBrowseState.selectedIndex = -1;
  renderOneDriveBrowserBreadcrumb();
  loadOneDriveChildren(driveId, item.id, item.path || null);
}

async function loadOneDriveDrives() {
  if (!oneDriveBrowseState) {
    return;
  }

  const token = { type: 'drives', at: Date.now() };
  oneDriveBrowseState.requestToken = token;
  setOneDriveBrowseLoading(true, 'Loading drives…');

  const driveFilter = oneDriveBrowseState.driveFilter || null;
  const query = driveFilter ? `?${new URLSearchParams({ driveId: driveFilter }).toString()}` : '';

  try {
    const response = await fetch(`/api/onedrive/drives${query}`);
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(payload?.error || 'Failed to load OneDrive drives.');
    }

    if (!oneDriveBrowseState || oneDriveBrowseState.requestToken !== token) {
      return;
    }

    const drives = Array.isArray(payload?.drives) ? payload.drives : [];
    const normalised = drives
      .map((drive) => ({
        kind: 'drive',
        id: drive?.id,
        name: drive?.name || drive?.id || 'Drive',
        owner: drive?.owner || null,
        driveType: drive?.driveType || null,
      }))
      .filter((drive) => drive.id)
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

    oneDriveBrowseState.stack = [];
    oneDriveBrowseState.items = normalised;
    oneDriveBrowseState.selectedIndex = -1;
    oneDriveBrowseState.currentDriveId = null;
    renderOneDriveBrowserBreadcrumb();
    renderOneDriveBrowserList(normalised);
    setOneDriveBrowseWarning(payload?.warning || '');

    if (!normalised.length && !payload?.warning) {
      const emptyMessage = driveFilter
        ? 'The OneDrive account or selected folder is unavailable or access is denied.'
        : 'No drives available for browsing.';
      setOneDriveBrowseStatus(emptyMessage, { tone: 'info' });
    } else if (!payload?.warning) {
      setOneDriveBrowseStatus('');
    }

    if (oneDriveBrowseBackButton) {
      oneDriveBrowseBackButton.disabled = true;
    }
  } catch (error) {
    setOneDriveBrowseWarning('');
    setOneDriveBrowseStatus(error.message || 'Failed to load OneDrive drives.', { tone: 'error' });
    renderOneDriveBrowserList([]);
  } finally {
    if (oneDriveBrowseState && oneDriveBrowseState.requestToken === token) {
      oneDriveBrowseState.requestToken = null;
    }
    setOneDriveBrowseLoading(false);
  }
}

async function loadOneDriveChildren(driveId, itemId = null, path = null) {
  if (!oneDriveBrowseState || !driveId) {
    return;
  }

  const token = { type: 'children', at: Date.now(), driveId, itemId };
  oneDriveBrowseState.requestToken = token;
  setOneDriveBrowseLoading(true, 'Loading folders…');

  const params = new URLSearchParams({ driveId });
  if (itemId) {
    params.set('itemId', itemId);
  } else if (path) {
    params.set('path', path);
  }

  try {
    const response = await fetch(`/api/onedrive/children?${params.toString()}`);
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(payload?.error || 'Failed to load OneDrive folder.');
    }

    if (!oneDriveBrowseState || oneDriveBrowseState.requestToken !== token) {
      return;
    }

    const items = Array.isArray(payload?.items) ? payload.items : [];
    const folders = items
      .filter((entry) => entry?.isFolder)
      .map((entry) => ({
        kind: 'folder',
        id: entry.id,
        name: entry.name || entry.id || 'Folder',
        path: entry.path || null,
        breadcrumb: deriveOneDriveItemBreadcrumb(entry) || null,
        driveId: entry.driveId || driveId,
        parentId: entry.parentId || null,
        childCount: Number.isFinite(entry.childCount) ? Number(entry.childCount) : null,
        webUrl: entry.webUrl || null,
      }))
      .filter((entry) => entry.id)
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

    oneDriveBrowseState.items = folders;
    oneDriveBrowseState.selectedIndex = -1;
    oneDriveBrowseState.currentDriveId = driveId;
    renderOneDriveBrowserList(folders);

    if (!folders.length) {
      setOneDriveBrowseStatus('No subfolders found.', { tone: 'info' });
    } else {
      setOneDriveBrowseStatus('');
    }

    if (oneDriveBrowseBackButton) {
      oneDriveBrowseBackButton.disabled = !oneDriveBrowseState.stack?.length;
    }
  } catch (error) {
    setOneDriveBrowseStatus(error.message || 'Failed to load OneDrive folder.', { tone: 'error' });
    renderOneDriveBrowserList([]);
  } finally {
    if (oneDriveBrowseState && oneDriveBrowseState.requestToken === token) {
      oneDriveBrowseState.requestToken = null;
    }
    setOneDriveBrowseLoading(false);
  }
}

function handleOneDriveBrowseBack() {
  if (!oneDriveBrowseState || !Array.isArray(oneDriveBrowseState.stack) || !oneDriveBrowseState.stack.length) {
    closeOneDriveBrowser();
    return;
  }

  oneDriveBrowseState.stack.pop();
  oneDriveBrowseState.selectedIndex = -1;

  const previous = oneDriveBrowseState.stack[oneDriveBrowseState.stack.length - 1] || null;

  if (!previous) {
    renderOneDriveBrowserBreadcrumb();
    loadOneDriveDrives();
    return;
  }

  renderOneDriveBrowserBreadcrumb();
  if (previous.kind === 'drive') {
    loadOneDriveChildren(previous.driveId || previous.id);
  } else {
    loadOneDriveChildren(previous.driveId, previous.id, previous.path || null);
  }
}

async function applyOneDriveBrowserSelection() {
  if (!oneDriveBrowseState || oneDriveBrowseState.selectedIndex < 0) {
    return;
  }

  const item = oneDriveBrowseState.items[oneDriveBrowseState.selectedIndex];
  if (!item || item.kind !== 'folder') {
    return;
  }

  if (oneDriveBrowseState.target === 'shared') {
    await persistSharedOneDriveSelection(item);
    return;
  }

  if (oneDriveBrowseState.target === 'processed') {
    if (oneDriveProcessedFolderIdInput) {
      oneDriveProcessedFolderIdInput.value = item.id || '';
      if (oneDriveProcessedFolderIdInput.dataset) {
        delete oneDriveProcessedFolderIdInput.dataset.cleared;
      }
    }
    if (oneDriveProcessedFolderPathInput) {
      oneDriveProcessedFolderPathInput.value = item.path || '';
      setOneDriveInputBreadcrumb(oneDriveProcessedFolderPathInput, item.breadcrumb || '');
    }
    if (oneDriveProcessedFolderNameInput) {
      oneDriveProcessedFolderNameInput.value = item.name || '';
    }
    if (oneDriveProcessedFolderWebUrlInput) {
      oneDriveProcessedFolderWebUrlInput.value = item.webUrl || '';
    }
    if (oneDriveProcessedFolderParentIdInput) {
      oneDriveProcessedFolderParentIdInput.value = item.parentId || '';
    }
    updateOneDriveProcessedSummaryFromInputs();
  } else {
    if (oneDriveFolderIdInput) {
      oneDriveFolderIdInput.value = item.id || '';
    }
    if (oneDriveFolderPathInput) {
      oneDriveFolderPathInput.value = item.path || '';
      setOneDriveInputBreadcrumb(oneDriveFolderPathInput, item.breadcrumb || '');
    }
    if (oneDriveFolderNameInput) {
      oneDriveFolderNameInput.value = item.name || '';
    }
    if (oneDriveFolderWebUrlInput) {
      oneDriveFolderWebUrlInput.value = item.webUrl || '';
    }
    if (oneDriveFolderParentIdInput) {
      oneDriveFolderParentIdInput.value = item.parentId || '';
    }
    updateOneDriveSelectionPreviewFromInputs();
  }

  closeOneDriveBrowser();
}

async function persistSharedOneDriveSelection(item) {
  if (!item || item.kind !== 'folder') {
    return;
  }

  const driveId = item.driveId || oneDriveBrowseState?.currentDriveId || sharedOneDriveSettings?.driveId || null;
  if (!driveId) {
    setOneDriveBrowseStatus('Unable to determine the selected drive.', { tone: 'error' });
    return;
  }

  const payload = {
    driveId,
  };

  if (item.id) {
    payload.folderId = item.id;
  }
  if (item.path) {
    payload.folderPath = item.path;
  }

  if (oneDriveBrowseConfirmButton) {
    oneDriveBrowseConfirmButton.disabled = true;
  }
  setOneDriveBrowseWarning('');
  setOneDriveBrowseLoading(true, 'Saving OneDrive connection…');

  try {
    const response = await fetch('/api/onedrive/settings', {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const body = await response.json().catch(() => ({}));

    if (!response.ok) {
      const message = body?.error || 'Failed to update the OneDrive connection.';
      throw new Error(message);
    }

    sharedOneDriveSettings = body?.settings || sharedOneDriveSettings;
    showStatus(globalStatus, 'OneDrive account connection updated.', 'success');
    closeOneDriveBrowser();
    renderSharedOneDriveSummary();
    renderOneDriveSettings();
  } catch (error) {
    console.error(error);
    setOneDriveBrowseStatus(error.message || 'Failed to update the OneDrive connection.', { tone: 'error' });
    if (oneDriveBrowseConfirmButton) {
      oneDriveBrowseConfirmButton.disabled = false;
    }
  } finally {
    setOneDriveBrowseLoading(false);
  }
}

function handleSharedOneDriveConnect() {
  openOneDriveBrowser('shared');
}

async function handleSharedOneDriveValidate() {
  if (!isSharedOneDriveConfigured()) {
    showStatus(globalStatus, 'Connect the OneDrive account before running a connection check.', 'error');
    return;
  }

  const payload = {
    driveId: sharedOneDriveSettings.driveId,
  };

  if (sharedOneDriveSettings.shareUrl) {
    payload.shareUrl = sharedOneDriveSettings.shareUrl;
  }
  if (sharedOneDriveSettings.folderId) {
    payload.folderId = sharedOneDriveSettings.folderId;
  } else if (sharedOneDriveSettings.folderPath) {
    payload.folderPath = sharedOneDriveSettings.folderPath;
  }

  const originalLabel = oneDriveSharedValidateButton ? oneDriveSharedValidateButton.textContent : '';
  if (oneDriveSharedValidateButton) {
    oneDriveSharedValidateButton.disabled = true;
    oneDriveSharedValidateButton.textContent = 'Validating…';
  }

  try {
    const response = await fetch('/api/onedrive/settings', {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const body = await response.json().catch(() => ({}));

    if (!response.ok) {
      const message = body?.error || 'Failed to run the OneDrive connection check.';
      throw new Error(message);
    }

    sharedOneDriveSettings = body?.settings || sharedOneDriveSettings;
    renderSharedOneDriveSummary();
    renderOneDriveSettings();
    showStatus(globalStatus, 'Connection check completed.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to run the OneDrive connection check.', 'error');
  } finally {
    if (oneDriveSharedValidateButton) {
      oneDriveSharedValidateButton.disabled = false;
      oneDriveSharedValidateButton.textContent = originalLabel || 'Run connection check';
    }
  }
}


async function refreshSharedOutlookSettings({ silent = false } = {}) {
  if (!outlookSharedCard) {
    sharedOutlookSettings = null;
    return;
  }

  try {
    const response = await fetch('/api/outlook/settings');
    if (!response.ok) {
      throw new Error('Unable to load Outlook mailbox settings.');
    }
    const payload = await response.json().catch(() => ({}));
    sharedOutlookSettings = payload?.settings || null;
  } catch (error) {
    if (!silent) {
      console.error('Failed to load shared Outlook settings', error);
      showStatus(globalStatus, error.message || 'Failed to load Outlook mailbox settings.', 'error');
    }
    sharedOutlookSettings = null;
  } finally {
    renderSharedOutlookSummary();
    renderOutlookSettings();
  }
}

function renderSharedOutlookSummary() {
  if (!outlookSharedCard) {
    return;
  }

  const settings = sharedOutlookSettings || null;
  const configured = Boolean(settings?.mailboxUserId);
  const status = settings?.status || (configured ? 'ready' : 'unconfigured');
  const baseFolder = settings?.baseFolder || null;

  setOutlookSharedStatusBadge(status);

  if (configured) {
    const mailboxLabel = settings.mailboxDisplayName || settings.mailboxUserId;
    outlookSharedSummaryText.textContent = `Monitoring mailbox ${mailboxLabel}.`;
    outlookSharedMailbox.textContent = mailboxLabel;
  } else {
    outlookSharedSummaryText.textContent = 'Shared Outlook mailbox is not configured yet.';
    outlookSharedMailbox.textContent = '—';
  }

  if (baseFolder?.path) {
    outlookSharedBaseFolder.textContent = formatOutlookFolderBreadcrumb(baseFolder.path);
  } else if (baseFolder?.displayName) {
    outlookSharedBaseFolder.textContent = baseFolder.displayName;
  } else {
    outlookSharedBaseFolder.textContent = configured ? 'Mailbox root' : '—';
  }

  if (settings?.lastValidatedAt) {
    outlookSharedLastValidated.textContent = formatTimestamp(settings.lastValidatedAt);
  } else {
    outlookSharedLastValidated.textContent = configured ? 'Not validated yet' : 'Never';
  }

  outlookSharedForm.hidden = true;
  outlookSharedForm.setAttribute('aria-hidden', 'true');
  outlookSharedForm.reset();
  outlookSharedMailboxInput.value = settings?.mailboxUserId || '';
  if (outlookSharedBrowseBaseButton) {
    outlookSharedBrowseBaseButton.disabled = !configured;
  }
  outlookSharedDisplayNameInput.value = settings?.mailboxDisplayName || '';
  outlookSharedBaseIdInput.value = baseFolder?.id || '';
  outlookSharedBasePathInput.value = baseFolder?.path || '';
  outlookSharedBaseNameInput.value = baseFolder?.displayName || '';
  outlookSharedBaseWebUrlInput.value = baseFolder?.webUrl || '';

  if (outlookSharedEditButton) {
    outlookSharedEditButton.disabled = false;
  }
}

function setOutlookSharedStatusBadge(status) {
  if (!outlookSharedStatus) {
    return;
  }

  outlookSharedStatus.classList.remove('status-pill--ready', 'status-pill--error', 'status-pill--warning', 'status-pill--muted');
  outlookSharedStatus.textContent = status === 'ready' ? 'Ready' : status === 'error' ? 'Error' : status === 'warning' ? 'Warning' : 'Unconfigured';

  if (status === 'ready') {
    outlookSharedStatus.classList.add('status-pill--ready');
  } else if (status === 'error') {
    outlookSharedStatus.classList.add('status-pill--error');
  } else if (status === 'warning') {
    outlookSharedStatus.classList.add('status-pill--warning');
  } else {
    outlookSharedStatus.classList.add('status-pill--muted');
  }
}

function handleOutlookSharedEdit() {
  if (!outlookSharedForm) {
    return;
  }

  outlookSharedForm.hidden = false;
  outlookSharedForm.setAttribute('aria-hidden', 'false');
  outlookSharedMailboxInput.focus();
}

function handleOutlookSharedCancel() {
  if (!outlookSharedForm) {
    return;
  }
  outlookSharedForm.hidden = true;
  outlookSharedForm.setAttribute('aria-hidden', 'true');
}

async function handleOutlookSharedSave(event) {
  event.preventDefault();
  if (!outlookSharedForm) {
    return;
  }

  const mailboxUserId = normaliseTextInput(outlookSharedMailboxInput?.value);
  const mailboxDisplayName = normaliseTextInput(outlookSharedDisplayNameInput?.value);
  const baseFolderId = normaliseTextInput(outlookSharedBaseIdInput?.value);
  const baseFolderPath = normaliseTextInput(outlookSharedBasePathInput?.value);
  const baseFolderName = normaliseTextInput(outlookSharedBaseNameInput?.value);
  const baseFolderWebUrl = normaliseTextInput(outlookSharedBaseWebUrlInput?.value);

  if (!mailboxUserId) {
    showStatus(globalStatus, 'Enter the mailbox user ID or UPN.', 'error');
    return;
  }

  const payload = {
    mailboxUserId,
    mailboxDisplayName: mailboxDisplayName || null,
  };

  if (baseFolderId || baseFolderPath) {
    payload.baseFolder = {
      id: baseFolderId || null,
      path: baseFolderPath || null,
      displayName: baseFolderName || null,
      webUrl: baseFolderWebUrl || null,
    };
  } else {
    payload.baseFolder = null;
  }

  const originalLabel = outlookSharedSaveButton ? outlookSharedSaveButton.textContent : '';
  if (outlookSharedSaveButton) {
    outlookSharedSaveButton.disabled = true;
    outlookSharedSaveButton.textContent = 'Saving…';
  }

  try {
    const response = await fetch('/api/outlook/settings', {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to update Outlook mailbox settings.';
      throw new Error(message);
    }

    sharedOutlookSettings = body?.settings || payload;
    renderSharedOutlookSummary();
    renderOutlookSettings();
    showStatus(globalStatus, 'Outlook mailbox settings saved.', 'success');
    outlookSharedForm.hidden = true;
    outlookSharedForm.setAttribute('aria-hidden', 'true');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update Outlook mailbox settings.', 'error');
  } finally {
    if (outlookSharedSaveButton) {
      outlookSharedSaveButton.disabled = false;
      outlookSharedSaveButton.textContent = originalLabel || 'Save mailbox';
    }
  }
}

function handleOutlookSharedBrowseBase() {
  if (!sharedOutlookSettings?.mailboxUserId) {
    showStatus(globalStatus, 'Configure the mailbox user before choosing a base folder.', 'error');
    return;
  }
  openOutlookBrowser('shared-base');
}

function renderOutlookSettings() {
  if (!outlookForm || !outlookEnabledInput) {
    return;
  }

  const company = getSelectedCompany();
  const config = company?.outlook || null;
  const hasCompany = Boolean(company);

  const inputs = [
    outlookPollIntervalInput,
    outlookMaxAttachmentInput,
    outlookMimeTypesInput,
    outlookEnabledInput,
    outlookBrowseButton,
    outlookClearFolderButton,
  ];

  inputs.forEach((input) => {
    if (input) {
      input.disabled = !hasCompany;
    }
  });

  if (outlookBrowseButton) {
    outlookBrowseButton.disabled = !hasCompany || !sharedOutlookSettings?.mailboxUserId;
  }
  if (outlookClearFolderButton) {
    outlookClearFolderButton.disabled = !hasCompany;
  }

  if (outlookSaveButton) {
    outlookSaveButton.disabled = !hasCompany;
  }
  if (outlookSyncButton) {
    outlookSyncButton.disabled = !hasCompany;
  }
  if (outlookResyncButton) {
    outlookResyncButton.disabled = !hasCompany;
  }
  if (outlookDisconnectButton) {
    outlookDisconnectButton.disabled = !hasCompany;
  }

  if (!hasCompany) {
    outlookForm.reset();
    outlookFolderSummary.textContent = MONITORED_DEFAULT_SUMMARY;
    if (outlookStatusContainer) {
      outlookStatusContainer.hidden = true;
    }
    return;
  }

  const monitoredFolder = config?.monitoredFolder || null;
  if (outlookFolderIdInput) {
    outlookFolderIdInput.value = monitoredFolder?.id || '';
  }
  if (outlookFolderPathInput) {
    outlookFolderPathInput.value = monitoredFolder?.path || '';
  }
  if (outlookFolderNameInput) {
    outlookFolderNameInput.value = monitoredFolder?.displayName || '';
  }
  if (outlookFolderWebUrlInput) {
    outlookFolderWebUrlInput.value = monitoredFolder?.webUrl || '';
  }
  if (outlookFolderParentIdInput) {
    outlookFolderParentIdInput.value = monitoredFolder?.parentId || '';
  }

  const pollIntervalMs = Number.parseInt(config?.pollIntervalMs, 10);
  if (outlookPollIntervalInput) {
    outlookPollIntervalInput.value = Number.isFinite(pollIntervalMs) ? Math.round(pollIntervalMs / 1000) : '';
  }

  const maxBytes = Number.parseInt(config?.maxAttachmentBytes, 10);
  if (outlookMaxAttachmentInput) {
    outlookMaxAttachmentInput.value = Number.isFinite(maxBytes)
      ? Math.max(1, Math.round(maxBytes / (1024 * 1024)))
      : '';
  }

  if (outlookMimeTypesInput) {
    const allowed = Array.isArray(config?.allowedMimeTypes) ? config.allowedMimeTypes : [];
    outlookMimeTypesInput.value = allowed.length ? allowed.join(', ') : '';
  }

  outlookEnabledInput.checked = config ? config.enabled !== false : false;
  updateOutlookFolderSummaryFromInputs();

  if (outlookSyncButton) {
    outlookSyncButton.disabled = !config || !outlookEnabledInput.checked;
  }
  if (outlookResyncButton) {
    outlookResyncButton.disabled = !config || !outlookEnabledInput.checked;
  }
  if (outlookDisconnectButton) {
    outlookDisconnectButton.disabled = !config;
  }

  if (!outlookStatusContainer) {
    return;
  }

  const hasStatus = Boolean(config);
  outlookStatusContainer.hidden = !hasStatus;
  if (!hasStatus) {
    return;
  }

  outlookStatusState.textContent = formatStatusLabel(config.status, outlookEnabledInput.checked ? 'Connected' : 'Disabled');

  const folderLabel = monitoredFolder?.path
    ? formatOutlookFolderBreadcrumb(monitoredFolder.path)
    : monitoredFolder?.displayName || 'Not selected';
  outlookStatusFolder.textContent = folderLabel;

  if (config.lastSyncAt) {
    outlookStatusLastSync.textContent = formatTimestamp(config.lastSyncAt);
  } else if (outlookEnabledInput.checked) {
    outlookStatusLastSync.textContent = 'Not run yet';
  } else {
    outlookStatusLastSync.textContent = 'Disabled';
  }

  outlookStatusResult.textContent = buildOutlookResultText(config, { enabled: outlookEnabledInput.checked });
}

function updateOutlookFolderSummaryFromInputs() {
  const id = normaliseTextInput(outlookFolderIdInput?.value);
  const path = normaliseTextInput(outlookFolderPathInput?.value);
  const name = normaliseTextInput(outlookFolderNameInput?.value);

  if (!id && !path) {
    outlookFolderSummary.textContent = MONITORED_DEFAULT_SUMMARY;
    return;
  }

  const label = formatOutlookFolderBreadcrumb(path || name || '');
  outlookFolderSummary.textContent = label || name || path || 'Selected folder';
}

function handleOutlookClearFolder() {
  if (!outlookFolderIdInput) {
    return;
  }
  outlookFolderIdInput.value = '';
  outlookFolderPathInput.value = '';
  outlookFolderNameInput.value = '';
  outlookFolderWebUrlInput.value = '';
  outlookFolderParentIdInput.value = '';
  updateOutlookFolderSummaryFromInputs();
}

async function handleOutlookSettingsSave(event) {
  event.preventDefault();
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before updating Outlook settings.', 'error');
    return;
  }

  const allowedMimeTypes = parseCommaSeparatedList(outlookMimeTypesInput?.value).map((value) => value.toLowerCase());
  const pollIntervalSeconds = outlookPollIntervalInput?.value
    ? Number.parseInt(outlookPollIntervalInput.value, 10)
    : null;
  const maxAttachmentMb = outlookMaxAttachmentInput?.value
    ? Number.parseInt(outlookMaxAttachmentInput.value, 10)
    : null;
  const enabled = outlookEnabledInput ? outlookEnabledInput.checked : false;

  const payload = {
    enabled,
    allowedMimeTypes,
  };

  if (enabled && !folderSelection?.id) {
    showStatus(globalStatus, 'Choose an Outlook folder before enabling automation.', 'error');
    return;
  }

  if (Number.isFinite(pollIntervalSeconds) && pollIntervalSeconds > 0) {
    payload.pollIntervalMs = Math.max(pollIntervalSeconds * 1000, 15000);
  }

  if (Number.isFinite(maxAttachmentMb) && maxAttachmentMb > 0) {
    payload.maxAttachmentBytes = Math.max(maxAttachmentMb * 1024 * 1024, 1024);
  }

  const folderSelection = getOutlookFolderSelectionFromInputs();
  payload.monitoredFolder = folderSelection;

  const originalLabel = outlookSaveButton ? outlookSaveButton.textContent : '';
  if (outlookSaveButton) {
    outlookSaveButton.disabled = true;
    outlookSaveButton.textContent = 'Saving…';
  }

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/outlook`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to update Outlook settings.';
      throw new Error(message);
    }

    if (body?.outlook !== undefined) {
      updateLocalCompanyOutlook(selectedRealmId, body.outlook);
    } else {
      await refreshQuickBooksCompanies(selectedRealmId);
    }

    showStatus(
      globalStatus,
      enabled ? 'Outlook settings saved. Sync will run shortly.' : 'Outlook monitoring disabled.',
      'success'
    );
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update Outlook settings.', 'error');
  } finally {
    if (outlookSaveButton) {
      outlookSaveButton.disabled = false;
      outlookSaveButton.textContent = originalLabel || 'Save Outlook settings';
    }
    renderOutlookSettings();
  }
}

async function handleOutlookSyncClick() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before triggering an Outlook sync.', 'error');
    return;
  }

  if (!outlookSyncButton || outlookSyncButton.disabled) {
    return;
  }

  const originalLabel = outlookSyncButton.textContent;
  outlookSyncButton.disabled = true;
  outlookSyncButton.textContent = 'Syncing…';

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/outlook/sync`, {
      method: 'POST',
    });

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to start Outlook sync.';
      throw new Error(message);
    }

    if (body?.outlook !== undefined) {
      updateLocalCompanyOutlook(selectedRealmId, body.outlook);
    }

    showStatus(globalStatus, 'Outlook sync queued. Check the status card for updates.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to start Outlook sync.', 'error');
  } finally {
    if (outlookSyncButton) {
      outlookSyncButton.disabled = false;
      outlookSyncButton.textContent = originalLabel || 'Sync now';
    }
    renderOutlookSettings();
  }
}

async function handleOutlookResync() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before requesting a full Outlook resync.', 'error');
    return;
  }

  if (!outlookResyncButton || outlookResyncButton.disabled) {
    return;
  }

  const originalLabel = outlookResyncButton.textContent;
  outlookResyncButton.disabled = true;
  outlookResyncButton.textContent = 'Queuing…';

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/outlook/resync`, {
      method: 'POST',
    });

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to start Outlook resync.';
      throw new Error(message);
    }

    showStatus(globalStatus, 'Outlook resync queued. This may take a few minutes.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to start Outlook resync.', 'error');
  } finally {
    if (outlookResyncButton) {
      outlookResyncButton.disabled = false;
      outlookResyncButton.textContent = originalLabel || 'Request full resync';
    }
    renderOutlookSettings();
  }
}

async function handleOutlookDisconnect() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before disconnecting Outlook.', 'error');
    return;
  }

  if (!outlookDisconnectButton || outlookDisconnectButton.disabled) {
    return;
  }

  const confirmed = window.confirm('Disconnect the Outlook folder for this company?');
  if (!confirmed) {
    return;
  }

  const originalLabel = outlookDisconnectButton.textContent;
  outlookDisconnectButton.disabled = true;
  outlookDisconnectButton.textContent = 'Disconnecting…';

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/outlook`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ enabled: false, monitoredFolder: null }),
    });

    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = body?.error || 'Failed to disconnect Outlook.';
      throw new Error(message);
    }

    updateLocalCompanyOutlook(selectedRealmId, body?.outlook || null);
    outlookFolderIdInput.value = '';
    outlookFolderPathInput.value = '';
    outlookFolderNameInput.value = '';
    outlookFolderWebUrlInput.value = '';
    outlookFolderParentIdInput.value = '';
    updateOutlookFolderSummaryFromInputs();
    showStatus(globalStatus, 'Outlook monitoring disabled for this company.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to disconnect Outlook.', 'error');
  } finally {
    if (outlookDisconnectButton) {
      outlookDisconnectButton.disabled = false;
      outlookDisconnectButton.textContent = originalLabel || 'Disconnect';
    }
    renderOutlookSettings();
  }
}

function openOutlookBrowser(mode) {
  if (!outlookBrowseModal) {
    return;
  }

  if (!sharedOutlookSettings?.mailboxUserId) {
    showStatus(globalStatus, 'Configure the shared Outlook mailbox before browsing folders.', 'error');
    return;
  }

  outlookBrowseState = {
    mode,
    stack: [],
    selectedItem: null,
    requestToken: 0,
  };

  const baseFolder = sharedOutlookSettings?.baseFolder || null;
  if (mode === 'company' && baseFolder?.id) {
    outlookBrowseState.stack.push({
      id: baseFolder.id,
      displayName: baseFolder.displayName || baseFolder.path || 'Base folder',
      path: baseFolder.path || '',
      webUrl: baseFolder.webUrl || null,
    });
  }

  if (mode === 'shared-base' && baseFolder?.id) {
    outlookSharedBaseIdInput.value = baseFolder.id;
    outlookSharedBasePathInput.value = baseFolder.path || '';
    outlookSharedBaseNameInput.value = baseFolder.displayName || '';
    outlookSharedBaseWebUrlInput.value = baseFolder.webUrl || '';
  }

  lastOutlookBrowseTrigger =
    mode === 'shared-base'
      ? outlookSharedBrowseBaseButton
      : outlookBrowseButton;

  if (outlookBrowseConfirmButton) {
    outlookBrowseConfirmButton.disabled = true;
    outlookBrowseConfirmButton.textContent = mode === 'shared-base' ? 'Set base folder' : 'Select folder';
  }

  if (outlookBrowseTitle) {
    outlookBrowseTitle.textContent = mode === 'shared-base' ? 'Choose Outlook base folder' : 'Choose Outlook folder';
  }

  outlookBrowseModal.hidden = false;
  outlookBrowseModal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  setOutlookBrowseEscapeHandler(true);

  renderOutlookBrowseBreadcrumb();
  renderOutlookBrowseList([]);

  const initialFolder = outlookBrowseState.stack.length
    ? outlookBrowseState.stack[outlookBrowseState.stack.length - 1]
    : null;
  loadOutlookFolderChildren(initialFolder?.id || null, { path: initialFolder?.path || null });
}

function closeOutlookBrowser() {
  if (!outlookBrowseModal) {
    return;
  }

  outlookBrowseModal.hidden = true;
  outlookBrowseModal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
  setOutlookBrowseEscapeHandler(false);
  outlookBrowseState = null;

  if (lastOutlookBrowseTrigger) {
    lastOutlookBrowseTrigger.focus();
  }
  lastOutlookBrowseTrigger = null;
}

function setOutlookBrowseEscapeHandler(enable) {
  if (enable && !outlookBrowseEscapeHandler) {
    outlookBrowseEscapeHandler = (event) => {
      if (event.key === 'Escape') {
        closeOutlookBrowser();
      }
    };
    window.addEventListener('keydown', outlookBrowseEscapeHandler);
  } else if (!enable && outlookBrowseEscapeHandler) {
    window.removeEventListener('keydown', outlookBrowseEscapeHandler);
    outlookBrowseEscapeHandler = null;
  }
}

function renderOutlookBrowseBreadcrumb() {
  if (!outlookBrowsePath) {
    return;
  }

  if (!outlookBrowseState || !outlookBrowseState.stack.length) {
    outlookBrowsePath.textContent = 'Mailbox root';
    return;
  }

  const segments = outlookBrowseState.stack.map((entry) => entry.displayName || formatOutlookFolderBreadcrumb(entry.path || ''));
  outlookBrowsePath.textContent = `Mailbox root / ${segments.join(' / ')}`;
}

function renderOutlookBrowseList(folders) {
  if (!outlookBrowseList) {
    return;
  }

  outlookBrowseList.innerHTML = '';
  outlookBrowseState.selectedItem = null;

  const fragment = document.createDocumentFragment();
  folders.forEach((folder, index) => {
    const item = document.createElement('li');
    item.className = 'outlook-browser-item';
    item.setAttribute('role', 'option');
    item.setAttribute('tabindex', '0');
    item.dataset.id = folder.id || '';
    item.dataset.path = folder.path || '';
    item.dataset.name = folder.displayName || '';
    item.dataset.parentId = folder.parentId || '';
    item.dataset.weburl = folder.webUrl || '';
    item.dataset.childCount = Number.isFinite(folder.childFolderCount) ? String(folder.childFolderCount) : '0';

    const entry = document.createElement('div');
    entry.className = 'outlook-browser-entry';

    const name = document.createElement('strong');
    name.className = 'outlook-entry-name';
    name.textContent = folder.displayName || (folder.path ? formatOutlookFolderBreadcrumb(folder.path) : '(Unnamed folder)');

    const meta = document.createElement('span');
    meta.className = 'outlook-entry-meta';
    if (folder.path) {
      meta.textContent = formatOutlookFolderBreadcrumb(folder.path);
    } else if (Number.isFinite(folder.childFolderCount)) {
      meta.textContent = `${folder.childFolderCount} subfolder${folder.childFolderCount === 1 ? '' : 's'}`;
    } else {
      meta.textContent = 'Folder';
    }

    entry.appendChild(name);
    entry.appendChild(meta);
    item.appendChild(entry);
    fragment.appendChild(item);

    if (index === 0 && outlookBrowseConfirmButton) {
      // focus the first item for keyboard navigation
      requestAnimationFrame(() => item.focus());
    }
  });

  if (!folders.length) {
    const empty = document.createElement('li');
    empty.className = 'outlook-browser-empty';
    empty.textContent = 'No subfolders in this location.';
    fragment.appendChild(empty);
  }

  outlookBrowseList.appendChild(fragment);
  if (outlookBrowseConfirmButton) {
    outlookBrowseConfirmButton.disabled = true;
  }
}

function handleOutlookBrowseListClick(event) {
  const item = event.target.closest('.outlook-browser-item');
  if (!item || !outlookBrowseList || !outlookBrowseList.contains(item)) {
    return;
  }
  selectOutlookBrowseItem(item);
}

function handleOutlookBrowseListDoubleClick(event) {
  const item = event.target.closest('.outlook-browser-item');
  if (!item || !outlookBrowseList || !outlookBrowseList.contains(item)) {
    return;
  }
  enterOutlookFolder(item.dataset);
}

function handleOutlookBrowseListKeydown(event) {
  if (!outlookBrowseState || !outlookBrowseList) {
    return;
  }

  const { key } = event;
  const items = Array.from(outlookBrowseList.querySelectorAll('.outlook-browser-item'));
  if (!items.length) {
    return;
  }

  const currentIndex = items.findIndex((item) => item.getAttribute('aria-selected') === 'true');
  let nextIndex = currentIndex;

  if (key === 'ArrowDown') {
    nextIndex = Math.min(items.length - 1, currentIndex + 1);
    event.preventDefault();
  } else if (key === 'ArrowUp') {
    nextIndex = Math.max(0, currentIndex - 1);
    event.preventDefault();
  } else if (key === 'Enter') {
    if (currentIndex >= 0) {
      const item = items[currentIndex];
      enterOutlookFolder(item.dataset);
    }
  } else if (key === 'Backspace') {
    handleOutlookBrowseBack();
  }

  if (nextIndex !== currentIndex && nextIndex >= 0) {
    const item = items[nextIndex];
    selectOutlookBrowseItem(item, { scrollIntoView: true });
  }
}

function selectOutlookBrowseItem(item, { scrollIntoView = false } = {}) {
  if (!outlookBrowseList || !item.classList.contains('outlook-browser-item')) {
    return;
  }

  const items = outlookBrowseList.querySelectorAll('.outlook-browser-item');
  items.forEach((entry) => entry.setAttribute('aria-selected', 'false'));
  item.setAttribute('aria-selected', 'true');
  outlookBrowseState.selectedItem = {
    id: item.dataset.id || null,
    path: item.dataset.path || null,
    displayName: item.dataset.name || null,
    parentId: item.dataset.parentId || null,
    webUrl: item.dataset.weburl || null,
  };

  if (outlookBrowseConfirmButton) {
    outlookBrowseConfirmButton.disabled = !outlookBrowseState.selectedItem?.id;
  }

  if (scrollIntoView) {
    item.scrollIntoView({ block: 'nearest' });
  }
}

function handleOutlookBrowseBack() {
  if (!outlookBrowseState || !outlookBrowseState.stack.length) {
    closeOutlookBrowser();
    return;
  }

  outlookBrowseState.stack.pop();
  const current = outlookBrowseState.stack.length
    ? outlookBrowseState.stack[outlookBrowseState.stack.length - 1]
    : null;
  renderOutlookBrowseBreadcrumb();
  loadOutlookFolderChildren(current?.id || null, { path: current?.path || null });
}

function enterOutlookFolder(dataset) {
  if (!dataset) {
    return;
  }
  const id = dataset.id || null;
  const name = dataset.name || dataset.path || '(Unnamed folder)';
  const path = dataset.path || '';

  if (!id) {
    return;
  }

  outlookBrowseState.stack.push({ id, displayName: name, path, webUrl: dataset.weburl || null });
  renderOutlookBrowseBreadcrumb();
  loadOutlookFolderChildren(id, { path });
}

async function loadOutlookFolderChildren(folderId, { path: parentPath = null } = {}) {
  if (!outlookBrowseState) {
    return;
  }

  const token = ++outlookBrowseState.requestToken;
  const query = new URLSearchParams();
  if (folderId) {
    query.set('folderId', folderId);
  } else if (parentPath) {
    query.set('path', parentPath);
  }

  setOutlookBrowseStatus('Loading folders…');

  try {
    const response = await fetch(`/api/outlook/folders?${query.toString()}`);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = payload?.error || payload?.message || 'Failed to load Outlook folders.';
      throw new Error(message);
    }
    if (token !== outlookBrowseState.requestToken) {
      return;
    }
    const folders = Array.isArray(payload?.folders) ? payload.folders : [];
    renderOutlookBrowseList(folders);
    setOutlookBrowseStatus(folders.length ? '' : 'No subfolders in this location.');
  } catch (error) {
    if (token !== outlookBrowseState.requestToken) {
      return;
    }
    console.error(error);
    setOutlookBrowseStatus(error.message || 'Failed to load folders.', { tone: 'error' });
    renderOutlookBrowseList([]);
  }
}

function setOutlookBrowseStatus(message, { tone = 'info' } = {}) {
  if (!outlookBrowseStatus) {
    return;
  }
  if (!message) {
    outlookBrowseStatus.hidden = true;
    outlookBrowseStatus.textContent = '';
    return;
  }
  outlookBrowseStatus.hidden = false;
  outlookBrowseStatus.textContent = message;
  outlookBrowseStatus.classList.toggle('is-error', tone === 'error');
}

function applyOutlookBrowserSelection() {
  if (!outlookBrowseState || !outlookBrowseState.selectedItem) {
    return;
  }

  const folder = outlookBrowseState.selectedItem;
  if (outlookBrowseState.mode === 'shared-base') {
    applyOutlookSharedBaseSelection(folder);
  } else {
    outlookFolderIdInput.value = folder.id || '';
    outlookFolderPathInput.value = folder.path || '';
    outlookFolderNameInput.value = folder.displayName || '';
    outlookFolderWebUrlInput.value = folder.webUrl || '';
    updateOutlookFolderSummaryFromInputs();
    if (outlookEnabledInput && outlookEnabledInput.checked) {
      outlookSyncButton.disabled = false;
      outlookResyncButton.disabled = false;
    }
  }

  closeOutlookBrowser();
}

function applyOutlookSharedBaseSelection(folder) {
  const id = folder.id || '';
  const path = folder.path || '';
  const name = folder.displayName || '';
  const webUrl = folder.webUrl || '';

  outlookSharedBaseIdInput.value = id;
  outlookSharedBasePathInput.value = path;
  outlookSharedBaseNameInput.value = name;
  outlookSharedBaseWebUrlInput.value = webUrl;

  const label = path ? formatOutlookFolderBreadcrumb(path) : name || 'Mailbox root';
  outlookSharedBaseFolder.textContent = label;
      outlookSharedSummaryText.textContent = sharedOutlookSettings?.mailboxUserId
        ? `Monitoring mailbox ${sharedOutlookSettings.mailboxDisplayName || sharedOutlookSettings.mailboxUserId}.`
        : outlookSharedSummaryText.textContent;

  outlookSharedForm.hidden = false;
  outlookSharedForm.setAttribute('aria-hidden', 'false');
  outlookSharedBrowseBaseButton?.focus();
}

function updateLocalCompanyOneDrive(realmId, state) {
  if (!realmId) {
    return;
  }

  const index = quickBooksCompanies.findIndex((entry) => entry.realmId === realmId);
  if (index >= 0) {
    quickBooksCompanies[index] = {
      ...quickBooksCompanies[index],
      oneDrive: state || null,
    };
  }

  if (selectedRealmId === realmId) {
    renderOneDriveSettings();
  }
}

function getOutlookFolderSelectionFromInputs() {
  const id = normaliseTextInput(outlookFolderIdInput?.value);
  const path = normaliseTextInput(outlookFolderPathInput?.value);
  const displayName = normaliseTextInput(outlookFolderNameInput?.value);
  const webUrl = normaliseTextInput(outlookFolderWebUrlInput?.value);
  const parentId = normaliseTextInput(outlookFolderParentIdInput?.value);

  if (!id && !path) {
    return null;
  }
  return {
    id: id || null,
    path: path || null,
    displayName: displayName || null,
    webUrl: webUrl || null,
    parentId: parentId || null,
  };
}

function formatOutlookFolderBreadcrumb(path) {
  if (!path) {
    return '';
  }
  return path.replace(/^\/+/, '').split('/').filter(Boolean).join(' / ');
}

function updateLocalCompanyOutlook(realmId, state) {
  if (!realmId) {
    return;
  }

  const index = quickBooksCompanies.findIndex((entry) => entry.realmId === realmId);
  if (index >= 0) {
    quickBooksCompanies[index] = {
      ...quickBooksCompanies[index],
      outlook: state || null,
    };
  }

  if (selectedRealmId === realmId) {
    renderOutlookSettings();
  }
}

function buildOutlookResultText(config, { enabled = true } = {}) {
  if (!config) {
    return enabled ? 'Not connected.' : 'Disabled.';
  }

  if (config.lastSyncError?.message) {
    const label = config.lastSyncStatus ? formatStatusLabel(config.lastSyncStatus, 'Error') : 'Error';
    return `${label} • ${config.lastSyncError.message}`;
  }

  if (!config.lastSyncStatus) {
    return enabled ? 'No syncs yet.' : 'Disabled.';
  }

  const metrics = config.lastSyncMetrics || {};
  const parts = [];

  if (typeof metrics.processedAttachments === 'number' && metrics.processedAttachments > 0) {
    parts.push(`${metrics.processedAttachments} attachment${metrics.processedAttachments === 1 ? '' : 's'} imported`);
  }

  if (typeof metrics.processedMessages === 'number' && metrics.processedMessages > 0) {
    parts.push(`${metrics.processedMessages} message${metrics.processedMessages === 1 ? '' : 's'} processed`);
  }

  if (typeof metrics.skippedAttachments === 'number' && metrics.skippedAttachments > 0) {
    parts.push(`${metrics.skippedAttachments} skipped`);
  }

  if (typeof metrics.errorCount === 'number' && metrics.errorCount > 0) {
    parts.push(`${metrics.errorCount} error${metrics.errorCount === 1 ? '' : 's'}`);
  }

  if (!parts.length && config.lastSyncStatus === 'success') {
    parts.push('Completed without errors');
  }

  const label = formatStatusLabel(config.lastSyncStatus, enabled ? 'Success' : 'Disabled');
  return parts.length ? `${label} • ${parts.join(', ')}` : label;
}

function sanitizeReviewSelectionId(value) {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return String(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed ? trimmed : null;
  }

  return null;
}

function normaliseMatchSuggestion(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const statusRaw = typeof entry.status === 'string' ? entry.status.trim().toLowerCase() : '';
  const status = ['exact', 'uncertain', 'unknown'].includes(statusRaw) ? statusRaw : null;
  const id = sanitizeReviewSelectionId(entry.id);
  const label = typeof entry.label === 'string' && entry.label.trim() ? entry.label.trim() : null;
  const reason = typeof entry.reason === 'string' && entry.reason.trim() ? entry.reason.trim() : null;

  return {
    id,
    label,
    status,
    reason,
  };
}

function extractInvoiceMatchHints(invoice) {
  const container =
    invoice?.matches && typeof invoice.matches === 'object'
      ? invoice.matches
      : invoice?.match && typeof invoice.match === 'object'
        ? invoice.match
        : invoice?.metadata?.matches && typeof invoice.metadata.matches === 'object'
          ? invoice.metadata.matches
          : invoice?.metadata?.match && typeof invoice.metadata.match === 'object'
            ? invoice.metadata.match
            : null;

  return {
    vendor: normaliseMatchSuggestion(container?.vendor) || null,
    account: normaliseMatchSuggestion(container?.account) || null,
    taxCode: normaliseMatchSuggestion(container?.taxCode || container?.tax) || null,
  };
}

function prepareMetadataSection(section) {
  const items = Array.isArray(section?.items) ? section.items : [];
  const normalised = [];
  const lookup = new Map();

  for (const item of items) {
    const id = sanitizeReviewSelectionId(item?.id);
    if (!id) {
      continue;
    }
    const entry = { ...item, id };
    normalised.push(entry);
    lookup.set(id, entry);
  }

  return {
    items: normalised,
    lookup,
    updatedAt: section?.updatedAt || null,
  };
}

function prepareVendorSettings(settings, sections = {}) {
  const entries = {};
  const vendorLookup = sections?.vendors?.lookup instanceof Map ? sections.vendors.lookup : null;
  const accountLookup = sections?.accounts?.lookup instanceof Map ? sections.accounts.lookup : null;
  const taxCodeLookup = sections?.taxCodes?.lookup instanceof Map ? sections.taxCodes.lookup : null;

  if (settings && typeof settings === 'object' && !Array.isArray(settings)) {
    for (const [rawVendorId, entry] of Object.entries(settings)) {
      const vendorId = sanitizeReviewSelectionId(rawVendorId);
      if (!vendorId || (vendorLookup && !vendorLookup.has(vendorId))) {
        continue;
      }

      const accountId = sanitizeReviewSelectionId(entry?.accountId);
      const validAccountId = accountId && accountLookup?.has(accountId) ? accountId : null;

      const taxCodeId = sanitizeReviewSelectionId(entry?.taxCodeId);
      const validTaxCodeId = taxCodeId && taxCodeLookup?.has(taxCodeId) ? taxCodeId : null;

      const vatTreatment = VENDOR_VAT_TREATMENT_VALUES.has(entry?.vatTreatment)
        ? entry.vatTreatment
        : null;

      if (validAccountId || validTaxCodeId || vatTreatment) {
        entries[vendorId] = {
          accountId: validAccountId,
          taxCodeId: validTaxCodeId,
          vatTreatment,
        };
      }
    }
  }

  return {
    entries,
    lookup: new Map(Object.entries(entries)),
  };
}

function applyVendorSettingsStructure(metadata, vendorSettingsMap) {
  if (!metadata) {
    return;
  }

  const sections = {
    vendors: metadata.vendors,
    accounts: metadata.accounts,
    taxCodes: metadata.taxCodes,
  };

  metadata.vendorSettings = prepareVendorSettings(vendorSettingsMap, sections);
}

function evaluateQuickBooksPreviewState(invoice, metadata) {
  const result = {
    missing: [],
    canPreview: true,
  };

  const selection = invoice?.reviewSelection || {};

  const vendorId = sanitizeReviewSelectionId(selection.vendorId);
  if (vendorId && metadata?.vendors?.lookup && !metadata.vendors.lookup.has(vendorId)) {
    result.missing.push('vendor');
  }

  const accountId = sanitizeReviewSelectionId(selection.accountId);
  if (accountId && metadata?.accounts?.lookup && !metadata.accounts.lookup.has(accountId)) {
    result.missing.push('account');
  }

  const taxCodeId = sanitizeReviewSelectionId(selection.taxCodeId);
  if (taxCodeId && metadata?.taxCodes?.lookup && !metadata.taxCodes.lookup.has(taxCodeId)) {
    result.missing.push('taxCode');
  }

  result.canPreview = result.missing.length === 0;
  return result;
}


function formatCurrencyAmount(value, currencyCode) {
  if (value === null || value === undefined || value === '') {
    return '—';
  }

  const number = Number(value);
  if (!Number.isFinite(number)) {
    return '—';
  }

  const amount = formatAmount(number);
  const symbol = currencySymbolFromCode(currencyCode);

  if (symbol) {
    return `${symbol}${amount}`;
  }

  const code = typeof currencyCode === 'string' && currencyCode.trim();
  if (code) {
    return `${code.toUpperCase()} ${amount}`;
  }

  return `£${amount}`;
}

function currencySymbolFromCode(code) {
  if (!code) {
    return '';
  }

  const upper = code.toString().trim().toUpperCase();
  if (upper === 'GBP') {
    return '£';
  }
  if (upper === 'USD') {
    return '$';
  }
  if (upper === 'EUR') {
    return '€';
  }

  return '';
}

function escapeHtml(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return value
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  return matrix[aLength][bLength];
}

function findAccountMatchByLabel(label, accounts) {
  if (!label || !Array.isArray(accounts) || !accounts.length) {
    return null;
  }

  const normalizedLabel = normaliseComparableText(label);
  if (!normalizedLabel) {
    return null;
  }

  let best = null;

  accounts.forEach((account) => {
    const candidates = [
      account?.fullyQualifiedName,
      account?.name,
    ].filter(Boolean);

    candidates.forEach((candidate) => {
      const normalizedCandidate = normaliseComparableText(candidate);
      if (!normalizedCandidate) {
        return;
      }

      if (normalizedCandidate === normalizedLabel) {
        best = { account, score: 1 };
        return;
      }

      const similarity = computeNormalisedSimilarity(normalizedLabel, normalizedCandidate);
      if (!best || similarity > best.score) {
        best = { account, score: similarity };
      }
    });
  });

  return best && best.score >= 0.88 ? best : null;
}

function findBestVendorMatch(vendorName, vendors) {
  if (!vendorName || !Array.isArray(vendors) || !vendors.length) {
    return null;
  }

  const normalizedVendor = normaliseComparableText(vendorName);
  if (!normalizedVendor) {
    return null;
  }

  let best = null;

  vendors.forEach((vendor) => {
    const labels = [
      vendor.displayName,
      vendor.name,
      vendor.companyName,
      vendor.fullyQualifiedName,
    ].filter(Boolean);

    labels.forEach((label) => {
      const normalizedCandidate = normaliseComparableText(label);
      if (!normalizedCandidate) {
        return;
      }

      const similarity = computeVendorSimilarity(normalizedVendor, normalizedCandidate);
      if (!best || similarity.score > best.score) {
        best = {
          vendor,
          score: similarity.score,
          matchedLabel: similarity.label,
        };
      }
    });
  });

  return best && best.score >= 0.4 ? best : null;
}

async function autoMatchInvoices(targetRealmId = selectedRealmId) {
  if (!storedInvoices.length) {
    return;
  }

  const realmId = targetRealmId || '';
  if (!realmId) {
    return;
  }

  const metadata = companyMetadataCache.get(realmId);
  if (!metadata) {
    return;
  }

  const relevantInvoices = storedInvoices.filter((invoice) => {
    if (resolveStoredInvoiceStatus(invoice) !== 'review') {
      return false;
    }
    const invoiceRealmId = invoice?.metadata?.companyProfile?.realmId || selectedRealmId || '';
    return !invoiceRealmId || invoiceRealmId === realmId;
  });

  for (const invoice of relevantInvoices) {
    // eslint-disable-next-line no-await-in-loop
    await autoMatchInvoice(invoice, metadata);
  }
}

async function autoMatchInvoice(invoice, metadata) {
  const checksum = invoice?.metadata?.checksum;
  if (!checksum || autoMatchInFlight.has(checksum) || autoMatchedChecksums.has(checksum)) {
    return;
  }

  const selection =
    invoice?.reviewSelection && typeof invoice.reviewSelection === 'object'
      ? invoice.reviewSelection
      : {};

  const currentVendorId = sanitizeReviewSelectionId(selection.vendorId);
  const currentAccountId = sanitizeReviewSelectionId(selection.accountId);
  const currentTaxCodeId = sanitizeReviewSelectionId(selection.taxCodeId);

  const invoiceVendor =
    invoice?.data?.vendor ||
    invoice?.metadata?.invoiceFilename ||
    invoice?.metadata?.originalName ||
    '';

  let updates = {};
  let canonicalVendorId = currentVendorId || null;

  if (!canonicalVendorId) {
    const vendorMatch = findBestVendorMatch(invoiceVendor, metadata?.vendors?.items || []);
    if (vendorMatch && vendorMatch.score >= 0.9) {
      canonicalVendorId = vendorMatch.vendor.id;
      updates.vendorId = canonicalVendorId;
    }
  }

  const vendorDefaults = canonicalVendorId
    ? metadata?.vendorSettings?.entries?.[canonicalVendorId] || null
    : null;

  if (!currentAccountId && vendorDefaults?.accountId) {
    updates.accountId = vendorDefaults.accountId;
  }

  if (!currentAccountId && !updates.accountId) {
    const suggestedAccount = invoice?.data?.suggestedAccount || null;
    const confidence = (suggestedAccount?.confidence || '').toString().toLowerCase();
    if (suggestedAccount?.name && confidence === 'high') {
      const matchedAccount = findAccountMatchByLabel(
        suggestedAccount.name,
        metadata?.accounts?.items || []
      );
      if (matchedAccount?.account?.id) {
        updates.accountId = matchedAccount.account.id;
      }
    }
  }

  if (!currentTaxCodeId && vendorDefaults?.taxCodeId) {
    updates.taxCodeId = vendorDefaults.taxCodeId;
  }

  if (!Object.keys(updates).length) {
    return;
  }

  autoMatchInFlight.add(checksum);
  try {
    const updatedInvoice = await patchInvoiceReviewSelection(checksum, updates);
    if (updatedInvoice && typeof updatedInvoice === 'object') {
      const index = storedInvoices.findIndex((entry) => entry?.metadata?.checksum === checksum);
      if (index !== -1) {
        storedInvoices[index] = updatedInvoice;
      }
    } else {
      const currentSelection = invoice.reviewSelection && typeof invoice.reviewSelection === 'object'
        ? invoice.reviewSelection
        : (invoice.reviewSelection = {});
      if (updates.vendorId) {
        currentSelection.vendorId = updates.vendorId;
      }
      if (updates.accountId) {
        currentSelection.accountId = updates.accountId;
      }
      if (updates.taxCodeId) {
        currentSelection.taxCodeId = updates.taxCodeId;
      }
    }

    autoMatchedChecksums.add(checksum);
    renderInvoices();
  } catch (error) {
    console.warn('Auto-match failed', error);
  } finally {
    autoMatchInFlight.delete(checksum);
  }
}

async function patchInvoiceReviewSelection(checksum, updates) {
  if (!checksum || !updates || !Object.keys(updates).length) {
    return null;
  }

  const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/review`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(updates),
  });

  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = body?.error || 'Failed to update invoice review selection.';
    throw new Error(message);
  }

  return body?.invoice || null;
}


function formatTimestamp(value) {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toLocaleString();
}

function formatDate(value) {
  if (!value) {
    return '—';
  }

  const isoMatch = /^\d{4}-\d{2}-\d{2}$/;
  if (isoMatch.test(value)) {
    return value;
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleDateString();
}

function formatAmount(value) {
  if (value === null || value === undefined || value === '') {
    return '—';
  }

  const number = Number(value);
  if (!Number.isFinite(number)) {
    return String(value);
  }

  return number.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function showStatus(element, message, state = 'info') {
  if (!element) {
    return;
  }
  element.textContent = message;
  element.dataset.state = state;
  show(element);
}

function show(element) {
  if (!element) {
    return;
  }
  element.hidden = false;
}

function hide(element) {
  if (!element) {
    return;
  }
  element.hidden = true;
}
