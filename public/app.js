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
const reviewBulkArchiveButton = document.getElementById('review-bulk-archive');
const reviewBulkDeleteButton = document.getElementById('review-bulk-delete');
const reviewSelectionCount = document.getElementById('review-selection-count');
const archiveTableBody = document.getElementById('archive-table-body');
const dropZone = document.getElementById('invoice-drop-zone');
const uploadInput = document.getElementById('invoice-upload-input');
const uploadStatusList = document.getElementById('upload-status-list');
const oneDriveForm = document.getElementById('onedrive-settings-form');
const oneDriveShareInput = document.getElementById('onedrive-share-url');
const oneDriveDriveIdInput = document.getElementById('onedrive-drive-id');
const oneDriveFolderIdInput = document.getElementById('onedrive-folder-id');
const oneDriveFolderPathInput = document.getElementById('onedrive-folder-path');
const oneDriveEnabledInput = document.getElementById('onedrive-enabled');
const oneDriveStatusContainer = document.getElementById('onedrive-status');
const oneDriveStatusState = document.getElementById('onedrive-status-state');
const oneDriveStatusFolder = document.getElementById('onedrive-status-folder');
const oneDriveStatusLastSync = document.getElementById('onedrive-status-last-sync');
const oneDriveStatusResult = document.getElementById('onedrive-status-result');
const oneDriveSyncButton = document.getElementById('onedrive-sync-now');
const oneDriveClearButton = document.getElementById('onedrive-clear');
const oneDriveSaveButton = document.getElementById('onedrive-save');

let quickBooksCompanies = [];
let selectedRealmId = '';
const companyMetadataCache = new Map();
const metadataRequests = new Map();
let storedInvoices = [];
const reviewSelectedChecksums = new Set();

const MATCH_BADGE_LABELS = {
  exact: 'Exact match',
  uncertain: 'Needs review',
  unknown: 'No match',
};

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
  renderBusinessProfile();
  renderOneDriveSettings();
  refreshQuickBooksCompanies();
  loadStoredInvoices();
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

  if (oneDriveSyncButton) {
    oneDriveSyncButton.addEventListener('click', handleOneDriveSyncClick);
  }

  if (oneDriveClearButton) {
    oneDriveClearButton.addEventListener('click', handleOneDriveDisconnect);
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

  if (reviewBulkArchiveButton) {
    reviewBulkArchiveButton.addEventListener('click', handleBulkArchiveSelected);
  }

  if (reviewBulkDeleteButton) {
    reviewBulkDeleteButton.addEventListener('click', handleBulkDeleteSelected);
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
  if (!params.has('quickbooks')) {
    return;
  }

  const status = params.get('quickbooks');
  const companyName = params.get('company');

  if (status === 'connected') {
    const name = companyName || 'QuickBooks company';
    showStatus(globalStatus, `Connected to ${name}.`, 'success');
  } else if (status === 'error') {
    showStatus(globalStatus, 'QuickBooks connection failed. Please try again.', 'error');
  }

  window.history.replaceState({}, document.title, window.location.pathname);
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
  if (!zone) {
    return;
  }

  event.preventDefault();
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
  if (!zone) {
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
    if (!response.ok) {
      throw new Error('Unable to load QuickBooks companies.');
    }

    const payload = await response.json();
    quickBooksCompanies = Array.isArray(payload?.companies) ? payload.companies : [];

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

function renderOneDriveSettings() {
  if (!oneDriveEnabledInput || !oneDriveForm) {
    return;
  }

  const company = getSelectedCompany();
  const hasCompany = Boolean(company);

  const inputs = [
    oneDriveShareInput,
    oneDriveDriveIdInput,
    oneDriveFolderIdInput,
    oneDriveFolderPathInput,
    oneDriveEnabledInput,
  ];

  inputs.forEach((input) => {
    if (input) {
      input.disabled = !hasCompany;
    }
  });

  if (oneDriveSaveButton) {
    oneDriveSaveButton.disabled = !hasCompany;
  }

  if (!hasCompany) {
    oneDriveForm.reset();
    if (oneDriveStatusContainer) {
      oneDriveStatusContainer.hidden = true;
    }
    if (oneDriveSyncButton) {
      oneDriveSyncButton.disabled = true;
    }
    if (oneDriveClearButton) {
      oneDriveClearButton.disabled = true;
    }
    return;
  }

  const config = company.oneDrive || null;

  if (oneDriveShareInput) {
    oneDriveShareInput.value = config?.shareUrl || '';
  }
  if (oneDriveDriveIdInput) {
    oneDriveDriveIdInput.value = config?.driveId || '';
  }
  if (oneDriveFolderIdInput) {
    oneDriveFolderIdInput.value = config?.folderId || '';
  }
  if (oneDriveFolderPathInput) {
    oneDriveFolderPathInput.value = config?.folderPath || '';
  }

  const isEnabled = config ? config.enabled !== false : false;
  oneDriveEnabledInput.checked = isEnabled;

  if (oneDriveSyncButton) {
    oneDriveSyncButton.disabled = !config || !isEnabled || !config.driveId || !config.folderId;
  }
  if (oneDriveClearButton) {
    oneDriveClearButton.disabled = !config;
  }

  if (!oneDriveStatusContainer) {
    return;
  }

  if (!config) {
    oneDriveStatusContainer.hidden = true;
    return;
  }

  oneDriveStatusContainer.hidden = false;

  if (oneDriveStatusState) {
    oneDriveStatusState.textContent = formatStatusLabel(config.status, isEnabled ? 'Connected' : 'Disabled');
  }

  if (oneDriveStatusFolder) {
    const folderParts = [];
    if (config.folderName) {
      folderParts.push(config.folderName);
    }
    if (config.folderPath) {
      folderParts.push(config.folderPath);
    }
    if (!folderParts.length && config.webUrl) {
      folderParts.push(config.webUrl);
    }
    if (!folderParts.length && config.shareUrl) {
      folderParts.push(config.shareUrl);
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
}

async function handleBusinessProfileSave(event) {
  event.preventDefault();

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before updating the business profile.', 'error');
    return;
  }

  const company = getSelectedCompany();
  const rawValue = businessTypeInput ? businessTypeInput.value : '';
  const trimmedValue = typeof rawValue === 'string' ? rawValue.trim() : '';
  const nextBusinessType = trimmedValue || null;
  const currentBusinessType = company?.businessType || null;

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
    } catch (error) {
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
        quickBooksCompanies[index] = updatedCompany;
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

async function handleOneDriveSettingsSave(event) {
  event.preventDefault();

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before updating OneDrive settings.', 'error');
    return;
  }

  const shareUrl = normaliseTextInput(oneDriveShareInput?.value);
  const driveId = normaliseTextInput(oneDriveDriveIdInput?.value);
  const folderId = normaliseTextInput(oneDriveFolderIdInput?.value);
  const folderPath = normaliseTextInput(oneDriveFolderPathInput?.value);
  const enabled = oneDriveEnabledInput ? oneDriveEnabledInput.checked : false;

  if (enabled && !shareUrl && !(driveId && (folderId || folderPath))) {
    showStatus(
      globalStatus,
      'Provide a sharing link or the drive and folder details before enabling OneDrive automation.',
      'error'
    );
    return;
  }

  const payload = { enabled };
  if (shareUrl) {
    payload.shareUrl = shareUrl;
  }
  if (driveId) {
    payload.driveId = driveId;
  }
  if (folderId) {
    payload.folderId = folderId;
  }
  if (folderPath) {
    payload.folderPath = folderPath;
  }

  const endpoint = `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive`;
  const originalLabel = oneDriveSaveButton ? oneDriveSaveButton.textContent : '';

  if (oneDriveSaveButton) {
    oneDriveSaveButton.disabled = true;
    oneDriveSaveButton.textContent = 'Saving…';
  }

  try {
    const response = await fetch(endpoint, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    let body = null;
    try {
      body = await response.json();
    } catch (error) {
      body = null;
    }

    if (!response.ok) {
      const message = body?.error || 'Failed to update OneDrive settings.';
      throw new Error(message);
    }

    if (body?.oneDrive !== undefined) {
      updateLocalCompanyOneDrive(selectedRealmId, body.oneDrive);
    } else {
      await refreshQuickBooksCompanies(selectedRealmId);
    }

    showStatus(
      globalStatus,
      enabled ? 'OneDrive folder connected. Sync will run shortly.' : 'OneDrive automation disabled.',
      'success'
    );
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to update OneDrive settings.', 'error');
  } finally {
    if (oneDriveSaveButton) {
      oneDriveSaveButton.disabled = !selectedRealmId;
      oneDriveSaveButton.textContent = originalLabel || 'Save OneDrive settings';
    }
    renderOneDriveSettings();
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
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive/sync`, {
      method: 'POST',
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to start OneDrive sync.';
      throw new Error(message);
    }

    showStatus(globalStatus, 'OneDrive sync queued. This may take a minute to complete.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to start OneDrive sync.', 'error');
  } finally {
    if (oneDriveSyncButton) {
      oneDriveSyncButton.textContent = originalLabel || 'Sync now';
      oneDriveSyncButton.disabled = false;
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

  const confirmed = window.confirm('Disconnect the configured OneDrive folder from this company?');
  if (!confirmed) {
    return;
  }

  const originalLabel = oneDriveClearButton.textContent;
  oneDriveClearButton.disabled = true;
  oneDriveClearButton.textContent = 'Disconnecting…';

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/onedrive`, {
      method: 'DELETE',
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to disconnect OneDrive.';
      throw new Error(message);
    }

    updateLocalCompanyOneDrive(selectedRealmId, null);
    showStatus(globalStatus, 'Disconnected the OneDrive folder for this company.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to disconnect OneDrive.', 'error');
  } finally {
    if (oneDriveClearButton) {
      oneDriveClearButton.textContent = originalLabel || 'Disconnect';
      oneDriveClearButton.disabled = !selectedRealmId;
    }
    renderOneDriveSettings();
  }
}

async function loadCompanyMetadata(realmId) {
  if (!realmId) {
    return;
  }

  renderVendorList(null, 'Loading QuickBooks vendors…');
  renderAccountList(null, 'Loading QuickBooks accounts…');

  try {
    const metadata = await ensureCompanyMetadata(realmId);
    renderVendorList(metadata, 'No vendors available for this company.');
    renderAccountList(metadata, 'No accounts available for this company.');
    renderReviewTable();
  } catch (error) {
    console.error(error);
    renderVendorList(null, 'Unable to load vendor list.');
    renderAccountList(null, 'Unable to load account list.');
    showStatus(globalStatus, error.message || 'Failed to load QuickBooks metadata.', 'error');
    renderReviewTable();
  }
}

async function ensureCompanyMetadata(realmId, { force = false } = {}) {
  if (!realmId) {
    return null;
  }

  if (!force && companyMetadataCache.has(realmId)) {
    return companyMetadataCache.get(realmId);
  }

  let pending = metadataRequests.get(realmId);
  if (!pending || force) {
    pending = fetchCompanyMetadata(realmId);
    metadataRequests.set(realmId, pending);
  }

  try {
    const metadata = await pending;
    return storeCompanyMetadata(realmId, metadata);
  } finally {
    metadataRequests.delete(realmId);
  }
}

async function fetchCompanyMetadata(realmId) {
  const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(realmId)}/metadata`);
  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    const message = payload?.error || 'Failed to load QuickBooks metadata.';
    throw new Error(message);
  }

  return payload?.metadata || {};
}

function storeCompanyMetadata(realmId, metadata) {
  const prepared = prepareMetadata(metadata);
  companyMetadataCache.set(realmId, prepared);
  return prepared;
}

function prepareMetadata(metadata) {
  const vendors = prepareMetadataSection(metadata?.vendors);
  const accounts = prepareMetadataSection(metadata?.accounts);
  const taxCodes = prepareMetadataSection(metadata?.taxCodes);

  return {
    vendors,
    accounts,
    taxCodes,
    vendorSettings: prepareVendorSettings(metadata?.vendorSettings, { vendors, accounts, taxCodes }),
  };
}

function prepareMetadataSection(section) {
  const items = Array.isArray(section?.items) ? section.items : [];
  const lookup = new Map(items.map((item) => [item.id, item]));
  return {
    updatedAt: section?.updatedAt || null,
    items,
    lookup,
  };
}

function prepareVendorSettings(settings, sections = {}) {
  const entries = {};
  const vendorLookup = sections?.vendors?.lookup;
  const accountLookup = sections?.accounts?.lookup;
  const taxCodeLookup = sections?.taxCodes?.lookup;

  if (settings && typeof settings === 'object' && !Array.isArray(settings)) {
    Object.entries(settings).forEach(([vendorId, entry]) => {
      if (!vendorLookup?.has(vendorId)) {
        return;
      }

      const accountId =
        typeof entry?.accountId === 'string' && accountLookup?.has(entry.accountId)
          ? entry.accountId
          : null;
      const taxCodeId =
        typeof entry?.taxCodeId === 'string' && taxCodeLookup?.has(entry.taxCodeId)
          ? entry.taxCodeId
          : null;
      const vatTreatment = VENDOR_VAT_TREATMENT_VALUES.has(entry?.vatTreatment)
        ? entry.vatTreatment
        : null;

      if (accountId || taxCodeId || vatTreatment) {
        entries[vendorId] = { accountId, taxCodeId, vatTreatment };
      }
    });
  }

  return {
    entries,
    lookup: new Map(Object.entries(entries)),
  };
}

function renderVendorList(metadata, emptyMessage) {
  if (!vendorList) {
    return;
  }

  vendorList.innerHTML = '';

  const items = metadata?.vendors?.items || [];
  if (!items.length) {
    vendorList.appendChild(createEmptyListItem(emptyMessage));
    return;
  }

  const vendorSettings = metadata?.vendorSettings?.entries || {};
  const accountOptions = buildAccountOptions(metadata?.accounts?.items || []);
  const taxCodeOptions = buildTaxCodeOptions(metadata?.taxCodes?.items || []);

  const sorted = [...items].sort((a, b) => {
    const nameA = (a.displayName || a.name || '').toLowerCase();
    const nameB = (b.displayName || b.name || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  sorted.forEach((item) => {
    const element = document.createElement('li');
    element.className = 'entity-item';

    const title = document.createElement('span');
    title.className = 'entity-name';
    const vendorName = item.displayName || item.name || `Vendor ${item.id}`;
    title.textContent = vendorName;
    element.appendChild(title);

    const metaParts = [];
    if (item.email) {
      metaParts.push(item.email);
    }
    if (item.phone) {
      metaParts.push(item.phone);
    }

    if (metaParts.length) {
      const meta = document.createElement('span');
      meta.className = 'entity-meta';
      meta.textContent = metaParts.join(' • ');
      element.appendChild(meta);
    }

    const defaults = vendorSettings[item.id] || {};
    const controls = createVendorSettingsControls({
      vendorId: item.id,
      vendorName,
      defaults,
      accountOptions,
      taxCodeOptions,
    });
    element.appendChild(controls);

    vendorList.appendChild(element);
  });
}

function buildAccountOptions(accounts) {
  const options = [{ value: '', label: 'No default' }];
  const sorted = [...accounts].sort((a, b) => {
    const labelA = (a.name || a.fullyQualifiedName || '').toLowerCase();
    const labelB = (b.name || b.fullyQualifiedName || '').toLowerCase();
    return labelA.localeCompare(labelB);
  });

  sorted.forEach((account) => {
    const baseLabel = account.name || account.fullyQualifiedName || `Account ${account.id}`;
    const labelParts = [baseLabel];
    if (account.accountType) {
      labelParts.push(account.accountType);
    }
    options.push({ value: account.id, label: labelParts.join(' • ') });
  });

  return options;
}

function buildReviewVendorOptions(vendors) {
  const options = [{ value: '', label: 'No vendor selected' }];
  const sorted = [...vendors].sort((a, b) => {
    const nameA = (a.displayName || a.name || '').toLowerCase();
    const nameB = (b.displayName || b.name || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  sorted.forEach((vendor) => {
    if (!vendor?.id) {
      return;
    }
    const label = vendor.displayName || vendor.name || `Vendor ${vendor.id}`;
    options.push({ value: vendor.id, label });
  });

  return options;
}

function buildReviewAccountOptions(accounts) {
  const options = [{ value: '', label: 'No account selected' }];
  const sorted = [...accounts].sort((a, b) => {
    const valueA = (a.name || a.fullyQualifiedName || '').toLowerCase();
    const valueB = (b.name || b.fullyQualifiedName || '').toLowerCase();
    return valueA.localeCompare(valueB);
  });

  sorted.forEach((account) => {
    if (!account?.id) {
      return;
    }
    const baseLabel = account.name || account.fullyQualifiedName || `Account ${account.id}`;
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
    const labelA = (a.name || '').toLowerCase();
    const labelB = (b.name || '').toLowerCase();
    return labelA.localeCompare(labelB);
  });

  sorted.forEach((code) => {
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

  container.appendChild(createVendorSettingGroup('CATEGORY', categorySelect));
  container.appendChild(createVendorSettingGroup('AMOUNTS ARE', vatBasisSelect));
  container.appendChild(createVendorSettingGroup('VAT', taxCodeSelect));

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
  const nextValue = target.value;

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
    showStatus(globalStatus, error.message || 'Failed to update vendor defaults.', 'error');
    target.value = previousValue;
  } finally {
    target.classList.remove('is-pending');
    target.removeAttribute('aria-busy');
    target.disabled = false;
  }
}

function resolveVendorSettingFieldValue(field, settings) {
  if (field === 'accountId') {
    return typeof settings?.accountId === 'string' && settings.accountId ? settings.accountId : '';
  }
  if (field === 'taxCodeId') {
    return typeof settings?.taxCodeId === 'string' && settings.taxCodeId ? settings.taxCodeId : '';
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

  const currentEntries = metadata.vendorSettings?.entries || {};
  const nextEntries = { ...currentEntries };
  const accountId =
    typeof settings?.accountId === 'string' && metadata.accounts?.lookup?.has(settings.accountId)
      ? settings.accountId
      : null;
  const taxCodeId =
    typeof settings?.taxCodeId === 'string' && metadata.taxCodes?.lookup?.has(settings.taxCodeId)
      ? settings.taxCodeId
      : null;
  const vatTreatment = VENDOR_VAT_TREATMENT_VALUES.has(settings?.vatTreatment)
    ? settings.vatTreatment
    : null;

  if (accountId || taxCodeId || vatTreatment) {
    nextEntries[vendorId] = { accountId, taxCodeId, vatTreatment };
  } else {
    delete nextEntries[vendorId];
  }

  metadata.vendorSettings = prepareVendorSettings(nextEntries, {
    vendors: metadata.vendors,
    accounts: metadata.accounts,
    taxCodes: metadata.taxCodes,
  });
}

async function handleImportVendorDefaults() {
  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a company before importing defaults.', 'error');
    return;
  }

  if (!importVendorDefaultsButton) {
    return;
  }

  const button = importVendorDefaultsButton;
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = 'Importing…';

  try {
    let metadata = companyMetadataCache.get(selectedRealmId);
    if (!metadata) {
      metadata = await ensureCompanyMetadata(selectedRealmId);
    }

    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/vendors/import-defaults`,
      {
        method: 'POST',
      }
    );

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to import vendor defaults.';
      throw new Error(message);
    }

    const vendorSettingsMap = payload?.vendorSettings || {};
    if (metadata) {
      metadata.vendorSettings = prepareVendorSettings(vendorSettingsMap, {
        vendors: metadata.vendors,
        accounts: metadata.accounts,
        taxCodes: metadata.taxCodes,
      });
      companyMetadataCache.set(selectedRealmId, metadata);
      renderVendorList(metadata, 'No vendors available for this company.');
    }

    const appliedCount = Array.isArray(payload?.applied) ? payload.applied.length : 0;
    if (appliedCount > 0) {
      showStatus(globalStatus, `Imported defaults for ${appliedCount} vendor${appliedCount === 1 ? '' : 's'}.`, 'success');
    } else {
      showStatus(globalStatus, 'No new defaults found for vendors.', 'info');
    }
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to import vendor defaults.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
  }
}

function renderAccountList(metadata, emptyMessage) {
  if (!accountList) {
    return;
  }

  accountList.innerHTML = '';

  const items = metadata?.accounts?.items || [];
  if (!items.length) {
    accountList.appendChild(createEmptyListItem(emptyMessage));
    return;
  }

  const sorted = [...items].sort((a, b) => {
    const nameA = (a.name || a.fullyQualifiedName || '').toLowerCase();
    const nameB = (b.name || b.fullyQualifiedName || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  sorted.forEach((item) => {
    const element = document.createElement('li');
    element.className = 'entity-item';

    const title = document.createElement('span');
    title.className = 'entity-name';
    title.textContent = item.name || item.fullyQualifiedName || `Account ${item.id}`;
    element.appendChild(title);

    const metaParts = [];
    if (item.accountType) {
      metaParts.push(item.accountType);
    }
    if (item.accountSubType) {
      metaParts.push(item.accountSubType);
    }

    if (metaParts.length) {
      const meta = document.createElement('span');
      meta.className = 'entity-meta';
      meta.textContent = metaParts.join(' • ');
      element.appendChild(meta);
    }

    accountList.appendChild(element);
  });
}

function createEmptyListItem(message) {
  const element = document.createElement('li');
  element.className = 'empty';
  element.textContent = message;
  return element;
}

async function refreshCompanyMetadata(realmId) {
  if (!realmId) {
    return;
  }

  refreshMetadataButton.disabled = true;
  const originalLabel = refreshMetadataButton.textContent;
  refreshMetadataButton.textContent = 'Refreshing…';

  try {
    const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(realmId)}/metadata/refresh`, {
      method: 'POST',
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to refresh QuickBooks metadata.';
      throw new Error(message);
    }

    showStatus(globalStatus, 'QuickBooks metadata refreshed.', 'success');
    await ensureCompanyMetadata(realmId, { force: true });
    renderVendorList(companyMetadataCache.get(realmId), 'No vendors available for this company.');
    renderAccountList(companyMetadataCache.get(realmId), 'No accounts available for this company.');
    renderReviewTable();
    await refreshQuickBooksCompanies(realmId);
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to refresh QuickBooks metadata.', 'error');
    renderReviewTable();
  } finally {
    refreshMetadataButton.textContent = originalLabel;
    refreshMetadataButton.disabled = !selectedRealmId;
  }
}

async function loadStoredInvoices() {
  try {
    const response = await fetch('/api/invoices');
    if (!response.ok) {
      throw new Error('Failed to load stored invoices.');
    }

    const payload = await response.json();
    storedInvoices = Array.isArray(payload?.invoices) ? payload.invoices : [];
    renderArchiveTable();
    renderReviewTable();
  } catch (error) {
    console.error(error);
    storedInvoices = [];
    renderArchiveTable();
    renderReviewTable();
    showStatus(globalStatus, error.message || 'Failed to load archived invoices.', 'error');
  }
}

function renderArchiveTable() {
  if (!archiveTableBody) {
    return;
  }

  archiveTableBody.innerHTML = '';

  const archiveItems = storedInvoices.filter((invoice) => invoice?.status === 'archive');
  if (!archiveItems.length) {
    const row = document.createElement('tr');
    row.className = 'empty-row';
    const cell = document.createElement('td');
    cell.colSpan = 5;
    cell.textContent = 'No archived invoices yet.';
    row.appendChild(cell);
    archiveTableBody.appendChild(row);
    return;
  }

  const sorted = [...archiveItems].sort((a, b) => {
    const aTime = new Date(a.parsedAt || 0).getTime();
    const bTime = new Date(b.parsedAt || 0).getTime();
    return bTime - aTime;
  });

  sorted.forEach((invoice) => {
    const row = document.createElement('tr');

    const invoiceCell = document.createElement('td');
    invoiceCell.appendChild(createCellTitle(invoice.data?.invoiceNumber || '—'));
    const subtitleParts = [];
    if (invoice.metadata?.originalName) {
      subtitleParts.push(invoice.metadata.originalName);
    }
    if (invoice.parsedAt) {
      subtitleParts.push(`Parsed ${formatTimestamp(invoice.parsedAt)}`);
    }
    if (subtitleParts.length) {
      invoiceCell.appendChild(createCellSubtitle(subtitleParts.join(' • ')));
    }
    row.appendChild(invoiceCell);

    const vendorCell = document.createElement('td');
    vendorCell.appendChild(createCellTitle(invoice.data?.vendor || '—'));
    row.appendChild(vendorCell);

    const invoiceDateCell = document.createElement('td');
    invoiceDateCell.appendChild(createCellTitle(formatDate(invoice.data?.invoiceDate)));
    row.appendChild(invoiceDateCell);

    const totalCell = document.createElement('td');
    totalCell.appendChild(createCellTitle(formatAmount(invoice.data?.totalAmount)));
    row.appendChild(totalCell);

    const actionsCell = document.createElement('td');
    actionsCell.className = 'table-actions';

    const moveButton = document.createElement('button');
    moveButton.type = 'button';
    moveButton.className = 'table-action';
    moveButton.dataset.action = 'move';
    moveButton.dataset.checksum = invoice.metadata?.checksum || '';
    moveButton.textContent = 'Move to review';
    actionsCell.appendChild(moveButton);

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'table-action destructive';
    deleteButton.dataset.action = 'delete';
    deleteButton.dataset.checksum = invoice.metadata?.checksum || '';
    deleteButton.textContent = 'Delete';
    actionsCell.appendChild(deleteButton);

    row.appendChild(actionsCell);
    archiveTableBody.appendChild(row);
  });
}

function renderReviewTable() {
  if (!reviewTableBody) {
    return;
  }

  reviewTableBody.innerHTML = '';

  const reviewItems = storedInvoices.filter((invoice) => invoice?.status === 'review');
  if (!reviewItems.length) {
    reviewSelectedChecksums.clear();
    updateReviewSelectionUi({ totalItems: 0 });
    const row = document.createElement('tr');
    row.className = 'empty-row';
    const cell = document.createElement('td');
    cell.colSpan = 9;
    cell.textContent = 'No invoices pending review.';
    row.appendChild(cell);
    reviewTableBody.appendChild(row);
    return;
  }

  const sorted = [...reviewItems].sort((a, b) => {
    const aTime = new Date(a.parsedAt || 0).getTime();
    const bTime = new Date(b.parsedAt || 0).getTime();
    return bTime - aTime;
  });

  const metadata = selectedRealmId ? companyMetadataCache.get(selectedRealmId) : null;
  const historical = buildReviewMatchHistory(storedInvoices, metadata);
  const checksums = sorted
    .map((invoice) => (typeof invoice?.metadata?.checksum === 'string' ? invoice.metadata.checksum : null))
    .filter(Boolean);
  const availableChecksumSet = new Set(checksums);

  pruneReviewSelection(availableChecksumSet);

  sorted.forEach((invoice) => {
    const row = createReviewTableRow(invoice, metadata, historical);
    reviewTableBody.appendChild(row);
  });

  updateReviewSelectionUi({ totalItems: availableChecksumSet.size });
}

function createReviewTableRow(invoice, metadata, historical) {
  try {
    return createEnhancedReviewRow(invoice, metadata, historical);
  } catch (error) {
    console.error('Failed to render enhanced review row', error, invoice);
    return createFallbackReviewRow(invoice);
  }
}

function createEnhancedReviewRow(invoice, metadata, historical) {
  const insights = buildInvoiceReviewInsights(invoice, metadata, historical);
  const row = document.createElement('tr');
  row.dataset.matchConfidence = insights.overallConfidence;
  row.classList.add(`match-${insights.overallConfidence}`);

  const checksum = typeof invoice?.metadata?.checksum === 'string' ? invoice.metadata.checksum : '';
  if (checksum) {
    row.dataset.checksum = checksum;
  }

  row.appendChild(createReviewSelectionCell(checksum));

  const invoiceCell = document.createElement('td');
  invoiceCell.className = 'cell-invoice';
  invoiceCell.appendChild(createCellTitle(invoice.data?.invoiceNumber || '—'));
  const invoiceSubtitleParts = [];
  if (invoice.metadata?.originalName) {
    invoiceSubtitleParts.push(invoice.metadata.originalName);
  }
  if (invoice.parsedAt) {
    invoiceSubtitleParts.push(`Parsed ${formatTimestamp(invoice.parsedAt)}`);
  }
  if (invoiceSubtitleParts.length) {
    invoiceCell.appendChild(createCellSubtitle(invoiceSubtitleParts.join(' • ')));
  }
  row.appendChild(invoiceCell);

  const dateCell = document.createElement('td');
  dateCell.appendChild(createCellTitle(formatDate(invoice.data?.invoiceDate)));
  row.appendChild(dateCell);

  const reviewSelection = getInvoiceReviewSelection(invoice);
  const vendorLookup = metadata?.vendors?.lookup || null;
  const accountLookup = metadata?.accounts?.lookup || null;
  const vendorSelectionId = reviewSelection?.vendorId || insights.vendor.quickBooksVendor?.id || '';
  const accountSelectionId = reviewSelection?.accountId || insights.account.id || '';
  const selectedVendor = vendorSelectionId && vendorLookup?.has(vendorSelectionId)
    ? vendorLookup.get(vendorSelectionId)
    : null;
  const selectedAccount = accountSelectionId && accountLookup?.has(accountSelectionId)
    ? accountLookup.get(accountSelectionId)
    : null;

  const vendorCell = document.createElement('td');
  vendorCell.className = 'cell-vendor';
  const vendorDisplayName = selectedVendor?.displayName || selectedVendor?.name || insights.vendor.displayName || '—';
  vendorCell.appendChild(createCellTitle(vendorDisplayName || '—'));
  if (insights.vendor.original && insights.vendor.displayName && insights.vendor.displayName !== insights.vendor.original) {
    vendorCell.appendChild(createCellSubtitle(`Extracted as ${insights.vendor.original}`));
  }
  if (reviewSelection?.vendorId && !selectedVendor) {
    vendorCell.appendChild(createCellSubtitle('Selected vendor is no longer available in QuickBooks metadata.'));
  }
  vendorCell.appendChild(createMatchBadge(insights.vendor.confidence));

  if (
    insights.vendor.confidence === 'unknown' &&
    selectedRealmId &&
    typeof invoice?.data?.vendor === 'string' &&
    invoice.data.vendor.trim()
  ) {
    const addVendorButton = document.createElement('button');
    addVendorButton.type = 'button';
    addVendorButton.className = 'inline-action';
    addVendorButton.dataset.action = 'create-vendor';
    addVendorButton.dataset.vendorName = invoice.data.vendor.trim();
    addVendorButton.textContent = 'Add vendor to QuickBooks';
    vendorCell.appendChild(addVendorButton);
  }

  const vendorSelect = createReviewVendorSelect({
    metadata,
    checksum,
    selectedVendorId: vendorLookup?.has(vendorSelectionId) ? vendorSelectionId : '',
  });
  if (vendorSelect) {
    vendorCell.appendChild(vendorSelect);
  } else if (selectedRealmId && (!metadata || !metadata?.vendors?.items?.length)) {
    vendorCell.appendChild(createCellSubtitle('Connect QuickBooks to manage vendors.'));
  }

  row.appendChild(vendorCell);

  const accountCell = document.createElement('td');
  accountCell.className = 'cell-account';
  const accountDisplayName = selectedAccount?.name || selectedAccount?.fullyQualifiedName || insights.account.name || '—';
  accountCell.appendChild(createCellTitle(accountDisplayName || '—'));
  const accountMetaParts = [];

  if (selectedAccount) {
    if (selectedAccount.accountType) {
      accountMetaParts.push(selectedAccount.accountType);
    }
    if (selectedAccount.accountSubType && selectedAccount.accountSubType !== selectedAccount.accountType) {
      accountMetaParts.push(selectedAccount.accountSubType);
    }
  } else {
    if (insights.account.accountType) {
      accountMetaParts.push(insights.account.accountType);
    }
    if (
      insights.account.accountSubType &&
      insights.account.accountSubType !== insights.account.accountType
    ) {
      accountMetaParts.push(insights.account.accountSubType);
    }
  }

  if (accountMetaParts.length) {
    accountCell.appendChild(createCellSubtitle(accountMetaParts.join(' • ')));
  }
  if (reviewSelection?.accountId && !selectedAccount) {
    accountCell.appendChild(createCellSubtitle('Selected account is no longer available in QuickBooks metadata.'));
  }
  if (insights.account.reason) {
    accountCell.appendChild(createCellSubtitle(insights.account.reason));
  }
  accountCell.appendChild(createMatchBadge(insights.account.confidence));

  if (insights.aiAccountSuggestion?.name) {
    const suggestionInfo = insights.aiAccountSuggestion;
    const parts = [];
    if (suggestionInfo.quickBooksMatch?.account) {
      const matchedAccount = suggestionInfo.quickBooksMatch.account;
      const similarity = suggestionInfo.quickBooksMatch.similarity
        ? ` (${Math.round(suggestionInfo.quickBooksMatch.similarity * 100)}% match)`
        : '';
      parts.push(`AI recommends ${matchedAccount.name || matchedAccount.fullyQualifiedName || suggestionInfo.name}${similarity}.`);
    } else {
      parts.push(`AI recommends new account "${suggestionInfo.name}".`);
    }

    if (suggestionInfo.reason) {
      parts.push(suggestionInfo.reason);
    }

    accountCell.appendChild(createCellSubtitle(parts.join(' ')));

    if (!suggestionInfo.quickBooksMatch?.account && selectedRealmId) {
      const addAccountButton = document.createElement('button');
      addAccountButton.type = 'button';
      addAccountButton.className = 'inline-action';
      addAccountButton.dataset.action = 'create-account';
      addAccountButton.dataset.accountName = suggestionInfo.name;
      if (suggestionInfo.accountType) {
        addAccountButton.dataset.accountType = suggestionInfo.accountType;
      }
      if (suggestionInfo.accountSubType) {
        addAccountButton.dataset.accountSubType = suggestionInfo.accountSubType;
      }
      addAccountButton.textContent = 'Add account to QuickBooks';
      accountCell.appendChild(addAccountButton);
    }
  }

  const accountSelect = createReviewAccountSelect({
    metadata,
    checksum,
    selectedAccountId: accountLookup?.has(accountSelectionId) ? accountSelectionId : '',
  });
  if (accountSelect) {
    accountCell.appendChild(accountSelect);
  } else if (selectedRealmId && (!metadata || !metadata?.accounts?.items?.length)) {
    accountCell.appendChild(createCellSubtitle('Connect QuickBooks to manage accounts.'));
  }

  row.appendChild(accountCell);

  const subtotalCell = document.createElement('td');
  subtotalCell.className = 'cell-amount';
  subtotalCell.appendChild(createCellTitle(formatAmount(insights.totals.net)));
  row.appendChild(subtotalCell);

  const vatCell = document.createElement('td');
  vatCell.className = 'cell-amount';
  vatCell.appendChild(createCellTitle(formatAmount(insights.totals.vat)));
  row.appendChild(vatCell);

  const totalCell = document.createElement('td');
  totalCell.className = 'cell-amount';
  totalCell.appendChild(createCellTitle(formatAmount(insights.totals.gross)));
  row.appendChild(totalCell);

  row.appendChild(createReviewActionsCell(invoice));
  return row;
}

function createFallbackReviewRow(invoice) {
  const row = document.createElement('tr');
  row.dataset.matchConfidence = 'unknown';
  row.classList.add('match-unknown');

  const checksum = typeof invoice?.metadata?.checksum === 'string' ? invoice.metadata.checksum : '';
  if (checksum) {
    row.dataset.checksum = checksum;
  }

  row.appendChild(createReviewSelectionCell(checksum));

  const invoiceCell = document.createElement('td');
  invoiceCell.className = 'cell-invoice';
  invoiceCell.appendChild(createCellTitle(invoice?.data?.invoiceNumber || '—'));
  row.appendChild(invoiceCell);

  const dateCell = document.createElement('td');
  dateCell.appendChild(createCellTitle(formatDate(invoice?.data?.invoiceDate)));
  row.appendChild(dateCell);

  const vendorCell = document.createElement('td');
  vendorCell.className = 'cell-vendor';
  vendorCell.appendChild(createCellTitle(invoice?.data?.vendor || '—'));
  row.appendChild(vendorCell);

  const accountCell = document.createElement('td');
  accountCell.className = 'cell-account';
  accountCell.appendChild(createCellTitle('—'));
  row.appendChild(accountCell);

  const subtotalCell = document.createElement('td');
  subtotalCell.className = 'cell-amount';
  subtotalCell.appendChild(createCellTitle(formatAmount(invoice?.data?.subtotal ?? null)));
  row.appendChild(subtotalCell);

  const vatCell = document.createElement('td');
  vatCell.className = 'cell-amount';
  vatCell.appendChild(createCellTitle(formatAmount(invoice?.data?.vatAmount ?? null)));
  row.appendChild(vatCell);

  const totalCell = document.createElement('td');
  totalCell.className = 'cell-amount';
  totalCell.appendChild(createCellTitle(formatAmount(invoice?.data?.totalAmount ?? null)));
  row.appendChild(totalCell);

  row.appendChild(createReviewActionsCell(invoice));
  return row;
}

function createReviewActionsCell(invoice) {
  const actionsCell = document.createElement('td');
  actionsCell.className = 'table-actions';

  const previewButton = document.createElement('button');
  previewButton.type = 'button';
  previewButton.className = 'table-action';
  previewButton.dataset.action = 'preview';
  previewButton.dataset.checksum = invoice?.metadata?.checksum || '';
  if (invoice?.metadata?.originalName) {
    previewButton.dataset.filename = invoice.metadata.originalName;
  }
  previewButton.textContent = 'Preview';
  actionsCell.appendChild(previewButton);

  const moveButton = document.createElement('button');
  moveButton.type = 'button';
  moveButton.className = 'table-action';
  moveButton.dataset.action = 'archive';
  moveButton.dataset.checksum = invoice?.metadata?.checksum || '';
  moveButton.textContent = 'Move to archive';
  actionsCell.appendChild(moveButton);

  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'table-action destructive';
  deleteButton.dataset.action = 'delete';
  deleteButton.dataset.checksum = invoice?.metadata?.checksum || '';
  deleteButton.textContent = 'Delete';
  actionsCell.appendChild(deleteButton);

  return actionsCell;
}

function createReviewSelectionCell(checksum) {
  const cell = document.createElement('td');
  cell.className = 'cell-select';

  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.className = 'review-select-checkbox';
  checkbox.dataset.checksum = checksum || '';
  checkbox.checked = checksum ? reviewSelectedChecksums.has(checksum) : false;
  checkbox.disabled = !checksum;
  checkbox.setAttribute('aria-label', 'Select invoice for bulk actions');

  cell.appendChild(checkbox);
  return cell;
}

function createReviewVendorSelect({ metadata, checksum, selectedVendorId }) {
  const vendors = metadata?.vendors?.items;
  if (!Array.isArray(vendors) || !vendors.length) {
    return null;
  }

  const select = document.createElement('select');
  select.className = 'review-field review-vendor-select';
  select.dataset.checksum = checksum || '';
  select.dataset.field = 'vendorId';
  select.setAttribute('aria-label', 'Select QuickBooks vendor');

  buildReviewVendorOptions(vendors).forEach((option) => {
    const element = document.createElement('option');
    element.value = option.value;
    element.textContent = option.label;
    select.appendChild(element);
  });

  const value = typeof selectedVendorId === 'string' ? selectedVendorId : '';
  select.value = value;
  select.dataset.currentValue = value;

  if (!checksum) {
    select.disabled = true;
  }

  return select;
}

function createReviewAccountSelect({ metadata, checksum, selectedAccountId }) {
  const accounts = metadata?.accounts?.items;
  if (!Array.isArray(accounts) || !accounts.length) {
    return null;
  }

  const select = document.createElement('select');
  select.className = 'review-field review-account-select';
  select.dataset.checksum = checksum || '';
  select.dataset.field = 'accountId';
  select.setAttribute('aria-label', 'Select QuickBooks account');

  buildReviewAccountOptions(accounts).forEach((option) => {
    const element = document.createElement('option');
    element.value = option.value;
    element.textContent = option.label;
    select.appendChild(element);
  });

  const value = typeof selectedAccountId === 'string' ? selectedAccountId : '';
  select.value = value;
  select.dataset.currentValue = value;

  if (!checksum) {
    select.disabled = true;
  }

  return select;
}

function getInvoiceReviewSelection(invoice) {
  if (!invoice || typeof invoice !== 'object') {
    return null;
  }

  const selection = invoice.reviewSelection;
  if (!selection || typeof selection !== 'object') {
    return null;
  }

  const vendorId = typeof selection.vendorId === 'string' && selection.vendorId ? selection.vendorId : null;
  const accountId = typeof selection.accountId === 'string' && selection.accountId ? selection.accountId : null;

  if (!vendorId && !accountId) {
    return null;
  }

  return {
    vendorId,
    accountId,
  };
}

function buildInvoiceReviewInsights(invoice, metadata, historical) {
  const normalizedVendor = normaliseComparableText(invoice?.data?.vendor);
  const vendorHistoryCount = normalizedVendor ? historical.vendorCounts.get(normalizedVendor) || 0 : 0;

  const vendor = suggestVendorForInvoice(invoice, metadata, vendorHistoryCount);
  const aiAccountSuggestion = deriveAiAccountSuggestion(invoice, metadata);
  const account = suggestAccountForInvoice(invoice, metadata, {
    vendorKey: normalizedVendor,
    matchedVendor: vendor,
    historicalVendorAccounts: historical.vendorAccountCounts,
    aiAccountSuggestion,
  });

  const overallConfidence = deriveOverallConfidence(vendor.confidence, account.confidence);

  return {
    vendor,
    account,
    aiAccountSuggestion,
    overallConfidence,
    totals: {
      net: invoice?.data?.subtotal ?? null,
      vat: invoice?.data?.vatAmount ?? null,
      gross: invoice?.data?.totalAmount ?? null,
    },
  };
}

function buildReviewMatchHistory(invoices, metadata) {
  const vendorCounts = new Map();
  const vendorAccountCounts = new Map();

  if (!Array.isArray(invoices)) {
    return { vendorCounts, vendorAccountCounts };
  }

  invoices.forEach((invoice) => {
    if (!invoice || invoice.status !== 'archive') {
      return;
    }

    const vendorKey = normaliseComparableText(invoice?.data?.vendor);
    if (!vendorKey) {
      return;
    }

    vendorCounts.set(vendorKey, (vendorCounts.get(vendorKey) || 0) + 1);

    if (!metadata?.accounts?.items?.length) {
      return;
    }

    const keywords = extractInvoiceKeywords(invoice);
    if (!keywords.length) {
      return;
    }

    const candidate = scoreAccountCandidates(metadata.accounts.items, keywords);
    const accountId = candidate?.account?.id;
    if (!accountId) {
      return;
    }

    const accountMap = vendorAccountCounts.get(vendorKey) || new Map();
    accountMap.set(accountId, (accountMap.get(accountId) || 0) + 1);
    vendorAccountCounts.set(vendorKey, accountMap);
  });

  return { vendorCounts, vendorAccountCounts };
}

function suggestVendorForInvoice(invoice, metadata, historyCount = 0) {
  const originalVendor = invoice?.data?.vendor || '';
  const normalizedVendor = normaliseComparableText(originalVendor);
  const quickBooksVendors = metadata?.vendors?.items || [];

  let bestMatch = null;

  if (normalizedVendor && quickBooksVendors.length) {
    quickBooksVendors.forEach((vendor) => {
      const candidateNames = [vendor.displayName, vendor.name].filter(Boolean);
      candidateNames.forEach((name) => {
        const normalizedCandidate = normaliseComparableText(name);
        if (!normalizedCandidate) {
          return;
        }

        const similarity = computeNormalisedSimilarity(normalizedVendor, normalizedCandidate);
        if (!bestMatch || similarity > bestMatch.similarity) {
          bestMatch = {
            vendor,
            similarity,
            normalizedName: normalizedCandidate,
          };
        }
      });
    });
  }

  let confidence = 'unknown';
  let displayName = originalVendor || null;
  let quickBooksVendor = null;

  if (bestMatch) {
    quickBooksVendor = bestMatch.vendor;
    displayName = quickBooksVendor.displayName || quickBooksVendor.name || displayName;

    const normalizedCandidate = bestMatch.normalizedName || '';
    const sharedLength = Math.min(normalizedVendor.length, normalizedCandidate.length);
    const hasMeaningfulOverlap =
      sharedLength >= 4 &&
      (normalizedVendor.includes(normalizedCandidate) || normalizedCandidate.includes(normalizedVendor));

    if (hasMeaningfulOverlap) {
      confidence = 'exact';
    } else if (normalizedCandidate === normalizedVendor || bestMatch.similarity >= 0.92) {
      confidence = 'exact';
    } else if (bestMatch.similarity >= 0.68) {
      confidence = 'uncertain';
    }
  }

  if (historyCount >= 2) {
    confidence = 'exact';
  } else if (historyCount === 1 && confidence === 'unknown') {
    confidence = 'uncertain';
  }

  return {
    original: originalVendor || null,
    displayName,
    confidence,
    quickBooksVendor,
    historyCount,
    normalized: normalizedVendor,
  };
}

function suggestAccountForInvoice(invoice, metadata, { vendorKey, matchedVendor, historicalVendorAccounts, aiAccountSuggestion } = {}) {
  const empty = {
    id: null,
    name: null,
    accountType: null,
    accountSubType: null,
    confidence: 'unknown',
    reason: null,
  };

  if (!metadata?.accounts?.items?.length) {
    return empty;
  }

  const lookup = metadata.accounts.lookup || new Map();
  const vendorSettingsLookup = metadata?.vendorSettings?.lookup;
  const matchedVendorId = matchedVendor?.quickBooksVendor?.id || matchedVendor?.id || null;
  const matchedVendorConfidence = matchedVendor?.confidence || 'unknown';

  if (matchedVendorId && matchedVendorConfidence !== 'unknown') {
    const defaults = vendorSettingsLookup?.get(matchedVendorId);
    if (defaults?.accountId) {
      const defaultAccount = lookup.get(defaults.accountId);
      return {
        id: defaultAccount?.id || defaults.accountId,
        name: defaultAccount?.name || defaultAccount?.fullyQualifiedName || null,
        accountType: defaultAccount?.accountType || null,
        accountSubType: defaultAccount?.accountSubType || null,
        confidence: matchedVendorConfidence === 'exact' ? 'exact' : 'uncertain',
        reason: 'Vendor default account.',
      };
    }
  }

  const keywords = extractInvoiceKeywords(invoice);
  let chosenAccount = null;
  let chosenScore = 0;
  let confidence = 'unknown';
  let reason = null;

  if (aiAccountSuggestion?.quickBooksMatch?.account) {
    chosenAccount = aiAccountSuggestion.quickBooksMatch.account;
    chosenScore = (aiAccountSuggestion.quickBooksMatch.similarity || 0) * 100;
    confidence = aiAccountSuggestion.finalConfidence || 'uncertain';
    reason = aiAccountSuggestion.message || `AI recommended ${chosenAccount.name || chosenAccount.fullyQualifiedName || 'an account'}.`;
  } else if (aiAccountSuggestion?.name && aiAccountSuggestion.message) {
    reason = aiAccountSuggestion.message;
  }

  if (vendorKey && historicalVendorAccounts?.has(vendorKey)) {
    const historyEntries = [...historicalVendorAccounts.get(vendorKey).entries()].sort((a, b) => b[1] - a[1]);
    if (historyEntries.length) {
      const [historyAccountId, occurrences] = historyEntries[0];
      const historyAccount = lookup.get(historyAccountId);
      if (historyAccount) {
        chosenAccount = historyAccount;
        chosenScore = occurrences;
        confidence = occurrences >= 2 ? 'exact' : 'uncertain';
        reason = `Matched historical account (${occurrences} occurrence${occurrences === 1 ? '' : 's'}).`;
      }
    }
  }

  const candidate = scoreAccountCandidates(metadata.accounts.items, keywords);
  if (candidate && (!chosenAccount || candidate.score > chosenScore)) {
    chosenAccount = candidate.account;
    chosenScore = candidate.score;
    reason = candidate.score >= 3 ? 'Strong keyword overlap with account name.' : 'Keyword overlap with account name.';
  }

  if (chosenAccount && confidence === 'unknown') {
    if (chosenScore >= 4) {
      confidence = 'exact';
    } else if (chosenScore >= 2) {
      confidence = 'uncertain';
    }
  }

  if (!chosenAccount && aiAccountSuggestion?.quickBooksMatch?.account) {
    chosenAccount = aiAccountSuggestion.quickBooksMatch.account;
    confidence = aiAccountSuggestion.finalConfidence || 'uncertain';
    reason = aiAccountSuggestion.message || `AI recommended ${aiAccountSuggestion.name}.`;
  }

  return {
    id: chosenAccount?.id || null,
    name: chosenAccount?.name || chosenAccount?.fullyQualifiedName || null,
    accountType: chosenAccount?.accountType || null,
    accountSubType: chosenAccount?.accountSubType || null,
    confidence,
    reason,
  };
}

function deriveAiAccountSuggestion(invoice, metadata) {
  const suggestion = invoice?.data?.suggestedAccount;
  if (!suggestion || !suggestion.name) {
    return null;
  }

  const accounts = metadata?.accounts?.items || [];
  const suggestionName = suggestion.name;
  const normalizedSuggestion = normaliseComparableText(suggestionName);
  let bestMatch = null;

  if (normalizedSuggestion && accounts.length) {
    accounts.forEach((account) => {
      const candidateNames = [account.name, account.fullyQualifiedName].filter(Boolean);
      candidateNames.forEach((candidateName) => {
        const normalizedCandidate = normaliseComparableText(candidateName);
        if (!normalizedCandidate) {
          return;
        }

        const similarity = computeNormalisedSimilarity(normalizedSuggestion, normalizedCandidate);
        if (!bestMatch || similarity > bestMatch.similarity) {
          bestMatch = {
            account,
            similarity,
            matchedName: candidateName,
          };
        }
      });
    });
  }

  const aiConfidenceRaw = typeof suggestion.confidence === 'string' ? suggestion.confidence.toLowerCase() : 'unknown';
  const aiConfidence = AI_ACCOUNT_CONFIDENCE_VALUES.has(aiConfidenceRaw) ? aiConfidenceRaw : 'unknown';

  let finalConfidence = 'unknown';
  let quickBooksMatch = null;
  if (bestMatch && bestMatch.similarity >= 0.68) {
    quickBooksMatch = {
      account: bestMatch.account,
      similarity: bestMatch.similarity,
      matchedName: bestMatch.matchedName,
    };

    if (bestMatch.similarity >= 0.92) {
      finalConfidence = 'exact';
    } else if (bestMatch.similarity >= 0.75) {
      finalConfidence = 'uncertain';
    } else {
      finalConfidence = 'unknown';
    }
  }

  if (finalConfidence === 'unknown') {
    if (aiConfidence === 'high' || aiConfidence === 'medium') {
      finalConfidence = 'uncertain';
    }
  }

  const similarityText = quickBooksMatch
    ? ` Closest QuickBooks match: ${quickBooksMatch.account.name || quickBooksMatch.account.fullyQualifiedName} (${Math.round(quickBooksMatch.similarity * 100)}% similarity).`
    : '';

  const reasonParts = [];
  if (suggestion.reason) {
    reasonParts.push(suggestion.reason);
  }
  reasonParts.push(`AI recommended account "${suggestionName}".`);
  const message = `${reasonParts.join(' ')}${similarityText}`.trim();

  return {
    name: suggestionName,
    accountType: suggestion.accountType || null,
    accountSubType: suggestion.accountSubType || null,
    aiConfidence,
    reason: suggestion.reason || null,
    message,
    quickBooksMatch,
    finalConfidence,
    needsCreation: !quickBooksMatch,
  };
}

function deriveOverallConfidence(vendorConfidence, accountConfidence) {
  const levels = [vendorConfidence, accountConfidence];
  if (levels.includes('unknown')) {
    return 'unknown';
  }
  if (levels.includes('uncertain')) {
    return 'uncertain';
  }
  return 'exact';
}

function createMatchBadge(confidence) {
  const level = MATCH_BADGE_LABELS[confidence] ? confidence : 'unknown';
  const badge = document.createElement('span');
  badge.className = `match-badge match-${level}`;
  badge.textContent = MATCH_BADGE_LABELS[level];
  return badge;
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

function scoreAccountCandidates(accounts, keywords) {
  if (!Array.isArray(accounts) || !accounts.length || !keywords.length) {
    return null;
  }

  let best = null;

  accounts.forEach((account) => {
    const haystack = normaliseComparableText([
      account.name,
      account.fullyQualifiedName,
      account.accountType,
      account.accountSubType,
    ].filter(Boolean).join(' '));

    if (!haystack) {
      return;
    }

    const score = computeKeywordScore(keywords, haystack);
    if (!score) {
      return;
    }

    if (!best || score > best.score) {
      best = { account, score };
    }
  });

  return best;
}

function computeKeywordScore(keywords, haystack) {
  return keywords.reduce((score, keyword) => {
    if (haystack.includes(keyword)) {
      return score + (keyword.length >= 6 ? 2 : 1);
    }
    return score;
  }, 0);
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

function createCellTitle(text) {
  const element = document.createElement('div');
  element.className = 'cell-primary';
  element.textContent = text || '—';
  return element;
}

function createCellSubtitle(text) {
  const element = document.createElement('div');
  element.className = 'cell-secondary';
  element.textContent = text;
  return element;
}

async function handleReviewAction(event) {
  const button = event.target.closest('button[data-action]');
  if (!button) {
    return;
  }

  const action = button.dataset.action;
  if (action === 'create-vendor') {
    await handleCreateVendor(button);
    return;
  }

  if (action === 'create-account') {
    await handleCreateAccount(button);
    return;
  }

  const checksum = button.dataset.checksum;
  if (!checksum) {
    return;
  }

  if (action === 'archive') {
    await moveInvoiceToArchive(checksum, button);
  } else if (action === 'delete') {
    await deleteInvoice(checksum, button);
  } else if (action === 'preview') {
    openInvoicePreview(checksum, button);
  }
}

function handleReviewChange(event) {
  const target = event.target;
  if (!target) {
    return;
  }

  if (target.classList.contains('review-select-checkbox')) {
    handleReviewSelectionToggle(target);
    return;
  }

  if (target.classList.contains('review-vendor-select')) {
    void handleInvoiceReviewFieldChange(target, 'vendorId');
    return;
  }

  if (target.classList.contains('review-account-select')) {
    void handleInvoiceReviewFieldChange(target, 'accountId');
  }
}

function handleReviewSelectionToggle(checkbox) {
  const checksum = typeof checkbox.dataset.checksum === 'string' ? checkbox.dataset.checksum : '';
  if (!checksum) {
    checkbox.checked = false;
    return;
  }

  if (checkbox.checked) {
    reviewSelectedChecksums.add(checksum);
  } else {
    reviewSelectedChecksums.delete(checksum);
  }

  updateReviewSelectionUi({ totalItems: getReviewRowCount() });
}

function handleReviewSelectAllChange(event) {
  if (!reviewTableBody) {
    return;
  }

  const checkbox = event.target;
  const shouldSelectAll = checkbox.checked;
  checkbox.indeterminate = false;

  const rowCheckboxes = reviewTableBody.querySelectorAll('input.review-select-checkbox');
  rowCheckboxes.forEach((rowCheckbox) => {
    if (rowCheckbox.disabled) {
      return;
    }
    rowCheckbox.checked = shouldSelectAll;
    const checksum = rowCheckbox.dataset.checksum;
    if (!checksum) {
      return;
    }
    if (shouldSelectAll) {
      reviewSelectedChecksums.add(checksum);
    } else {
      reviewSelectedChecksums.delete(checksum);
    }
  });

  updateReviewSelectionUi({ totalItems: getReviewRowCount() });
}

function getReviewRowCount() {
  if (!reviewTableBody) {
    return 0;
  }
  return reviewTableBody.querySelectorAll('tr[data-checksum]').length;
}

function pruneReviewSelection(availableChecksums) {
  if (!(availableChecksums instanceof Set)) {
    return;
  }

  reviewSelectedChecksums.forEach((checksum) => {
    if (!availableChecksums.has(checksum)) {
      reviewSelectedChecksums.delete(checksum);
    }
  });
}

function updateReviewSelectionUi({ totalItems = getReviewRowCount() } = {}) {
  const selectedCount = reviewSelectedChecksums.size;

  if (reviewSelectionCount) {
    reviewSelectionCount.textContent = selectedCount
      ? `${selectedCount} invoice${selectedCount === 1 ? '' : 's'} selected`
      : 'No invoices selected.';
  }

  if (reviewBulkActions) {
    reviewBulkActions.hidden = selectedCount === 0;
  }

  if (reviewBulkArchiveButton) {
    reviewBulkArchiveButton.disabled = selectedCount === 0;
  }

  if (reviewBulkDeleteButton) {
    reviewBulkDeleteButton.disabled = selectedCount === 0;
  }

  if (reviewSelectAllCheckbox) {
    reviewSelectAllCheckbox.disabled = totalItems === 0;
    reviewSelectAllCheckbox.checked = totalItems > 0 && selectedCount === totalItems;
    reviewSelectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < totalItems;
  }
}

async function handleInvoiceReviewFieldChange(select, field) {
  if (!select || select.disabled) {
    return;
  }

  const checksum = typeof select.dataset.checksum === 'string' ? select.dataset.checksum : '';
  if (!checksum) {
    select.value = '';
    return;
  }

  const previousValue = typeof select.dataset.currentValue === 'string' ? select.dataset.currentValue : '';
  const nextValue = typeof select.value === 'string' ? select.value : '';

  if (nextValue === previousValue) {
    return;
  }

  const payload = field === 'vendorId'
    ? { vendorId: nextValue || null }
    : { accountId: nextValue || null };

  const originalDisabled = select.disabled;
  select.disabled = true;
  select.classList.add('is-pending');

  try {
    await updateInvoiceReviewSelectionRequest(checksum, payload);
    applyInvoiceReviewSelection(checksum, payload);
    select.dataset.currentValue = nextValue;
    renderReviewTable();
    const fieldLabel = field === 'vendorId' ? 'Vendor' : 'Account';
    showStatus(globalStatus, `${fieldLabel} updated for invoice.`, 'success');
  } catch (error) {
    console.error('Failed to update invoice review selection', error);
    select.value = previousValue;
    showStatus(globalStatus, error.message || 'Failed to update invoice.', 'error');
  } finally {
    select.disabled = originalDisabled;
    select.classList.remove('is-pending');
  }
}

async function handleBulkArchiveSelected() {
  if (!reviewSelectedChecksums.size) {
    return;
  }

  const checksums = Array.from(reviewSelectedChecksums);
  const originalArchiveLabel = reviewBulkArchiveButton ? reviewBulkArchiveButton.textContent : '';

  if (reviewBulkArchiveButton) {
    reviewBulkArchiveButton.disabled = true;
    reviewBulkArchiveButton.classList.add('is-pending');
    reviewBulkArchiveButton.textContent = 'Archiving…';
  }

  if (reviewBulkDeleteButton) {
    reviewBulkDeleteButton.disabled = true;
  }

  const errors = [];
  let successCount = 0;

  try {
    for (const checksum of checksums) {
      try {
        await updateInvoiceStatusRequest(checksum, 'archive', {
          errorMessage: 'Failed to move invoice to archive.',
        });
        successCount += 1;
      } catch (error) {
        errors.push(error);
        console.error('Failed to archive invoice', checksum, error);
      }
    }

    if (successCount) {
      showStatus(globalStatus, `Moved ${successCount} invoice${successCount === 1 ? '' : 's'} to archive.`, 'success');
    }

    if (errors.length) {
      const message = errors.length === 1
        ? errors[0].message || 'Failed to move invoice to archive.'
        : `Failed to move ${errors.length} invoices to archive.`;
      showStatus(globalStatus, message, 'error');
    }

    reviewSelectedChecksums.clear();
    await loadStoredInvoices();
  } catch (error) {
    console.error('Bulk archive operation failed', error);
    showStatus(globalStatus, error.message || 'Failed to archive selected invoices.', 'error');
  } finally {
    if (reviewBulkArchiveButton) {
      reviewBulkArchiveButton.textContent = originalArchiveLabel;
      reviewBulkArchiveButton.disabled = false;
      reviewBulkArchiveButton.classList.remove('is-pending');
    }

    if (reviewBulkDeleteButton) {
      reviewBulkDeleteButton.disabled = false;
    }

    updateReviewSelectionUi({ totalItems: getReviewRowCount() });
  }
}

async function handleBulkDeleteSelected() {
  if (!reviewSelectedChecksums.size) {
    return;
  }

  const checksums = Array.from(reviewSelectedChecksums);
  const originalDeleteLabel = reviewBulkDeleteButton ? reviewBulkDeleteButton.textContent : '';

  if (reviewBulkDeleteButton) {
    reviewBulkDeleteButton.disabled = true;
    reviewBulkDeleteButton.classList.add('is-pending');
    reviewBulkDeleteButton.textContent = 'Deleting…';
  }

  if (reviewBulkArchiveButton) {
    reviewBulkArchiveButton.disabled = true;
  }

  const errors = [];
  let successCount = 0;

  try {
    for (const checksum of checksums) {
      try {
        await deleteInvoiceRequest(checksum, {
          errorMessage: 'Failed to delete invoice.',
        });
        successCount += 1;
      } catch (error) {
        errors.push(error);
        console.error('Failed to delete invoice', checksum, error);
      }
    }

    if (successCount) {
      showStatus(globalStatus, `Deleted ${successCount} invoice${successCount === 1 ? '' : 's'}.`, 'success');
    }

    if (errors.length) {
      const message = errors.length === 1
        ? errors[0].message || 'Failed to delete invoice.'
        : `Failed to delete ${errors.length} invoices.`;
      showStatus(globalStatus, message, 'error');
    }

    reviewSelectedChecksums.clear();
    await loadStoredInvoices();
  } catch (error) {
    console.error('Bulk delete operation failed', error);
    showStatus(globalStatus, error.message || 'Failed to delete selected invoices.', 'error');
  } finally {
    if (reviewBulkDeleteButton) {
      reviewBulkDeleteButton.textContent = originalDeleteLabel;
      reviewBulkDeleteButton.disabled = false;
      reviewBulkDeleteButton.classList.remove('is-pending');
    }

    if (reviewBulkArchiveButton) {
      reviewBulkArchiveButton.disabled = false;
    }

    updateReviewSelectionUi({ totalItems: getReviewRowCount() });
  }
}

function applyInvoiceReviewSelection(checksum, updates = {}) {
  if (!checksum || !Array.isArray(storedInvoices)) {
    return;
  }

  const index = storedInvoices.findIndex((entry) => entry?.metadata?.checksum === checksum);
  if (index === -1) {
    return;
  }

  const current = storedInvoices[index];
  const existingSelection = current.reviewSelection && typeof current.reviewSelection === 'object'
    ? { ...current.reviewSelection }
    : {};

  if (Object.prototype.hasOwnProperty.call(updates, 'vendorId')) {
    existingSelection.vendorId = sanitizeReviewSelectionId(updates.vendorId);
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'accountId')) {
    existingSelection.accountId = sanitizeReviewSelectionId(updates.accountId);
  }

  const nextSelection = normalizeClientReviewSelection(existingSelection);

  storedInvoices[index] = {
    ...current,
    reviewSelection: nextSelection,
  };
}

function sanitizeReviewSelectionId(value) {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed || null;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return String(value);
  }
  if (typeof value === 'bigint') {
    return value.toString();
  }
  return null;
}

function normalizeClientReviewSelection(selection) {
  if (!selection || typeof selection !== 'object') {
    return null;
  }

  const vendorId = sanitizeReviewSelectionId(selection.vendorId);
  const accountId = sanitizeReviewSelectionId(selection.accountId);

  if (!vendorId && !accountId) {
    return null;
  }

  return {
    vendorId,
    accountId,
  };
}

async function updateInvoiceReviewSelectionRequest(checksum, updates = {}, { errorMessage } = {}) {
  if (!checksum) {
    throw new Error(errorMessage || 'Invoice checksum is required.');
  }

  const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/review`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(updates),
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    const message = payload?.error || errorMessage || 'Failed to update invoice review selection.';
    throw new Error(message);
  }

  return payload?.invoice || null;
}

async function updateInvoiceStatusRequest(checksum, status, { errorMessage } = {}) {
  if (!checksum) {
    throw new Error(errorMessage || 'Invoice checksum is required.');
  }

  const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/status`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ status }),
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    const message = payload?.error || errorMessage || 'Failed to update invoice status.';
    throw new Error(message);
  }

  return payload?.invoice || null;
}

async function deleteInvoiceRequest(checksum, { errorMessage } = {}) {
  if (!checksum) {
    throw new Error(errorMessage || 'Invoice checksum is required.');
  }

  const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}`, {
    method: 'DELETE',
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    const message = payload?.error || errorMessage || 'Failed to delete invoice.';
    throw new Error(message);
  }

  return payload || {};
}

async function handleCreateVendor(button) {
  if (button.disabled) {
    return;
  }

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a QuickBooks company before adding vendors.', 'error');
    return;
  }

  const vendorName = typeof button.dataset.vendorName === 'string' ? button.dataset.vendorName.trim() : '';
  if (!vendorName) {
    showStatus(globalStatus, 'Invoice is missing a vendor name to add.', 'error');
    return;
  }

  const originalLabel = button.textContent;
  button.disabled = true;
  button.classList.add('is-pending');
  button.textContent = 'Adding…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/vendors`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ displayName: vendorName }),
      }
    );

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to create vendor in QuickBooks.';
      throw new Error(message);
    }

    await ensureCompanyMetadata(selectedRealmId, { force: true });
    const metadata = companyMetadataCache.get(selectedRealmId) || null;
    if (metadata) {
      renderVendorList(metadata, 'No vendors available for this company.');
      renderAccountList(metadata, 'No accounts available for this company.');
    }
    renderReviewTable();

    const createdName = payload?.vendor?.displayName || vendorName;
    showStatus(globalStatus, `Added ${createdName} to QuickBooks.`, 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to create vendor in QuickBooks.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
    button.classList.remove('is-pending');
  }
}

async function handleCreateAccount(button) {
  if (button.disabled) {
    return;
  }

  if (!selectedRealmId) {
    showStatus(globalStatus, 'Select a QuickBooks company before adding accounts.', 'error');
    return;
  }

  const accountName = typeof button.dataset.accountName === 'string' ? button.dataset.accountName.trim() : '';
  if (!accountName) {
    showStatus(globalStatus, 'Invoice is missing an account name to add.', 'error');
    return;
  }

  const accountType = typeof button.dataset.accountType === 'string' ? button.dataset.accountType.trim() : '';
  const accountSubType = typeof button.dataset.accountSubType === 'string' ? button.dataset.accountSubType.trim() : '';

  const originalLabel = button.textContent;
  button.disabled = true;
  button.classList.add('is-pending');
  button.textContent = 'Adding…';

  try {
    const response = await fetch(
      `/api/quickbooks/companies/${encodeURIComponent(selectedRealmId)}/accounts`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: accountName,
          accountType: accountType || null,
          accountSubType: accountSubType || null,
        }),
      }
    );

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to create account in QuickBooks.';
      throw new Error(message);
    }

    await ensureCompanyMetadata(selectedRealmId, { force: true });
    const metadata = companyMetadataCache.get(selectedRealmId) || null;
    if (metadata) {
      renderAccountList(metadata, 'No accounts available for this company.');
      renderVendorList(metadata, 'No vendors available for this company.');
    }
    renderReviewTable();

    const createdName = payload?.account?.name || accountName;
    showStatus(globalStatus, `Added ${createdName} to QuickBooks.`, 'success');
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to create account in QuickBooks.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
    button.classList.remove('is-pending');
  }
}

function openInvoicePreview(checksum, button) {
  if (!checksum) {
    return;
  }

  const previewUrl = buildInvoicePreviewUrl(checksum);
  const previewWindow = window.open(previewUrl, '_blank', 'noopener,noreferrer');

  if (!previewWindow) {
    const fileName = button?.dataset?.filename || 'invoice';
    showStatus(globalStatus, `Allow pop-ups to preview ${fileName}.`, 'error');
  }
}

function buildInvoicePreviewUrl(checksum) {
  const encoded = encodeURIComponent(checksum);
  return `/api/invoices/${encoded}/file`;
}

async function handleArchiveAction(event) {
  const button = event.target.closest('button[data-action]');
  if (!button) {
    return;
  }

  const action = button.dataset.action;
  const checksum = button.dataset.checksum;
  if (!checksum) {
    return;
  }

  if (action === 'move') {
    await moveInvoiceToReview(checksum, button);
  } else if (action === 'delete') {
    await deleteInvoice(checksum, button);
  }
}

async function moveInvoiceToReview(checksum, button) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = 'Moving…';

  try {
    await updateInvoiceStatusRequest(checksum, 'review', {
      errorMessage: 'Failed to move invoice to review.',
    });
    showStatus(globalStatus, 'Invoice moved to review.', 'success');
    await loadStoredInvoices();
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to move invoice to review.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
  }
}

async function moveInvoiceToArchive(checksum, button) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = 'Moving…';

  try {
    await updateInvoiceStatusRequest(checksum, 'archive', {
      errorMessage: 'Failed to move invoice to archive.',
    });
    showStatus(globalStatus, 'Invoice moved to archive.', 'success');
    await loadStoredInvoices();
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to move invoice to archive.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
  }
}

async function deleteInvoice(checksum, button) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = 'Deleting…';

  try {
    await deleteInvoiceRequest(checksum, { errorMessage: 'Failed to delete invoice.' });
    showStatus(globalStatus, 'Invoice removed from archive.', 'success');
    await loadStoredInvoices();
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to delete invoice.', 'error');
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
  }
}

function normaliseTextInput(value) {
  if (typeof value !== 'string') {
    return '';
  }
  return value.trim();
}

function updateLocalCompanyOneDrive(realmId, state) {
  if (!realmId) {
    return;
  }

  const index = quickBooksCompanies.findIndex((entry) => entry.realmId === realmId);
  if (index === -1) {
    return;
  }

  quickBooksCompanies[index] = {
    ...quickBooksCompanies[index],
    oneDrive: state || null,
  };

  if (selectedRealmId === realmId) {
    renderOneDriveSettings();
  }
}

function formatStatusLabel(value, fallback = 'Unknown') {
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) {
    return fallback;
  }

  return text
    .split(/[_\s]+/)
    .filter(Boolean)
    .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
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

  if (!parts.length && config.lastSyncStatus === 'success') {
    parts.push('Completed without errors');
  }

  const label = formatStatusLabel(config.lastSyncStatus, enabled ? 'Success' : 'Disabled');
  return parts.length ? `${label} • ${parts.join(', ')}` : label;
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
