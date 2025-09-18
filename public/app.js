const globalStatus = document.getElementById('global-status');
const companyDashboard = document.getElementById('company-dashboard');
const companySelect = document.getElementById('company-select');
const addCompanyButton = document.getElementById('add-company');
const tabButtons = document.querySelectorAll('[data-company-tab]');
const companyPanels = document.querySelectorAll('.company-panel');
const connectionStatus = document.getElementById('connection-status');
const connectCompanyButton = document.getElementById('connect-company');
const refreshMetadataButton = document.getElementById('refresh-metadata');
const vendorList = document.getElementById('vendor-list');
const accountList = document.getElementById('account-list');
const reviewTableBody = document.getElementById('review-table-body');
const archiveTableBody = document.getElementById('archive-table-body');
const dropZone = document.getElementById('invoice-drop-zone');
const uploadInput = document.getElementById('invoice-upload-input');
const uploadStatusList = document.getElementById('upload-status-list');

let quickBooksCompanies = [];
let selectedRealmId = '';
const companyMetadataCache = new Map();
const metadataRequests = new Map();
let storedInvoices = [];

const MATCH_BADGE_LABELS = {
  exact: 'Exact match',
  uncertain: 'Needs review',
  unknown: 'No match',
};

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
]);

bootstrap();

function bootstrap() {
  attachEventListeners();
  handleQuickBooksCallback();
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
  }

  tabButtons.forEach((button) => {
    button.addEventListener('click', () => {
      activateCompanyTab(button.dataset.companyTab);
    });
  });
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
  if (dropZone) {
    dropZone.classList.remove('dragover');
  }
  if (event.dataTransfer?.files?.length) {
    processSelectedFiles(event.dataTransfer.files);
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

  setUploadStatus(statusEntry, 'Uploading to Gemini…');

  const formData = new FormData();
  formData.append('invoice', file);

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
  return {
    vendors: prepareMetadataSection(metadata?.vendors),
    accounts: prepareMetadataSection(metadata?.accounts),
    taxCodes: prepareMetadataSection(metadata?.taxCodes),
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
    title.textContent = item.displayName || item.name || `Vendor ${item.id}`;
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

    vendorList.appendChild(element);
  });
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
    const row = document.createElement('tr');
    row.className = 'empty-row';
    const cell = document.createElement('td');
    cell.colSpan = 8;
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

  sorted.forEach((invoice) => {
    const row = createReviewTableRow(invoice, metadata, historical);
    reviewTableBody.appendChild(row);
  });
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

  const vendorCell = document.createElement('td');
  vendorCell.className = 'cell-vendor';
  vendorCell.appendChild(createCellTitle(insights.vendor.displayName || '—'));
  if (insights.vendor.original && insights.vendor.displayName && insights.vendor.displayName !== insights.vendor.original) {
    vendorCell.appendChild(createCellSubtitle(`Extracted as ${insights.vendor.original}`));
  }
  vendorCell.appendChild(createMatchBadge(insights.vendor.confidence));
  row.appendChild(vendorCell);

  const accountCell = document.createElement('td');
  accountCell.className = 'cell-account';
  accountCell.appendChild(createCellTitle(insights.account.name || '—'));
  const accountMetaParts = [];
  if (insights.account.accountType) {
    accountMetaParts.push(insights.account.accountType);
  }
  if (insights.account.accountSubType && insights.account.accountSubType !== insights.account.accountType) {
    accountMetaParts.push(insights.account.accountSubType);
  }
  if (accountMetaParts.length) {
    accountCell.appendChild(createCellSubtitle(accountMetaParts.join(' • ')));
  }
  accountCell.appendChild(createMatchBadge(insights.account.confidence));
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

function buildInvoiceReviewInsights(invoice, metadata, historical) {
  const normalizedVendor = normaliseComparableText(invoice?.data?.vendor);
  const vendorHistoryCount = normalizedVendor ? historical.vendorCounts.get(normalizedVendor) || 0 : 0;

  const vendor = suggestVendorForInvoice(invoice, metadata, vendorHistoryCount);
  const account = suggestAccountForInvoice(invoice, metadata, {
    vendorKey: normalizedVendor,
    historicalVendorAccounts: historical.vendorAccountCounts,
  });

  const overallConfidence = deriveOverallConfidence(vendor.confidence, account.confidence);

  return {
    vendor,
    account,
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
    if (bestMatch.normalizedName === normalizedVendor || bestMatch.similarity >= 0.92) {
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

function suggestAccountForInvoice(invoice, metadata, { vendorKey, historicalVendorAccounts } = {}) {
  const empty = {
    id: null,
    name: null,
    accountType: null,
    accountSubType: null,
    confidence: 'unknown',
  };

  if (!metadata?.accounts?.items?.length) {
    return empty;
  }

  const lookup = metadata.accounts.lookup || new Map();
  const keywords = extractInvoiceKeywords(invoice);
  let chosenAccount = null;
  let chosenScore = 0;
  let confidence = 'unknown';

  if (vendorKey && historicalVendorAccounts?.has(vendorKey)) {
    const historyEntries = [...historicalVendorAccounts.get(vendorKey).entries()].sort((a, b) => b[1] - a[1]);
    if (historyEntries.length) {
      const [historyAccountId, occurrences] = historyEntries[0];
      const historyAccount = lookup.get(historyAccountId);
      if (historyAccount) {
        chosenAccount = historyAccount;
        chosenScore = occurrences;
        confidence = occurrences >= 2 ? 'exact' : 'uncertain';
      }
    }
  }

  const candidate = scoreAccountCandidates(metadata.accounts.items, keywords);
  if (candidate && (!chosenAccount || candidate.score > chosenScore)) {
    chosenAccount = candidate.account;
    chosenScore = candidate.score;
  }

  if (chosenAccount && confidence === 'unknown') {
    if (chosenScore >= 4) {
      confidence = 'exact';
    } else if (chosenScore >= 2) {
      confidence = 'uncertain';
    }
  }

  return {
    id: chosenAccount?.id || null,
    name: chosenAccount?.name || chosenAccount?.fullyQualifiedName || null,
    accountType: chosenAccount?.accountType || null,
    accountSubType: chosenAccount?.accountSubType || null,
    confidence,
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
  const checksum = button.dataset.checksum;
  if (!checksum) {
    return;
  }

  if (action === 'archive') {
    await moveInvoiceToArchive(checksum, button);
  } else if (action === 'delete') {
    await deleteInvoice(checksum, button);
  }
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
    const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/status`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ status: 'review' }),
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to move invoice to review.';
      throw new Error(message);
    }

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
    const response = await fetch(`/api/invoices/${encodeURIComponent(checksum)}/status`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ status: 'archive' }),
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      const message = payload?.error || 'Failed to move invoice to archive.';
      throw new Error(message);
    }

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
      const message = payload?.error || 'Failed to delete invoice.';
      throw new Error(message);
    }

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
