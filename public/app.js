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
const archiveTableBody = document.getElementById('archive-table-body');

let quickBooksCompanies = [];
let selectedRealmId = '';
const companyMetadataCache = new Map();
const metadataRequests = new Map();
let storedInvoices = [];

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

  if (archiveTableBody) {
    archiveTableBody.addEventListener('click', handleArchiveAction);
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

    populateCompanySelect(previousSelection);
  } catch (error) {
    console.error(error);
    quickBooksCompanies = [];
    populateCompanySelect('');
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
  } catch (error) {
    console.error(error);
    renderVendorList(null, 'Unable to load vendor list.');
    renderAccountList(null, 'Unable to load account list.');
    showStatus(globalStatus, error.message || 'Failed to load QuickBooks metadata.', 'error');
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
    await refreshQuickBooksCompanies(realmId);
  } catch (error) {
    console.error(error);
    showStatus(globalStatus, error.message || 'Failed to refresh QuickBooks metadata.', 'error');
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
  } catch (error) {
    console.error(error);
    storedInvoices = [];
    renderArchiveTable();
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
