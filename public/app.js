const tabs = document.querySelectorAll('.tab');
const tabPanels = document.querySelectorAll('.tab-panel');

const form = document.getElementById('invoice-form');
const fileInput = document.getElementById('invoice');
const statusPanel = document.getElementById('status');
const resultsPanel = document.getElementById('results');
const duplicatePanel = document.getElementById('duplicate');
const invoicePre = document.getElementById('invoice-json');
const duplicateReason = document.getElementById('duplicate-reason');
const duplicatePre = document.getElementById('duplicate-json');

const qbAddButton = document.getElementById('qb-add-company');
const qbStatus = document.getElementById('qb-status');
const qbCompanyList = document.getElementById('qb-company-list');
const companyManagementList = document.getElementById('company-management-list');
const companiesStatus = document.getElementById('companies-status');

const matchForm = document.getElementById('match-form');
const matchStatus = document.getElementById('match-status');
const matchCompanySelect = document.getElementById('match-company');
const matchVendorSelect = document.getElementById('match-vendor');
const matchVendorManualInput = document.getElementById('match-vendor-manual');
const matchAccountSelect = document.getElementById('match-account');
const matchAccountManualInput = document.getElementById('match-account-manual');
const matchTaxRateSelect = document.getElementById('match-tax-rate');
const matchTaxRateManualInput = document.getElementById('match-tax-rate-manual');
const matchInvoiceNumberInput = document.getElementById('match-invoice-number');
const matchInvoiceNameInput = document.getElementById('match-invoice-name');
const matchSubtotalInput = document.getElementById('match-subtotal');
const matchTaxInput = document.getElementById('match-tax');
const matchTotalInput = document.getElementById('match-total');

const MATCH_STORAGE_KEY = 'invoiceMatches';
let quickBooksCompanies = [];
let currentInvoiceData = null;
let currentInvoiceMetadata = null;
let currentMatchRecord = null;
let savedMatches = loadSavedMatches();
const companyMetadataCache = new Map();
const metadataRequests = new Map();
const matcherEditableInputs = [
  matchCompanySelect,
  matchVendorSelect,
  matchVendorManualInput,
  matchAccountSelect,
  matchAccountManualInput,
  matchTaxRateSelect,
  matchTaxRateManualInput,
];
const matcherReadonlyInputs = [matchInvoiceNumberInput, matchInvoiceNameInput, matchSubtotalInput, matchTaxInput, matchTotalInput];
const matchSubmitButton = matchForm ? matchForm.querySelector('button[type="submit"]') : null;

tabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    activateTab(tab.dataset.target);
  });
});

if (form) {
  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    hide(resultsPanel);
    hide(duplicatePanel);

    const file = fileInput.files[0];
    if (!file) {
      showStatus('Choose an invoice file before submitting.', 'error');
      return;
    }

    try {
      showStatus('Uploading invoice and requesting Gemini parsing…');
      const formData = new FormData();
      formData.append('invoice', file);

      const response = await fetch('/api/parse-invoice', {
        method: 'POST',
        body: formData,
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.error || 'Invoice parsing failed');
      }

      if (payload.invoice?.data) {
        currentInvoiceData = payload.invoice.data;
        currentInvoiceMetadata = payload.invoice.metadata || null;
        invoicePre.textContent = JSON.stringify(payload.invoice.data, null, 2);
        show(resultsPanel);
        await populateMatcherFromInvoice(payload.invoice);
      }

      if (payload.duplicate) {
        duplicateReason.textContent = payload.duplicate.reason;
        duplicatePre.textContent = JSON.stringify(payload.duplicate.match, null, 2);
        show(duplicatePanel);
      }

      showStatus('Invoice processed successfully. Stored for future duplicate checks.');
      showMatchStatus('Review QuickBooks mappings in the Matcher tab.', 'info');
    } catch (error) {
      console.error(error);
      showStatus(error.message || 'Something went wrong while parsing the invoice.', 'error');
      showMatchStatus('Invoice parsing failed. Try again before matching.', 'error');
    }
  });
}

if (qbAddButton) {
  qbAddButton.addEventListener('click', () => {
    showQuickBooksStatus('Redirecting to QuickBooks for authentication…');
    window.location.href = '/api/quickbooks/connect';
  });
}

if (matchForm) {
  matchForm.addEventListener('submit', handleMatchSubmit);
}

if (matchCompanySelect) {
  matchCompanySelect.addEventListener('change', handleMatchCompanyChange);
}

if (matchVendorSelect) {
  matchVendorSelect.addEventListener('change', () => handleMatchSelectChange('vendor'));
}

if (matchAccountSelect) {
  matchAccountSelect.addEventListener('change', () => handleMatchSelectChange('account'));
}

if (matchTaxRateSelect) {
  matchTaxRateSelect.addEventListener('change', () => handleMatchSelectChange('taxRate'));
}

activateTab('parser-tab');
resetMatcherForm();
refreshQuickBooksCompanies();

const params = new URLSearchParams(window.location.search);
if (params.has('quickbooks')) {
  const status = params.get('quickbooks');
  const companyName = params.get('company');
  if (status === 'connected') {
    const name = companyName || 'QuickBooks company';
    showQuickBooksStatus(`Connected to ${name}.`);
    activateTab('quickbooks-tab');
    refreshQuickBooksCompanies();
  } else if (status === 'error') {
    showQuickBooksStatus('QuickBooks connection failed. Please try again.', 'error');
    activateTab('quickbooks-tab');
  }

  window.history.replaceState({}, document.title, window.location.pathname);
}

function activateTab(targetId) {
  tabPanels.forEach((panel) => {
    const isActive = panel.id === targetId;
    panel.classList.toggle('active', isActive);
    if (isActive) {
      show(panel);
    } else {
      hide(panel);
    }
  });

  tabs.forEach((tab) => {
    const isActive = tab.dataset.target === targetId;
    tab.classList.toggle('active', isActive);
    tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });
}

async function refreshQuickBooksCompanies() {
  const preferredRealmId = getPreferredMatchRealmId();

  try {
    const response = await fetch('/api/quickbooks/companies');
    if (!response.ok) {
      throw new Error('Unable to load QuickBooks companies.');
    }

    const payload = await response.json();
    quickBooksCompanies = Array.isArray(payload?.companies) ? payload.companies : [];

    renderQuickBooksConnectList(quickBooksCompanies);
    renderCompanyManagementList(quickBooksCompanies);
    renderMatchCompanyOptions(quickBooksCompanies, preferredRealmId);

    const activeRealmIds = new Set(quickBooksCompanies.map((company) => company.realmId));
    Array.from(companyMetadataCache.keys()).forEach((realmId) => {
      if (!activeRealmIds.has(realmId)) {
        companyMetadataCache.delete(realmId);
      }
    });

    hide(qbStatus);
    hide(companiesStatus);

    if (currentInvoiceData && matchCompanySelect) {
      const activeRealmId = matchCompanySelect.value || null;
      populateMatcherFields(currentInvoiceData, currentMatchRecord, activeRealmId);
    }
  } catch (error) {
    console.error(error);
    showQuickBooksStatus(error.message || 'Failed to load QuickBooks connections.', 'error');
    showCompaniesStatus(error.message || 'Failed to load QuickBooks connections.', 'error');
    quickBooksCompanies = [];
    renderQuickBooksConnectList(quickBooksCompanies);
    renderCompanyManagementList(quickBooksCompanies);
    renderMatchCompanyOptions(quickBooksCompanies, preferredRealmId);
  }
}

function renderQuickBooksConnectList(companies) {
  if (!qbCompanyList) {
    return;
  }

  qbCompanyList.innerHTML = '';

  if (!companies.length) {
    const empty = document.createElement('li');
    empty.className = 'empty';
    empty.textContent = 'No companies connected yet.';
    qbCompanyList.appendChild(empty);
    return;
  }

  companies.forEach((company) => {
    const item = document.createElement('li');
    item.className = 'company';

    const name = document.createElement('span');
    name.className = 'company-name';
    name.textContent = company.companyName || company.realmId;

    const meta = document.createElement('span');
    meta.className = 'company-meta';
    meta.textContent = formatQuickBooksMeta(company);

    item.append(name, meta);
    qbCompanyList.appendChild(item);
  });
}

function renderCompanyManagementList(companies) {
  if (!companyManagementList) {
    return;
  }

  companyManagementList.innerHTML = '';

  if (!companies.length) {
    const empty = document.createElement('li');
    empty.className = 'empty';
    empty.textContent = 'Connect a QuickBooks company to edit its name.';
    companyManagementList.appendChild(empty);
    return;
  }

  companies.forEach((company) => {
    const card = document.createElement('li');
    card.className = 'company-edit-card';

    const title = document.createElement('div');
    title.className = 'company-edit-title';
    title.textContent = company.companyName || company.legalName || company.realmId;

    const meta = document.createElement('div');
    meta.className = 'company-edit-meta';
    const detailLines = [formatQuickBooksMeta(company)];
    if (company.vendorsCount || company.accountsCount || company.taxCodesCount || company.taxAgenciesCount) {
      const counts = [];
      if (company.vendorsCount) {
        counts.push(`${company.vendorsCount} vendors`);
      }
      if (company.accountsCount) {
        counts.push(`${company.accountsCount} accounts`);
      }
      if (company.taxCodesCount) {
        counts.push(`${company.taxCodesCount} tax codes`);
      }
      if (company.taxAgenciesCount) {
        counts.push(`${company.taxAgenciesCount} tax agencies`);
      }
      detailLines.push(counts.join(' • '));
    }
    meta.textContent = detailLines.filter(Boolean).join('\n');

    const form = document.createElement('form');
    form.className = 'company-edit-form';
    form.dataset.realmId = company.realmId;

    const displayGroup = document.createElement('label');
    displayGroup.className = 'input-group';
    const displaySpan = document.createElement('span');
    displaySpan.textContent = 'Display Name';
    const displayInput = document.createElement('input');
    displayInput.type = 'text';
    displayInput.placeholder = company.realmId;
    displayInput.value = company.companyName || '';
    displayInput.required = true;
    displayGroup.append(displaySpan, displayInput);

    const legalGroup = document.createElement('label');
    legalGroup.className = 'input-group';
    const legalSpan = document.createElement('span');
    legalSpan.textContent = 'Legal Name';
    const legalInput = document.createElement('input');
    legalInput.type = 'text';
    legalInput.placeholder = company.companyName || 'Registered QuickBooks name';
    legalInput.value = company.legalName || '';
    legalGroup.append(legalSpan, legalInput);

    const actions = document.createElement('div');
    actions.className = 'company-edit-actions';

    const saveButton = document.createElement('button');
    saveButton.type = 'submit';
    saveButton.textContent = 'Save Names';

    const refreshButton = document.createElement('button');
    refreshButton.type = 'button';
    refreshButton.className = 'company-refresh';
    refreshButton.textContent = 'Refresh QuickBooks Data';
    refreshButton.addEventListener('click', () => {
      refreshCompanyMetadata(company.realmId, refreshButton).catch((error) => {
        console.error(error);
        showCompaniesStatus(error.message || 'Failed to refresh QuickBooks data.', 'error');
      });
    });

    actions.append(saveButton, refreshButton);

    form.append(displayGroup, legalGroup, actions);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      handleCompanyRename(event.currentTarget);
    });

    card.append(title, meta, form);
    companyManagementList.appendChild(card);
  });
}

async function handleCompanyRename(form) {
  if (!form) {
    return;
  }

  const realmId = form.dataset.realmId;
  const inputs = form.querySelectorAll('input');
  const button = form.querySelector('button[type="submit"]');

  if (!realmId || inputs.length < 2) {
    return;
  }

  const [displayInput, legalInput] = inputs;
  const displayName = displayInput?.value?.trim();
  const legalName = legalInput?.value?.trim();

  if (!displayName) {
    showCompaniesStatus('Display name cannot be empty.', 'error');
    displayInput?.focus();
    return;
  }

  if (button) {
    button.disabled = true;
    button.textContent = 'Saving…';
  }

  try {
    const updated = await updateCompanyDetails(realmId, {
      companyName: displayName,
      legalName: legalName || null,
    });
    updateSavedMatchesWithCompany(updated);
    showCompaniesStatus(`Updated ${updated.companyName || updated.realmId}.`);
    await refreshQuickBooksCompanies();
  } catch (error) {
    console.error(error);
    showCompaniesStatus(error.message || 'Failed to update company name.', 'error');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = 'Save Names';
    }
  }
}

async function updateCompanyDetails(realmId, updates) {
  const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(realmId)}`, {
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
    const message = payload?.error || 'Failed to update company name.';
    throw new Error(message);
  }

  return payload?.company || { realmId, ...updates };
}

function renderMatchCompanyOptions(companies, preferredRealmId = null) {
  if (!matchCompanySelect) {
    return;
  }

  const previousValue = preferredRealmId ?? matchCompanySelect.value;
  matchCompanySelect.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = companies.length
    ? 'Select a QuickBooks company'
    : 'Connect a QuickBooks company to select it';
  matchCompanySelect.appendChild(placeholder);

  companies.forEach((company) => {
    const option = document.createElement('option');
    option.value = company.realmId;
    option.textContent = company.companyName || company.legalName || `Realm ${company.realmId}`;
    matchCompanySelect.appendChild(option);
  });

  const hasPreferred = previousValue && companies.some((company) => company.realmId === previousValue);
  matchCompanySelect.value = hasPreferred ? previousValue : '';
  matchCompanySelect.disabled = !companies.length;

  if (!companies.length && currentInvoiceData) {
    showMatchStatus('Connect a QuickBooks company to enable matching.', 'info');
  }
}

function ensureCompanyOption(realmId, label) {
  if (!matchCompanySelect || !realmId) {
    return;
  }

  const options = Array.from(matchCompanySelect.options || []);
  const exists = options.some((option) => option.value === realmId);
  if (exists) {
    return;
  }

  const option = document.createElement('option');
  option.value = realmId;
  option.textContent = label || `Realm ${realmId}`;
  matchCompanySelect.appendChild(option);
}

function getPreferredMatchRealmId() {
  const checksum = currentInvoiceMetadata?.checksum;
  if (checksum) {
    const saved = savedMatches[checksum];
    if (saved?.companyRealmId) {
      return saved.companyRealmId;
    }
  }

  if (matchCompanySelect && matchCompanySelect.value) {
    return matchCompanySelect.value;
  }

  return null;
}


async function populateMatcherFromInvoice(invoicePayload) {
  if (!matchForm) {
    return;
  }

  const invoice = invoicePayload?.data || currentInvoiceData;
  const metadata = invoicePayload?.metadata || currentInvoiceMetadata;

  if (!invoice) {
    resetMatcherForm();
    return;
  }

  currentInvoiceData = invoice;
  currentInvoiceMetadata = metadata || null;

  setMatcherDisabled(false);

  const checksum = currentInvoiceMetadata?.checksum;
  const savedMatchRaw = checksum ? savedMatches[checksum] : null;
  const savedMatch = normalizeMatchRecord(savedMatchRaw);
  currentMatchRecord = savedMatch;

  const preferredRealmId = savedMatch?.companyRealmId || (quickBooksCompanies.length === 1 ? quickBooksCompanies[0].realmId : null);
  renderMatchCompanyOptions(quickBooksCompanies, preferredRealmId);

  if (preferredRealmId && Array.from(matchCompanySelect?.options || []).some((option) => option.value === preferredRealmId)) {
    matchCompanySelect.value = preferredRealmId;
  }

  const activeRealmId = matchCompanySelect?.value || null;
  if (activeRealmId) {
    await ensureCompanyMetadata(activeRealmId).catch((error) => {
      console.error(error);
      showMatchStatus(error.message || 'Failed to load QuickBooks metadata.', 'error');
    });
  }

  populateMatcherFields(invoice, savedMatch, activeRealmId);

  if (savedMatch?.companyRealmId && !quickBooksCompanies.some((company) => company.realmId === savedMatch.companyRealmId)) {
    ensureCompanyOption(savedMatch.companyRealmId, savedMatch.companyName || `Realm ${savedMatch.companyRealmId}`);
    matchCompanySelect.disabled = false;
    matchCompanySelect.value = savedMatch.companyRealmId;
  }

  if (savedMatchRaw) {
    showMatchStatus('Loaded saved QuickBooks match for this invoice.', 'info');
  } else {
    showMatchStatus('Review QuickBooks mappings before export.', 'info');
  }
}

function resetMatcherForm() {
  currentInvoiceData = null;
  currentInvoiceMetadata = null;
  currentMatchRecord = null;
  setMatcherDisabled(true);

  matcherEditableInputs.forEach((input) => {
    if (input) {
      input.value = '';
    }
  });
  matcherReadonlyInputs.forEach((input) => {
    if (input) {
      input.value = '';
      input.readOnly = true;
    }
  });

  if (matchCompanySelect) {
    renderMatchCompanyOptions(quickBooksCompanies, null);
  }

  showMatchStatus('Parse an invoice to begin matching.', 'info');
}

function setMatcherDisabled(disabled) {
  matcherEditableInputs.forEach((input) => {
    if (input) {
      input.disabled = disabled;
    }
  });

  if (matchSubmitButton) {
    matchSubmitButton.disabled = disabled;
  }
}

function handleMatchSelectChange(type) {
  const controls = getMatchControls(type);
  if (!controls) {
    return;
  }

  const { select, manual, metadataKey } = controls;
  if (!select || !manual) {
    return;
  }

  const selectedValue = select.value;
  const metadata = getActiveMetadata();

  if (selectedValue && selectedValue !== '__manual__' && metadata?.[metadataKey]?.lookup) {
    const item = metadata[metadataKey].lookup.get(selectedValue);
    const label = item?.displayName || item?.name || manual.value;
    if (label && !manual.value) {
      manual.value = label;
    } else if (label && selectedValue) {
      manual.value = label;
    }
    manual.disabled = true;
  } else {
    manual.disabled = false;
    if (!manual.value) {
      if (type === 'vendor') {
        manual.value = currentInvoiceData?.vendor || '';
      } else if (type === 'taxRate') {
        manual.value = currentInvoiceData?.taxCode || '';
      }
    }
  }
}

function handleMatchSubmit(event) {
  event.preventDefault();

  const checksum = currentInvoiceMetadata?.checksum;
  if (!checksum) {
    showMatchStatus('Parse an invoice before saving a match.', 'error');
    return;
  }

  if (!matchCompanySelect || matchCompanySelect.disabled || !matchCompanySelect.value) {
    showMatchStatus('Select a QuickBooks company before saving.', 'error');
    if (matchCompanySelect) {
      matchCompanySelect.focus();
    }
    return;
  }

  const companyRealmId = matchCompanySelect.value;
  const company = quickBooksCompanies.find((entry) => entry.realmId === companyRealmId);
  const metadata = companyMetadataCache.get(companyRealmId) || null;

  const vendorField = buildMatchField('vendor', metadata, currentInvoiceData?.vendor || '', currentMatchRecord?.vendor);
  const accountField = buildMatchField('account', metadata, '', currentMatchRecord?.account);
  const taxRateField = buildMatchField('taxRate', metadata, currentInvoiceData?.taxCode || '', currentMatchRecord?.taxRate);

  const record = {
    checksum,
    updatedAt: new Date().toISOString(),
    companyRealmId,
    companyName: company?.companyName || company?.legalName || company?.realmId || null,
    vendor: vendorField,
    account: accountField,
    taxRate: taxRateField,
    invoice: {
      invoiceNumber: currentInvoiceData?.invoiceNumber || null,
      invoiceName: matchInvoiceNameInput?.value || null,
      subtotal: currentInvoiceData?.subtotal ?? null,
      tax: currentInvoiceData?.vatAmount ?? currentInvoiceData?.taxAmount ?? null,
      total: currentInvoiceData?.totalAmount ?? null,
      originalName: currentInvoiceMetadata?.originalName || null,
    },
  };

  currentMatchRecord = record;
  saveMatchRecord(record);

  if (record.companyName) {
    showMatchStatus(`Saved match for ${record.companyName}.`);
  } else {
    showMatchStatus('Saved match configuration.');
  }
}

async function handleMatchCompanyChange() {
  if (!matchCompanySelect) {
    return;
  }

  const realmId = matchCompanySelect.value;
  if (!realmId) {
    updateMatchSelectOptions(null);
    applyMatchFieldState('vendor', currentInvoiceData?.vendor || '', null, null);
    applyMatchFieldState('account', '', null, null);
    applyMatchFieldState('taxRate', currentInvoiceData?.taxCode || '', null, null);
    return;
  }

  try {
    const metadata = await ensureCompanyMetadata(realmId);
    populateMatcherFields(currentInvoiceData || {}, currentMatchRecord, realmId);
    if (currentMatchRecord) {
      currentMatchRecord.companyRealmId = realmId;
    }
    if (
      !metadata?.vendors?.items?.length &&
      !metadata?.accounts?.items?.length &&
      !metadata?.taxCodes?.items?.length &&
      !metadata?.taxAgencies?.items?.length &&
      currentInvoiceData
    ) {
      showMatchStatus('QuickBooks lists are incomplete. Refresh metadata from the Companies tab if needed.', 'info');
    }
  } catch (error) {
    console.error(error);
    showMatchStatus(error.message || 'Failed to load QuickBooks metadata.', 'error');
  }
}

function formatQuickBooksMeta(company) {
  const parts = [];
  if (company.legalName && company.legalName !== company.companyName) {
    parts.push(company.legalName);
  }
  if (company.environment) {
    parts.push(company.environment === 'production' ? 'Production' : 'Sandbox');
  }

  if (company.connectedAt) {
    const connected = new Date(company.connectedAt);
    if (!Number.isNaN(connected.getTime())) {
      parts.push(`Connected ${connected.toLocaleString()}`);
    }
  }

  parts.push(`Realm ${company.realmId}`);

  if (company.updatedAt && company.updatedAt !== company.connectedAt) {
    const updated = new Date(company.updatedAt);
    if (!Number.isNaN(updated.getTime())) {
      parts.push(`Updated ${updated.toLocaleString()}`);
    }
  }

  if (company.vendorsUpdatedAt) {
    const vendorsDate = new Date(company.vendorsUpdatedAt);
    if (!Number.isNaN(vendorsDate.getTime())) {
      parts.push(`Vendors updated ${vendorsDate.toLocaleDateString()}`);
    }
  }

  if (company.taxAgenciesUpdatedAt) {
    const agenciesDate = new Date(company.taxAgenciesUpdatedAt);
    if (!Number.isNaN(agenciesDate.getTime())) {
      parts.push(`Tax agencies updated ${agenciesDate.toLocaleDateString()}`);
    }
  }

  return parts.join(' • ');
}

function showQuickBooksStatus(message, state = 'info') {
  if (!qbStatus) {
    return;
  }
  qbStatus.textContent = message;
  qbStatus.dataset.state = state;
  show(qbStatus);
}

function showStatus(message, state = 'info') {
  if (!statusPanel) {
    return;
  }
  statusPanel.textContent = message;
  statusPanel.dataset.state = state;
  show(statusPanel);
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

function showCompaniesStatus(message, state = 'info') {
  if (!companiesStatus) {
    return;
  }
  companiesStatus.textContent = message;
  companiesStatus.dataset.state = state;
  show(companiesStatus);
}

function showMatchStatus(message, state = 'info') {
  if (!matchStatus) {
    return;
  }
  matchStatus.textContent = message;
  matchStatus.dataset.state = state;
  show(matchStatus);
}

function loadSavedMatches() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return {};
  }

  try {
    const raw = window.localStorage.getItem(MATCH_STORAGE_KEY);
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      return migrateSavedMatches(parsed);
    }
  } catch (error) {
    console.warn('Failed to load saved matches', error);
  }

  return {};
}

function migrateSavedMatches(records) {
  if (!records || typeof records !== 'object') {
    return {};
  }

  const migrated = {};
  Object.entries(records).forEach(([checksum, record]) => {
    migrated[checksum] = normalizeMatchRecord(record);
  });
  return migrated;
}

function persistSavedMatches() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }

  try {
    window.localStorage.setItem(MATCH_STORAGE_KEY, JSON.stringify(savedMatches));
  } catch (error) {
    console.warn('Failed to persist saved matches', error);
  }
}

function saveMatchRecord(record) {
  if (!record?.checksum) {
    return;
  }

  savedMatches[record.checksum] = normalizeMatchRecord(record);
  persistSavedMatches();
}

function getMatchControls(type) {
  switch (type) {
    case 'vendor':
      return { select: matchVendorSelect, manual: matchVendorManualInput, metadataKey: 'vendors' };
    case 'account':
      return { select: matchAccountSelect, manual: matchAccountManualInput, metadataKey: 'accounts' };
    case 'taxRate':
      return { select: matchTaxRateSelect, manual: matchTaxRateManualInput, metadataKey: 'taxCodes' };
    default:
      return null;
  }
}

function getActiveMetadata() {
  const realmId = matchCompanySelect?.value;
  if (!realmId) {
    return null;
  }
  return companyMetadataCache.get(realmId) || null;
}

function populateMatcherFields(invoice, savedMatch, realmId) {
  const metadata = realmId ? companyMetadataCache.get(realmId) : null;
  updateMatchSelectOptions(metadata);

  const invoiceName = deriveInvoiceName({ metadata: currentInvoiceMetadata, data: invoice });

  if (matchInvoiceNumberInput) {
    matchInvoiceNumberInput.value = invoice.invoiceNumber || '';
    matchInvoiceNumberInput.readOnly = true;
  }

  if (matchInvoiceNameInput) {
    matchInvoiceNameInput.value = invoiceName;
    matchInvoiceNameInput.readOnly = true;
  }

  if (matchSubtotalInput) {
    matchSubtotalInput.value = formatAmount(invoice.subtotal);
    matchSubtotalInput.readOnly = true;
  }

  const taxAmount = invoice.vatAmount ?? invoice.taxAmount ?? null;
  if (matchTaxInput) {
    matchTaxInput.value = formatAmount(taxAmount);
    matchTaxInput.readOnly = true;
  }

  if (matchTotalInput) {
    matchTotalInput.value = formatAmount(invoice.totalAmount);
    matchTotalInput.readOnly = true;
  }

  applyMatchFieldState('vendor', invoice.vendor || '', savedMatch?.vendor, metadata);
  applyMatchFieldState('account', '', savedMatch?.account, metadata);
  applyMatchFieldState('taxRate', invoice.taxCode || '', savedMatch?.taxRate, metadata);

  handleMatchSelectChange('vendor');
  handleMatchSelectChange('account');
  handleMatchSelectChange('taxRate');
}

function updateMatchSelectOptions(metadata) {
  const vendorItems = metadata?.vendors?.items || [];
  const accountItems = metadata?.accounts?.items || [];
  const taxCodeItems = metadata?.taxCodes?.items || [];

  populateSelectFromMetadata(matchVendorSelect, 'Select a QuickBooks vendor', 'Use manual value', vendorItems);
  populateSelectFromMetadata(matchAccountSelect, 'Select a QuickBooks account', 'Use manual value', accountItems);
  populateSelectFromMetadata(matchTaxRateSelect, 'Select a QuickBooks tax code', 'Use manual value', taxCodeItems);
}

function populateSelectFromMetadata(select, placeholder, manualLabel, items) {
  if (!select) {
    return;
  }

  const previousValue = select.value;
  select.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = placeholder;
  select.appendChild(placeholderOption);

  const manualOption = document.createElement('option');
  manualOption.value = '__manual__';
  manualOption.textContent = manualLabel;
  select.appendChild(manualOption);

  items.forEach((item) => {
    const option = document.createElement('option');
    option.value = item.id;
    option.textContent = item.displayName || item.name || item.fullyQualifiedName || `ID ${item.id}`;
    select.appendChild(option);
  });

  const hasItems = items.length > 0;
  select.disabled = !hasItems;

  if (!hasItems) {
    select.value = '__manual__';
  } else if (previousValue && Array.from(select.options).some((option) => option.value === previousValue)) {
    select.value = previousValue;
  } else {
    select.value = '';
  }
}

function applyMatchFieldState(type, invoiceValue, savedField, metadata) {
  const controls = getMatchControls(type);
  if (!controls) {
    return;
  }

  const { select, manual, metadataKey } = controls;
  if (!manual) {
    return;
  }

  const normalizedField = normalizeMatchField(savedField);
  const lookup = metadata?.[metadataKey]?.lookup;
  const hasMetadata = Boolean(metadata?.[metadataKey]?.items?.length);

  if (normalizedField.source === 'quickbooks' && normalizedField.id) {
    const item = lookup?.get(normalizedField.id);
    if (select) {
      ensureSelectOption(select, normalizedField.id, normalizedField.name || item?.displayName || item?.name || `ID ${normalizedField.id}`);
      select.disabled = false;
      select.value = normalizedField.id;
    }
    manual.value = normalizedField.name || item?.displayName || item?.name || invoiceValue || '';
    manual.disabled = true;
  } else {
    const manualValue = normalizedField.name || invoiceValue || '';
    manual.value = manualValue;
    manual.disabled = false;
    if (select) {
      if (!hasMetadata) {
        select.value = '__manual__';
        select.disabled = true;
      } else {
        select.disabled = false;
        select.value = '__manual__';
      }
    }
  }
}

function ensureSelectOption(select, value, label) {
  if (!select || !value) {
    return;
  }

  const exists = Array.from(select.options || []).some((option) => option.value === value);
  if (!exists) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label || `ID ${value}`;
    select.appendChild(option);
  }
}

function buildMatchField(type, metadata, invoiceValue, existingField) {
  const controls = getMatchControls(type);
  if (!controls) {
    return normalizeMatchField(existingField);
  }

  const { select, manual, metadataKey } = controls;
  const manualValue = manual?.value?.trim() || '';
  const selection = select?.value;

  if (selection && selection !== '__manual__' && metadata?.[metadataKey]?.lookup) {
    const item = metadata[metadataKey].lookup.get(selection);
    const snapshot = getMetadataSnapshot(type, item);
    return {
      id: selection,
      name: snapshot?.name || manualValue || invoiceValue || '',
      source: 'quickbooks',
      details: snapshot,
    };
  }

  if (manualValue) {
    return {
      id: null,
      name: manualValue,
      source: 'manual',
    };
  }

  if (existingField?.source === 'manual' && existingField?.name) {
    return normalizeMatchField(existingField);
  }

  return {
    id: null,
    name: invoiceValue || null,
    source: 'manual',
  };
}

function getMetadataSnapshot(type, item) {
  if (!item) {
    return null;
  }

  if (type === 'vendor') {
    return {
      name: item.displayName || item.name || null,
      email: item.email || null,
      phone: item.phone || null,
    };
  }

  if (type === 'account') {
    return {
      name: item.name || item.fullyQualifiedName || null,
      type: item.accountType || item.type || null,
      subType: item.accountSubType || item.subType || null,
      fullyQualifiedName: item.fullyQualifiedName || null,
    };
  }

  if (type === 'taxRate') {
    return {
      name: item.name || null,
      rate: typeof item.rate === 'number' ? item.rate : item.rateValue ?? null,
      active: item.active ?? null,
      agency: item.agency || null,
    };
  }

  return {
    name: item.name || item.displayName || null,
  };
}

function normalizeMatchRecord(record) {
  if (!record || typeof record !== 'object') {
    return null;
  }

  return {
    checksum: record.checksum || null,
    updatedAt: record.updatedAt || new Date().toISOString(),
    companyRealmId: record.companyRealmId || null,
    companyName: record.companyName || null,
    vendor: normalizeMatchField(record.vendor),
    account: normalizeMatchField(record.account),
    taxRate: normalizeMatchField(record.taxRate),
    invoice: record.invoice || {},
  };
}

function normalizeMatchField(field) {
  if (!field) {
    return { id: null, name: null, source: 'manual' };
  }

  if (typeof field === 'string') {
    return { id: null, name: field, source: 'manual' };
  }

  if (typeof field === 'object') {
    const source = field.source === 'quickbooks' ? 'quickbooks' : 'manual';
    const normalized = {
      id: field.id || null,
      name: field.name || field.displayName || field.value || null,
      source,
    };

    if (source === 'quickbooks') {
      normalized.details = field.details || {
        type: field.type || field.accountType || null,
        subType: field.subType || field.accountSubType || null,
        rate: field.rate ?? field.rateValue ?? null,
        agency: field.agency || null,
      };
    }

    return normalized;
  }

  return { id: null, name: null, source: 'manual' };
}

function updateSavedMatchesWithCompany(company) {
  if (!company?.realmId) {
    return;
  }

  let mutated = false;
  Object.values(savedMatches).forEach((record) => {
    if (record.companyRealmId === company.realmId) {
      record.companyName = company.companyName || company.realmId;
      mutated = true;
    }
  });

  if (mutated) {
    persistSavedMatches();
  }
}

function deriveInvoiceName(payload) {
  if (!payload) {
    return '';
  }

  const metadataName = payload.metadata?.originalName;
  if (metadataName) {
    return metadataName;
  }

  const vendor = payload.data?.vendor;
  if (vendor) {
    return vendor;
  }

  const billTo = payload.data?.billTo;
  if (billTo) {
    return billTo;
  }

  return 'Invoice';
}

function formatAmount(value) {
  if (value === null || value === undefined || value === '') {
    return '';
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
    const data = await pending;
    return storeCompanyMetadata(realmId, data);
  } catch (error) {
    companyMetadataCache.delete(realmId);
    throw error;
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

  return payload?.metadata || {
    vendors: { items: [] },
    accounts: { items: [] },
    taxCodes: { items: [] },
    taxAgencies: { items: [] },
  };
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
    taxAgencies: prepareMetadataSection(metadata?.taxAgencies),
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

async function refreshCompanyMetadata(realmId, button) {
  if (!realmId) {
    return;
  }

  if (button) {
    button.disabled = true;
    button.textContent = 'Refreshing…';
  }

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

    await ensureCompanyMetadata(realmId, { force: true });
    await refreshQuickBooksCompanies();
    showCompaniesStatus('QuickBooks metadata refreshed.');
  } catch (error) {
    console.error(error);
    showCompaniesStatus(error.message || 'Failed to refresh QuickBooks data.', 'error');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = 'Refresh QuickBooks Data';
    }
  }
}
