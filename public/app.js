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
const matchVendorInput = document.getElementById('match-vendor');
const matchAccountInput = document.getElementById('match-account');
const matchTaxRateInput = document.getElementById('match-tax-rate');
const matchInvoiceNumberInput = document.getElementById('match-invoice-number');
const matchInvoiceNameInput = document.getElementById('match-invoice-name');
const matchSubtotalInput = document.getElementById('match-subtotal');
const matchTaxInput = document.getElementById('match-tax');
const matchTotalInput = document.getElementById('match-total');

const MATCH_STORAGE_KEY = 'invoiceMatches';
let quickBooksCompanies = [];
let currentInvoiceData = null;
let currentInvoiceMetadata = null;
let savedMatches = loadSavedMatches();
const matcherEditableInputs = [matchCompanySelect, matchVendorInput, matchAccountInput, matchTaxRateInput];
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
        populateMatcherFromInvoice(payload.invoice);
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

    hide(qbStatus);
    hide(companiesStatus);
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
    title.textContent = company.companyName || company.realmId;

    const meta = document.createElement('div');
    meta.className = 'company-edit-meta';
    meta.textContent = formatQuickBooksMeta(company);

    const form = document.createElement('form');
    form.className = 'company-edit-form';
    form.dataset.realmId = company.realmId;

    const label = document.createElement('label');
    label.className = 'input-group';

    const span = document.createElement('span');
    span.textContent = 'Display Name';

    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = company.realmId;
    input.value = company.companyName || '';
    input.required = true;

    label.append(span, input);

    const button = document.createElement('button');
    button.type = 'submit';
    button.textContent = 'Save Name';

    form.append(label, button);
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
  const input = form.querySelector('input');
  const button = form.querySelector('button[type="submit"]');

  if (!realmId || !input) {
    return;
  }

  const newName = input.value.trim();
  if (!newName) {
    showCompaniesStatus('Company name cannot be empty.', 'error');
    input.focus();
    return;
  }

  if (button) {
    button.disabled = true;
    button.textContent = 'Saving…';
  }

  try {
    const updated = await updateCompanyName(realmId, newName);
    updateSavedMatchesWithCompany(updated);
    showCompaniesStatus(`Updated ${updated.companyName || updated.realmId}.`);
    await refreshQuickBooksCompanies();
  } catch (error) {
    console.error(error);
    showCompaniesStatus(error.message || 'Failed to update company name.', 'error');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = 'Save Name';
    }
  }
}

async function updateCompanyName(realmId, companyName) {
  const response = await fetch(`/api/quickbooks/companies/${encodeURIComponent(realmId)}`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ companyName }),
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

  return payload?.company || { realmId, companyName };
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
    option.textContent = company.companyName || `Realm ${company.realmId}`;
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

function populateMatcherFromInvoice(invoicePayload) {
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
  const savedMatch = checksum ? savedMatches[checksum] : null;

  const preferredRealmId = savedMatch?.companyRealmId || null;
  renderMatchCompanyOptions(quickBooksCompanies, preferredRealmId);

  if (savedMatch?.companyRealmId && !quickBooksCompanies.some((company) => company.realmId === savedMatch.companyRealmId)) {
    ensureCompanyOption(savedMatch.companyRealmId, savedMatch.companyName || `Realm ${savedMatch.companyRealmId}`);
    matchCompanySelect.disabled = false;
    matchCompanySelect.value = savedMatch.companyRealmId;
  }

  if (matchVendorInput) {
    matchVendorInput.value = savedMatch?.vendor ?? invoice.vendor ?? '';
  }

  if (matchAccountInput) {
    matchAccountInput.value = savedMatch?.account ?? '';
  }

  if (matchTaxRateInput) {
    matchTaxRateInput.value = savedMatch?.taxRate ?? invoice.taxCode ?? '';
  }

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

  if (savedMatch?.companyRealmId && Array.from(matchCompanySelect?.options || []).some((option) => option.value === savedMatch.companyRealmId)) {
    matchCompanySelect.value = savedMatch.companyRealmId;
  }

  if (savedMatch) {
    showMatchStatus('Loaded saved QuickBooks match for this invoice.', 'info');
  } else {
    showMatchStatus('Review QuickBooks mappings before export.', 'info');
  }
}

function resetMatcherForm() {
  currentInvoiceData = null;
  currentInvoiceMetadata = null;
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

  const record = {
    checksum,
    updatedAt: new Date().toISOString(),
    companyRealmId,
    companyName: company?.companyName || company?.realmId || null,
    vendor: matchVendorInput?.value?.trim() || null,
    account: matchAccountInput?.value?.trim() || null,
    taxRate: matchTaxRateInput?.value?.trim() || null,
    invoice: {
      invoiceNumber: currentInvoiceData?.invoiceNumber || null,
      invoiceName: matchInvoiceNameInput?.value || null,
      subtotal: currentInvoiceData?.subtotal ?? null,
      tax: currentInvoiceData?.vatAmount ?? currentInvoiceData?.taxAmount ?? null,
      total: currentInvoiceData?.totalAmount ?? null,
      originalName: currentInvoiceMetadata?.originalName || null,
    },
  };

  saveMatchRecord(record);

  if (record.companyName) {
    showMatchStatus(`Saved match for ${record.companyName}.`);
  } else {
    showMatchStatus('Saved match configuration.');
  }
}

function formatQuickBooksMeta(company) {
  const parts = [];
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
      return parsed;
    }
  } catch (error) {
    console.warn('Failed to load saved matches', error);
  }

  return {};
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

  savedMatches[record.checksum] = record;
  persistSavedMatches();
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
