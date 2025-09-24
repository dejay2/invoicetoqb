const { test, expect } = require('@playwright/test');
const path = require('node:path');
const fs = require('node:fs/promises');
const { spawn } = require('node:child_process');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const DATA_DIR = path.join(PROJECT_ROOT, 'data');
const QUICKBOOKS_DIR = path.join(DATA_DIR, 'quickbooks');
const REALM_ID = 'playwright-realm';
const PORT = 5180;
const CHECKSUM = 'playwright-checksum';

const parsedInvoicesPath = path.join(DATA_DIR, 'parsed_invoices.json');
const companiesPath = path.join(DATA_DIR, 'quickbooks_companies.json');
const metadataDir = path.join(QUICKBOOKS_DIR, REALM_ID);
const vendorsPath = path.join(metadataDir, 'vendors.json');
const accountsPath = path.join(metadataDir, 'accounts.json');
const taxCodesPath = path.join(metadataDir, 'taxCodes.json');
const vendorSettingsPath = path.join(metadataDir, 'vendor-settings.json');

async function readOriginal(pathname) {
  try {
    return await fs.readFile(pathname, 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') {
      return null;
    }
    throw error;
  }
}

async function restoreFile(pathname, contents) {
  if (contents === null) {
    await fs.rm(pathname, { force: true });
    return;
  }

  await fs.mkdir(path.dirname(pathname), { recursive: true });
  await fs.writeFile(pathname, contents);
}

async function writeJson(pathname, data) {
  await fs.mkdir(path.dirname(pathname), { recursive: true });
  await fs.writeFile(pathname, JSON.stringify(data, null, 2));
}

async function prepareFixtures() {
  const now = new Date().toISOString();

  await writeJson(parsedInvoicesPath, [
    {
      parsedAt: now,
      status: 'review',
      metadata: {
        checksum: CHECKSUM,
        originalName: 'playwright-invoice.pdf',
        companyProfile: {
          realmId: REALM_ID,
        },
      },
      data: {
        invoiceNumber: 'PW-1001',
        invoiceDate: '2024-10-10',
        vendor: 'Vendor One',
        subtotal: 100,
        vatAmount: 20,
        totalAmount: 120,
      },
      reviewSelection: {
        vendorId: 'VEN-1',
        accountId: 'ACC-1',
      },
    },
  ]);

  await writeJson(companiesPath, [
    {
      realmId: REALM_ID,
      companyName: 'Playwright Realm',
      legalName: 'Playwright Realm LLC',
      environment: 'sandbox',
      connectedAt: now,
      updatedAt: now,
      vendorsCount: 1,
      accountsCount: 1,
      taxCodesCount: 1,
      oneDrive: null,
      outlook: null,
    },
  ]);

  await writeJson(vendorsPath, {
    updatedAt: now,
    items: [
      { id: 'VEN-1', displayName: 'Vendor One' },
    ],
  });

  await writeJson(accountsPath, {
    updatedAt: now,
    items: [
      { id: 'ACC-1', name: 'Expenses' },
    ],
  });

  await writeJson(taxCodesPath, {
    updatedAt: now,
    items: [
      { id: 'TAX-1', name: 'Standard VAT', rate: 20 },
    ],
  });

  await writeJson(vendorSettingsPath, {});
}

async function cleanupFixtures(originals) {
  await restoreFile(parsedInvoicesPath, originals.parsedInvoices);
  await restoreFile(companiesPath, originals.companies);

  const noMetadataFiles =
    originals.vendors === null &&
    originals.accounts === null &&
    originals.taxCodes === null &&
    originals.vendorSettings === null;

  if (noMetadataFiles) {
    await fs.rm(metadataDir, { recursive: true, force: true });
    return;
  }

  if (originals.vendors === null) {
    await fs.rm(vendorsPath, { force: true });
  } else {
    await writeJson(vendorsPath, JSON.parse(originals.vendors));
  }

  if (originals.accounts === null) {
    await fs.rm(accountsPath, { force: true });
  } else {
    await writeJson(accountsPath, JSON.parse(originals.accounts));
  }

  if (originals.taxCodes === null) {
    await fs.rm(taxCodesPath, { force: true });
  } else {
    await writeJson(taxCodesPath, JSON.parse(originals.taxCodes));
  }

  if (originals.vendorSettings === null) {
    await fs.rm(vendorSettingsPath, { force: true });
  } else {
    await writeJson(vendorSettingsPath, JSON.parse(originals.vendorSettings));
  }
}

async function waitForServerReady(process) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error('Server start timed out'));
    }, 10000);

    function handleData(chunk) {
      if (chunk.toString().includes('Invoice parser listening')) {
        cleanup();
        resolve();
      }
    }

    function handleExit(code) {
      cleanup();
      reject(new Error(`Server exited early with code ${code}`));
    }

    function cleanup() {
      clearTimeout(timeout);
      process.stdout.off('data', handleData);
      process.stderr.off('data', handleData);
      process.off('exit', handleExit);
    }

    process.stdout.on('data', handleData);
    process.stderr.on('data', handleData);
    process.on('exit', handleExit);
  });
}

test.describe('QuickBooks review flow', () => {
  /** @type {import('node:child_process').ChildProcessWithoutNullStreams | null} */
  let serverProcess = null;
  const originals = {};

  test.beforeAll(async () => {
    originals.parsedInvoices = await readOriginal(parsedInvoicesPath);
    originals.companies = await readOriginal(companiesPath);
    originals.vendors = await readOriginal(vendorsPath);
    originals.accounts = await readOriginal(accountsPath);
    originals.taxCodes = await readOriginal(taxCodesPath);
    originals.vendorSettings = await readOriginal(vendorSettingsPath);

    await prepareFixtures();

    serverProcess = spawn('node', ['src/server.js'], {
      cwd: PROJECT_ROOT,
      env: {
        ...process.env,
        PORT: String(PORT),
        NODE_ENV: 'test',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    serverProcess.stdout.setEncoding('utf8');
    serverProcess.stderr.setEncoding('utf8');

    await waitForServerReady(serverProcess);
  });

  test.afterAll(async () => {
    if (serverProcess) {
      serverProcess.kill();
      await new Promise((resolve) => {
        serverProcess?.once('exit', () => resolve());
      });
      serverProcess = null;
    }

    await cleanupFixtures(originals);
  });

  test('enables QuickBooks preview after assigning tax code', async ({ page }) => {
    await page.goto(`http://127.0.0.1:${PORT}/`);

    await page.getByRole('tab', { name: 'To review' }).click();

    const taxSelect = page.locator('select.review-tax-select');
    await expect(taxSelect).toBeVisible();

    const previewButton = page.getByRole('button', { name: 'Preview QB' });
    await expect(previewButton).toBeDisabled();

    await taxSelect.selectOption('TAX-1');

    await expect(previewButton).toBeEnabled();

    const file = await fs.readFile(parsedInvoicesPath, 'utf8');
    const invoices = JSON.parse(file);
    expect(invoices[0]?.reviewSelection?.taxCodeId).toBe('TAX-1');
  });
});
