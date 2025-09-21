const assert = require('node:assert');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

async function withTempCompanyStore(testFn) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'qb-company-store-'));
  const storePath = path.join(tempDir, 'quickbooks_companies.json');
  const originalEnv = process.env.QUICKBOOKS_COMPANIES_FILE;

  try {
    process.env.QUICKBOOKS_COMPANIES_FILE = storePath;
    delete require.cache[require.resolve('../src/server')];
    const {
      persistQuickBooksCompanies,
      readQuickBooksCompanies,
      getHealthMetricHistory,
    } = require('../src/server');

    await testFn({
      storePath,
      persistQuickBooksCompanies,
      readQuickBooksCompanies,
      getHealthMetricHistory,
    });
  } finally {
    delete require.cache[require.resolve('../src/server')];
    if (originalEnv === undefined) {
      delete process.env.QUICKBOOKS_COMPANIES_FILE;
    } else {
      process.env.QUICKBOOKS_COMPANIES_FILE = originalEnv;
    }
  }
}

function buildCompany(index) {
  return {
    realmId: `realm-${index}`,
    companyName: `Company ${index}`,
    legalName: `Company ${index} Legal`,
    tokens: null,
    environment: 'sandbox',
    connectedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

async function testConcurrentWrites(context) {
  const { storePath, persistQuickBooksCompanies } = context;
  const batches = Array.from({ length: 5 }, (_, batchIndex) =>
    Array.from({ length: 3 }, (_, itemIndex) => buildCompany(`${batchIndex}-${itemIndex}`))
  );

  await Promise.all(batches.map((companies) => persistQuickBooksCompanies(companies)));

  const raw = await fs.readFile(storePath, 'utf8');
  assert.ok(raw.endsWith('\n'), 'expected file to end with newline');
  const parsed = JSON.parse(raw);
  assert.deepStrictEqual(parsed, batches[batches.length - 1], 'latest write should win deterministically');
}

async function testRepairLogic(context) {
  const { storePath, persistQuickBooksCompanies, readQuickBooksCompanies, getHealthMetricHistory } = context;
  const baseline = [buildCompany('A'), buildCompany('B')];
  await persistQuickBooksCompanies(baseline);

  const validPayload = `${JSON.stringify(baseline, null, 2)}\n`;
  const corruptPayload = `${validPayload}]: "2025-09-20T21:27:51.646Z", "extra": true`;
  await fs.writeFile(storePath, corruptPayload, 'utf8');

  const companies = await readQuickBooksCompanies();
  assert.strictEqual(companies.length, baseline.length, 'repair should preserve original entries');

  const repairedRaw = await fs.readFile(storePath, 'utf8');
  assert.strictEqual(repairedRaw, validPayload, 'repair should truncate to the first balanced array closing bracket');

  const entries = await fs.readdir(path.dirname(storePath));
  const backupFile = entries.find((entry) => entry.startsWith('quickbooks_companies.json.corrupt-'));
  assert.ok(backupFile, 'expected corrupt backup file to be created');

  const metrics = getHealthMetricHistory();
  assert.ok(
    metrics.some((entry) => entry.name === 'quickbooks.company_file.repaired'),
    'repair should emit health metric'
  );
}

(async () => {
  await withTempCompanyStore(async (context) => {
    await testConcurrentWrites(context);
    await testRepairLogic(context);
  });

  console.log('quickbooks-companies-store regression passed');
})();
