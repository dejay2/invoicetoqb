const assert = require('node:assert');
const fs = require('node:fs/promises');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');

async function withCompanyStore(testFn) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'qb-companies-endpoint-'));
  const storePath = path.join(tempDir, 'quickbooks_companies.json');
  const originalEnv = process.env.QUICKBOOKS_COMPANIES_FILE;

  const seededCompanies = [
    {
      realmId: 'seed-1',
      companyName: 'Seed Company 1',
      legalName: 'Seed Company 1 Legal',
      tokens: { accessToken: 'secret-token', refreshToken: 'refresh-token' },
      environment: 'sandbox',
      connectedAt: '2024-09-20T10:00:00.000Z',
      updatedAt: '2024-09-20T10:00:00.000Z',
    },
    {
      realmId: 'seed-2',
      companyName: 'Seed Company 2',
      legalName: 'Seed Company 2 Legal',
      tokens: null,
      environment: 'production',
      connectedAt: '2024-09-21T10:00:00.000Z',
      updatedAt: '2024-09-21T10:00:00.000Z',
    },
  ];

  await fs.writeFile(storePath, `${JSON.stringify(seededCompanies, null, 2)}\n`, 'utf8');

  try {
    process.env.QUICKBOOKS_COMPANIES_FILE = storePath;
    delete require.cache[require.resolve('../src/server')];
    const { app, readQuickBooksCompanies } = require('../src/server');

    const server = http.createServer(app);
    await new Promise((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, () => resolve());
    });

    try {
      const port = server.address().port;
      await testFn({ port, readQuickBooksCompanies });
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  } finally {
    delete require.cache[require.resolve('../src/server')];
    if (originalEnv === undefined) {
      delete process.env.QUICKBOOKS_COMPANIES_FILE;
    } else {
      process.env.QUICKBOOKS_COMPANIES_FILE = originalEnv;
    }
  }
}

(async () => {
  await withCompanyStore(async ({ port, readQuickBooksCompanies }) => {
    const companies = await readQuickBooksCompanies();
    assert.strictEqual(companies.length, 2, 'helper should return seeded companies');

    const response = await fetch(`http://127.0.0.1:${port}/api/quickbooks/companies`);
    const raw = await response.text();
    assert.ok(raw.length > 0, 'expected non-empty response body');
    assert.strictEqual(response.status, 200, 'expected 200 response');

    const payload = JSON.parse(raw);
    assert.ok(Array.isArray(payload.companies), 'payload.companies should be an array');
    assert.strictEqual(payload.companies.length, 2, 'expected both seeded companies in response');
    payload.companies.forEach((company) => {
      assert.strictEqual(Object.prototype.hasOwnProperty.call(company, 'tokens'), false, 'tokens should not be exposed');
    });
  });

  console.log('quickbooks-companies endpoint regression passed');
})();
