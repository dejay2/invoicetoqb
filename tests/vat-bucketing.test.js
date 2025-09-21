const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const {
  buildQuickBooksExpenseLines,
  resolveQuickBooksTaxCodes,
  sumLineAmounts,
} = require('../src/server');

function loadFixture(filename) {
  const filePath = path.join(__dirname, 'fixtures', 'vat', filename);
  const raw = fs.readFileSync(filePath, 'utf8');
  return JSON.parse(raw);
}

const TAX_CODES = [
  { id: 'TAX_STD', name: 'Standard VAT', rate: 20 },
  { id: 'TAX_ZERO', name: 'Zero VAT', rate: 0 },
];

const TAX_LOOKUP = new Map(TAX_CODES.map((entry) => [entry.id, entry]));
const ACCOUNT = { id: 'ACCT_EXP', name: 'Expenses' };
const VENDOR_DEFAULTS = {};

const scenarios = [
  {
    name: 'single-rate invoice uses one QuickBooks line',
    fixture: 'single-rate-invoice.json',
    expectedLineCount: 1,
    expectRequiresSecondary: false,
    expectedBuckets: ['standard'],
  },
  {
    name: 'mixed-rate invoice with Gemini hints collapses into two VAT buckets',
    fixture: 'mixed-rate-with-hints.json',
    expectedLineCount: 2,
    expectRequiresSecondary: false,
    expectedBuckets: ['standard', 'zero'],
    expectedTaxCodes: {
      standard: 'TAX_STD',
      zero: 'TAX_ZERO',
    },
  },
  {
    name: 'mixed-rate invoice without Gemini hints falls back to vendor defaults',
    fixture: 'mixed-rate-missing-hints.json',
    expectedLineCount: 1,
    expectRequiresSecondary: false,
    expectedBuckets: ['standard'],
  },
];

scenarios.forEach((scenario) => {
  const { invoice, expectedTotal } = loadFixture(scenario.fixture);

  const taxResolution = resolveQuickBooksTaxCodes(invoice, TAX_LOOKUP, VENDOR_DEFAULTS);
  assert.ok(taxResolution.primaryTaxCodeId, `${scenario.name}: primary tax code should resolve.`);

  const result = buildQuickBooksExpenseLines(invoice, {
    account: ACCOUNT,
    taxLookup: TAX_LOOKUP,
    taxResolution,
  });

  assert.ok(Array.isArray(result.lines), `${scenario.name}: lines array should be produced.`);
  assert.ok(result.lines.length <= 2, `${scenario.name}: QuickBooks lines capped at two buckets.`);
  assert.strictEqual(
    result.lines.length,
    scenario.expectedLineCount,
    `${scenario.name}: unexpected number of QuickBooks lines.`
  );

  const totalAmount = sumLineAmounts(result.lines);
  assert.strictEqual(totalAmount, expectedTotal, `${scenario.name}: total amount should be preserved.`);

  if (typeof scenario.expectRequiresSecondary === 'boolean') {
    assert.strictEqual(
      result.requiresSecondaryTaxCode,
      scenario.expectRequiresSecondary,
      `${scenario.name}: secondary tax requirement mismatch.`
    );
  }

  if (Array.isArray(scenario.expectedBuckets)) {
    const bucketKeys = result.vatBuckets.map((bucket) => bucket.key);
    assert.deepStrictEqual(
      bucketKeys,
      scenario.expectedBuckets,
      `${scenario.name}: VAT buckets did not match expected order.`
    );
  }

  if (scenario.expectedTaxCodes) {
    scenario.expectedBuckets.forEach((bucketKey) => {
      const bucket = result.vatBuckets.find((entry) => entry.key === bucketKey);
      assert.ok(bucket, `${scenario.name}: expected ${bucketKey} bucket to be present.`);
      if (scenario.expectedTaxCodes[bucketKey]) {
        assert.strictEqual(
          bucket.taxCodeId,
          scenario.expectedTaxCodes[bucketKey],
          `${scenario.name}: ${bucketKey} bucket tax code mismatch.`
        );
      }
    });
  }
});

console.log('vat-bucketing regression tests passed');
