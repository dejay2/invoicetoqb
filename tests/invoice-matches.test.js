const assert = require('assert');
const { buildInvoiceMatches } = require('../src/server');

function createMetadata() {
  return {
    vendors: {
      items: [
        { id: 'vendor-1', displayName: 'Beach Restaurant' },
        { id: 'vendor-2', displayName: 'Harbor Cafe' },
      ],
    },
    accounts: {
      items: [
        { id: 'acct-1', fullyQualifiedName: 'Expenses:Meals' },
        { id: 'acct-2', fullyQualifiedName: 'Expenses:Supplies' },
      ],
    },
    taxCodes: {
      items: [
        { id: 'tax-std', name: 'Standard VAT' },
        { id: 'tax-zero', name: 'Zero VAT' },
      ],
    },
    vendorSettings: {
      'vendor-2': {
        accountId: 'acct-2',
        taxCodeId: 'tax-zero',
      },
    },
  };
}

(function shouldHonorStoredSelections() {
  const invoice = {
    data: {
      vendor: 'Beach Restaurant',
      suggestedAccount: null,
    },
    metadata: {
      companyProfile: { realmId: 'realm-test' },
      invoiceFilename: 'beach-restaurant-invoice.pdf',
    },
    reviewSelection: {
      vendorId: 'vendor-1',
      accountId: 'acct-1',
      taxCodeId: 'tax-std',
    },
  };

  const matches = buildInvoiceMatches(invoice, createMetadata());

  assert.ok(matches);
  assert.deepStrictEqual(matches.vendor.status, 'exact');
  assert.strictEqual(matches.vendor.id, 'vendor-1');
  assert.strictEqual(matches.vendor.label, 'Beach Restaurant');

  assert.deepStrictEqual(matches.account.status, 'exact');
  assert.strictEqual(matches.account.id, 'acct-1');
  assert.strictEqual(matches.account.label, 'Expenses:Meals');

  assert.deepStrictEqual(matches.taxCode.status, 'exact');
  assert.strictEqual(matches.taxCode.id, 'tax-std');
  assert.strictEqual(matches.taxCode.label, 'Standard VAT');

  console.log('✓ matches respect stored review selections');
})();

(function shouldInferMatchesFromMetadata() {
  const metadata = createMetadata();
  const invoice = {
    data: {
      vendor: 'Harbor Cafe',
      suggestedAccount: {
        name: 'Expenses:Supplies',
        confidence: 'high',
        reason: 'Strong keyword overlap',
      },
      taxCode: 'Zero VAT',
    },
    metadata: {
      companyProfile: { realmId: 'realm-test' },
      invoiceFilename: 'harbor-cafe.pdf',
    },
    reviewSelection: null,
  };

  const matches = buildInvoiceMatches(invoice, metadata);

  assert.ok(matches);
  assert.deepStrictEqual(matches.vendor.status, 'exact');
  assert.strictEqual(matches.vendor.id, 'vendor-2');
  assert.strictEqual(matches.vendor.label, 'Harbor Cafe');

  assert.deepStrictEqual(matches.account.status, 'exact');
  assert.strictEqual(matches.account.id, 'acct-2');
  assert.strictEqual(matches.account.label, 'Expenses:Supplies');
  assert.match(matches.account.reason || '', /default|AI/i);

  assert.deepStrictEqual(matches.taxCode.status, 'exact');
  assert.strictEqual(matches.taxCode.id, 'tax-zero');
  assert.strictEqual(matches.taxCode.label, 'Zero VAT');

  console.log('✓ matches derive vendor defaults and suggestions');
})();
