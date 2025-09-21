const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const APP_SOURCE = fs.readFileSync(path.join(__dirname, '..', 'public', 'app.js'), 'utf8');

function extractFunctionSource(name) {
  const signature = `function ${name}`;
  const start = APP_SOURCE.indexOf(signature);
  if (start === -1) {
    throw new Error(`Unable to locate ${name} in public/app.js`);
  }

  const bodyStart = APP_SOURCE.indexOf('{', start);
  if (bodyStart === -1) {
    throw new Error(`Unable to locate body for ${name}`);
  }

  let depth = 0;
  let index = bodyStart;
  while (index < APP_SOURCE.length) {
    const char = APP_SOURCE[index];
    if (char === '{') {
      depth += 1;
    } else if (char === '}') {
      depth -= 1;
      if (depth === 0) {
        return APP_SOURCE.slice(start, index + 1);
      }
    }
    index += 1;
  }

  throw new Error(`Unable to extract complete function for ${name}`);
}

function instantiateFunction(name, dependencies = {}) {
  const source = extractFunctionSource(name);
  const dependencyNames = Object.keys(dependencies);
  const factory = new Function(
    ...dependencyNames,
    `${source};\nreturn ${name};`
  );
  return factory(...dependencyNames.map((dependencyName) => dependencies[dependencyName]));
}

const sanitizeReviewSelectionId = instantiateFunction('sanitizeReviewSelectionId');
const prepareMetadataSection = instantiateFunction('prepareMetadataSection', {
  sanitizeReviewSelectionId,
});
const evaluateQuickBooksPreviewState = instantiateFunction('evaluateQuickBooksPreviewState', {
  sanitizeReviewSelectionId,
  findInvoiceTaxCodeId: () => null,
  analyzeInvoiceVatBuckets: () => ({ hasSplit: false, requiresSecondary: false }),
  selectedRealmId: 'test-realm',
});

// The numeric QuickBooks IDs should be normalised to strings in the metadata lookup.
const metadata = {
  vendors: prepareMetadataSection({
    items: [
      { id: 'vendor-001', displayName: 'Vendor 001' },
    ],
  }),
  accounts: prepareMetadataSection({
    items: [
      { id: 101, name: 'Office Supplies' },
    ],
  }),
  taxCodes: prepareMetadataSection({
    items: [
      { id: 29, name: 'Standard VAT' },
    ],
  }),
  vendorSettings: { entries: {} },
};

assert.strictEqual(metadata.accounts.items[0].id, '101');
assert.strictEqual(metadata.taxCodes.items[0].id, '29');
assert.ok(metadata.accounts.lookup.has('101'));
assert.ok(metadata.taxCodes.lookup.has('29'));

const invoice = {
  reviewSelection: {
    vendorId: 'vendor-001',
    accountId: 101,
    taxCodeId: 29,
  },
};

const state = evaluateQuickBooksPreviewState(invoice, metadata);

assert.deepStrictEqual(state.missing, []);
assert.strictEqual(state.canPreview, true);

console.log('quickbooks-preview-state regression passed');
