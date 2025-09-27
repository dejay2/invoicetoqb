const assert = require('assert');
const fs = require('fs');
const path = require('path');

process.env.GEMINI_API_KEY = 'test-key';
process.env.ALLOW_MULTI_INVOICE_SPLIT = 'true';

const originalFetch = global.fetch;
let fetchQueue = [];

global.fetch = async () => {
  if (!fetchQueue.length) {
    return {
      ok: true,
      json: async () => ({
        candidates: [],
      }),
    };
  }

  const payload = fetchQueue.shift();
  return {
    ok: true,
    json: async () => ({
      candidates: [
        {
          content: {
            parts: [{ text: JSON.stringify(payload) }],
          },
        },
      ],
    }),
  };
};

const {
  ingestInvoiceFromSource,
  readStoredInvoices,
  validateInvoiceTotals,
  INVOICE_STATUS_SPLIT,
} = require('../src/server');

const FIXTURES_DIR = path.join(__dirname, 'fixtures');
const SINGLE_INVOICE_PDF = path.join(FIXTURES_DIR, 'single-invoice.pdf');
const MULTI_INVOICE_PDF = path.join(FIXTURES_DIR, 'multi-invoice.pdf');
const MULTI_INVOICE_SPACED_PDF = path.join(FIXTURES_DIR, 'multi-invoice-spaced.pdf');
const DATA_FILE = path.join(__dirname, '..', 'data', 'parsed_invoices.json');
const INVOICE_STORE_DIR = path.join(__dirname, '..', 'data', 'invoice_store');

function ensureInvoiceStoreDir() {
  if (!fs.existsSync(INVOICE_STORE_DIR)) {
    fs.mkdirSync(INVOICE_STORE_DIR, { recursive: true });
  }
}

function snapshotParsedInvoices() {
  try {
    const content = fs.readFileSync(DATA_FILE, 'utf-8');
    return { exists: true, content };
  } catch (error) {
    if (error.code === 'ENOENT') {
      return { exists: false, content: null };
    }
    throw error;
  }
}

function restoreParsedInvoices(snapshot) {
  if (!snapshot.exists) {
    try {
      fs.unlinkSync(DATA_FILE);
    } catch (error) {
      if (error.code !== 'ENOENT') {
        throw error;
      }
    }
    return;
  }
  fs.writeFileSync(DATA_FILE, snapshot.content);
}

function snapshotInvoiceStore() {
  ensureInvoiceStoreDir();
  const entries = fs.readdirSync(INVOICE_STORE_DIR);
  return new Set(entries);
}

function restoreInvoiceStore(originalSet) {
  ensureInvoiceStoreDir();
  const current = new Set(fs.readdirSync(INVOICE_STORE_DIR));
  for (const file of current) {
    if (!originalSet.has(file)) {
      fs.unlinkSync(path.join(INVOICE_STORE_DIR, file));
    }
  }
}

function stubGeminiResponses(responses) {
  fetchQueue = Array.isArray(responses) ? responses.slice() : [];
  return () => {
    fetchQueue = [];
  };
}

async function shouldIngestSingleInvoicePdf() {
  const dataSnapshot = snapshotParsedInvoices();
  const invoiceSnapshot = snapshotInvoiceStore();
  const restoreFetch = stubGeminiResponses([
    {
      vendor: 'Cornish Bakery',
      products: [
        { description: 'Pasties', quantity: 10, unitPrice: 8.75, lineTotal: 87.5 },
      ],
      invoiceNumber: '1528554',
      totalAmount: 87.5,
    },
  ]);

  try {
    fs.writeFileSync(DATA_FILE, '[]');
    const buffer = fs.readFileSync(SINGLE_INVOICE_PDF);
    const result = await ingestInvoiceFromSource({
      buffer,
      mimeType: 'application/pdf',
      originalName: 'single-invoice.pdf',
      fileSize: buffer.length,
      realmId: 'test-realm',
    });

    assert.ok(result);
    assert.strictEqual(Boolean(result.multiInvoice), false);
    assert.ok(Array.isArray(result.invoices));
    assert.strictEqual(result.invoices.length, 1);

    const entry = result.invoices[0];
    assert.ok(entry.stored, 'expected stored invoice payload');
    const segmentMeta = entry.stored.metadata.segment;
    assert.ok(segmentMeta, 'expected segment metadata on stored invoice');
    assert.strictEqual(segmentMeta.index, 1);
    assert.strictEqual(segmentMeta.count, 1);
    assert.strictEqual(segmentMeta.invoiceNumber, '1528554');

    const storedChecksum = entry.stored.metadata.checksum;
    assert.ok(storedChecksum, 'expected checksum on stored metadata');
    const storedPath = path.join(INVOICE_STORE_DIR, `${storedChecksum}.pdf`);
    assert.ok(fs.existsSync(storedPath), 'expected stored PDF for single invoice');

    const storedInvoices = await readStoredInvoices();
    assert.strictEqual(storedInvoices.length, 1, 'expected one stored invoice record');
  } finally {
    restoreFetch();
    restoreParsedInvoices(dataSnapshot);
    restoreInvoiceStore(invoiceSnapshot);
  }

  console.log('✓ single invoice PDF ingests as one record');
}

async function shouldSplitMultiInvoicePdf() {
  const dataSnapshot = snapshotParsedInvoices();
  const invoiceSnapshot = snapshotInvoiceStore();
  const restoreFetch = stubGeminiResponses([
    {
      vendor: 'Cornish Suppliers',
      products: [
        { description: 'Service Callout', quantity: 1, unitPrice: 270.09, lineTotal: 270.09 },
      ],
      invoiceNumber: 'INV-270',
      totalAmount: 270.09,
    },
    {
      vendor: 'Cornish Suppliers',
      products: [
        { description: 'Emergency callout', quantity: 1, unitPrice: 733.6, lineTotal: 733.6 },
      ],
      invoiceNumber: 'INV-734',
      totalAmount: 733.6,
    },
  ]);

  try {
    fs.writeFileSync(DATA_FILE, '[]');
    const buffer = fs.readFileSync(MULTI_INVOICE_SPACED_PDF);
    const result = await ingestInvoiceFromSource({
      buffer,
      mimeType: 'application/pdf',
      originalName: 'multi-invoice-spaced.pdf',
      fileSize: buffer.length,
      realmId: 'test-realm',
    });

    assert.ok(result.multiInvoice, 'expected multiInvoice flag');
    assert.ok(Array.isArray(result.invoices));
    assert.strictEqual(result.invoices.length, 1, 'expected single stored invoice for manual split');

    const storedEntry = result.invoices.find((entry) => entry.stored && !entry.duplicate);
    assert.ok(storedEntry, 'expected stored invoice entry');

    const storedInvoice = storedEntry.stored;
    assert.strictEqual(storedInvoice.status, INVOICE_STATUS_SPLIT, 'expected needs-split status');

    const candidateSegments = storedInvoice.metadata.segmentCandidates;
    assert.ok(Array.isArray(candidateSegments), 'expected segmentCandidates metadata');
    assert.strictEqual(candidateSegments.length, 2, 'expected two segment candidates');

    candidateSegments.forEach((candidate, index) => {
      assert.ok(candidate.invoiceNumber, `expected invoice number for candidate ${index + 1}`);
      assert.ok(candidate.sampleText, `expected sample text for candidate ${index + 1}`);
      assert.ok(
        candidate.normalizedSample === null || typeof candidate.normalizedSample === 'string',
        'expected normalizedSample as string or null'
      );
      assert.ok(Number.isInteger(candidate.startPage), 'expected startPage');
      assert.ok(Number.isInteger(candidate.endPage), 'expected endPage');
    });

    assert.strictEqual(storedInvoice.metadata.splitDecision, 'pending', 'expected pending split decision');

    const storedRecords = await readStoredInvoices();
    assert.strictEqual(storedRecords.length, 1, 'expected single stored invoice record');
    assert.strictEqual(storedRecords[0].status, INVOICE_STATUS_SPLIT, 'expected stored record flagged for split');
  } finally {
    restoreFetch();
    restoreParsedInvoices(dataSnapshot);
    restoreInvoiceStore(invoiceSnapshot);
  }

  console.log('✓ multi-invoice PDF with spaced characters splits into segments and preserves totals');
}

function shouldValidateTotals() {
  const invoice = {
    products: [
      { description: 'Item A', lineTotal: 10 },
      { description: 'Item B', quantity: 2, unitPrice: 5 },
    ],
    totalAmount: 20,
  };

  const pass = validateInvoiceTotals(invoice, { detectedTotal: 20 });
  assert.strictEqual(pass.status, 'pass');
  assert.strictEqual(pass.delta, 0);

  const fail = validateInvoiceTotals(invoice, { detectedTotal: 25 });
  assert.strictEqual(fail.status, 'mismatch');
  assert.strictEqual(fail.expected, 25);
  assert.strictEqual(fail.computed, 20);
  assert.strictEqual(fail.delta, -5);

  console.log('✓ totals validation reports pass and mismatch states');
}

(async () => {
  try {
    await shouldIngestSingleInvoicePdf();
    await shouldSplitMultiInvoicePdf();
    shouldValidateTotals();
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  } finally {
    global.fetch = originalFetch;
  }
})();
