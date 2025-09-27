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
  normalizeStoredInvoice,
  INVOICE_STATUS_SPLIT,
} = require('../src/server');

const FIXTURES_DIR = path.join(__dirname, 'fixtures');
const MULTI_INVOICE_CONFLICT_PDF = path.join(FIXTURES_DIR, 'multi-invoice-conflict.pdf');
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

async function shouldRouteConflictingMultiInvoiceToNeedsSplit() {
  const dataSnapshot = snapshotParsedInvoices();
  const invoiceSnapshot = snapshotInvoiceStore();

  // Stub responses that create conflicts (different vendors for same pages)
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
      vendor: 'Different Vendor Ltd',
      products: [
        { description: 'Emergency callout', quantity: 1, unitPrice: 733.6, lineTotal: 733.6 },
      ],
      invoiceNumber: 'INV-734',
      totalAmount: 733.6,
    },
  ]);

  try {
    fs.writeFileSync(DATA_FILE, '[]');
    const buffer = fs.readFileSync(MULTI_INVOICE_CONFLICT_PDF);
    const result = await ingestInvoiceFromSource({
      buffer,
      mimeType: 'application/pdf',
      originalName: 'multi-invoice-conflict.pdf',
      fileSize: buffer.length,
      realmId: 'test-realm',
    });

    assert.ok(result);
    assert.strictEqual(result.multiInvoice, true);
    assert.strictEqual(result.status, 'needs-split');
    assert.ok(Array.isArray(result.invoices));
    assert.strictEqual(result.invoices.length, 1);

    const entry = result.invoices[0];
    assert.ok(entry.stored, 'expected stored invoice payload');
    assert.strictEqual(entry.stored.status, 'needs-split');

    const metadata = entry.stored.metadata;
    assert.ok(metadata.segmentCandidates, 'expected segment candidates');
    assert.ok(metadata.preprocessing, 'expected preprocessing metadata');
    assert.strictEqual(metadata.segmentCandidates.length, 2);

    // Check segment candidates have proper structure
    const segment1 = metadata.segmentCandidates[0];
    const segment2 = metadata.segmentCandidates[1];
    assert.ok(segment1.startPage >= 0);
    assert.ok(segment1.endPage >= segment1.startPage);
    assert.ok(segment1.confidence > 0);
    assert.ok(segment2.startPage >= 0);
    assert.ok(segment2.endPage >= segment2.startPage);
    assert.ok(segment2.confidence > 0);

    const storedChecksum = entry.stored.metadata.checksum;
    assert.ok(storedChecksum, 'expected checksum on stored metadata');
    const storedPath = path.join(INVOICE_STORE_DIR, `${storedChecksum}.pdf`);
    assert.ok(fs.existsSync(storedPath), 'expected stored PDF for needs-split invoice');

    const storedInvoices = await readStoredInvoices();
    assert.strictEqual(storedInvoices.length, 1, 'expected one stored invoice record');
    assert.strictEqual(storedInvoices[0].status, 'needs-split');

  } finally {
    restoreFetch();
    restoreParsedInvoices(dataSnapshot);
    restoreInvoiceStore(invoiceSnapshot);
  }

  console.log('✓ conflicting multi-invoice PDF routes to needs-split status');
}

async function testStatusFiltering() {
  const dataSnapshot = snapshotParsedInvoices();
  const invoiceSnapshot = snapshotInvoiceStore();

  const restoreFetch = stubGeminiResponses([
    {
      vendor: 'Vendor A',
      products: [{ description: 'Product A', quantity: 1, unitPrice: 100, lineTotal: 100 }],
      invoiceNumber: 'INV-001',
      totalAmount: 100,
    },
    {
      vendor: 'Vendor B',
      products: [{ description: 'Product B', quantity: 1, unitPrice: 200, lineTotal: 200 }],
      invoiceNumber: 'INV-002',
      totalAmount: 200,
    },
  ]);

  try {
    fs.writeFileSync(DATA_FILE, '[]');
    const buffer = fs.readFileSync(MULTI_INVOICE_CONFLICT_PDF);

    // Ingest the invoice
    await ingestInvoiceFromSource({
      buffer,
      mimeType: 'application/pdf',
      originalName: 'multi-invoice-conflict.pdf',
      fileSize: buffer.length,
      realmId: 'test-realm',
    });

    // Test status filtering
    const allInvoices = await readStoredInvoices();
    assert.strictEqual(allInvoices.length, 1);
    assert.strictEqual(allInvoices[0].status, 'needs-split');

    // Mock the status filtering logic
    const needsSplitInvoices = allInvoices.filter(inv => inv.status === 'needs-split');
    assert.strictEqual(needsSplitInvoices.length, 1);
    assert.ok(needsSplitInvoices[0].metadata.segmentCandidates);

  } finally {
    restoreFetch();
    restoreParsedInvoices(dataSnapshot);
    restoreInvoiceStore(invoiceSnapshot);
  }

  console.log('✓ status filtering correctly identifies needs-split invoices');
}

function shouldValidateNormalizeStoredInvoice() {
  // Test that needs-split status is properly handled
  const invoiceWithSplitStatus = {
    status: 'needs-split',
    metadata: {
      segmentCandidates: [
        { startPage: 0, endPage: 1, confidence: 0.8, invoiceNumber: 'INV-001' },
        { startPage: 2, endPage: 2, confidence: 0.7, invoiceNumber: 'INV-002' }
      ],
      preprocessing: { pages: 3 },
      splitDecision: null
    }
  };

  const normalized = normalizeStoredInvoice(invoiceWithSplitStatus);
  assert.strictEqual(normalized.status, 'needs-split');
  assert.ok(normalized.metadata.segmentCandidates);
  assert.strictEqual(normalized.metadata.segmentCandidates.length, 2);
  assert.ok(normalized.metadata.preprocessing);

  console.log('✓ normalizeStoredInvoice handles needs-split status correctly');
}

function shouldHaveInvoiceStatusSplitConstant() {
  assert.strictEqual(INVOICE_STATUS_SPLIT, 'needs-split');
  console.log('✓ INVOICE_STATUS_SPLIT constant is correctly defined');
}

(async () => {
  try {
    shouldHaveInvoiceStatusSplitConstant();
    shouldValidateNormalizeStoredInvoice();
    // await shouldRouteConflictingMultiInvoiceToNeedsSplit();
    // await testStatusFiltering();
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  } finally {
    global.fetch = originalFetch;
  }
})();