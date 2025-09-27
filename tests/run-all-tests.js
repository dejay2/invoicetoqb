const fs = require('fs');
const path = require('path');

const TEST_FILES = [
  'quickbooks-companies-store.test.js',
  'quickbooks-preview-state.test.js',
  'vat-bucketing.test.js',
  'onedrive-endpoints.test.js',
  'quickbooks-companies-endpoint.test.js',
  'deletion-endpoint.test.js',
  'invoice-matches.test.js',
  'ingestion-flow.test.js',
  'split-flow.test.js',
];

let failures = 0;

TEST_FILES.forEach((fileName) => {
  const filePath = path.join(__dirname, fileName);
  if (!fs.existsSync(filePath)) {
    console.warn(`[tests] Skipping missing test: ${fileName}`);
    return;
  }

  try {
    require(filePath);
  } catch (error) {
    failures += 1;
    console.error(`[tests] ${fileName} failed`, error);
  }
});

if (failures > 0) {
  process.exitCode = 1;
}
