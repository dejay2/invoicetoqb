#!/usr/bin/env node

const fs = require('node:fs/promises');
const path = require('node:path');

function printUsage() {
  console.log(`Usage: node scripts/quickbooks-companies.js [inspect|repair] [--file <path>]\n\n` +
    `Commands:\n` +
    `  inspect        Validate the QuickBooks company store without modifying it.\n` +
    `  repair         Attempt to repair the QuickBooks company store by truncating corrupt data.\n\n` +
    `Options:\n` +
    `  --file <path>  Override the quickbooks_companies.json location. Defaults to the server setting.\n`);
}

function parseArguments(argv) {
  const args = argv.slice(2);
  let command = 'inspect';
  let fileOverride = null;

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === 'inspect' || arg === 'repair') {
      command = arg;
      continue;
    }
    if (arg === '--file' && index + 1 < args.length) {
      fileOverride = args[index + 1];
      index += 1;
      continue;
    }
    if (arg.startsWith('--file=')) {
      fileOverride = arg.slice('--file='.length);
      continue;
    }
    if (arg === '--help' || arg === '-h') {
      return { command: 'help' };
    }
    throw new Error(`Unknown argument: ${arg}`);
  }

  return { command, fileOverride };
}

async function ensureFileExists(filePath) {
  try {
    await fs.access(filePath);
  } catch (error) {
    if (error.code === 'ENOENT') {
      throw new Error(`No QuickBooks company store found at ${filePath}.`);
    }
    throw error;
  }
}

async function createCorruptBackup(filePath, contents) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = `${filePath}.corrupt-${timestamp}.json`;
  try {
    await fs.writeFile(backupPath, contents, { flag: 'wx' });
    return backupPath;
  } catch (error) {
    if (error.code === 'EEXIST') {
      return backupPath;
    }
    throw error;
  }
}

async function inspectFile(filePath, { readQuickBooksCompanies }) {
  try {
    await ensureFileExists(filePath);
    const companies = await readQuickBooksCompanies({ allowRepair: false });
    const count = Array.isArray(companies) ? companies.length : 0;
    console.log(`✅ ${filePath} is a valid QuickBooks company store (${count} compan${count === 1 ? 'y' : 'ies'}).`);
  } catch (error) {
    if (error.code === 'QUICKBOOKS_COMPANY_FILE_CORRUPT') {
      console.error(`❌ ${error.message}`);
      if (error.backupPath) {
        console.error(`Backup created at: ${error.backupPath}`);
      }
      process.exitCode = 1;
      return;
    }
    console.error(error.message || error);
    process.exitCode = 1;
  }
}

async function repairFile(filePath, { readQuickBooksCompanies, attemptQuickBooksCompaniesRepair }) {
  let raw;
  try {
    raw = await fs.readFile(filePath, 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') {
      console.error(`❌ No QuickBooks company store found at ${filePath}.`);
      process.exitCode = 1;
      return;
    }
    throw error;
  }

  try {
    JSON.parse(raw);
    const companies = await readQuickBooksCompanies({ allowRepair: false });
    const count = Array.isArray(companies) ? companies.length : 0;
    console.log(`✅ ${filePath} is already valid. No repair needed (${count} compan${
      count === 1 ? 'y' : 'ies'
    }).`);
    return;
  } catch (parseError) {
    let backupPath;
    try {
      backupPath = await createCorruptBackup(filePath, raw);
    } catch (error) {
      console.error(`❌ Failed to write corrupt backup: ${error.message}`);
      process.exitCode = 1;
      return;
    }

    const result = await attemptQuickBooksCompaniesRepair({ raw, parseError, backupPath });
    if (result && result.companies) {
      const count = Array.isArray(result.companies) ? result.companies.length : 0;
      console.log(
        `✅ Repaired ${filePath}. Truncated ${result.truncatedBytes} trailing byte${
          result.truncatedBytes === 1 ? '' : 's'
        }. Backup saved at ${backupPath}. Store now contains ${count} compan${count === 1 ? 'y' : 'ies'}.`
      );
      return;
    }

    console.error('❌ Repair failed. Restore the corrupt backup manually.');
    process.exitCode = 1;
  }
}

(async () => {
  try {
    const { command, fileOverride } = parseArguments(process.argv);
    if (command === 'help') {
      printUsage();
      return;
    }
    if (fileOverride) {
      process.env.QUICKBOOKS_COMPANIES_FILE = path.resolve(fileOverride);
    }

    // Ensure the server module picks up the override before it is required.
    delete require.cache[require.resolve('../src/server')];
    const {
      QUICKBOOKS_COMPANIES_FILE,
      readQuickBooksCompanies,
      attemptQuickBooksCompaniesRepair,
    } = require('../src/server');

    const targetPath = fileOverride ? path.resolve(fileOverride) : QUICKBOOKS_COMPANIES_FILE;

    if (command === 'inspect') {
      await inspectFile(targetPath, { readQuickBooksCompanies });
      return;
    }
    if (command === 'repair') {
      await repairFile(targetPath, { readQuickBooksCompanies, attemptQuickBooksCompaniesRepair });
      return;
    }

    printUsage();
    process.exitCode = 1;
   } catch (error) {
     printUsage();
     console.error(error.message || error);
     process.exitCode = 1;
   }
 })();
