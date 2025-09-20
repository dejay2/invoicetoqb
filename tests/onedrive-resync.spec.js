const { test, expect } = require('@playwright/test');
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const PUBLIC_DIR = path.join(PROJECT_ROOT, 'public');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
};

test.describe('OneDrive full resync', () => {
  /** @type {http.Server} */
  let server;
  /** @type {number} */
  let port;

  let oneDriveState;

  test.beforeAll(async () => {
    oneDriveState = {
      enabled: true,
      status: 'connected',
      driveId: 'drive-id',
      folderId: 'folder-id',
      folderName: 'Incoming invoices',
      folderPath: '/Invoices',
      webUrl: 'https://example.com/invoices',
      lastSyncStatus: 'partial',
      lastSyncError: {
        message:
          "Resync required. Replace any local items with the server's version (including deletes) if you're sure that the service was up to date with your local changes when you last sync'd. Upload any local changes that the server doesn't know about.",
      },
      lastSyncReason: 'manual',
      lastSyncMetrics: null,
      lastSyncAt: '2024-10-01T12:00:00.000Z',
    };

    server = http.createServer((req, res) => {
      const url = new URL(req.url || '/', `http://127.0.0.1:${port || 0}`);

      if (req.method === 'GET' && url.pathname === '/api/quickbooks/companies') {
        const payload = {
          companies: [
            {
              realmId: 'realm-1',
              companyName: 'Acme Corp',
              oneDrive: oneDriveState,
            },
          ],
        };
        res.writeHead(200, {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-store',
        });
        res.end(JSON.stringify(payload));
        return;
      }

      if (req.method === 'GET' && url.pathname === '/api/quickbooks/companies/realm-1/metadata') {
        const payload = {
          metadata: {
            vendors: { items: [] },
            accounts: { items: [] },
            taxCodes: { items: [] },
            vendorSettings: {},
          },
        };
        res.writeHead(200, {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-store',
        });
        res.end(JSON.stringify(payload));
        return;
      }

      if (req.method === 'GET' && url.pathname === '/api/invoices') {
        res.writeHead(200, {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-store',
        });
        res.end(JSON.stringify({ invoices: [] }));
        return;
      }

      if (req.method === 'POST' && url.pathname === '/api/quickbooks/companies/realm-1/onedrive/sync') {
        let raw = '';
        req.on('data', (chunk) => {
          raw += chunk;
        });
        req.on('end', () => {
          let payload;
          try {
            payload = JSON.parse(raw || '{}');
          } catch (error) {
            payload = {};
          }

          if (payload?.forceFull) {
            oneDriveState = {
              ...oneDriveState,
              lastSyncStatus: null,
              lastSyncError: null,
              lastSyncMetrics: null,
              lastSyncReason: 'full-resync',
            };
          }

          res.writeHead(202, {
            'Content-Type': 'application/json',
            'Cache-Control': 'no-store',
          });
          res.end(JSON.stringify({ accepted: true, oneDrive: oneDriveState }));
        });
        return;
      }

      const relativePath = url.pathname === '/' ? 'index.html' : url.pathname.slice(1);
      const safePath = path.normalize(relativePath).replace(/^(\.\.(?:[\\/]|$))+/, '');
      const filePath = path.join(PUBLIC_DIR, safePath);

      if (!filePath.startsWith(PUBLIC_DIR)) {
        res.writeHead(403);
        res.end('Forbidden');
        return;
      }

      fs.readFile(filePath, (error, content) => {
        if (error) {
          res.writeHead(404);
          res.end('Not Found');
          return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const type = MIME_TYPES[ext] || 'application/octet-stream';
        res.writeHead(200, {
          'Content-Type': type,
          'Cache-Control': 'no-store',
        });
        res.end(content);
      });
    });

    await new Promise((resolve) => {
      server.listen(0, '127.0.0.1', () => {
        const address = server.address();
        port = address && typeof address === 'object' ? address.port : 0;
        resolve();
      });
    });
  });

  test.afterAll(async () => {
    await new Promise((resolve, reject) => {
      server.close((error) => {
        if (error) {
          reject(error);
        } else {
          resolve();
        }
      });
    });
  });

  test('clears stale OneDrive status after full resync', async ({ page }) => {
    await page.goto(`http://127.0.0.1:${port}/`);

    const statusLocator = page.locator('#onedrive-status-result');
    await expect(statusLocator).toContainText('Resync required');

    await page.getByRole('button', { name: 'Full resync' }).click();

    await expect(statusLocator).toHaveText('No syncs yet.');
  });
});
