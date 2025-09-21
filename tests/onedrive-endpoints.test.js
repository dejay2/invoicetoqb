const assert = require('node:assert');
const fs = require('node:fs/promises');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');

const ENV_KEYS = [
  'ONEDRIVE_SETTINGS_FILE',
  'MS_GRAPH_CLIENT_ID',
  'MS_GRAPH_CLIENT_SECRET',
  'MS_GRAPH_TENANT_ID',
  'MS_GRAPH_SERVICE_USER_ID',
  'MS_GRAPH_SHAREPOINT_SITE_ID',
];

function createGraphStub(nativeFetch, Response) {
  const driveId = 'drive-123';
  const invoicesFolderId = 'folder-invoices';

  const drive = {
    id: driveId,
    name: 'Finance Drive',
    driveType: 'documentLibrary',
    webUrl: 'https://example.com/drives/drive-123',
    owner: {
      user: {
        displayName: 'Finance Ops',
      },
    },
  };

  const rootPath = `/drives/${driveId}/root:`;
  const invoicesFolder = {
    id: invoicesFolderId,
    name: 'Invoices',
    parentReference: {
      id: 'root',
      driveId,
      path: rootPath,
    },
    webUrl: 'https://example.com/drives/drive-123/root:/Invoices',
    folder: {
      childCount: 2,
    },
  };

  const rootItems = [
    invoicesFolder,
    {
      id: 'file-readme',
      name: 'Readme.txt',
      parentReference: {
        id: 'root',
        driveId,
        path: rootPath,
      },
      webUrl: 'https://example.com/drives/drive-123/root:/Readme.txt',
    },
  ];

  const invoicesChildren = [
    {
      id: 'sub-folder-2024',
      name: '2024',
      parentReference: {
        id: invoicesFolderId,
        driveId,
        path: `${rootPath}/Invoices`,
      },
      webUrl: 'https://example.com/drives/drive-123/root:/Invoices/2024',
      folder: {
        childCount: 0,
      },
    },
    {
      id: 'file-jan',
      name: 'jan.pdf',
      parentReference: {
        id: invoicesFolderId,
        driveId,
        path: `${rootPath}/Invoices`,
      },
      webUrl: 'https://example.com/drives/drive-123/root:/Invoices/jan.pdf',
    },
  ];

  const itemsByDrive = {
    [driveId]: {
      root: rootItems,
      [invoicesFolderId]: invoicesChildren,
    },
  };

  const itemsById = {
    [invoicesFolderId]: invoicesFolder,
  };

  const itemsByPath = {
    [`${driveId}:/`]: {
      id: 'root',
      parentReference: {
        id: null,
        driveId,
        path: rootPath,
      },
      name: 'root',
      folder: {},
    },
    [`${driveId}:/Invoices`]: invoicesFolder,
  };

  return async function graphAwareFetch(input, init = {}) {
    const urlString = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input?.url || '';
    if (typeof urlString !== 'string' || !urlString) {
      return nativeFetch(input, init);
    }

    if (urlString.startsWith('https://login.microsoftonline.com/') && urlString.includes('/oauth2/v2.0/token')) {
      return new Response(
        JSON.stringify({ access_token: 'stub-token', expires_in: 3600 }),
        {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }
      );
    }

    if (urlString.startsWith('https://graph.microsoft.com')) {
      const url = new URL(urlString);
      const strippedPath = url.pathname.replace(/^\/v1\.0/, '');

      const respondJson = (payload) =>
        new Response(JSON.stringify(payload), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });

      if (/^\/users\/.+\/drives$/.test(strippedPath)) {
        return respondJson({ value: [drive] });
      }

      const driveMatch = strippedPath.match(/^\/drives\/([^/]+)$/);
      if (driveMatch) {
        return respondJson(drive);
      }

      const rootChildrenMatch = strippedPath.match(/^\/drives\/([^/]+)\/root\/children$/);
      if (rootChildrenMatch) {
        const id = rootChildrenMatch[1];
        const entries = itemsByDrive[id]?.root || [];
        return respondJson({ value: entries });
      }

      const itemChildrenMatch = strippedPath.match(/^\/drives\/([^/]+)\/items\/([^/]+)\/children$/);
      if (itemChildrenMatch) {
        const id = itemChildrenMatch[1];
        const itemId = itemChildrenMatch[2];
        const entries = itemsByDrive[id]?.[itemId] || [];
        return respondJson({ value: entries });
      }

      const itemByIdMatch = strippedPath.match(/^\/drives\/([^/]+)\/items\/([^/]+)$/);
      if (itemByIdMatch) {
        const itemId = itemByIdMatch[2];
        if (itemsById[itemId]) {
          return respondJson(itemsById[itemId]);
        }
        return new Response('Not found', { status: 404 });
      }

      const childrenByPathMatch = strippedPath.match(/^\/drives\/([^/]+)\/root:(.*):\/children$/);
      if (childrenByPathMatch) {
        const id = childrenByPathMatch[1];
        const rawPath = childrenByPathMatch[2] || '';
        const normalised = `/${decodeURIComponent(rawPath).replace(/^\/+/, '')}`;
        const parent = itemsByPath[`${id}:${normalised}`];
        const entries = parent ? itemsByDrive[id]?.[parent.id] || [] : [];
        return respondJson({ value: entries });
      }

      const itemByPathMatch = strippedPath.match(/^\/drives\/([^/]+)\/root:(.*):$/);
      if (itemByPathMatch) {
        const id = itemByPathMatch[1];
        const rawPath = itemByPathMatch[2] || '';
        const normalised = `/${decodeURIComponent(rawPath).replace(/^\/+/, '')}`;
        const item = itemsByPath[`${id}:${normalised}`];
        if (item) {
          return respondJson(item);
        }
        return new Response('Not found', { status: 404 });
      }

      throw new Error(`Unexpected Microsoft Graph request: ${urlString}`);
    }

    return nativeFetch(input, init);
  };
}

(async () => {
  const nativeFetch = globalThis.fetch;
  const Response = globalThis.Response;

  if (typeof nativeFetch !== 'function' || typeof Response !== 'function') {
    throw new Error('Global fetch/Response are required for the OneDrive endpoint regression suite.');
  }

  const originalEnv = {};
  for (const key of ENV_KEYS) {
    originalEnv[key] = process.env[key];
  }

  let server;
  let settingsFile;
  try {
    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'onedrive-endpoints-'));
    settingsFile = path.join(tempDir, 'onedrive_settings.json');

    process.env.ONEDRIVE_SETTINGS_FILE = settingsFile;
    process.env.MS_GRAPH_CLIENT_ID = 'client-id';
    process.env.MS_GRAPH_CLIENT_SECRET = 'client-secret';
    process.env.MS_GRAPH_TENANT_ID = 'tenant-id';
    process.env.MS_GRAPH_SERVICE_USER_ID = 'service-user';
    process.env.MS_GRAPH_SHAREPOINT_SITE_ID = '';

    delete require.cache[require.resolve('../src/server')];
    globalThis.fetch = createGraphStub(nativeFetch, Response);

    const { app } = require('../src/server');
    server = http.createServer(app);
    await new Promise((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, '127.0.0.1', () => resolve());
    });

    const address = server.address();
    const port = typeof address === 'object' && address ? address.port : 0;
    const baseUrl = `http://127.0.0.1:${port}`;

    const requestJson = async (route, init = {}, expectedStatus = 200) => {
      const response = await fetch(`${baseUrl}${route}`, {
        ...init,
        headers: {
          Accept: 'application/json',
          ...(init.headers || {}),
        },
      });

      const raw = await response.text();
      let payload = null;
      if (raw) {
        try {
          payload = JSON.parse(raw);
        } catch (error) {
          throw new Error(`Failed to parse JSON response for ${route}: ${error.message}\n${raw}`);
        }
      }

      assert.strictEqual(
        response.status,
        expectedStatus,
        `Expected ${expectedStatus} response for ${route} but received ${response.status}: ${raw}`
      );

      return payload;
    };

    const initialSettings = await requestJson('/api/onedrive/settings');
    assert.ok(initialSettings);
    assert.ok(initialSettings.settings, 'settings payload missing');
    assert.strictEqual(initialSettings.settings.status, 'unconfigured');

    const drivesPayload = await requestJson('/api/onedrive/drives');
    assert.ok(Array.isArray(drivesPayload.drives), 'drives payload missing drives array');
    assert.strictEqual(drivesPayload.drives.length, 1, 'expected single drive from stub');
    assert.strictEqual(drivesPayload.drives[0].id, 'drive-123');
    assert.strictEqual(drivesPayload.drives[0].owner, 'Finance Ops');

    const rootChildren = await requestJson(`/api/onedrive/children?driveId=${encodeURIComponent('drive-123')}`);
    assert.ok(Array.isArray(rootChildren.items), 'root children payload missing items');
    assert.strictEqual(rootChildren.items.length, 2, 'expected two items in root listing');
    assert.strictEqual(rootChildren.items[0].id, 'folder-invoices');
    assert.strictEqual(rootChildren.items[0].isFolder, true);

    const pathChildren = await requestJson(
      `/api/onedrive/children?driveId=drive-123&path=${encodeURIComponent('/Invoices')}`
    );
    assert.ok(Array.isArray(pathChildren.items), 'path listing missing items array');
    assert.strictEqual(pathChildren.items.length, 2, 'expected two items inside invoices folder');
    assert.strictEqual(pathChildren.items[0].parentId, 'folder-invoices');

    const resolvedFolder = await requestJson('/api/onedrive/resolve', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ driveId: 'drive-123', folderId: 'folder-invoices' }),
    });
    assert.ok(resolvedFolder.item, 'resolve payload missing item');
    assert.strictEqual(resolvedFolder.item.id, 'folder-invoices');
    assert.strictEqual(resolvedFolder.item.isFolder, true);

    const updatedSettings = await requestJson('/api/onedrive/settings', {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ driveId: 'drive-123', folderId: 'folder-invoices' }),
    });
    assert.ok(updatedSettings.settings, 'PUT settings missing settings payload');
    assert.strictEqual(updatedSettings.settings.driveId, 'drive-123');
    assert.strictEqual(updatedSettings.settings.folderId, 'folder-invoices');
    assert.strictEqual(updatedSettings.settings.status, 'ready');
    assert.ok(updatedSettings.settings.lastValidatedAt, 'expected lastValidatedAt to be set');

    const persistedRaw = await fs.readFile(settingsFile, 'utf8');
    const persisted = JSON.parse(persistedRaw);
    assert.strictEqual(persisted.driveId, 'drive-123');
    assert.strictEqual(persisted.folderId, 'folder-invoices');

    console.log('onedrive endpoints regression passed');
  } finally {
    if (server) {
      await new Promise((resolve) => server.close(() => resolve()));
    }
    delete require.cache[require.resolve('../src/server')];
    globalThis.fetch = nativeFetch;

    for (const key of ENV_KEYS) {
      const value = originalEnv[key];
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
})();
