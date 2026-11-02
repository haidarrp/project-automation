const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { TextEncoder } = require('node:util');
const source = fs.readFileSync(require('node:path').join(__dirname, '../assets/js/sharepoint-storage.js'), 'utf8');

function setup(fetchImpl) {
  const calls = [];
  let account = 'haidar@example.org';
  const config = {
    shareUrl: ' https://example.sharepoint.com/:f:/g/example?e=test  ',
    rootFolderName: 'TUKIN',
    graphBaseUrl: 'https://graph.microsoft.com/v1.0'
  };
  const window = {
    MICROSOFT_CONFIG: config,
    MicrosoftAuth: {
      isConfigured: () => true,
      initialize: async () => {},
      connect: async () => {},
      acquireToken: async () => 'test-token',
      getStatus: () => ({ connected: true, email: account })
    }
  };
  const ctx = {
    window, TextEncoder, btoa, Headers, Blob,
    fetch: async (url, init) => {
      calls.push({url, init});
      return fetchImpl(url, init);
    }
  };
  vm.runInNewContext(source, ctx);
  return { storage: window.SharePointStorage, calls, setAccount: value => { account = value; } };
}

const json = (object, status = 200) => new Response(JSON.stringify(object), {
  status,
  headers: { 'Content-Type': 'application/json' }
});

test('404 from the old sharing link is an actionable error, not a successful connection', async () => {
  const { storage, calls } = setup(async () => json({
    error: { code: 'itemNotFound', message: 'Requested sharing link could not be found' }
  }, 404));
  await assert.rejects(storage.ensureReady(true), /sharing link folder tidak ditemukan/);
  assert.match(calls[0].url, /\/shares\/u!/);
  assert.equal(calls.length, 1);
});

test('valid link and writable root folder permit SharePoint ready state', async () => {
  const { storage, calls } = setup(async (url) => {
    if (url.includes('/shares/')) return json({
      id: 'shared-folder-id', name: 'Project automation', folder: {},
      webUrl: 'https://example.sharepoint.com/shared', parentReference: { driveId: 'drive-id' }
    });
    if (url.includes('/children')) return json({ value: [{
      id: 'tukin-folder-id', name: 'TUKIN', folder: {}
    }] });
    throw Error('Unexpected Graph call ' + url);
  });
  assert.equal(await storage.ensureReady(true), true);
  assert.equal(calls.length, 2);
  const folder = await storage.resolveAppRoot(false);
  assert.equal(folder.itemId, 'tukin-folder-id');
  assert.equal(calls.length, 2, 'cache must avoid unnecessary requests');
});

test('account switch invalidates cached remote folder', async () => {
  const { storage, calls, setAccount } = setup(async (url) => {
    if (url.includes('/shares/')) return json({
      id: 'shared-folder-id', name: 'Project automation', folder: {},
      webUrl: 'https://example.sharepoint.com/shared', parentReference: { driveId: 'drive-id' }
    });
    if (url.includes('/children')) return json({ value: [{
      id: 'tukin-folder-id', name: 'TUKIN', folder: {}
    }] });
    throw Error('Unexpected Graph call ' + url);
  });
  await storage.ensureReady(true);
  assert.equal(calls.length, 2);
  setAccount('ayuni@example.org');
  await storage.ensureReady(true);
  assert.equal(calls.length, 4);
});
