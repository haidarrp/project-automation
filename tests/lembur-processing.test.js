'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../assets/js/app.js'), 'utf8');

function setup({ assignSatkerCodes = async () => [], apply = async () => ({}) } = {}) {
  const window = {
    APP_CONFIG: {INDONESIAN_MONTHS: ['Januari'], RULES: {}},
    BusinessRules: { assignSatkerCodes },
    LemburDocumentIntegration: { apply },
    ResourceLoader: { warm: () => {} },
    __uiHistory: [],
    addEventListener: () => {}
  };
  const document = {getElementById: () => ({}), addEventListener: () => {}};
  const sandbox = { window, document, console, Date, Set, Map, Promise, setTimeout, clearTimeout };
  const adapted = source
    .replace('const MASTER_LOOKUP_TIMEOUT_MS = 25000;', 'const MASTER_LOOKUP_TIMEOUT_MS = 35;')
    .replace('const DOCUMENT_LOOKUP_TIMEOUT_MS = 35000;', 'const DOCUMENT_LOOKUP_TIMEOUT_MS = 35;')
    .replace('  init();\n})();', `
      render = () => window.__uiHistory.push(renderProcessing());
      window.__processingTestHooks = {state, processData, renderProcessing, withTimeout, resetProcess};
    })();`);
  assert.notEqual(adapted, source, 'Hooks must replace init in unit tests');
  vm.runInNewContext(adapted, sandbox, {filename:'app.js'});
  const hooks = window.__processingTestHooks;
  hooks.state.files = [{name:'presensi-pegawai.xlsx'}];
  hooks.state.validation = {
    errors: [],
    employees: [{name: 'Pegawai Contoh', nip: '19990101', records: {'2026-10-01': {overtimeHours: 2}}}]
  };
  return {hooks, window};
}

test('Firestore permission-denied pada 38% menampilkan galat dan tombol ulang, bukan spinner permanen', async () => {
  let applyCalls = 0;
  const {hooks} = setup({
    assignSatkerCodes: async () => { throw Object.assign(new Error('Missing or insufficient permissions.'), {code:'permission-denied'}); },
    apply: async () => { applyCalls++; }
  });
  await hooks.processData();
  assert.equal(hooks.state.processing.percent, 38);
  assert.equal(hooks.state.processStep, 'processing');
  assert.equal(hooks.state.busy, false);
  assert.match(hooks.state.processing.error, /Firestore ditolak/);
  assert.match(hooks.renderProcessing(), /data-action="retry-process"/);
  assert.match(hooks.renderProcessing(), /Kembali ke Validasi/);
  assert.equal(applyCalls, 0, 'Mandatory leave/assignment verification must not run without master lookup');
});

test('Jika Master Pegawai tidak merespons, timeout mengembalikan kontrol pengguna', async () => {
  const {hooks} = setup({assignSatkerCodes: () => new Promise(() => {})});
  await hooks.processData();
  assert.equal(hooks.state.processing.percent, 38);
  assert.match(hooks.state.processing.error, /tidak selesai dalam/);
  assert.equal(hooks.state.busy, false);
});

test('Kesalahan sinkronisasi Cuti/Surat Tugas tidak diabaikan', async () => {
  const {hooks} = setup({
    assignSatkerCodes: async () => [],
    apply: async () => { throw new Error('Dokumen Cuti belum bisa diperiksa'); }
  });
  await hooks.processData();
  assert.equal(hooks.state.processing.percent, 92);
  assert.match(hooks.state.processing.error, /Dokumen Cuti belum bisa diperiksa/);
  assert.equal(hooks.state.processStep, 'processing');
  assert.equal(hooks.state.busy, false);
});

test('Coba Lagi melanjutkan dari data validasi tanpa upload ulang, serta menjaga data dari proses lama', async () => {
  let times = 0;
  const {hooks} = setup({
    assignSatkerCodes: async employees => {
      times++;
      if (times === 1) throw new Error('Koneksi terputus');
      employees[0].satkerCode = '123';
    },
    apply: async employees => {
      employees[0].records['2026-10-01'].documentChecked = true;
    }
  });
  const original = hooks.state.validation.employees[0];
  await hooks.processData();
  assert.equal(hooks.state.processing.percent, 38);
  assert.equal(hooks.state.busy, false);
  await hooks.processData();
  assert.equal(times, 2);
  assert.equal(hooks.state.processing.percent, 100);
  assert.equal(hooks.state.processStep, 'review');
  assert.equal(hooks.state.busy, false);
  assert.equal(hooks.state.employees[0].satkerCode, '123');
  assert.equal(hooks.state.employees[0].records['2026-10-01'].documentChecked, true);
  assert.equal(original.satkerCode, undefined, 'Original validation records should be untouched');
  assert.equal(original.records['2026-10-01'].documentChecked, undefined);
});

test('Reset membatalkan hasil proses tertunda yang datang setelah navigasi', async () => {
  let release;
  const {hooks} = setup({
    assignSatkerCodes: () => new Promise(resolve => {release = resolve;}),
    apply: async () => {}
  });
  const task = hooks.processData();
  await new Promise(resolve => setTimeout(resolve, 380));
  hooks.resetProcess();
  release();
  await task;
  assert.equal(hooks.state.processStep, 'period');
  assert.equal(hooks.state.busy, false);
  assert.equal(hooks.state.employees.length, 0);
});
