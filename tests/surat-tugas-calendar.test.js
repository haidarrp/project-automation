'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'assets', 'js', 'surat-tugas.js'), 'utf8');
const now = new Date();
const year = now.getFullYear();
const month = now.getMonth() + 1;
const iso = (day) => `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
const before = new Date(year, month - 1, 0);
const beforeIso = `${before.getFullYear()}-${String(before.getMonth() + 1).padStart(2, '0')}-${String(before.getDate()).padStart(2, '0')}`;
const employee = (n, unit = 'Bidang A') => ({ id: `P${n}`, name: `Pegawai ${n}`, unit, nip: '', active: true });
const people = [...[1, 2].map((n) => employee(n)), ...[3, 4, 5].map((n) => employee(n, 'Bidang B')), employee(6)];
const records = [
  {
    id: 'ST-A', letterNumber: 'ST-001', startDate: iso(2), endDate: iso(4), activity: 'Rapat Koordinasi',
    assignmentType: 'official_travel', employees: people.slice(0, 5).map((p) => ({ employeeId: p.id }))
  },
  {
    id: 'ST-B', letterNumber: 'ST-002', startDate: iso(3), endDate: iso(3), activity: 'Pelatihan',
    assignmentType: 'work_arrangement', employees: [{ employeeId: 'P6' }]
  },
  {
    id: 'ST-C', letterNumber: 'ST-003', startDate: beforeIso, endDate: iso(2), activity: 'Lintas Bulan',
    assignmentType: 'other', employees: [{ employeeId: 'P6' }]
  }
];

async function mountCalendar() {
  const handlers = {};
  const app = { innerHTML: '', addEventListener(type, fn) { handlers[type] = fn; } };
  const context = {
    document: { getElementById(id) { return id === 'app' ? app : null; }, body: { classList: { toggle() {} } } },
    location: { hash: '#kalender', search: '' },
    history: { replaceState() {} },
    window: {
      SharePointStorage: {},
      AppShell: { render({ content = '', overlays = '' }) { return content + overlays; } },
      UserAccess: { async requireModule() { return true; }, canView() { return true; }, firstView() { return 'kalender'; } },
      MasterDataService: { async getBundle() { return { employees: people, directory: new Map(people.map((p) => [p.id, p])) }; } },
      FirebaseClient: {
        getDb() { return { collection() { return { async get() { return { docs: records.map((r) => ({ id: r.id, data() { return r; } })) }; } }; } }; },
        getCurrentUser() { return { uid: 'test' }; }
      },
      addEventListener() {}, requestAnimationFrame() {}
    },
    URLSearchParams,
    console,
    setTimeout,
    clearTimeout
  };
  vm.runInNewContext(source, context, { filename: 'surat-tugas.js' });
  await new Promise((resolve) => setImmediate(resolve));
  return { app, handlers };
}

function change(handlers, id, value) {
  handlers.change({ target: { id, value, matches() { return false; } } });
}
function click(handlers, action, dataset = {}) {
  const element = { dataset: { action, ...dataset }, classList: { contains() { return false; } } };
  handlers.click({ target: { closest(selector) { return selector === '[data-action]' ? element : null; } } });
}
function mode(handlers, value) {
  const element = { dataset: { assignmentCalendarMode: value } };
  handlers.click({ target: { closest(selector) { return selector === '[data-assignment-calendar-mode]' ? element : null; } } });
}

test('Kalender Surat Tugas renders 7-day grid and employees for every assigned day', async () => {
  const { app } = await mountCalendar();
  assert.match(app.innerHTML, /class="assignment-calendar-grid"/);
  assert.match(app.innerHTML, /<span>Sen<\/span><span>Sel<\/span>/);
  assert.match(app.innerHTML, /data-assignment-detail="ST-A"/);
  assert.match(app.innerHTML, /<strong>Pegawai 1<\/strong>/);
  assert.equal((app.innerHTML.match(/<strong>Pegawai 1<\/strong>/g) || []).length, 3);
  assert.match(app.innerHTML, /\+\d+ pegawai lainnya/);
  assert.match(app.innerHTML, /Lintas Bulan/);
});

test('More employees can be expanded and collapsed for one date', async () => {
  const { app, handlers } = await mountCalendar();
  assert.equal((app.innerHTML.match(/<strong>Pegawai 5<\/strong>/g) || []).length, 0);
  click(handlers, 'calendar-toggle-day', { calendarDay: iso(2) });
  assert.match(app.innerHTML, /<strong>Pegawai 5<\/strong>/);
  assert.match(app.innerHTML, /Tampilkan lebih sedikit/);
  click(handlers, 'calendar-toggle-day', { calendarDay: iso(2) });
  assert.equal((app.innerHTML.match(/<strong>Pegawai 5<\/strong>/g) || []).length, 0);
});

test('Unit filter limits the names displayed, not only the Surat Tugas records', async () => {
  const { app, handlers } = await mountCalendar();
  change(handlers, 'assignment-calendar-unit', 'Bidang A');
  assert.match(app.innerHTML, /<strong>Pegawai 1<\/strong>/);
  assert.doesNotMatch(app.innerHTML, /<strong>Pegawai 3<\/strong>/);
  assert.doesNotMatch(app.innerHTML, /<strong>Pegawai 4<\/strong>/);
});

test('Type filters, list toggle, month navigation and month input work', async () => {
  const { app, handlers } = await mountCalendar();
  change(handlers, 'assignment-calendar-type', 'work_arrangement');
  assert.match(app.innerHTML, /<strong>Pegawai 6<\/strong>/);
  assert.doesNotMatch(app.innerHTML, /data-assignment-detail="ST-A"/);
  mode(handlers, 'list');
  assert.match(app.innerHTML, /class="assignment-calendar-list"/);
  assert.match(app.innerHTML, /data-assignment-detail="ST-B"/);
  mode(handlers, 'calendar');
  assert.match(app.innerHTML, /class="assignment-calendar-grid"/);
  click(handlers, 'calendar-next');
  assert.match(app.innerHTML, /0 Surat Tugas pada bulan dan filter aktif/);
  change(handlers, 'assignment-calendar-month', month);
  assert.match(app.innerHTML, /1 Surat Tugas pada bulan dan filter aktif/);
});
