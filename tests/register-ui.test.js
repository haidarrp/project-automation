const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const E = require('../assets/js/register-engine.js');

function snapshot(rows) {return {docs:rows.map(row => ({id:row.id, data:() => row}))};}
async function boot() {
  const requestedCollections = [];
  const app = {innerHTML:'', listeners:{}, addEventListener(name, handler) {this.listeners[name] = handler;}};
  const items = {
    leaveRecords:[{id:'c1', employeeId:'p1', startDate:'2026-10-12', endDate:'2026-10-14', leaveYear:2026, documentNumber:'CT/1', leaveType:'Cuti Tahunan'}],
    assignmentRecords:[{id:'st1', startDate:'2026-10-13', endDate:'2026-10-16', assignmentYear:2026, letterNumber:'ST/1', employees:[{employeeId:'p1'}]}]
  };
  const mockDb = {collection(name) {
    requestedCollections.push(name);
    if (!Object.hasOwn(items, name)) throw Error('Monitoring tidak boleh membaca '+name);
    return {where:() => ({get:async() => snapshot(items[name])})};
  }};
  const window = {
    AttendanceRegisterEngine:E,
    UserAccess:{canView:(module,view) => view === 'data', requireModule:async() => true},
    FirebaseClient:{getDb:() => mockDb},
    MasterDataService:{getBundle:async() => ({employees:[{id:'p1'}], directory:new Map([['p1',{name:'Pegawai Contoh'}]])})},
    AppShell:{render:({content}) => content}
  };
  const context = {
    window,document:{getElementById:() => app},location:{hash:'#cases'},Date,console,
    setTimeout,Map,Set,Promise,encodeURIComponent
  };
  vm.runInNewContext(fs.readFileSync('assets/js/register.js','utf8'), context);
  for(let attempt=0;attempt<30&&!app.innerHTML.includes('Pegawai Contoh');attempt++) await new Promise(resolve => setTimeout(resolve,5));
  return {app,requestedCollections};
}

test('Monitoring menampilkan irisan Cuti–Surat Tugas tanpa verifikasi atau rekonsiliasi',async() => {
  const {app} = await boot();
  assert.match(app.innerHTML,/Monitoring Konflik Kehadiran/);
  assert.match(app.innerHTML,/Pegawai Contoh/);
  assert.match(app.innerHTML,/Cuti dan Surat Tugas/);
  assert.match(app.innerHTML,/13 Okt 2026/);
  assert.match(app.innerHTML,/14 Okt 2026/);
  assert.doesNotMatch(app.innerHTML,/name="selectedSourceKeys"|Simpan — Sudah Verifikasi|rg-reconcile-form|Catat realisasi/);
  assert.doesNotMatch(app.innerHTML,/Data tidak dapat dimuat/);
});

test('Tiap dokumen beririsan memiliki tautan langsung ke detail di menu sumber',async() => {
  const {app} = await boot();
  assert.match(app.innerHTML,/cuti\.html\?record=c1&amp;year=2026#data/);
  assert.match(app.innerHTML,/surat-tugas\.html\?record=st1&amp;year=2026#data/);
  assert.match(app.innerHTML,/Buka Data Cuti/);
  assert.match(app.innerHTML,/Buka Data Surat Tugas/);
});

test('Monitoring hanya membaca dokumen sumber, tidak membaca keputusan kasus lama',async() => {
  const {requestedCollections} = await boot();
  assert.deepEqual([...new Set(requestedCollections)].sort(), ['assignmentRecords','leaveRecords']);
});

test('Filter jenis dan bulan memperbarui daftar tanpa membuat dokumen atau kasus',async() => {
  const {app} = await boot();
  app.listeners.change({target:{id:'rg-kind', value:'surat_ganda'}});
  assert.match(app.innerHTML,/Tidak ada konflik atau irisan yang sesuai/);
  app.listeners.change({target:{id:'rg-kind', value:'all'}});
  app.listeners.change({target:{id:'rg-month', value:'11'}});
  assert.match(app.innerHTML,/Tidak ada konflik atau irisan yang sesuai/);
  app.listeners.change({target:{id:'rg-month', value:'10'}});
  assert.match(app.innerHTML,/Pegawai Contoh/);
});
