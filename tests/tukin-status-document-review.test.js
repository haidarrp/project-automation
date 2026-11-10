'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const scripts = name => fs.readFileSync(path.join(root, 'assets/js', name), 'utf8');
const day = '2026-11-04';
const period = {month: 12, year: 2026}; // 11 Oktober s.d. 10 November

function harness() {
  const collections = {leaveRecords: [], assignmentRecords: []};
  const db = {collection(name) {
    assert.ok(Object.hasOwn(collections, name), `Tidak boleh membaca ${name}`);
    return {where(field, op, year) {
      assert.equal(op, '==');
      return {get: async () => ({docs: collections[name]
        .filter(row => row[field] === year)
        .map(row => ({id: row.id, data: () => row}))})};
    }};
  }};
  const win = {FirebaseClient: {getDb: () => db}, addEventListener() {}};
  const ctx = vm.createContext({window: win, console, Date, Intl, Set, Map});
  vm.runInContext(scripts('tukin-config.js'), ctx);
  vm.runInContext(scripts('tukin-rules.js'), ctx);
  vm.runInContext(scripts('tukin-parser.js'), ctx);
  win.MasterDataService = {findByIdentity: async () => ({id: 'pegawai-1'})};
  win.AttendanceDocumentPolicy = {fetch: async () => ({stale: [],
    exportable: (type, row) => type !== 'surat' || row.accountable !== false,
    reason: () => ''})};
  vm.runInContext(scripts('tukin-supporting-data-integration.js'), ctx);
  const appCode = scripts('tukin-app.js');
  const testCode = appCode.replace('  init();\n})();',
    '  window.__testHooks = {state, calendarType, calendarDayCaption, renderCalendarDetail, renderCalendarTable, supportIntegrationAlert};\n})();');
  assert.notEqual(testCode, appCode, 'Injeksi hook UI harus berhasil');
  vm.runInContext(testCode, vm.createContext({...ctx, window: win, document: {getElementById: () => ({})}, console, Date, Intl, Set, Map}));
  const r = win.TukinRules;
  function make(status='Offsite', options={}) {
    const date = r.dateFromKey(options.date || day);
    const calc = r.calculateDay(date, 7*60+35, 18*60, status, {holidays: []});
    const rec = {date, key:r.dateKey(date), status, flags: [], inMinutes: 455, outMinutes: 1080,
      autoTotalPercent: calc.totalPercent, adjustedPercent: options.adjustedPercent ?? null,
      adjustmentSource: options.adjustmentSource || '', adjustmentNote: '', evidence: [],
      reason: calc.reason, isWeekend: r.isWeekend(date), isHoliday: calc.holiday,
      tlCategory: calc.tl.category, tlPercent: calc.tl.percent,
      pswCategory: calc.psw.category, pswPercent: calc.psw.percent};
    r.setStatusDocumentReview(rec, null);
    const emp = {name:'Pegawai Uji', nip:'12345678', masterEmployeeId: options.unmatched ? '' : 'pegawai-1',
      records: {[rec.key]:rec}, sourceFiles: []};
    return {rec, emp};
  }
  async function sync(employee) {return win.TukinSupportingDataIntegration.apply([employee], period);}
  return {win, r, collections, make, sync, ui:win.__testHooks};
}

const assignment = (extra={}) => ({id:'tugas1', assignmentYear:2026, startDate:day,
  endDate:day, employees:[{employeeId:'pegawai-1'}],
  assignmentType:'official_travel', accountable:true, letterNumber:'ST-01', ...extra});
const leave = (extra={}) => ({id:'cuti1', leaveYear:2026, startDate:day,
  endDate:day, employeeId:'pegawai-1', leaveType:'Cuti Tahunan', documentNumber:'C-01', ...extra});

test('Offsite dengan jam kerja normal dan potongan 0% tetap Perlu Cek tanpa dokumen', async () => {
  const h = harness(), {rec, emp} = h.make();
  assert.equal(rec.autoTotalPercent, 0);
  assert.equal(rec.needsVerification, true);
  const result = await h.sync(emp);
  assert.equal(rec.statusDocumentReview.state, 'missing');
  assert.equal(rec.needsVerification, true);
  assert.equal(rec.adjustedPercent, null);
  assert.equal(h.r.recordFinalPercent(rec), 0);
  assert.equal(h.r.summarizeEmployee(emp).flaggedRecords, 1);
  assert.equal(result.missingStatusDocumentRecords, 1);
  assert.equal(h.ui.calendarType(rec), 'review');
  assert.equal(h.ui.calendarDayCaption(rec), 'Perlu cek');
  const detail = h.ui.renderCalendarDetail(rec);
  assert.match(detail,/Perlu Cek — Dokumen Status Presensi/);
  assert.match(detail,/belum ditemukan Surat Tugas atau data Cuti/);
  assert.match(detail,/0,00%/);
  const table = h.ui.renderCalendarTable(emp,[{key: day, date: rec.date}]);
  assert.match(table,/Perlu cek/);
  assert.match(table,/belum ditemukan Surat Tugas atau data Cuti/);
  h.ui.state.supportIntegration = result;
  assert.match(h.ui.supportIntegrationAlert(),/1 tanggal status Offsite\/Tidak Hadir perlu cek/);
});

test('tidak_hadir juga Perlu Cek meski jam masuk dan pulang normal',async()=>{
  const h=harness(),{rec,emp}=h.make('tidak_hadir');
  const result=await h.sync(emp);
  assert.equal(result.missingStatusDocumentRecords,1);
  assert.equal(rec.needsVerification,true);
  assert.equal(rec.autoTotalPercent,0);
  assert.match(h.r.statusDocumentReviewMessage(rec),/Tidak Hadir/);
});

test('status umum WFO, WFH, Dinas, Libur tidak mendapat temuan dokumen status',async()=>{
  const h=harness();
  for (const status of ['WFO','WFH','Dinas','Libur']) {
    const {rec,emp}=h.make(status);
    await h.sync(emp);
    assert.equal(h.r.statusDocumentNeedsReview(rec),false,status);
    assert.equal(rec.needsVerification,false,status);
    assert.equal(h.ui.calendarType(rec),status==='Libur'?'holiday':'normal',status);
  }
});

test('Surat Tugas dipertanggungjawabkan pada tanggal & pegawai yang sama menghapus Perlu Cek',async()=>{
  const h=harness();h.collections.assignmentRecords.push(assignment());
  const {rec,emp}=h.make();const result=await h.sync(emp);
  assert.equal(result.assignmentLinkedRecords,1);
  assert.equal(result.missingStatusDocumentRecords,0);
  assert.equal(rec.statusDocumentReview.state,'supported');
  assert.equal(rec.needsVerification,false);
  assert.equal(rec.adjustedPercent,null);
  assert.equal(h.ui.calendarType(rec),'normal');
});

test('Cuti yang mencakup tanggal Offsite menghapus Perlu Cek tanpa mengubah TL/PSW 0%',async()=>{
  const h=harness();h.collections.leaveRecords.push(leave());
  const {rec,emp}=h.make('Offsite');await h.sync(emp);
  assert.equal(rec.statusDocumentReview.state,'supported');
  assert.equal(rec.needsVerification,false);
  assert.equal(rec.adjustedPercent,null);
  assert.equal(h.ui.calendarType(rec),'normal');
});

test('Surat Tugas tidak dipertanggungjawabkan tidak menghilangkan temuan',async()=>{
  const h=harness();h.collections.assignmentRecords.push(assignment({accountable:false}));
  const {rec,emp}=h.make();const result=await h.sync(emp);
  assert.equal(result.missingStatusDocumentRecords,1);
  assert.equal(rec.statusDocumentReview.state,'missing');
  assert.equal(h.ui.calendarType(rec),'review');
  assert.equal(rec.adjustedPercent,null);
});

test('Dokumen pegawai/tanggal lain tidak dianggap cocok',async()=>{
  const h=harness();
  h.collections.leaveRecords.push(leave({employeeId:'pegawai-lain'}));
  h.collections.assignmentRecords.push(assignment({startDate:'2026-11-05',endDate:'2026-11-05'}));
  const {rec,emp}=h.make();await h.sync(emp);
  assert.equal(rec.statusDocumentReview.state,'missing');
  assert.equal(rec.needsVerification,true);
});

test('Penghapusan Surat Tugas dan sinkronisasi ulang memunculkan kembali Perlu Cek',async()=>{
  const h=harness();h.collections.assignmentRecords.push(assignment());
  const {rec,emp}=h.make();await h.sync(emp);
  assert.equal(rec.needsVerification,false);
  h.collections.assignmentRecords.length=0;
  await h.sync(emp);
  assert.equal(rec.statusDocumentReview.state,'missing');
  assert.equal(rec.needsVerification,true);
  assert.equal(h.ui.calendarType(rec),'review');
});

test('Koreksi manual 0% tanpa dokumen tetap ditandai Perlu Cek',async()=>{
  const h=harness();const {rec,emp}=h.make('Offsite',{adjustedPercent:0,adjustmentSource:'manual'});
  await h.sync(emp);
  assert.equal(rec.adjustmentSource,'manual');
  assert.equal(rec.adjustedPercent,0);
  assert.equal(h.ui.calendarType(rec),'review');
  assert.equal(h.r.summarizeEmployee(emp).flaggedRecords,1);
});

test('Tukin: Sabtu dan Minggu Offsite/Tidak Hadir diabaikan, tanpa perlu Surat Tugas/Cuti',async()=>{
  for(const date of ['2026-11-07','2026-11-08']){
    for(const status of ['Offsite','tidak_hadir']){
      const h=harness();const {rec,emp}=h.make(status,{date,adjustedPercent:2.5});
      rec.flags=['TK'];
      h.r.recalculateVerification(rec);
      const result=await h.sync(emp);
      assert.equal(rec.isHoliday,true);
      assert.equal(rec.autoTotalPercent,0);
      assert.equal(rec.statusDocumentReview,null);
      assert.equal(rec.needsVerification,false);
      assert.equal(h.r.statusDocumentNeedsReview(rec),false);
      assert.equal(h.r.statusDocumentReviewMessage(rec),'');
      assert.equal(h.r.recordFinalPercent(rec),0);
      assert.equal(h.r.summarizeEmployee(emp).attendancePercent,0);
      assert.equal(result.missingStatusDocumentRecords,0);
      assert.equal(result.pendingStatusDocumentRecords,0);
      assert.equal(h.ui.calendarType(rec),'holiday');
    }
  }
});

test('Tukin: tanggal merah pada Senin bukan Sabtu/Minggu tetap mengikuti validasi dokumen',async()=>{
  const h=harness();const {rec,emp}=h.make('Offsite',{date:'2026-11-09'});
  h.win.TukinParser.recalculateEmployee(emp,{holidays:['2026-11-09']});
  await h.sync(emp);
  assert.equal(rec.isHoliday,true);
  assert.equal(rec.isWeekend,false);
  assert.equal(h.r.statusDocumentNeedsReview(rec),true);
  assert.equal(h.ui.calendarType(rec),'review');
});

test('Perubahan pengaturan kerja tidak menghapus flag dokumen',async()=>{
  const h=harness();const {rec,emp}=h.make();
  await h.sync(emp);
  h.win.TukinParser.recalculateEmployee(emp,{holidays:[],ramadanEnabled:false});
  assert.equal(rec.needsVerification,true);
  assert.equal(rec.autoTotalPercent,0);
});

test('Jika identitas belum cocok, status menunggu verifikasi, bukan dinyatakan punya dokumen',async()=>{
  const h=harness();
  h.win.MasterDataService.findByIdentity=async()=>null;
  const {rec,emp}=h.make('Offsite',{unmatched:true});
  const result=await h.sync(emp);
  assert.equal(result.pendingStatusDocumentRecords,1);
  assert.equal(rec.statusDocumentReview.state,'pending');
  assert.equal(rec.needsVerification,true);
});

test('Variasi penulisan status Offsite dan Tidak Hadir dikenali',()=>{
  const h=harness();
  for(const text of [' Offsite ','OFFSITE','off-site','tidak_hadir','Tidak Hadir','TIDAK-HADIR']) {
    assert.equal(h.r.requiresStatusDocument(text),true,text);
  }
});


test('Riwayat lama tanpa metadata review tetap menandai Offsite bila dokumen belum ada',()=>{
  const h=harness(); const {rec,emp}=h.make();
  delete rec.statusDocumentReview;
  rec.needsVerification=false; // Versi lama tidak memeriksa status
  h.win.TukinParser.recalculateEmployee(emp,{holidays:[]});
  assert.equal(rec.needsVerification,true);
  assert.equal(h.ui.calendarType(rec),'review');
});

test('Riwayat lama dengan metadata Cuti/Surat Tugas tidak ditandai sebagai dokumen hilang',()=>{
  const h=harness(); const {rec,emp}=h.make();
  delete rec.statusDocumentReview;
  rec.cutiAdjustment={automatic:true,leaveType:'Cuti Tahunan'};
  h.win.TukinParser.recalculateEmployee(emp,{holidays:[]});
  assert.equal(rec.needsVerification,false);
  assert.equal(h.ui.calendarType(rec),'normal');
});
