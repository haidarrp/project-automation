'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname,'..');
const app = fs.readFileSync(path.join(root,'assets/js/app.js'),'utf8');
const css = fs.readFileSync(path.join(root,'assets/css/lembur-calendar.css'),'utf8');
const html = fs.readFileSync(path.join(root,'lembur.html'),'utf8');
const docs = require('../assets/js/lembur-document-integration.js');

function boot(){
  const months=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const rules={
    employeeKey:e=>e.nip,
    dateFromKey:key=>new Date(Number(key.slice(0,4)),Number(key.slice(5,7))-1,Number(key.slice(8,10)),12),
    pad2:num=>String(num).padStart(2,'0'),
    formatIndonesianDate:d=>`${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`,
    isWeekend:d=>[0,6].includes(d.getDay()), isHoliday:()=>false,
    isWeekendLike:d=>[0,6].includes(d.getDay()),
    formatMinutes:m=>m==null?'-':`${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`,
    summarize:emps=>({employees:emps.length,overtimeEmployees:0,totalHours:0,mealDays:0}),
    normalizeHolidays:h=>h
  };
  const doc={getElementById:()=>({}),addEventListener:()=>{}};
  const win={APP_CONFIG:{INDONESIAN_MONTHS:months,RULES:{MEAL_ALLOWANCE_MIN_HOURS:2}},BusinessRules:rules,
    LemburDocumentIntegration:docs,addEventListener:()=>{},scrollTo:()=>{},scrollX:0,scrollY:0};
  const sandbox={window:win,document:doc,console,Date,Set,Map};
  const injected=app.replace('  init();\n})();',`  window.__lemburTestHooks={state,renderReview,renderReviewEmployeeList,renderReviewEmployeeCalendar,renderReviewDayPanel,renderReviewDateCalendar,renderReviewTable,openInlineDocumentDetail,employeePendingDocuments};
})();`);
  assert.notEqual(injected,app,'sisipan pengujian harus berhasil');
  vm.runInNewContext(injected,sandbox,{filename:'app.js'});
  return {hooks:win.__lemburTestHooks,rules};
}
function recordsFor(days){
  return Object.fromEntries(days.map(d=>[`2026-10-${String(d).padStart(2,'0')}`,{
    date:new Date(2026,9,d,12),inMinutes:450,outMinutes:1080,status:'wfo',
    originalOvertimeHours:2,overtimeHours:2,documentReviews:{}
  }]));
}
function sample(){
 const a={name:'Pegawai A',nip:'101',records:recordsFor([8,9,12,13])};
 const b={name:'Pegawai B',nip:'202',records:recordsFor([8,9])};
 const c={name:'Pegawai C',nip:'303',records:recordsFor([8])};
 const leaves=[{kind:'cuti',id:'cuti1',employeeId:'101',startDate:'2026-10-12',endDate:'2026-10-13',leaveType:'Alasan Penting'}];
 const assigns=[{kind:'surat',id:'st1',employeeIds:['101','202'],startDate:'2026-10-08',endDate:'2026-10-09',accountable:true,letterNumber:'253/ST'}];
 const employees=[a,b,c];
 docs.applyDocuments(employees,leaves,assigns,{allows:()=>true},e=>e.nip);
 return {employees,a,b,c};
}

test('Verifikasi dipasang langsung setelah status presensi; tidak ada panel atau tombol periode terpisah',()=>{
 assert.match(app,/function renderInlineDocumentReview\(record\)/);
 assert.match(app,/Status presensi:.*?\$\{renderInlineDocumentReview\(record\)\}/s);
 assert.doesNotMatch(app,/Buka Verifikasi Periode|data-action="open-period-document-review"|renderDocumentPeriodReview/);
 assert.doesNotMatch(app,/id="lm-verify-document"/);
 assert.match(app,/data-lm-document-save/);
});

test('Dokumen Surat Tugas dan Cuti ditampilkan pada tanggal bersangkutan tanpa memunculkan dokumen tanggal lain',()=>{
 const {hooks}=boot();
 const {employees,a}=sample();
 hooks.state.employees=employees;
 const day8=hooks.renderReviewDayPanel(a,'2026-10-08');
 const day12=hooks.renderReviewDayPanel(a,'2026-10-12');
 assert.match(day8,/Status presensi: wfo/);
 assert.match(day8,/253\/ST/);
 assert.doesNotMatch(day8,/Alasan Penting/);
 assert.match(day12,/Alasan Penting/);
 assert.doesNotMatch(day12,/253\/ST/);
 assert.match(day8,/2 tanggal, 2 pegawai \(4 tanggal-pegawai\)/);
 assert.match(day8,/value="deny" selected/);
 assert.match(day8,/value="allow"[^>]*disabled/);
});

test('Daftar seluruh pegawai membutuhkan verifikasi ditandai mandiri, tidak hanya terpilih',()=>{
 const {hooks}=boot(); const {employees}=sample();
 hooks.state.employees=employees;
 hooks.state.reviewEmployeeKey='101';
 const list=hooks.renderReviewEmployeeList();
 assert.match(list,/class="lm-person lm-person-pending selected" data-review-person="101"/);
 assert.match(list,/class="lm-person lm-person-pending " data-review-person="202"/);
 assert.match(list,/class="lm-person  " data-review-person="303"/);
 assert.match(list,/2 Perlu cek/);
 assert.match(list,/1 Perlu cek/);
 assert.match(css,/\.lm-person\.lm-person-pending/);
 assert.match(css,/\.lm-person-badge\.pending/);
});

test('Simpan keputusan satu Surat Tugas memperbarui tanggal dan pegawai lain, lalu hilangkan penanda jingga',()=>{
 const {hooks}=boot(); const {employees,a,b}=sample();
 hooks.state.employees=employees;
 hooks.state.reviewEmployeeKey='101';
 const group=docs.reviewGroups(employees).find(g=>g.kind==='surat');
 assert.equal(hooks.employeePendingDocuments(a),2);
 assert.equal(hooks.employeePendingDocuments(b),1);
 docs.verifySource(employees,group.key,group.signature,'deny',{displayName:'Operator'});
 assert.equal(hooks.employeePendingDocuments(a),1);
 assert.equal(hooks.employeePendingDocuments(b),0);
 assert.equal(a.records['2026-10-08'].documentReviewPending,false);
 assert.equal(b.records['2026-10-09'].documentReviewPending,false);
 assert.match(hooks.renderReviewDayPanel(a,'2026-10-08'),/Sudah diverifikasi/);
 assert.match(hooks.renderReviewEmployeeList(),/lm-person-verified[^\"]*" data-review-person="202"/);
 assert.match(hooks.renderReviewEmployeeCalendar(),/lm-doc-verified/);
});

test('Tabel dan tampilan per tanggal membuka detail tanggal terkait, bukan panel periode lain',()=>{
 const {hooks}=boot(); const {employees}=sample();
 hooks.state.employees=employees;
 hooks.state.period={month:10,year:2026};
 hooks.state.reviewDateKey='2026-10-08';
 assert.match(hooks.renderReviewDateCalendar(),/data-review-inspect="101\|2026-10-08"/);
 assert.match(hooks.renderReviewTable(),/data-review-inspect="101\|2026-10-08"/);
 assert.match(app,/openInlineDocumentDetail\(employeeKey,dateKey\)/);
});

test('Kalender memiliki warna berbeda untuk perlu verifikasi dan sudah diverifikasi',()=>{
 assert.match(css,/\.lm-cal-day\.lm-needs-review/);
 assert.match(css,/\.lm-cal-day\.lm-doc-verified/);
 assert.match(css,/\.lm-person\.selected/);
 assert.match(app,/record\.documentReviewPending/);
});

test('HTML mengandung cache-busting baru untuk CSS dan JavaScript serta SW dinaikkan',()=>{
 assert.match(html,/lembur-calendar\.css\?v=3/);
 assert.match(html,/lembur-document-integration\.js\?v=2/);
 assert.match(html,/assets\/js\/app\.js\?v=15/);
 assert.match(fs.readFileSync(path.join(root,'sw.js'),'utf8'),/pusdatin-static-v34/);
});

test('Dua dokumen beririsan tetap memiliki kontrol tersendiri dalam panel tanggal yang sama',()=>{
  const {hooks}=boot();
  const {employees,a}=sample();
  docs.applyDocuments(employees,
    [{kind:'cuti',id:'cuti2',employeeId:'101',startDate:'2026-10-08',endDate:'2026-10-08',leaveType:'Tahunan'}],
    [{kind:'surat',id:'st1',employeeIds:['101','202'],startDate:'2026-10-08',endDate:'2026-10-09',accountable:true,letterNumber:'253/ST'}],
    {allows:()=>true},e=>e.nip);
  hooks.state.employees=employees;
  const before=hooks.renderReviewDayPanel(a,'2026-10-08');
  assert.equal((before.match(/data-lm-document-save=/g)||[]).length,2);
  assert.match(before,/2 perlu verifikasi/);
  const st=docs.reviewGroups(employees).find(g=>g.kind==='surat');
  docs.verifySource(employees,st.key,st.signature,'deny');
  const after=hooks.renderReviewDayPanel(a,'2026-10-08');
  assert.match(after,/1 perlu verifikasi/);
  assert.match(after,/Sudah diverifikasi/);
  assert.match(after,/Perlu verifikasi/);
});
