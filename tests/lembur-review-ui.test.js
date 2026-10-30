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
  const sandbox={window:win,document:doc,console,Date,Set,Map,URL};
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

test('Daftar seluruh pegawai dengan dokumen terkait ditandai mandiri, tidak hanya terpilih',()=>{
 const {hooks}=boot(); const {employees}=sample();
 hooks.state.employees=employees;
 hooks.state.reviewEmployeeKey='101';
 const list=hooks.renderReviewEmployeeList();
 assert.match(list,/class="lm-person lm-person-pending selected" data-review-person="101"/);
 assert.match(list,/class="lm-person lm-person-pending " data-review-person="202"/);
 assert.match(list,/class="lm-person  " data-review-person="303"/);
 assert.match(list,/2 Dokumen/);
 assert.match(list,/1 Dokumen/);
 assert.match(css,/\.lm-person\.lm-person-pending/);
 assert.match(css,/\.lm-person-badge\.pending/);
});

test('Penyesuaian per dokumen tetap mengubah hasil seluruh tanggal/pegawai tetapi tanda informatif tidak hilang',()=>{
 const {hooks}=boot(); const {employees,a,b}=sample();
 hooks.state.employees=employees;
 const group=docs.reviewGroups(employees).find(g=>g.kind==='surat');
 assert.equal(hooks.employeePendingDocuments(a),2);
 assert.equal(hooks.employeePendingDocuments(b),1);
 docs.verifySource(employees,group.key,group.signature,'deny',{displayName:'Operator'});
 assert.equal(hooks.employeePendingDocuments(a),2);
 assert.equal(hooks.employeePendingDocuments(b),1);
 assert.equal(a.records['2026-10-08'].documentReviewPending,false);
 assert.equal(b.records['2026-10-09'].documentReviewPending,false);
 assert.match(hooks.renderReviewDayPanel(a,'2026-10-08'),/Sudah disesuaikan/);
 assert.match(hooks.renderReviewEmployeeList(),/lm-person-pending[^"]*" data-review-person="202"/);
 assert.match(hooks.renderReviewEmployeeCalendar(),/lm-needs-review/);
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

test('Kalender menggunakan warna dokumen secara informatif dan koreksi manual berbeda',()=>{
 assert.match(css,/\.lm-cal-day\.lm-needs-review/);
 assert.match(css,/\.lm-cal-day\.lm-doc-verified/);
 assert.match(css,/\.lm-person\.selected/);
 assert.match(app,/record\.documentReviewPending/);
});

test('HTML mengandung cache-busting baru untuk CSS dan JavaScript serta SW dinaikkan',()=>{
 assert.match(html,/lembur-calendar\.css\?v=4/);
 assert.match(html,/lembur-document-integration\.js\?v=3/);
 assert.match(html,/assets\/js\/app\.js\?v=16/);
 assert.match(fs.readFileSync(path.join(root,'sw.js'),'utf8'),/pusdatin-static-v35/);
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
  assert.match(before,/2 dokumen terdeteksi/);
  const st=docs.reviewGroups(employees).find(g=>g.kind==='surat');
  docs.verifySource(employees,st.key,st.signature,'deny');
  const after=hooks.renderReviewDayPanel(a,'2026-10-08');
  assert.match(after,/2 dokumen terdeteksi/);
  assert.match(after,/Sudah disesuaikan/);
  assert.match(after,/Rekomendasi otomatis/);
});

test('Penyesuaian opsional, lanjut dan ekspor tidak diblokir oleh dokumen belum disesuaikan',()=>{
 const {hooks}=boot();const {employees}=sample();
 hooks.state.employees=employees;
 hooks.state.period={month:10,year:2026};
 const out=hooks.renderReview();
 assert.match(out,/data-action="to-confirm">Konfirmasi & Lanjut/);
 assert.doesNotMatch(app,/if\(pendingDocumentReviews\(\)\)\s*(?:return|throw)/);
 assert.doesNotMatch(app,/state\.busy\|\|pendingDocumentReviews\(\)/);
 assert.match(app,/data-action="generate" \$\{state\.busy\?'disabled':''\}/);
});

test('Metadata berkas Surat Tugas dan Cuti tampil di bawah Penyesuaian, tanpa memuat iframe otomatis',()=>{
 const {hooks}=boot();const {employees,a}=sample();
 const st={id:'st1',employeeIds:['101','202'],startDate:'2026-10-08',endDate:'2026-10-09',accountable:true,letterNumber:'253/ST',
   documents:[{id:'f1',name:'Berkas Surat Tugas',type:'Surat Tugas',file:{name:'Surat Tugas.pdf',driveId:'drive1',itemId:'item1',webUrl:'https://tenant.sharepoint.com/docs/st.pdf'}}]};
 const leave={id:'cuti1',employeeId:'101',startDate:'2026-10-12',endDate:'2026-10-13',leaveType:'Alasan Penting',
   documents:[{id:'f2',file:{name:'Dokumen Cuti.pdf',driveId:'drive2',itemId:'item2',webUrl:'https://tenant.sharepoint.com/docs/cuti.pdf'}}]};
 docs.applyDocuments(employees,[leave],[st],{allows:()=>true},e=>e.nip);
 hooks.state.employees=employees;
 assert.equal(a.records['2026-10-08'].documentSources[0].documents[0].file.itemId,'item1');
 const html8=hooks.renderReviewDayPanel(a,'2026-10-08');
 const html12=hooks.renderReviewDayPanel(a,'2026-10-12');
 assert.match(html8,/Berkas Pendukung Cuti \/ Surat Tugas/);
 assert.match(html8,/Surat Tugas\.pdf/);
 assert.match(html8,/Buka SharePoint/);
 assert.match(html8,/data-lm-preview-load="surat:st1\|0"/);
 assert.doesNotMatch(html8,/<iframe/);
 assert.doesNotMatch(html8,/Dokumen Cuti\.pdf/);
 assert.match(html12,/Dokumen Cuti\.pdf/);
 assert.doesNotMatch(html12,/Surat Tugas\.pdf/);
 assert.match(app,/window\.SharePointStorage\.getPreviewInfo/);
 assert.match(html,/sharepoint-storage\.js\?v=4/);
});

test('Tidak ada auto-scroll ke detail; dokumen tetap ditandai meskipun sudah disesuaikan',()=>{
 assert.doesNotMatch(app,/scrollIntoView\s*\(/);
 assert.match(app,/record\?\.documentSources\?\.length\) return 'needs-review'/);
 assert.match(app,/state\.reviewFilter === 'documents' && !record\?\.documentSources\?\.length/);
 assert.match(css,/\.lm-support-file-frame/);
 assert.match(css,/overflow-anchor:none/);
});
