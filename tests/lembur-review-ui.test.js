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
  const injected=app.replace('  init();\n})();',`  window.__lemburTestHooks={state,renderReview,renderReviewEmployeeList,renderReviewEmployeeCalendar,renderReviewDayPanel,renderReviewDateCalendar,renderReviewTable,openInlineDocumentDetail,employeePendingDocuments,documentRelevantToOvertime,loadDocumentPreview};
})();`);
  assert.notEqual(injected,app,'sisipan pengujian harus berhasil');
  vm.runInNewContext(injected,sandbox,{filename:'app.js'});
  return {hooks:win.__lemburTestHooks,rules,win,doc};
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
 assert.match(html,/lembur-calendar\.css\?v=5/);
 assert.match(html,/lembur-document-integration\.js\?v=3/);
 assert.match(html,/assets\/js\/app\.js\?v=17/);
 assert.match(fs.readFileSync(path.join(root,'sw.js'),'utf8'),/pusdatin-static-v37/);
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

test('Metadata berkas Surat Tugas dan Cuti tampil dengan loader pratinjau otomatis',()=>{
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
 assert.match(html8,/data-lm-preview-auto="surat:st1\|0"/);
 assert.match(html8,/Memuat pratinjau berkas/);
 assert.doesNotMatch(html8,/Tampilkan Pratinjau/);
 assert.doesNotMatch(html8,/<iframe/);
 assert.doesNotMatch(html8,/Dokumen Cuti\.pdf/);
 assert.match(html12,/Dokumen Cuti\.pdf/);
 assert.doesNotMatch(html12,/Surat Tugas\.pdf/);
 assert.match(app,/window\.SharePointStorage\.getPreviewInfo/);
 assert.match(app,/querySelectorAll\('\[data-lm-preview-auto\]'\)/);
 assert.match(app,/keepPreviewScrollPosition\(iframe, x, y\)/);
 assert.match(css,/\.lm-support-file-preview/);
 assert.match(html,/sharepoint-storage\.js\?v=4/);
});

test('Tidak ada auto-scroll ke detail; dokumen tetap ditandai meskipun sudah disesuaikan',()=>{
 assert.doesNotMatch(app,/scrollIntoView\s*\(/);
 assert.match(app,/documentRelevantToOvertime\(record\)\) return 'needs-review'/);
 assert.match(app,/state\.reviewFilter === 'documents' && !documentRelevantToOvertime\(record\)/);
 assert.match(css,/\.lm-support-file-frame/);
 assert.match(css,/overflow-anchor:none/);
});


test('Dokumen pada hari tanpa jam lembur tidak diberi tanda; dokumen yang memblokir jam potensial tetap ditandai',()=>{
 const {hooks}=boot();const {employees,a,b}=sample();
 const noOvertime={date:new Date(2026,9,27,12),status:'tidak_hadir',inMinutes:null,outMinutes:null,
   originalOvertimeHours:0,overtimeHours:0,documentReviews:{}};
 a.records['2026-10-27']=noOvertime;
 b.records['2026-10-27']={...noOvertime,documentReviews:{}};
 docs.applyDocuments(employees,[],[{kind:'surat',id:'withoutHours',employeeIds:['101','202'],
   startDate:'2026-10-27',endDate:'2026-10-28',accountable:true,letterNumber:'ST/Tanpa Jam'}],
   {allows:()=>true},e=>e.nip);
 hooks.state.employees=employees; hooks.state.period={month:10,year:2026};
 assert.equal(hooks.documentRelevantToOvertime(a.records['2026-10-27']),false);
 assert.equal(hooks.documentRelevantToOvertime(a.records['2026-10-08']),false); // Dokumen lama dihapus saat sinkron ulang
 const day=hooks.renderReviewDayPanel(a,'2026-10-27');
 assert.doesNotMatch(day,/Penyesuaian Cuti \/ Surat Tugas/);
 assert.doesNotMatch(day,/Berkas Pendukung/);
 assert.doesNotMatch(day,/Dokumen terkait/);
 const cal=hooks.renderReviewEmployeeCalendar();
 const b27=cal.match(/<button[^>]+data-review-day="2026-10-27"[\s\S]*?<\/button>/)?.[0]||'';
 assert.match(b27,/0 jam/);
 assert.doesNotMatch(b27,/Ada dokumen|lm-needs-review/);
 const stats=hooks.renderReviewDateCalendar();
 const day27=stats.match(/<button[^>]+data-review-day="2026-10-27"[\s\S]*?<\/button>/)?.[0]||'';
 assert.doesNotMatch(day27,/ada dokumen|lm-needs-review/);
 assert.doesNotMatch(hooks.renderReviewTable(),/data-review-inspect="101\|2026-10-27"/);
 // Sumber tetap tersimpan untuk penghitungan/konsistensi, tetapi tidak ditandai.
 assert.equal(a.records['2026-10-27'].documentSources.length,1);
 a.records['2026-10-27'].originalOvertimeHours=2;
 assert.equal(hooks.documentRelevantToOvertime(a.records['2026-10-27']),true);
 assert.match(hooks.renderReviewDayPanel(a,'2026-10-27'),/Penyesuaian Cuti \/ Surat Tugas/);
 assert.match(hooks.renderReviewEmployeeCalendar(),/Ada dokumen/);
 assert.match(hooks.renderReviewTable(),/data-review-inspect="101\|2026-10-27"/);
});

test('Status 0 jam karena pembatasan dokumen tetap memunculkan panel untuk jam potensial',()=>{
 const {hooks}=boot();const {employees,a}=sample();hooks.state.employees=employees;
 const rec=a.records['2026-10-08'];
 assert.equal(rec.originalOvertimeHours,2);assert.equal(rec.overtimeHours,0);
 assert.equal(hooks.documentRelevantToOvertime(rec),true);
 assert.match(hooks.renderReviewDayPanel(a,'2026-10-08'),/Penyesuaian Cuti \/ Surat Tugas/);
});

test('Auto-preview memasang iframe tanpa fokus dan memulihkan posisi scroll setelah selesai dimuat',async()=>{
 const {hooks,win,doc}=boot();
 const {employees,a}=sample();
 docs.applyDocuments(employees,[],[{id:'st1',employeeIds:['101','202'],startDate:'2026-10-08',endDate:'2026-10-09',accountable:true,letterNumber:'253/ST',documents:[{file:{name:'Lampiran.pdf',driveId:'drive1',itemId:'item1'}}]}],{allows:()=>true},e=>e.nip);
 hooks.state.employees=employees;
 hooks.state.period={month:10,year:2026};
 hooks.state.reviewEmployeeKey='101';hooks.state.reviewDateKey='2026-10-08';
 const calls=[];
 win.scrollX=16;win.scrollY=350;
 win.scrollTo=(x,y)=>calls.push([x,y]);
 win.addEventListener=()=>{};win.removeEventListener=()=>{};
 win.setTimeout=()=>0;win.requestAnimationFrame=fn=>fn();
 win.SharePointStorage={getPreviewInfo:async()=>({getUrl:'https://tenant.sharepoint.com/preview'})};
 doc.createElement=(name)=>({tagName:name,attributes:{},isConnected:true,
   addEventListener(type,cb){this['on_'+type]=cb;},setAttribute(key,val){this.attributes[key]=val;}});
 const body={isConnected:true,children:[],replaceChildren(...els){this.children=els;},appendChild(el){this.children.push(el);}};
 const card={dataset:{lmFileKey:'surat:st1|0'},querySelector:()=>body};
 const placeholder={dataset:{lmPreviewAuto:'surat:st1|0'},closest:()=>card,isConnected:true};
 await hooks.loadDocumentPreview(placeholder);
 assert.equal(body.children.length,1);
 const iframe=body.children[0];
 assert.equal(iframe.src,'https://tenant.sharepoint.com/preview');
 assert.equal(iframe.tabIndex,-1);
 assert.equal(iframe.loading,'eager');
 assert.deepEqual(calls.at(-1),[16,350]);
 iframe.on_load();
 assert.deepEqual(calls.at(-1),[16,350]);
});
