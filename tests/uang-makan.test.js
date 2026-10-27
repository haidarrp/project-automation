'use strict';
const assert=require('node:assert/strict');
const {test}=require('node:test');
const vm=require('node:vm');
const fs=require('node:fs');
const path=require('node:path');
const scripts=path.resolve(__dirname,'../assets/js');
function sandbox(code,overrides={}){
 const ctx={window:{},Date,console,URL,URLSearchParams,...overrides};
 vm.runInNewContext(code,ctx);
 return ctx;
}
const ruleCode=fs.readFileSync(path.join(scripts,'uang-makan-rules.js'),'utf8');
function appHarness(R){
 const code=fs.readFileSync(path.join(scripts,'uang-makan-app.js'),'utf8')
  .replace(/\ninit\(\);\s*\}\)\(\);\s*$/,`\nwindow.__umTest={state,previewPanel,dailySupporting,supportingFilesSection,supportingRecordsSection,supportDocuments,documentKey,modal,loadSupportingPreviews,getSupportingData};\n})();`);
 const root={addEventListener:()=>{},querySelector:()=>null};
 const ctx=sandbox(code,{
  window:{UangMakanRules:R,addEventListener:()=>{},UserAccess:{can:()=>true}},
  document:{getElementById:()=>root},
  localStorage:{getItem:()=>null},
  location:{hash:''}
 });
 return Object.assign(ctx.window.__umTest,{client:ctx.window});
} 
function rules(){return sandbox(ruleCode).window.UangMakanRules;}
function attendance(key,extras={}){
 return {key,inMinutes:8*60,outMinutes:17*60,status:'Hadir',flags:[],...extras};
}

test('WFA/WFH Senin-Kamis otomatis 0 meskipun terdapat presensi',()=>{
 const R=rules();
 for(const status of ['WFH','WFA','WORK FROM HOME','work from anywhere','Bekerja dari rumah']){
  const d=R.classifyDay(attendance('2026-09-14',{status}));
  assert.equal(d.automatic,0,status);
  assert.equal(d.value,0,status);
  assert.equal(d.remoteExcluded,true,status);
  assert.equal(d.issues.length,0,status);
 }
});
test('WFH/WFA hari Jumat boleh 1 bila presensi tersedia',()=>{
 const R=rules();
 for(const status of ['WFH','WFA','Work From Home']){
  const d=R.classifyDay(attendance('2026-09-18',{status}));
  assert.equal(d.automatic,1,status);
  assert.equal(d.value,1,status);
  assert.equal(d.remoteExcluded,false,status);
  assert.deepEqual(Array.from(d.issues),[]);
 }
 assert.equal(R.classifyDay({key:'2026-09-18',status:'WFH'}).automatic,0);
});
test('Koreksi manual lama tidak bisa mengalahkan aturan remote atau libur',()=>{
 const R=rules();
 const remote=R.classifyDay(attendance('2026-09-16',{status:'WFA',manualValue:1}));
 const weekend=R.classifyDay(attendance('2026-09-19',{manualValue:1}));
 const holiday=R.classifyDay(attendance('2026-09-17',{isHoliday:true,manualValue:1}));
 for(const d of [remote,weekend,holiday]){
  assert.equal(d.value,0);assert.equal(d.manualValue,null);
 }
 assert.equal(R.classifyDay(attendance('2026-09-17')).automatic,1);
});
test('Konflik dinas atau cuti masih membutuhkan verifikasi, berapapun hasil WFH',()=>{
 const R=rules();
 const leaves=[{id:'cuti1',employeeId:'u1'}];
 const official=[{id:'dinas1',assignmentType:'official_travel'}];
 const d=R.classifyDay(attendance('2026-09-14',{status:'WFH'}),leaves,official);
 assert.equal(d.automatic,0);
 assert.ok(d.issues.includes('Presensi bertentangan dengan data cuti'));
 assert.ok(d.issues.includes('Presensi bertentangan dengan dinas luar'));
 const e={name:'Tes',jenisPegawai:'PNS',days:{'2026-09-14':d}};
 assert.equal(R.overview(e).status,'Perlu verifikasi');
 assert.equal(R.overview(e).present,0);
 assert.equal(R.overview(e).days,0);
});
test('Pembaca Excel mendeteksi kolom WFH dengan tanda 1 dan isi status WFA',async()=>{
 const R=rules();
 const header=['No','Hari','Tanggal','Masuk','Keluar','Keterangan','WFH'];
 const row=(date,status,mode='')=>[1,'Senin',date,'08:00','17:00',status,mode];
 const raw=[[],[],header,row('2026-09-14','Hadir',1),row('2026-09-15','WFA',0),row('2026-09-16','Hadir','tidak'),row('2026-09-18','WFH',0)];
 const wb={SheetNames:['Sheet1'],Sheets:{Sheet1:{A2:{w:'Pegawai Tes - 198901012010011001'},_raw:raw,_display:raw}}};
 const ctx=sandbox(fs.readFileSync(path.join(scripts,'uang-makan-parser.js'),'utf8'),{
  window:{UangMakanRules:R,TukinParser:{
   parseDate:(raw)=>{const d=new Date(String(raw)+'T12:00:00');return Number.isNaN(d.getTime())?null:d;},
   parseTimeMinutes:(raw)=>{const [h,m]=String(raw).split(':').map(Number);return Number.isNaN(h)?null:h*60+m;}
  },XLSX:{read:()=>wb,utils:{sheet_to_json:(sheet,config)=>config.raw?sheet._raw:sheet._display}}},
  XLSX:{read:()=>wb,utils:{sheet_to_json:(sheet,config)=>config.raw?sheet._raw:sheet._display}}
 });
 const file={name:'laporan_presensi_Tes.xlsx',arrayBuffer:async()=>new ArrayBuffer(0)};
 const parsed=await ctx.window.UangMakanParser.parseFile(file,2026,9);
 assert.equal(parsed.rows['2026-09-14'].workMode,'WFH');
 assert.equal(parsed.rows['2026-09-15'].workMode,'WFA');
 assert.equal(parsed.rows['2026-09-16'].workMode,'');
 assert.equal(parsed.rows['2026-09-18'].workMode,'WFH');
 for(const [date,expected] of [['2026-09-14',0],['2026-09-15',0],['2026-09-16',1],['2026-09-18',1]]){
  assert.equal(R.classifyDay(parsed.rows[date]).automatic,expected,date);
 }
});

test('Tabel pratinjau menampilkan semua pegawai dan iframe berkas pendukung',()=>{
 const R=rules();
 const T=appHarness(R);
 assert.ok(T);
 T.state.year=2026;T.state.month=9;
 T.state.employees=Array.from({length:47},(_,i)=>({id:`id-${i}`,name:`Pegawai ${i+1}`,nip:`123${i}`,jenisPegawai:'PNS',days:{}}));
 const html=T.previewPanel();
 assert.ok(html.includes('Menampilkan seluruh 47 pegawai'));
 assert.ok(html.includes('Pegawai 47'));
 assert.equal((html.match(/<td class="name">/g)||[]).length,47);
 const docs=T.supportDocuments([{id:'doc1',name:'Surat Perintah',type:'Surat Tugas',file:{name:'Surat.pdf',driveId:'drive',itemId:'item',webUrl:'https://example.sharepoint.com/surat.pdf'}}]);
 const group={kind:'tugas',id:'a1',documents:docs};
 const key=T.documentKey(group,docs[0]);
 T.state.documentPreviews[key]={loading:false,getUrl:'https://example.sharepoint.com/preview.pdf'};
 const conflict=R.classifyDay(attendance('2026-09-16'),[],[{id:'a1',assignmentType:'official_travel',activity:'Rapat',documents:docs}]);
 const filesHtml=T.supportingFilesSection(conflict);
 assert.ok(filesHtml.includes('Berkas Pendukung Cuti / Surat Tugas'));
 assert.ok(filesHtml.includes('Surat.pdf'));
 assert.ok(filesHtml.includes('um-document-frame'));
 assert.ok(filesHtml.includes('Buka SharePoint'));
 assert.ok(filesHtml.includes('preview.pdf'));
});


test('Data cuti tanpa konflik tetap muncul pada kalender, detail, dan pratinjau dokumen',()=>{
 const R=rules(),T=appHarness(R);
 const docs=T.supportDocuments([{id:'doc-cuti',type:'Surat Cuti',name:'Persetujuan Cuti',number:'ABC/2026',
  file:{name:'Persetujuan.pdf',driveId:'drive',itemId:'item',webUrl:'https://example.sharepoint.com/cuti.pdf'}}]);
 const leave={id:'cuti1',employeeId:'e1',startDate:'2026-09-01',endDate:'2026-09-02',leaveType:'Cuti Tahunan',
  documentNumber:'SK-CUTI/2026',documentDate:'2026-08-29',note:'Cuti yang telah disetujui',documents:docs};
 const e=R.applySupporting({id:'e1',name:'Pegawai Contoh',nip:'123',jenisPegawai:'PNS',days:{'2026-09-01':{key:'2026-09-01',inMinutes:null,outMinutes:null,status:'Cuti',flags:['C']}}},[leave],[],2026,9);
 T.state.year=2026;T.state.month=9;T.state.employees=[e];T.state.modalId='e1';T.state.modalDate='2026-09-01';
 const day=e.days['2026-09-01'];
 assert.equal(day.issues.length,0);
 assert.equal(day.value,0);
 assert.equal(T.dailySupporting(day).length,1);
 const key=T.documentKey(T.dailySupporting(day)[0],docs[0]);
 T.state.documentPreviews[key]={loading:false,getUrl:'https://example.sharepoint.com/preview-cuti.pdf'};
 const html=T.modal();
 assert.ok(html.includes('um-day-source cuti')); // Sel 1 Sep: tidak ada konflik
 assert.ok(html.includes('2 tanggal cuti'));
 assert.ok(html.includes('Tidak ditemukan konflik otomatis. Data cuti/surat tugas terkait tetap ditampilkan'));
 assert.ok(html.includes('Data Cuti &amp; Surat Tugas Terintegrasi'));
 assert.ok(html.includes('Cuti Tahunan'));
 assert.ok(html.includes('SK-CUTI/2026'));
 assert.ok(html.includes('Cuti yang telah disetujui'));
 assert.ok(html.includes('Persetujuan.pdf'));
 assert.ok(html.includes('preview-cuti.pdf'));
 assert.ok(html.includes('Buka SharePoint'));
});

test('Surat tugas tanpa konflik, termasuk pengaturan kerja, terlihat meskipun tidak ada presensi',()=>{
 const R=rules(),T=appHarness(R);
 const task={id:'st1',employeeIds:['e1'],startDate:'2026-09-03',endDate:'2026-09-04',
  assignmentType:'work_arrangement',activity:'Pengaturan kerja pelayanan',letterNumber:'ST/123/2026',
  letterDate:'2026-09-02',location:'Kantor Pusdatin',note:'Tanpa presensi fisik',
  documents:T.supportDocuments([{id:'st-doc',name:'Surat Tugas',file:{name:'ST-123.pdf',driveId:'d',itemId:'i'}}])};
 const e=R.applySupporting({id:'e1',name:'Pegawai Contoh',nip:'123',jenisPegawai:'PNS',days:{}},[],[task],2026,9);
 const day=e.days['2026-09-03'];
 assert.equal(day.issues.length,0);
 assert.equal(day.value,0);
 assert.equal(T.dailySupporting(day).length,1);
 const html=T.supportingRecordsSection(day);
 assert.ok(html.includes('Pengaturan Kerja'));
 assert.ok(html.includes('Pengaturan kerja pelayanan'));
 assert.ok(html.includes('ST/123/2026'));
 assert.ok(html.includes('Kantor Pusdatin'));
 assert.ok(T.supportingFilesSection(day).includes('ST-123.pdf'));
 T.state.year=2026;T.state.month=9;T.state.employees=[e];T.state.modalId='e1';T.state.modalDate='2026-09-03';
 assert.ok(T.modal().includes('um-day-source tugas'));
 assert.ok(T.modal().includes('2 tanggal surat tugas'));
});

test('Dokumen tetap terbaca saat tanggal libur, tanpa benturan, serta kosong jika tanpa catatan',()=>{
 const R=rules(),T=appHarness(R);
 const leave={id:'cuti1',employeeId:'e1',startDate:'2026-09-19',endDate:'2026-09-20',leaveType:'Cuti Tahunan',documents:[]};
 const task={id:'st1',employeeIds:['e1'],startDate:'2026-09-19',endDate:'2026-09-20',assignmentType:'official_travel',activity:'Tugas khusus',documents:[]};
 const e=R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{}},[leave],[task],2026,9);
 const holiday=e.days['2026-09-19'];
 assert.equal(holiday.nonWorking,true);
 assert.equal(holiday.issues.length,0);
 assert.equal(T.dailySupporting(holiday).length,2);
 assert.ok(T.supportingRecordsSection(holiday).includes('Tugas khusus'));
 assert.ok(T.supportingFilesSection(holiday).includes('Belum terdapat berkas digital'));
 const empty=e.days['2026-09-21'];
 assert.equal(T.dailySupporting(empty).length,0);
 assert.equal(T.supportingFilesSection(empty),'');
 assert.ok(T.supportingRecordsSection(empty).includes('Tidak ada data cuti atau surat tugas'));
 T.state.year=2026;T.state.month=9;T.state.employees=[e];T.state.modalId='e1';T.state.modalDate='2026-09-19';
 const modal=T.modal();
 assert.ok(modal.includes('Cuti'));
 assert.ok(modal.includes('Surat Tugas'));
 assert.ok(modal.includes('Sabtu/Minggu'));
});

test('Riwayat menyimpan metadata dokumen cuti dan surat tugas tanpa URL pratinjau sementara',()=>{
 const storageCode=fs.readFileSync(path.join(scripts,'uang-makan-storage.js'),'utf8').replace(
  'window.UangMakanStorage=Object.freeze({list,save,get,remove});',
  'window.__umStorageTest={sanitizeEmployee};window.UangMakanStorage=Object.freeze({list,save,get,remove});');
 const storage=sandbox(storageCode).window.__umStorageTest;
 const doc={id:'d1',name:'Surat',type:'PDF',number:'N-1',file:{name:'Surat.pdf',driveId:'drive',itemId:'item',
  webUrl:'https://example.sharepoint.com/doc.pdf',getUrl:'https://temporary.example.com/preview'}};
 const row=storage.sanitizeEmployee({id:'e1',name:'Tes',nip:'123',days:{'2026-09-01':{
  leaves:[{id:'l1',leaveType:'Cuti Tahunan',documentNumber:'C-1',documentDate:'2026-08-30',note:'Disetujui',documents:[doc]}],
  assignments:[{id:'s1',assignmentType:'official_travel',activity:'Rapat',letterNumber:'ST-1',letterDate:'2026-08-30',
   location:'Jakarta',documents:[doc]}]
 }}});
 const day=row.days['2026-09-01'];
 assert.equal(day.leaves[0].documents[0].file.itemId,'item');
 assert.equal(day.assignments[0].documents[0].file.webUrl,'https://example.sharepoint.com/doc.pdf');
 assert.equal(day.assignments[0].location,'Jakarta');
 assert.equal(day.leaves[0].documentDate,'2026-08-30');
 assert.equal(day.leaves[0].documents[0].file.getUrl,undefined);
});

test('Metadata sumber di-escape sebelum dirender ke HTML',()=>{
 const R=rules(),T=appHarness(R);
 const leave={id:'l',leaveType:'<img src=x onerror=alert(1)>',documentNumber:'<script>alert(1)</script>',documents:[]};
 const day=R.classifyDay({key:'2026-09-01'},[leave],[]);
 const html=T.supportingRecordsSection(day);
 assert.ok(html.includes('&lt;img'));
 assert.ok(html.includes('&lt;script&gt;'));
 assert.ok(!html.includes('<script>'));
});


test('Pratinjau Microsoft Graph diminta untuk surat tugas tanpa konflik',async()=>{
 const R=rules(),T=appHarness(R);
 const docs=T.supportDocuments([{id:'doc1',name:'Surat Tugas',file:{name:'Penugasan.pdf',driveId:'drive1',itemId:'item1'}}]);
 const task={id:'st1',employeeIds:['e1'],startDate:'2026-09-03',endDate:'2026-09-03',assignmentType:'official_travel',documents:docs};
 const e=R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{}},[],[task],2026,9);
 assert.equal(e.days['2026-09-03'].issues.length,0);
 T.state.year=2026;T.state.month=9;T.state.employees=[e];T.state.modalId='e1';T.state.modalDate='2026-09-03';
 let requests=0;
 T.client.SharePointStorage={getPreviewInfo:async(file,interactive)=>{
  requests++;
  assert.equal(file.driveId,'drive1');
  assert.equal(file.itemId,'item1');
  assert.equal(interactive,true);
  return {getUrl:'https://example.sharepoint.com/preview'};
 }};
 await T.loadSupportingPreviews();
 const key=T.documentKey(T.dailySupporting(e.days['2026-09-03'])[0],docs[0]);
 assert.equal(requests,1);
 assert.equal(T.state.documentPreviews[key].getUrl,'https://example.sharepoint.com/preview');
});

test('Sinkronisasi memuat metadata rinci dari kedua koleksi sumber',async()=>{
 const R=rules(),T=appHarness(R);
 T.state.year=2026;T.state.month=9;
 const docs=[{id:'file1',type:'Surat',number:'N-1',file:{name:'Lampiran.pdf',driveId:'d',itemId:'i'}}];
 const dbData={
  leaveRecords:[{id:'leave1',data:()=>({employeeId:'e1',leaveYear:2026,startDate:'2026-09-01',endDate:'2026-09-05',
   leaveType:'Cuti Tahunan',documentNumber:'C1',documentDate:'2026-08-30',note:'Disetujui',documents:docs})}],
  assignmentRecords:[{id:'task1',data:()=>({assignmentYear:2026,startDate:'2026-09-07',endDate:'2026-09-08',
   assignmentType:'other',employees:[{employeeId:'e1'}],activity:'Kegiatan Internal',letterNumber:'ST1',letterDate:'2026-09-06',
   location:'Jakarta',note:'Catatan tugas',documents:docs})}]
 };
 T.client.FirebaseClient={getDb:()=>({collection:(name)=>({where:()=>({get:async()=>({docs:dbData[name]||[]})})})})};
 const result=await T.getSupportingData();
 assert.equal(result.leaves.length,1);
 assert.equal(result.assignments.length,1);
 assert.equal(result.leaves[0].documentDate,'2026-08-30');
 assert.equal(result.leaves[0].documents[0].number,'N-1');
 assert.equal(result.assignments[0].location,'Jakarta');
 assert.equal(result.assignments[0].note,'Catatan tugas');
 assert.equal(result.assignments[0].documents[0].file.itemId,'i');
});
