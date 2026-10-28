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
 const warnings=[];
 const code=fs.readFileSync(path.join(scripts,'uang-makan-app.js'),'utf8')
  .replace(/\ninit\(\);\s*\}\)\(\);\s*$/,`\nwindow.__umTest={state,previewPanel,dailySupporting,supportingFilesSection,supportingRecordsSection,supportDocuments,documentKey,modal,loadSupportingPreviews,getSupportingData,archiveAbsences,historyTable,computeEmployees,saveDay,saveSupport,reviewParticipants,periodRecommendation,storedSupporting,hydrateArchivedEmployee,sourceReviewPanel,sharePointPanel,checkSharePoint,connectSharePoint};\n})();`);
 const root={addEventListener:()=>{},querySelector:()=>null};
 const ctx=sandbox(code,{
  window:{UangMakanRules:R,addEventListener:()=>{},UserAccess:{can:()=>true,assert:()=>true}},
  document:{getElementById:()=>root},
  localStorage:{getItem:()=>null},
  alert:message=>warnings.push(String(message)),
  location:{hash:''}
 });
 return Object.assign(ctx.window.__umTest,{client:ctx.window,doc:ctx.document,warnings});
} 
function rules(){return sandbox(ruleCode).window.UangMakanRules;}
function attendance(key,extras={}){
 return {key,inMinutes:8*60,outMinutes:17*60,status:'Hadir',flags:[],...extras};
}

test('WFA/WFH Senin-Kamis direkomendasikan 0 tetapi wajib verifikasi',()=>{
 const R=rules();
 for(const status of ['WFH','WFA','WORK FROM HOME','work from anywhere','Bekerja dari rumah']){
  const d=R.classifyDay(attendance('2026-09-14',{status}));
  assert.equal(d.automatic,0,status);
  assert.equal(d.value,0,status);
  assert.equal(d.remoteExcluded,true,status);
  assert.ok(d.issues.some(issue=>issue.includes('WFH/WFA selain Jumat')),status);
  assert.equal(d.reviewed,false,status);
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
test('WFH/WFA di luar Jumat bisa dikoreksi menjadi mendapat, tetapi libur tetap terkunci nol',()=>{
 const R=rules();
 const remote=R.classifyDay(attendance('2026-09-16',{status:'WFA',manualValue:1,reviewed:true}));
 assert.equal(remote.automatic,0);
 assert.equal(remote.value,1);
 assert.equal(remote.manualValue,1);
 assert.equal(remote.reviewed,true);
 for(const d of [R.classifyDay(attendance('2026-09-19',{manualValue:1})),
                  R.classifyDay(attendance('2026-09-17',{isHoliday:true,manualValue:1}))]){
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


test('Cuti tanpa konflik wajib diverifikasi, tetap muncul pada kalender dan dokumen',()=>{
 const R=rules(),T=appHarness(R);
 const docs=T.supportDocuments([{id:'doc-cuti',type:'Surat Cuti',name:'Persetujuan Cuti',number:'ABC/2026',
  file:{name:'Persetujuan.pdf',driveId:'drive',itemId:'item',webUrl:'https://example.sharepoint.com/cuti.pdf'}}]);
 const leave={id:'cuti1',employeeId:'e1',startDate:'2026-09-01',endDate:'2026-09-02',leaveType:'Cuti Tahunan',
  documentNumber:'SK-CUTI/2026',documentDate:'2026-08-29',note:'Cuti yang telah disetujui',documents:docs};
 const e=R.applySupporting({id:'e1',name:'Pegawai Contoh',nip:'123',jenisPegawai:'PNS',days:{'2026-09-01':{key:'2026-09-01',inMinutes:null,outMinutes:null,status:'Cuti',flags:['C']}}},[leave],[],2026,9);
 T.state.year=2026;T.state.month=9;T.state.employees=[e];T.state.modalId='e1';T.state.modalDate='2026-09-01';
 const day=e.days['2026-09-01'];
 assert.ok(day.issues.some(x=>x.includes('Konfirmasi pelaksanaan cuti')));
 assert.equal(day.value,0);
 assert.equal(T.dailySupporting(day).length,1);
 const key=T.documentKey(T.dailySupporting(day)[0],docs[0]);
 T.state.documentPreviews[key]={loading:false,getUrl:'https://example.sharepoint.com/preview-cuti.pdf'};
 const html=T.modal();
 assert.ok(html.includes('um-day-source cuti')); // Sel 1 Sep: tidak ada konflik
 assert.ok(html.includes('2 tanggal cuti'));
 assert.ok(html.includes('Verifikasi Cuti / Surat Tugas per Periode'));
 assert.ok(html.includes('Perlu verifikasi'));
 assert.ok(html.includes('2 tanggal kerja pada September 2026'));
 assert.ok(html.includes('Data Cuti &amp; Surat Tugas Terintegrasi'));
 assert.ok(html.includes('Cuti Tahunan'));
 assert.ok(html.includes('SK-CUTI/2026'));
 assert.ok(html.includes('Cuti yang telah disetujui'));
 assert.ok(html.includes('Persetujuan.pdf'));
 assert.ok(html.includes('preview-cuti.pdf'));
 assert.ok(html.includes('Buka SharePoint'));
});

test('Surat tugas tersinkron, termasuk pengaturan kerja, perlu verifikasi meskipun tidak ada presensi',()=>{
 const R=rules(),T=appHarness(R);
 const task={id:'st1',employeeIds:['e1'],startDate:'2026-09-03',endDate:'2026-09-04',
  assignmentType:'work_arrangement',activity:'Pengaturan kerja pelayanan',letterNumber:'ST/123/2026',
  letterDate:'2026-09-02',location:'Kantor Pusdatin',note:'Tanpa presensi fisik',
  documents:T.supportDocuments([{id:'st-doc',name:'Surat Tugas',file:{name:'ST-123.pdf',driveId:'d',itemId:'i'}}])};
 const e=R.applySupporting({id:'e1',name:'Pegawai Contoh',nip:'123',jenisPegawai:'PNS',days:{}},[],[task],2026,9);
 const day=e.days['2026-09-03'];
 assert.equal(day.issues.length,1);
 assert.match(day.issues[0],/Konfirmasi pelaksanaan surat tugas/);
 assert.equal(day.value,0);
 assert.equal(R.overview(e).status,'Perlu verifikasi');
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
 assert.ok(T.modal().includes('Periksa'));
 assert.ok(!T.modal().includes('Ikuti rekomendasi sistem per tanggal'));
 assert.ok(T.modal().includes('Rekomendasi awal sistem:'));
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


test('Pratinjau Microsoft Graph tetap tersedia untuk surat tugas yang memerlukan konfirmasi',async()=>{
 const R=rules(),T=appHarness(R);
 const docs=T.supportDocuments([{id:'doc1',name:'Surat Tugas',file:{name:'Penugasan.pdf',driveId:'drive1',itemId:'item1'}}]);
 const task={id:'st1',employeeIds:['e1'],startDate:'2026-09-03',endDate:'2026-09-03',assignmentType:'official_travel',documents:docs};
 const e=R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{}},[],[task],2026,9);
 assert.equal(e.days['2026-09-03'].issues.length,1);
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


test('Pratinjau dan riwayat menampilkan tombol Unduh ZIP',()=>{
 const R=rules(),T=appHarness(R);
 T.state.year=2026;T.state.month=9;
 T.state.employees=[{id:'1',name:'A',nip:'123',files:['A.xlsx'],days:{}}];
 assert.ok(T.previewPanel().includes('data-um-action="export-zip"'));
 T.state.history=[{id:'h1',year:2026,month:9,createdAt:new Date().toISOString(),summary:{employees:1,days:20}}];
 assert.ok(T.historyTable().includes('data-um-action="history-download-zip"'));
});

test('Arsip Uang Makan mengunggah absensi ke folder SharePoint per pegawai dan menyimpan referensi',async()=>{
 const R=rules(),T=appHarness(R);
 T.state.year=2026;T.state.month=9;
 T.state.employees=[{id:'u1',name:'Pegawai A',nip:'198000000000000001',
  files:['Asli.xlsx'],sourceFiles:[{name:'Asli.xlsx',file:{name:'Asli.xlsx',size:25}}]}];
 let calls=0;
 T.client.SharePointStorage={ensureReady:async(interactive)=>assert.equal(interactive,true),
  uploadAttachment:async(source,options,interactive)=>{
   assert.equal(interactive,true);
   assert.equal(options.path[0],'Uang Makan');
   assert.equal(options.path[1],'2026');
   assert.equal(options.path[2],'09_September');
   assert.equal(options.path[4],'Pegawai');
   assert.equal(options.path[5],'198000000000000001_Pegawai A');
   assert.equal(options.remoteName,'1_Asli.xlsx');
   calls++;
   return {...source,driveId:'d',itemId:'i',file:source.file};
  }};
 await T.archiveAbsences();
 assert.equal(calls,1);
 assert.equal(T.state.employees[0].sourceFiles[0].itemId,'i');
 await T.archiveAbsences();
 assert.equal(calls,1,'Berkas yang sudah terarsip tidak boleh diunggah ulang');
});

test('Riwayat menyimpan referensi absensi SharePoint tanpa menyimpan binary lokal',()=>{
 const code=fs.readFileSync(path.join(scripts,'uang-makan-storage.js'),'utf8').replace(
  'window.UangMakanStorage=Object.freeze({list,save,get,remove});',
  'window.__umStorageTest={sanitizeEmployee};window.UangMakanStorage=Object.freeze({list,save,get,remove});');
 const storage=sandbox(code).window.__umStorageTest;
 const row=storage.sanitizeEmployee({id:'a',name:'A',nip:'123',files:['X.xlsx'],sourceFiles:[{
  name:'X.xlsx',size:100,driveId:'d',itemId:'i',webUrl:'https://example.sharepoint.com/x',
  file:{content:'Rahasia!'},getUrl:'https://temporary.example.com/preview'
 }],days:{}});
 assert.equal(row.sourceFiles[0].itemId,'i');
 assert.equal(row.sourceFiles[0].file,undefined);
 assert.equal(row.sourceFiles[0].getUrl,undefined);
 assert.equal(row.sourceFiles[0].webUrl,'https://example.sharepoint.com/x');
});

test('Semua jenis surat tugas wajib diperiksa pada hari kerja, walaupun tidak ada benturan absensi',()=>{
 const R=rules();
 for(const assignmentType of ['official_travel','work_arrangement','other']){
  const task={id:'st-01',assignmentType,activity:'Agenda',startDate:'2026-09-24',endDate:'2026-09-24'};
  for(const record of [attendance('2026-09-24'),attendance('2026-09-24',{inMinutes:null,outMinutes:null,status:'WFO'})]){
   const d=R.classifyDay(record,[],[task]);
   assert.ok(d.issues.some(x=>x.includes('Konfirmasi pelaksanaan surat tugas')),assignmentType);
   assert.equal(d.reviewed,false);
   assert.equal(R.overview({jenisPegawai:'PNS',days:{[d.key]:d}}).status,'Perlu verifikasi');
  }
 }
 const weekend=R.classifyDay(attendance('2026-09-26'),[],[{id:'s',assignmentType:'official_travel'}]);
 const holiday=R.classifyDay(attendance('2026-09-24',{isHoliday:true}),[],[{id:'s',assignmentType:'official_travel'}]);
 assert.equal(weekend.issues.length,0);
 assert.equal(holiday.issues.length,0);
 assert.equal(weekend.value,0);
 assert.equal(holiday.value,0);
});

test('Koreksi tidak dihitung jika pilihan manual sama dengan rekomendasi otomatis',()=>{
 const R=rules();
 const d=R.classifyDay(attendance('2026-09-24',{manualValue:1,reviewed:true}),[],[{id:'s',assignmentType:'other'}]);
 assert.equal(d.reviewed,true);
 assert.equal(R.overview({jenisPegawai:'PNS',days:{[d.key]:d}}).issues,0);
 assert.equal(R.overview({jenisPegawai:'PNS',days:{[d.key]:d}}).manual,0);
 const corrected=R.classifyDay(attendance('2026-09-24',{manualValue:0,reviewed:true}),[],[{id:'s',assignmentType:'other'}]);
 assert.equal(R.overview({jenisPegawai:'PNS',days:{[corrected.key]:corrected}}).manual,1);
});

test('Sinkronisasi ulang mempertahankan verifikasi dokumen yang identik, membatalkan yang berubah atau dihapus',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const original={id:'s1',employeeIds:['e1'],startDate:'2026-09-24',endDate:'2026-09-25',
  assignmentType:'other',activity:'Agenda Awal',letterNumber:'ST-12'};
 const e=R.applySupporting({id:'e1',jenisPegawai:'PNS',days:{'2026-09-24':attendance('2026-09-24')}},[],[original],2026,9);
 const k=R.supportKey('tugas',original);
 e.supportReviews[k]={choice:0,signature:R.supportSignature('tugas',original),reviewedAt:'2026-09-25T01:00:00Z',reviewedBy:'operator@example.com'};
 const verified=R.applySupporting(e,[],[original],2026,9);
 for(const date of ['2026-09-24','2026-09-25'])assert.equal(verified.days[date].reviewed,true);
 const same=T.computeEmployees([verified],[],[original])[0];
 assert.equal(same.days['2026-09-24'].reviewed,true);
 assert.equal(same.days['2026-09-24'].value,0);
 const changed=T.computeEmployees([verified],[],[{...original,activity:'Agenda Batal Dilaksanakan'}])[0];
 assert.equal(changed.days['2026-09-24'].reviewed,false);
 assert.equal(changed.days['2026-09-25'].reviewed,false);
 assert.equal(changed.days['2026-09-24'].manualValue,null);
 assert.equal(changed.days['2026-09-24'].value,1,'Kembali ke perhitungan absensi sampai diverifikasi ulang');
 assert.equal(Object.keys(changed.supportReviews).length,0);
 const cancelled=T.computeEmployees([verified],[],[])[0];
 assert.equal(cancelled.days['2026-09-24'].manualValue,null);
 assert.equal(cancelled.days['2026-09-24'].value,1);
 assert.equal(cancelled.days['2026-09-24'].issues.length,0);
 assert.equal(Object.keys(cancelled.supportReviews).length,0);
});

test('Pilihan hasil per dokumen tidak menampilkan angka 1/0, tanpa alasan koreksi tertulis',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'s1',employeeIds:['e1'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel'};
 T.state.assignments=[task];
 T.state.employees=[R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{
  '2026-09-24':attendance('2026-09-24'),'2026-09-25':attendance('2026-09-25')
 }},[],[task],2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 const html=T.modal();
 assert.ok(!html.includes('Ikuti rekomendasi sistem per tanggal'));
 assert.ok(html.includes('Rekomendasi awal sistem:'));
 assert.ok(html.includes('value="1"'));
 assert.ok(html.includes('value="0"'));
 assert.ok(html.includes('>Memperoleh uang makan</option>'));
 assert.ok(html.includes('>Tidak memperoleh uang makan</option>'));
 assert.ok(html.includes('2 tanggal kerja pada September 2026'));
 assert.ok(!html.includes('>1 - Memperoleh'));
 assert.ok(!html.includes('id="um-note"'));
 assert.ok(!html.includes('Catatan Verifikasi / Alasan Koreksi'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'operator@example.com'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'0'}:null;
 T.saveSupport('tugas','s1','0');
 const employee=T.state.employees[0];
 for(const date of ['2026-09-24','2026-09-25']){
  assert.equal(employee.days[date].reviewed,true);
  assert.equal(employee.days[date].value,0);
  assert.equal(employee.days[date].manualValue,0);
  assert.equal(employee.days[date].reviewedBy,'operator@example.com');
 }
 assert.equal(employee.days['2026-09-24'].manualNote,'');
 assert.equal(R.overview(employee).status,'Siap');
});

test('Panel SharePoint menunjukkan kondisi belum terhubung dan terhubung setelah autentikasi',async()=>{
 const R=rules(),T=appHarness(R);
 const status={configured:true,libraryLoaded:true,connected:false,email:'',name:''};
 let authCalls=0,rootCalls=0;
 T.client.MicrosoftAuth={
  getStatus:()=>status,
  isConfigured:()=>true,
  initialize:async()=>true,
  connect:async()=>{authCalls++;status.connected=true;status.email='pegawai@pkp.go.id';status.name='Pegawai PKP';}
 };
 T.client.SharePointStorage={ensureReady:async interactive=>{rootCalls++;assert.equal(interactive,true);}};
 assert.ok(T.sharePointPanel().includes('data-um-action="connect-sharepoint"'));
 assert.ok(T.sharePointPanel().includes('Hubungkan SharePoint'));
 await T.connectSharePoint();
 assert.equal(authCalls,1);
 assert.equal(rootCalls,1);
 assert.equal(T.state.sharePointReady,true);
 assert.ok(T.sharePointPanel().includes('SharePoint terhubung'));
 assert.ok(T.sharePointPanel().includes('pegawai@pkp.go.id'));
 assert.ok(!T.sharePointPanel().includes('data-um-action="connect-sharepoint"'));
});

test('Status SharePoint diperiksa ulang tanpa popup dan kegagalan akses tidak dinyatakan terhubung',async()=>{
 const R=rules(),T=appHarness(R);
 const status={configured:true,libraryLoaded:true,connected:true,email:'a@pkp.go.id',name:'A'};
 let promptCalls=0,checks=0;
 T.client.MicrosoftAuth={getStatus:()=>status,isConfigured:()=>true,initialize:async()=>{promptCalls++;}};
 T.client.SharePointStorage={ensureReady:async interactive=>{assert.equal(interactive,false);checks++;throw new Error('Akses ditolak');}};
 await T.checkSharePoint();
 assert.equal(promptCalls,1);
 assert.equal(checks,1);
 assert.equal(T.state.sharePointReady,false);
 assert.ok(T.sharePointPanel().includes('Koneksi SharePoint perlu diperiksa'));
 assert.ok(T.sharePointPanel().includes('Akses ditolak'));
 assert.ok(T.sharePointPanel().includes('Periksa Koneksi SharePoint'));
});

test('Satu verifikasi cuti berlaku untuk seluruh tanggal kerja dalam rentang, termasuk bila tidak ada konflik presensi',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 // 1–7 September mencakup Sabtu 5 dan Minggu 6; 3 September tanggal merah tambahan.
 const leave={id:'leave-7',employeeId:'e1',startDate:'2026-09-01',endDate:'2026-09-07',
  leaveType:'Cuti Tahunan',documentNumber:'C-7'};
 T.state.leaves=[leave];T.state.holidays=['2026-09-03'];
 T.state.employees=[R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{
  '2026-09-01':{key:'2026-09-01',status:'Cuti',flags:['C']},
  '2026-09-02':attendance('2026-09-02') // cuti didaftarkan tetapi ada presensi
 }},[leave],[],2026,9,T.state.holidays)];
 let e=T.state.employees[0];
 assert.equal(R.overview(e).issues,4,'Hanya empat hari kerja yang memerlukan verifikasi');
 T.state.modalId='e1';T.state.modalDate='2026-09-02';
 const html=T.modal();
 assert.ok(html.includes('Cuti Tahunan'));
 assert.ok(html.includes('4 tanggal kerja pada September 2026'));
 assert.ok(html.includes('Simpan Verifikasi Periode'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'reviewer@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'0'}:null;
 T.saveSupport('cuti','leave-7','0');
 e=T.state.employees[0];
 for(const date of ['2026-09-01','2026-09-02','2026-09-04','2026-09-07']){
  assert.equal(e.days[date].value,0,date);assert.equal(e.days[date].reviewed,true,date);
 }
 for(const date of ['2026-09-03','2026-09-05','2026-09-06']){
  assert.equal(e.days[date].value,0,date);assert.equal(e.days[date].nonWorking,true,date);
 }
 assert.equal(R.overview(e).issues,0);
 assert.equal(R.overview(e).status,'Siap');
 assert.equal(Object.keys(e.supportReviews).length,1,'Cukup satu keputusan untuk seluruh tanggal di rentang');
});

test('Satu verifikasi surat tugas lintas bulan hanya diterapkan ke tanggal bulan yang sedang dihitung',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'cross-month',employeeIds:['e1'],startDate:'2026-08-30',endDate:'2026-09-04',
  assignmentType:'official_travel',activity:'Dinas luar'};
 T.state.assignments=[task];
 T.state.employees=[R.applySupporting({id:'e1',name:'Pegawai',nip:'123',jenisPegawai:'PNS',days:{}},[],[task],2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-01';
 assert.equal(R.overview(T.state.employees[0]).issues,4);
 assert.ok(T.modal().includes('4 tanggal kerja pada September 2026'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'reviewer@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:null;
 T.saveSupport('tugas','cross-month','0');
 const e=T.state.employees[0];
 for(const date of ['2026-09-01','2026-09-02','2026-09-03','2026-09-04']){
  assert.equal(e.days[date].reviewed,true,date);assert.equal(e.days[date].value,1,date);
 }
 assert.equal(R.overview(e).status,'Siap');
 assert.equal(R.overview(e).days,4);
});

test('Pilihan periode hanya dua hasil dengan rekomendasi konservatif sebagai pilihan awal',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'mixed',employeeIds:['e1'],startDate:'2026-09-14',endDate:'2026-09-15',assignmentType:'other'};
 T.state.assignments=[task];
 T.state.employees=[R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{
  '2026-09-14':attendance('2026-09-14')
 }},[],[task],2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-14';
 const html=T.modal();
 assert.ok(html.includes('<option value="0" selected>Tidak memperoleh uang makan</option>'));
 assert.ok(html.includes('Hasil otomatis antar tanggal atau pegawai tidak seragam'));
 assert.ok(!html.includes('<option value="auto"'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'reviewer@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'0'}:null;
 T.saveSupport('tugas','mixed','0');
 const e=T.state.employees[0];
 assert.equal(e.days['2026-09-14'].value,0);
 assert.equal(e.days['2026-09-15'].value,0);
 assert.equal(e.days['2026-09-14'].reviewed,true);
 assert.equal(e.days['2026-09-15'].reviewed,true);
 assert.equal(e.supportReviews['tugas:mixed'].choice,0);
});

test('Penugasan tersinkron berbeda yang tumpang tindih harus dikonfirmasi masing-masing',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task1={id:'a',employeeIds:['e1'],startDate:'2026-09-23',endDate:'2026-09-25',assignmentType:'other',activity:'Kegiatan A'};
 const task2={id:'b',employeeIds:['e1'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'other',activity:'Kegiatan B'};
 T.state.assignments=[task1,task2];
 T.state.employees=[R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{}},[],T.state.assignments,2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 assert.equal((T.modal().match(/data-um-action="save-support"/g)||[]).length,2);
 T.client.FirebaseClient={getCurrentUser:()=>({email:'reviewer@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'0'}:null;
 T.saveSupport('tugas','a','0');
 assert.equal(T.state.employees[0].days['2026-09-23'].reviewed,true,'Tanggal yang hanya terkait surat A selesai');
 assert.equal(T.state.employees[0].days['2026-09-24'].reviewed,false,'Surat B belum dikonfirmasi');
 assert.equal(R.overview(T.state.employees[0]).issues,2);
 T.doc.getElementById=id=>id==='um-source-choice-1'?{value:'1'}:null;
 T.saveSupport('tugas','b','1');
 const e=T.state.employees[0];
 assert.equal(e.days['2026-09-24'].reviewConflict,true);
 assert.equal(e.days['2026-09-24'].reviewed,false);
 assert.equal(e.days['2026-09-25'].reviewed,false);
 assert.equal(R.overview(e).status,'Perlu verifikasi');
 const dialog=T.modal();
 assert.ok(!dialog.includes('Koreksi Khusus Tanggal'));
 assert.ok(!dialog.includes('data-um-action="save-day"'));
 assert.ok(!dialog.includes('data-um-action="clear-day"'));
 assert.ok(dialog.includes('Dokumen yang tumpang tindih memiliki keputusan berbeda'));
 assert.ok(dialog.includes('Sinkronkan Ulang'));
 T.doc.getElementById=id=>id==='um-manual'?{value:'0'}:null;
 T.saveDay(); // Bahkan pemanggilan langsung tidak boleh mengubah keputusan dokumen.
 const unresolved=T.state.employees[0];
 assert.equal(unresolved.days['2026-09-24'].reviewed,false);
 assert.equal(unresolved.days['2026-09-25'].reviewed,false);
 assert.equal(R.overview(unresolved).status,'Perlu verifikasi');
 assert.match(T.warnings[0],/dokumen koreksi/);
 // Setelah pembetulan dokumen B menghasilkan keputusan yang sama,
 // kedua hari bersama otomatis selesai tanpa koreksi per tanggal.
 T.doc.getElementById=id=>id==='um-source-choice-1'?{value:'0'}:null;
 T.saveSupport('tugas','b','1');
 const fixed=T.state.employees[0];
 assert.equal(fixed.days['2026-09-24'].reviewConflict,false);
 assert.equal(fixed.days['2026-09-24'].reviewed,true);
 assert.equal(fixed.days['2026-09-25'].reviewed,true);
 assert.equal(R.overview(fixed).status,'Siap');
});

test('WFH/WFA non-Jumat dapat memperoleh lewat koreksi periode dan Jumat tetap terhitung',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const leave={id:'l',employeeId:'e1',startDate:'2026-09-17',endDate:'2026-09-18',leaveType:'Cuti'};
 T.state.leaves=[leave];
 T.state.employees=[R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{
  '2026-09-17':attendance('2026-09-17',{status:'WFH'}),
  '2026-09-18':attendance('2026-09-18',{status:'WFA'})
 }},[leave],[],2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-17';
 T.client.FirebaseClient={getCurrentUser:()=>({email:'operator@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:null;
 T.saveSupport('cuti','l','0');
 const e=T.state.employees[0];
 assert.equal(e.days['2026-09-17'].automatic,0);
 assert.equal(e.days['2026-09-17'].value,1);
 assert.equal(e.days['2026-09-18'].value,1);
 assert.equal(e.days['2026-09-17'].reviewed,true);
 assert.equal(e.days['2026-09-18'].reviewed,true);
});

test('Keputusan periode disimpan dalam skema riwayat; tidak tersimpan URL preview dan dapat direkonstruksi',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const leave={id:'l',employeeId:'e1',startDate:'2026-09-14',endDate:'2026-09-16',leaveType:'Cuti',
  documents:T.supportDocuments([{id:'doc',name:'Cuti.pdf',file:{name:'Cuti.pdf',driveId:'d',itemId:'i'}}])};
 const e=R.applySupporting({id:'e1',name:'Tes',nip:'123',jenisPegawai:'PNS',days:{}},[leave],[],2026,9);
 e.supportReviews[R.supportKey('cuti',leave)]={choice:0,signature:R.supportSignature('cuti',leave),reviewedAt:'2026-09-16T00:00:00Z',reviewedBy:'actor'};
 const final=R.applySupporting(e,[leave],[],2026,9);
 const storageCode=fs.readFileSync(path.join(scripts,'uang-makan-storage.js'),'utf8').replace(
   'window.UangMakanStorage=Object.freeze({list,save,get,remove});',
   'window.__umStorageTest={sanitizeEmployee};window.UangMakanStorage=Object.freeze({list,save,get,remove});');
 const stored=sandbox(storageCode).window.__umStorageTest.sanitizeEmployee(final);
 assert.equal(Object.keys(stored.supportReviews).length,1);
 assert.equal(stored.supportReviews['cuti:l'].choice,0);
 assert.equal(stored.days['2026-09-14'].dateManualValue,null);
 assert.equal(stored.days['2026-09-14'].manualValue,0);
 const sources=T.storedSupporting([stored]);
 const reopened=R.applySupporting(stored,sources.leaves,sources.assignments,2026,9);
 assert.equal(reopened.days['2026-09-14'].reviewed,true);
 assert.equal(reopened.days['2026-09-15'].reviewed,true);
 assert.equal(reopened.days['2026-09-16'].reviewed,true);
 assert.equal(R.overview(reopened).issues,0);
});

test('Hasil dua dokumen berbeda tetap perlu perbaikan sumber; koreksi tanggal lama diabaikan',()=>{
 const R=rules();
 const a={id:'a',employeeIds:['e1'],startDate:'2026-09-14',endDate:'2026-09-14',assignmentType:'other'};
 const b={id:'b',employeeIds:['e1'],startDate:'2026-09-14',endDate:'2026-09-14',assignmentType:'other'};
 const initial={id:'e1',jenisPegawai:'PNS',days:{'2026-09-14':attendance('2026-09-14',{status:'WFH'})}};
 initial.supportReviews={
  'tugas:a':{choice:0,signature:R.supportSignature('tugas',a),reviewedAt:'2026-09-15T00:00:00Z'},
  'tugas:b':{choice:1,signature:R.supportSignature('tugas',b),reviewedAt:'2026-09-15T00:00:00Z'}
 };
 const result=R.applySupporting(initial,[],[a,b],2026,9);
 assert.equal(result.days['2026-09-14'].value,0);
 assert.equal(result.days['2026-09-14'].reviewConflict,true);
 assert.equal(result.days['2026-09-14'].reviewed,false);
 assert.equal(R.overview(result).status,'Perlu verifikasi');
 const resolved=R.applySupporting({...result,days:{...result.days,'2026-09-14':{
  ...result.days['2026-09-14'],dateManualValue:1,dateReviewed:true,
  dateReviewedAt:'2026-09-16T00:00:00Z',dateReviewedBy:'operator'
 }}},[],[a,b],2026,9);
 assert.equal(resolved.days['2026-09-14'].reviewed,false);
 assert.equal(resolved.days['2026-09-14'].value,0);
 assert.equal(resolved.days['2026-09-14'].dateManualValue,null);
 assert.equal(resolved.days['2026-09-14'].dateReviewed,false);
 assert.equal(resolved.days['2026-09-14'].dateReviewedAt,'');
 assert.equal(resolved.days['2026-09-14'].reviewConflict,true);
});

test('Jika satu dokumen tumpang tindih diubah, persetujuan dokumen lain tidak ikut hilang',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const leave={id:'leave',employeeId:'e1',startDate:'2026-09-24',endDate:'2026-09-25',leaveType:'Cuti Tahunan'};
 const task={id:'task',employeeIds:['e1'],startDate:'2026-09-24',endDate:'2026-09-25',
  assignmentType:'official_travel',activity:'Pelatihan'};
 const employee={id:'e1',name:'Tes',jenisPegawai:'PNS',days:{}};
 employee.supportReviews={
  'cuti:leave':{choice:0,signature:R.supportSignature('cuti',leave),reviewedAt:'2026-09-26T00:00:00Z'},
  'tugas:task':{choice:0,signature:R.supportSignature('tugas',task),reviewedAt:'2026-09-26T00:00:00Z'}
 };
 const prior=R.applySupporting(employee,[leave],[task],2026,9);
 assert.equal(R.overview(prior).issues,0);
 const updated=T.computeEmployees([prior],[{...leave,leaveType:'Cuti Besar'}],[task])[0];
 assert.equal(Object.keys(updated.supportReviews).length,1);
 assert.ok(updated.supportReviews['tugas:task']);
 assert.equal(updated.days['2026-09-24'].reviewed,false);
 assert.equal(updated.days['2026-09-24'].value,0,'Dokumen tugas lain yang masih valid tetap dapat memberi hasil sementara');
 assert.equal(R.overview(updated).issues,2);
});

test('WFH Senin tanpa dokumen wajib diverifikasi, hasil default nol tetapi dapat memperoleh',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 T.state.employees=[R.applySupporting({id:'p1',name:'WFH Tes',nip:'111',jenisPegawai:'PNS',days:{
  '2026-09-14':attendance('2026-09-14',{status:'WFH'})
 }},[],[],2026,9)];
 T.state.modalId='p1';T.state.modalDate='2026-09-14';
 const before=T.state.employees[0];
 assert.equal(R.overview(before).status,'Perlu verifikasi');
 assert.equal(before.days['2026-09-14'].automatic,0);
 const html=T.modal();
 assert.match(html,/<select id="um-manual"><option value="1" >Memperoleh uang makan<\/option><option value="0" selected>Tidak memperoleh uang makan<\/option><\/select>/);
 assert.ok(html.includes('hasil akhir tidak dikunci'));
 assert.ok(!html.includes('value="auto"'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'reviewer@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-manual'?{value:'1'}:null;
 T.saveDay();
 const corrected=T.state.employees[0];
 assert.equal(corrected.days['2026-09-14'].automatic,0);
 assert.equal(corrected.days['2026-09-14'].value,1);
 assert.equal(corrected.days['2026-09-14'].reviewed,true);
 assert.equal(R.overview(corrected).status,'Siap');
 const synced=T.computeEmployees([corrected],[],[])[0];
 assert.equal(synced.days['2026-09-14'].value,1);
 assert.equal(synced.days['2026-09-14'].reviewed,true);
});

test('Verifikasi satu Surat Tugas berlaku bagi seluruh pegawai dengan ID dokumen identik, tetapi tidak untuk surat lain bernomor sama',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const shared={id:'task-utama',employeeIds:['e1','e2','e3'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel',letterNumber:'ST/99/2026',activity:'Kegiatan Bersama'};
 const other={id:'task-lain',employeeIds:['e4'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel',letterNumber:'ST/99/2026',activity:'Penugasan Berbeda'};
 T.state.assignments=[shared,other];
 T.state.employees=[
  R.applySupporting({id:'e1',name:'Satu',jenisPegawai:'PNS',days:{}},[],[shared,other],2026,9),
  R.applySupporting({id:'e2',name:'Dua',jenisPegawai:'PNS',days:{'2026-09-24':attendance('2026-09-24')}},[],[shared,other],2026,9),
  R.applySupporting({id:'e3',name:'Tiga',jenisPegawai:'PNS',days:{}},[],[shared,other],2026,9),
  R.applySupporting({id:'e4',name:'Empat',jenisPegawai:'PNS',days:{}},[],[shared,other],2026,9)
 ];
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 const html=T.modal();
 assert.ok(html.includes('Penerapan bersama:'));
 assert.ok(html.includes('3 pegawai, 6 tanggal-pegawai'));
 assert.ok(html.includes('value="0" selected>Tidak memperoleh uang makan'));
 T.client.FirebaseClient={getCurrentUser:()=>({email:'admin@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:null;
 T.saveSupport('tugas','task-utama','0');
 for(const e of T.state.employees.slice(0,3)){
  for(const day of ['2026-09-24','2026-09-25']){
   assert.equal(e.days[day].value,1,`${e.id} ${day}`);
   assert.equal(e.days[day].reviewed,true,`${e.id} ${day}`);
  }
  assert.equal(e.supportReviews['tugas:task-utama'].choice,1);
  assert.equal(R.overview(e).issues,0);
 }
 const unaffected=T.state.employees[3];
 assert.equal(unaffected.supportReviews?.['tugas:task-utama'],undefined);
 assert.equal(unaffected.days['2026-09-24'].reviewed,false);
 assert.equal(R.overview(unaffected).issues,2);
 const reconstructed=T.storedSupporting(T.state.employees);
 assert.equal(reconstructed.assignments.length,2);
 const reopened=T.state.employees.slice(0,3).map(e=>R.applySupporting(e,[],reconstructed.assignments,2026,9));
 assert.ok(reopened.every(e=>e.days['2026-09-24'].reviewed&&e.days['2026-09-25'].value===1));
});

test('Keputusan surat tugas bersama tidak mengesampingkan verifikasi cuti individual yang bertumpang tindih',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'task-bersama',employeeIds:['e1','e2'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel'};
 const leave={id:'cuti-dua',employeeId:'e2',startDate:'2026-09-24',endDate:'2026-09-24',leaveType:'Cuti Tahunan'};
 T.state.assignments=[task];T.state.leaves=[leave];
 T.state.employees=[
  R.applySupporting({id:'e1',name:'Satu',jenisPegawai:'PNS',days:{}},[leave],[task],2026,9),
  R.applySupporting({id:'e2',name:'Dua',jenisPegawai:'PNS',days:{}},[leave],[task],2026,9)
 ];
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 T.client.FirebaseClient={getCurrentUser:()=>({email:'admin@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:null;
 T.saveSupport('tugas','task-bersama','0');
 assert.equal(R.overview(T.state.employees[0]).status,'Siap');
 const e2=T.state.employees[1];
 assert.equal(e2.supportReviews['tugas:task-bersama'].choice,1);
 assert.equal(e2.days['2026-09-24'].value,1);
 assert.equal(e2.days['2026-09-24'].reviewed,false);
 assert.equal(e2.days['2026-09-25'].reviewed,true);
 assert.equal(e2.pendingSourceKeys?.length,undefined);
 assert.equal(e2.days['2026-09-24'].pendingSourceKeys[0],'cuti:cuti-dua');
});

test('Sinkronisasi mempertahankan hasil bersama selama dokumen identik, membatalkan seluruhnya bila berubah',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'task',employeeIds:['e1','e2'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel',activity:'Aktivitas Awal'};
 T.state.assignments=[task];
 T.state.employees=['e1','e2'].map(id=>R.applySupporting({id,name:id,jenisPegawai:'PNS',days:{}},[],[task],2026,9));
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 T.client.FirebaseClient={getCurrentUser:()=>({email:'admin@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:null;
 T.saveSupport('tugas','task','0');
 const resynced=T.computeEmployees(T.state.employees,[],[task]);
 assert.ok(resynced.every(e=>e.days['2026-09-24'].reviewed&&e.supportReviews['tugas:task']));
 const updated={...task,activity:'Aktivitas Diubah'};
 const changed=T.computeEmployees(resynced,[],[updated]);
 assert.ok(changed.every(e=>!e.days['2026-09-24'].reviewed));
 assert.ok(changed.every(e=>!e.supportReviews['tugas:task']));
 assert.ok(changed.every(e=>R.overview(e).status==='Perlu verifikasi'));
});

test('Tanpa koreksi khusus tanggal: keputusan Surat Tugas bersama tidak bisa dioverride individu',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'s',employeeIds:['e1','e2'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel'};
 T.state.assignments=[task];
 T.state.employees=['e1','e2'].map(id=>R.applySupporting({id,name:id,jenisPegawai:'PNS',days:{}},[],[task],2026,9));
 T.state.modalId='e1';T.state.modalDate='2026-09-24';
 T.client.FirebaseClient={getCurrentUser:()=>({email:'admin@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'1'}:id==='um-manual'?{value:'0'}:null;
 T.saveSupport('tugas','s','0');
 T.state.modalId='e2';
 const before=T.modal();
 assert.ok(before.includes('value="1" selected>Memperoleh uang makan'));
 assert.ok(!before.includes('Koreksi Khusus Tanggal'));
 assert.ok(!before.includes('data-um-action="save-day"'));
 assert.ok(!before.includes('data-um-action="clear-day"'));
 assert.ok(!before.includes('id="um-manual"'));
 T.saveDay();
 assert.match(T.warnings.at(-1),/tidak dapat dikoreksi secara manual/);
 for(const e of T.state.employees)for(const date of ['2026-09-24','2026-09-25']){
  assert.equal(e.days[date].value,1,`${e.id}/${date}`);
  assert.equal(e.days[date].reviewed,true,`${e.id}/${date}`);
 }
 // Koreksi tanggal dari run lama tidak lagi memengaruhi hasil saat diproses ulang.
 const former=T.state.employees[1];
 const prior={...former,days:{...former.days,
  '2026-09-24':{...former.days['2026-09-24'],dateManualValue:0,dateReviewed:true}
 }};
 const recalculated=R.applySupporting(prior,[],[task],2026,9);
 assert.equal(recalculated.days['2026-09-24'].value,1);
 assert.equal(recalculated.days['2026-09-24'].dateManualValue,null);
 assert.equal(recalculated.days['2026-09-24'].dateReviewed,false);
 assert.equal(recalculated.days['2026-09-25'].value,1);
});

test('Tanggal dalam satu periode Cuti tidak dapat diubah per hari meskipun sudah diverifikasi',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const leave={id:'cuti-bersama',employeeId:'u1',startDate:'2026-09-14',endDate:'2026-09-16',leaveType:'Cuti Tahunan'};
 T.state.leaves=[leave];
 T.state.employees=[R.applySupporting({id:'u1',name:'Pegawai',jenisPegawai:'PNS',days:{
  '2026-09-15':attendance('2026-09-15')
 }},[leave],[],2026,9)];
 T.state.modalId='u1';T.state.modalDate='2026-09-15';
 T.client.FirebaseClient={getCurrentUser:()=>({email:'operator@pkp.go.id'})};
 T.doc.getElementById=id=>id==='um-source-choice-0'?{value:'0'}:id==='um-manual'?{value:'1'}:null;
 T.saveSupport('cuti','cuti-bersama','0');
 const modal=T.modal();
 assert.ok(modal.includes('Verifikasi Cuti / Surat Tugas per Periode'));
 assert.ok(!modal.includes('Koreksi Khusus Tanggal'));
 assert.ok(!modal.includes('data-um-action="save-day"'));
 T.saveDay();
 assert.match(T.warnings.at(-1),/dokumen koreksi/);
 for(const date of ['2026-09-14','2026-09-15','2026-09-16']){
  assert.equal(T.state.employees[0].days[date].value,0,date);
  assert.equal(T.state.employees[0].days[date].reviewed,true,date);
 }
 // Bila data Cuti diperbarui di sumber, perlu dilakukan verifikasi ulang.
 const correction={...leave,note:'Dokumen pembetulan tanggal pelaksanaan'};
 const corrected=T.computeEmployees(T.state.employees,[correction],[])[0];
 assert.equal(corrected.days['2026-09-15'].reviewed,false);
 assert.equal(corrected.supportReviews['cuti:cuti-bersama'],undefined);
});

test('Riwayat baru tidak menyimpan koreksi khusus tanggal untuk Cuti/Surat Tugas',()=>{
 const R=rules();
 const task={id:'history-task',employeeIds:['u1'],startDate:'2026-09-14',endDate:'2026-09-15',assignmentType:'official_travel'};
 const e={id:'u1',jenisPegawai:'PNS',days:{'2026-09-14':attendance('2026-09-14')}};
 e.supportReviews={'tugas:history-task':{choice:1,signature:R.supportSignature('tugas',task),reviewedAt:'2026-09-20T01:00:00Z'}};
 const review=R.applySupporting(e,[],[task],2026,9);
 review.days['2026-09-14'].dateManualValue=0;
 review.days['2026-09-14'].dateReviewed=true;
 const code=fs.readFileSync(path.join(scripts,'uang-makan-storage.js'),'utf8').replace(
  'window.UangMakanStorage=Object.freeze({list,save,get,remove});',
  'window.__umStorageTest={sanitizeEmployee};window.UangMakanStorage=Object.freeze({list,save,get,remove});');
 const saved=sandbox(code).window.__umStorageTest.sanitizeEmployee(review);
 assert.equal(saved.days['2026-09-14'].dateManualValue,null);
 assert.equal(saved.days['2026-09-14'].dateReviewed,false);
 assert.equal(saved.days['2026-09-14'].dateReviewedAt,'');
 assert.equal(saved.days['2026-09-14'].manualValue,1,'Keputusan periode tetap dicatat');
 const restored=R.applySupporting(saved,[],[task],2026,9);
 assert.equal(restored.days['2026-09-14'].value,1);
 assert.equal(restored.days['2026-09-15'].value,1);
});

test('Riwayat skema 5 tetap menampilkan koreksi historis tanpa membolehkan koreksi di proses baru',()=>{
 const R=rules(),T=appHarness(R);
 const task={id:'archive',employeeIds:['e1'],startDate:'2026-09-24',endDate:'2026-09-25',assignmentType:'official_travel'};
 const original={id:'e1',jenisPegawai:'PNS',days:{}};
 original.supportReviews={'tugas:archive':{choice:1,signature:R.supportSignature('tugas',task),reviewedAt:'2026-09-26T00:00:00Z'}};
 const current=R.applySupporting(original,[],[task],2026,9);
 const stored={...current,days:{...current.days,
  '2026-09-24':{...current.days['2026-09-24'],dateManualValue:0,dateReviewed:true,
   manualValue:0,reviewed:true,dateReviewedAt:'2026-09-26T01:00:00Z',reviewedAt:'2026-09-26T01:00:00Z'}
 }};
 const archived=T.hydrateArchivedEmployee(stored,{leaves:[],assignments:[task]},2026,9,[],5);
 assert.equal(archived.days['2026-09-24'].value,0,'Nilai lama pada arsip tidak boleh berubah');
 assert.equal(archived.days['2026-09-25'].value,1);
 assert.equal(archived.days['2026-09-24'].reviewed,true);
 const newRun=R.applySupporting(stored,[],[task],2026,9);
 assert.equal(newRun.days['2026-09-24'].value,1,'Perhitungan baru wajib mengikuti dokumen');
 const v6=T.hydrateArchivedEmployee(stored,{leaves:[],assignments:[task]},2026,9,[],6);
 assert.equal(v6.days['2026-09-24'].value,1,'Arsip versi baru tidak mendukung pengecualian tanggal');
});

test('Verifikasi lama bernilai auto tetap dapat dibaca, tetapi UI baru tidak menawarkan pilihan otomatis',()=>{
 const R=rules(),T=appHarness(R);T.state.year=2026;T.state.month=9;
 const task={id:'lama',employeeIds:['e1'],startDate:'2026-09-14',endDate:'2026-09-15',assignmentType:'other'};
 T.state.assignments=[task];
 const e={id:'e1',name:'Lama',jenisPegawai:'PNS',days:{'2026-09-14':attendance('2026-09-14')}};
 e.supportReviews={'tugas:lama':{choice:'auto',signature:R.supportSignature('tugas',task),reviewedAt:'2026-09-16T00:00:00Z'}};
 T.state.employees=[R.applySupporting(e,[],[task],2026,9)];
 T.state.modalId='e1';T.state.modalDate='2026-09-14';
 const restored=T.state.employees[0];
 assert.equal(restored.days['2026-09-14'].value,1);
 assert.equal(restored.days['2026-09-15'].value,0);
 assert.equal(restored.days['2026-09-14'].reviewed,true);
 assert.equal(restored.days['2026-09-15'].reviewed,true);
 assert.ok(!T.modal().includes('<option value="auto"'));
});
