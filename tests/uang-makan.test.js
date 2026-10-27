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

test('Tabel pratinjau menampilkan semua pegawai dan iframe berkas konflik',()=>{
 const R=rules();
 const appCode=fs.readFileSync(path.join(scripts,'uang-makan-app.js'),'utf8')
  .replace(/\ninit\(\);\s*\}\)\(\);\s*$/,`\nwindow.__umTest={state,previewPanel,conflictSupporting,supportingFilesSection,supportDocuments,documentKey};\n})();`);
 const root={addEventListener:()=>{},querySelector:()=>null};
 const ctx=sandbox(appCode,{
  window:{UangMakanRules:R,addEventListener:()=>{}},
  document:{getElementById:()=>root},
  localStorage:{getItem:()=>null},
  location:{hash:''}
 });
 const T=ctx.window.__umTest;
 assert.ok(T);
 T.state.year=2026;T.state.month=9;
 T.state.employees=Array.from({length:47},(_,i)=>({id:`id-${i}`,name:`Pegawai ${i+1}`,nip:`123${i}`,jenisPegawai:'PNS',days:{}}));
 const html=T.previewPanel();
 assert.ok(html.includes('Menampilkan seluruh 47 pegawai'));
 assert.ok(html.includes('Pegawai 47'));
 assert.equal((html.match(/<td class="name">/g)||[]).length,47);
 const docs=T.supportDocuments([{id:'doc1',name:'Surat Perintah',type:'Surat Tugas',file:{name:'Surat.pdf',driveId:'drive',itemId:'item',webUrl:'https://example.sharepoint.com/surat.pdf'}}]);
 const group={kind:'dinas',id:'a1',documents:docs};
 const key=T.documentKey(group,docs[0]);
 T.state.documentPreviews[key]={loading:false,getUrl:'https://example.sharepoint.com/preview.pdf'};
 const conflict=R.classifyDay(attendance('2026-09-16'),[],[{id:'a1',assignmentType:'official_travel',activity:'Rapat',documents:docs}]);
 const filesHtml=T.supportingFilesSection(conflict);
 assert.ok(filesHtml.includes('Berkas Pendukung Cuti / Dinas'));
 assert.ok(filesHtml.includes('Surat.pdf'));
 assert.ok(filesHtml.includes('um-document-frame'));
 assert.ok(filesHtml.includes('Buka SharePoint'));
 assert.ok(filesHtml.includes('preview.pdf'));
});
