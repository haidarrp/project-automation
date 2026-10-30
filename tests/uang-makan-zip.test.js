'use strict';
const assert=require('node:assert/strict');
const {test}=require('node:test');
const vm=require('node:vm');
const fs=require('node:fs');
const path=require('node:path');
const JSZip=require('../assets/vendor/jszip.min.js');
const code=fs.readFileSync(path.resolve(__dirname,'../assets/js/uang-makan-zip.js'),'utf8');

function fixture(){
 const requested=[];
 const window={
  JSZip,
  UangMakanRules:{MONTHS:['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember']},
  UangMakanExport:{generate:async (year,month,employees,options)=>{
   assert.equal(options.download,false);
   assert.equal(year,2026);
   assert.equal(month,9);
   assert.equal(employees.length,2);
   return {filename:'Rekap_September.xlsx',blob:new Blob(['Isi rekap Excel'])};
  }},
  AttendanceDocumentPolicy:{fetch:async()=>({stale:[],exportable:()=>true})},
  SharePointStorage:{downloadAttachment:async (item,interactive)=>{
   assert.equal(interactive,true);
   requested.push(`${item.driveId}:${item.itemId}`);
   return new Blob([`File-${item.itemId}`]);
  }}
 };
 vm.runInNewContext(code,{window,Blob,URL,document:{},setTimeout});
 return {gen:window.UangMakanZip,window,requested};
}
function employee(name,id,source,days={}){
 return {id,name,nip:`19800000000000000${id}`,files:[source.name],sourceFiles:[source],days};
}

test('ZIP lengkap memuat Excel, absensi asli tiap pegawai, bukti Cuti & Surat Tugas tanpa duplikasi tanggal',async()=>{
 const {gen,requested}=fixture();
 const docCuti={id:'c',file:{name:'Surat-Cuti.pdf',driveId:'drive',itemId:'cuti'}};
 const docTugas={id:'t',file:{name:'Surat-Tugas.pdf',driveId:'drive',itemId:'tugas'}};
 const days={
  '2026-09-01':{leaves:[{id:'leave-1',documents:[docCuti]}],assignments:[]},
  '2026-09-02':{leaves:[{id:'leave-1',documents:[docCuti]}],assignments:[{id:'tugas-1',documents:[docTugas]}]}
 };
 const data=[
  employee('Pegawai A','1',{name:'Absensi_A.xlsx',file:new Blob(['absensi-a'])},days),
  employee('Pegawai B','2',{name:'Absensi_B.xlsx',driveId:'drive',itemId:'absen-b'})
 ];
 const result=await gen.generate(2026,9,data);
 assert.equal(result.name,'Uang_Makan_PUSDATIN_September_2026.zip');
 assert.equal(result.attendanceCount,2);
 assert.equal(result.evidenceCount,2);
 const zip=await JSZip.loadAsync(await result.blob.arrayBuffer());
 const filenames=Object.keys(zip.files).filter(key=>!zip.files[key].dir).sort();
 assert.deepEqual(filenames,[
  'Pegawai/Pegawai A/Absensi_A.xlsx',
  'Pegawai/Pegawai A/Surat-Cuti.pdf',
  'Pegawai/Pegawai A/Surat-Tugas.pdf',
  'Pegawai/Pegawai B/Absensi_B.xlsx',
  'Rekap_September.xlsx'
 ].sort());
 assert.equal(await zip.file('Pegawai/Pegawai A/Absensi_A.xlsx').async('string'),'absensi-a');
 assert.equal(await zip.file('Pegawai/Pegawai B/Absensi_B.xlsx').async('string'),'File-absen-b');
 assert.deepEqual(requested.sort(),['drive:absen-b','drive:cuti','drive:tugas'].sort());
});

test('Folder/berkas diamankan dari nama path traversal dan nama pegawai ganda',async()=>{
 const {gen}=fixture();
 const results=await gen.generate(2026,9,[
  employee('../Pegawai\\Rahasia','1',{name:'../absensi\\input.xlsx',file:new Blob(['1'])}),
  employee('../Pegawai\\Rahasia','2',{name:'../absensi\\input.xlsx',file:new Blob(['2'])})
 ]);
 const zip=await JSZip.loadAsync(await results.blob.arrayBuffer());
 const files=Object.keys(zip.files).filter(key=>!zip.files[key].dir);
 assert.equal(files.length,3);
 assert.ok(files.some(x=>x.startsWith('Pegawai/_Pegawai_Rahasia/')));
 assert.ok(files.some(x=>x.startsWith('Pegawai/_Pegawai_Rahasia (2)/')));
 assert.ok(files.every(x=>!x.includes('../')&&!x.includes('\\')));
});

test('Tidak membuat ZIP tanpa sumber absensi lama atau saat bukti SharePoint gagal diunduh',async()=>{
 const {gen,window}=fixture();
 await assert.rejects(gen.generate(2026,9,[
  employee('A','1',{name:'A.xlsx',file:new Blob(['A'])}),
  {id:'2',name:'B',nip:'123',files:['B.xlsx'],days:{}}
 ]),/Riwayat lama hanya menyimpan nama file/);
 const data=[
  employee('A','1',{name:'A.xlsx',file:new Blob(['A'])},{'2026-09-01':{
   leaves:[{documents:[{file:{name:'Terlindungi.pdf',driveId:'d',itemId:'forbidden'}}]}],assignments:[]
  }}),
  employee('B','2',{name:'B.xlsx',file:new Blob(['B'])})
 ];
 window.SharePointStorage.downloadAttachment=async()=>{throw new Error('403 Forbidden');};
 await assert.rejects(gen.generate(2026,9,data),/Gagal mengambil bukti dukung A.*403 Forbidden/);
});


test('Bukti dukung yang baru memiliki metadata dicatat sebagai berkas belum tersedia dalam ZIP',async()=>{
 const {gen}=fixture();
 const data=[
  employee('Pegawai A','1',{name:'A.xlsx',file:new Blob(['data'])},{
   '2026-09-03':{leaves:[{id:'cuti-1',documents:[{id:'doc-1',name:'SK Cuti Tahunan',file:{}}]}],assignments:[]},
   '2026-09-04':{leaves:[{id:'cuti-1',documents:[{id:'doc-1',name:'SK Cuti Tahunan',file:{}}]}],assignments:[]}
  }),
  employee('Pegawai B','2',{name:'B.xlsx',file:new Blob(['data'])})
 ];
 const output=await gen.generate(2026,9,data);
 assert.equal(output.missingEvidenceCount,1);
 const zip=await JSZip.loadAsync(await output.blob.arrayBuffer());
 const note=zip.file('Pegawai/Pegawai A/BERKAS_BELUM_TERSEDIA.txt');
 assert.ok(note);
 assert.match(await note.async('string'),/SK Cuti Tahunan/);
});

test('Beberapa absensi dengan nama sama pada seorang pegawai tidak saling menimpa',async()=>{
 const {gen}=fixture();
 const data=[employee('Pegawai A','1',{name:'Absen.xlsx',file:new Blob(['satu'])}),employee('Pegawai B','2',{name:'B.xlsx',file:new Blob(['lain'])})];
 data[0].files.push('Absen.xlsx');
 data[0].sourceFiles.push({name:'Absen.xlsx',file:new Blob(['dua'])});
 const output=await gen.generate(2026,9,data);
 const zip=await JSZip.loadAsync(await output.blob.arrayBuffer());
 assert.equal(output.attendanceCount,3);
 assert.equal(await zip.file('Pegawai/Pegawai A/Absen.xlsx').async('string'),'satu');
 assert.equal(await zip.file('Pegawai/Pegawai A/Absen (2).xlsx').async('string'),'dua');
});


test('Nama berkas identik pada absensi, cuti, dan surat tugas tidak saling menimpa',async()=>{
 const {gen}=fixture();
 const days={'2026-09-01':{
  leaves:[{id:'cuti',documents:[{id:'c',file:{name:'Dokumen.pdf',driveId:'d',itemId:'cuti'}}]}],
  assignments:[{id:'tugas',documents:[{id:'t',file:{name:'Dokumen.pdf',driveId:'d',itemId:'tugas'}}]}]
 }};
 const data=[
  employee('Pegawai A','1',{name:'Dokumen.pdf',file:new Blob(['absensi'])},days),
  employee('Pegawai B','2',{name:'B.xlsx',file:new Blob(['B'])})
 ];
 const output=await gen.generate(2026,9,data);
 const zip=await JSZip.loadAsync(await output.blob.arrayBuffer());
 assert.equal(await zip.file('Pegawai/Pegawai A/Dokumen.pdf').async('string'),'absensi');
 assert.equal(await zip.file('Pegawai/Pegawai A/Dokumen (2).pdf').async('string'),'File-cuti');
 assert.equal(await zip.file('Pegawai/Pegawai A/Dokumen (3).pdf').async('string'),'File-tugas');
 const employeeFiles=Object.keys(zip.files).filter(name=>!zip.files[name].dir&&name.startsWith('Pegawai/Pegawai A/'));
 assert.ok(employeeFiles.every(name=>name.split('/').length===3));
});

test('Dokumen tereliminasi tetap ada dalam metadata pratinjau namun tidak ditambahkan ke ZIP',async()=>{
 const {gen,window}=fixture();
 window.AttendanceDocumentPolicy.fetch=async()=>({stale:[],exportable:(kind,row,employeeId)=>!(kind==='surat'&&row.id==='st-eliminated'&&employeeId==='1')});
 const doc={id:'stfile',file:{name:'ST-dieliminasi.pdf',driveId:'d',itemId:'st-dieliminasi'}};
 const archived={
  '2026-09-14':{leaves:[],assignments:[],previewLeaves:[],previewAssignments:[{id:'st-eliminated',documents:[doc]}]}
 };
 const data=[employee('Pegawai A','1',{name:'Absen_A.xlsx',file:new Blob(['a'])},archived),
  employee('Pegawai B','2',{name:'Absen_B.xlsx',file:new Blob(['b'])})];
 const result=await gen.generate(2026,9,data);
 const zip=await JSZip.loadAsync(await result.blob.arrayBuffer());
 assert.equal(result.evidenceCount,0);
 assert.equal(zip.file('Pegawai/Pegawai A/ST-dieliminasi.pdf'),null);
 assert.equal(archived['2026-09-14'].previewAssignments.length,1,'Metadata pratinjau tidak dihapus');
});

test('Surat Tugas tidak dipertanggungjawabkan dihapus dari daftar berkas ZIP',async()=>{
 const {gen,window}=fixture();
 window.AttendanceDocumentPolicy.fetch=async()=>({stale:[],exportable:(kind,row)=>kind!=='surat'||row.accountable!==false});
 const doc={id:'stfile',file:{name:'ST-jalan.pdf',driveId:'d',itemId:'jalan'}};
 const data=[employee('Pegawai A','1',{name:'Absen_A.xlsx',file:new Blob(['a'])},{
  '2026-09-14':{assignments:[{id:'st-no',accountable:false,documents:[doc]}]}
 }),employee('Pegawai B','2',{name:'Absen_B.xlsx',file:new Blob(['b'])})];
 const result=await gen.generate(2026,9,data);
 assert.equal(result.evidenceCount,0);
});
