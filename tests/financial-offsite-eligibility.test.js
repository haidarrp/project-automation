'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const main=path.join(__dirname,'..','assets','js');
const ctx={window:{},console,Date};
vm.runInNewContext(fs.readFileSync(path.join(main,'uang-makan-rules.js'),'utf8'),ctx);
const UM=ctx.window.UangMakanRules;
const LM=require('../assets/js/lembur-document-integration.js');
const date='2026-11-04';
const statusRecord=(status,extra={})=>({key:date,inMinutes:7*60+35,outMinutes:18*60,status,flags:[],...extra});
const employee=(status,extra={})=>({id:'pegawai-1',name:'Tes Pegawai',jenisPegawai:'PNS',days:{[date]:statusRecord(status,extra)}});
const leave={id:'cuti-1',employeeId:'pegawai-1',startDate:date,endDate:date,leaveType:'Cuti Tahunan'};
const task={id:'st-1',employeeIds:['pegawai-1'],startDate:date,endDate:date,accountable:true,assignmentType:'other'};
const leavePolicy={allows:()=>true};
const overtime=(status,extra={})=>({id:'pegawai-1',name:'Tes Pegawai',records:{[date]:{date:new Date(2026,10,4),status,originalOvertimeHours:3,overtimeHours:3,manualOverride:true,...extra}}});
const verify=(persons,leaves=[],assignments=[],holidays=[])=>LM.applyDocuments(persons,leaves,assignments,leavePolicy,e=>e.id,holidays);

test('Uang Makan: Offsite/tidak_hadir hari kerja tanpa dokumen selalu Perlu Cek, 0 hari walau jam normal dan manual 1',()=>{
 for(const status of ['Offsite','off_site','Off Site','tidak_hadir','Tidak Hadir']){
  const result=UM.applySupporting(employee(status,{manualValue:1,reviewed:true,dateManualValue:1,dateReviewed:true}),[],[],2026,11);
  const day=result.days[date];
  assert.equal(day.statusDocumentMissing,true,status);
  assert.equal(day.automatic,0,status);
  assert.equal(day.value,0,status);
  assert.equal(day.manualValue,null,status);
  assert.equal(day.dateManualValue,null,status);
  assert.equal(day.reviewed,false,status);
  assert.equal(UM.overview(result).status,'Perlu Cek',status);
  assert.equal(UM.overview(result).days,0,status);
  assert.equal(UM.overview(result).present,0,status);
  assert.match(day.issues.join(' '),/Perlu Cek/);
 }
});

test('Uang Makan: dokumen pegawai dan tanggal yang cocok menghapus temuan; dokumen orang/tanggal lain tidak',()=>{
 const a=UM.applySupporting(employee('Offsite'),[leave],[],2026,11).days[date];
 assert.equal(a.statusDocumentMissing,false);
 const b=UM.applySupporting(employee('Offsite'),[],[task],2026,11).days[date];
 assert.equal(b.statusDocumentMissing,false);
 const wrongPerson=UM.applySupporting(employee('Offsite'),[{...leave,employeeId:'pegawai-2'}],[],2026,11).days[date];
 assert.equal(wrongPerson.statusDocumentMissing,true);
 const wrongDate=UM.applySupporting(employee('Offsite'),[],[{...task,startDate:'2026-11-05',endDate:'2026-11-05'}],2026,11).days[date];
 assert.equal(wrongDate.statusDocumentMissing,true);
});

test('Uang Makan: Surat Tugas non-pertanggungjawaban tidak cukup sebagai bukti status keuangan',()=>{
 const nonAccountable={...task,accountable:false};
 const day=UM.applySupporting(employee('Offsite'),[],[nonAccountable],2026,11).days[date];
 assert.equal(day.statusDocumentSupported,false);
 assert.equal(day.statusDocumentMissing,true);
 assert.equal(day.assignments.length,0);
});

test('Uang Makan: weekend/tanggal merah tidak berstatus Perlu Cek; WFO/WFH/Dinas tetap seperti semula',()=>{
 const weekend=UM.classifyDay(statusRecord('Offsite',{key:'2026-11-07',manualValue:1}));
 const holiday=UM.classifyDay(statusRecord('Offsite',{isHoliday:true,manualValue:1}));
 for(const d of [weekend,holiday]){assert.equal(d.statusDocumentMissing,false);assert.equal(d.value,0);}
 for(const status of ['WFO','WFH','Dinas','Libur'])assert.equal(UM.classifyDay(statusRecord(status)).statusDocumentMissing,false);
 assert.equal(UM.classifyDay(statusRecord('WFO')).value,1);
});

test('Lembur: Offsite/Tidak Hadir tanpa dokumen dinonaktifkan, status Perlu Cek bahkan jika jam presensi valid',()=>{
 for(const status of ['Offsite','off_site','tidak_hadir','Tidak Hadir']){
  const e=overtime(status);
  const summary=verify([e]);
  const day=e.records[date];
  assert.equal(summary.missingStatus,1);
  assert.equal(summary.held,1);
  assert.equal(day.statusDocumentNeedsReview,true);
  assert.equal(day.statusDocumentReview.state,'missing');
  assert.equal(day.documentBlocksOvertime,true);
  assert.equal(day.originalOvertimeHours,3);
  assert.equal(day.overtimeHours,0);
  assert.equal(day.manualOverride,false);
  assert.match(day.documentBlockNote,/Perlu Cek/);
 }
});

test('Lembur: dokumen orang lain atau luar tanggal tidak membolehkan klaim',()=>{
 for(const source of [{...leave,employeeId:'pegawai-lain'}, {...task,startDate:'2026-11-05',endDate:'2026-11-05'}]){
  const e=overtime('Offsite');
  verify([e],source.employeeId?[source]:[],source.employeeIds?[source]:[]);
  assert.equal(e.records[date].overtimeHours,0);
  assert.equal(e.records[date].statusDocumentNeedsReview,true);
 }
});

test('Lembur: ST non-pertanggungjawaban masih Perlu Cek; ST accountable menghapus temuan tetapi tetap membatasi klaim',()=>{
 const e=overtime('Offsite');
 verify([e]);
 assert.equal(e.records[date].statusDocumentNeedsReview,true);
 verify([e],[],[{...task,accountable:false}]);
 assert.equal(e.records[date].statusDocumentNeedsReview,true);
 assert.equal(e.records[date].statusDocumentReview.state,'missing');
 assert.equal(e.records[date].documentBlocksOvertime,true);
 verify([e],[],[task]);
 assert.equal(e.records[date].statusDocumentNeedsReview,false);
 assert.equal(e.records[date].documentBlocksOvertime,true);
 assert.equal(e.records[date].overtimeHours,0);
});

test('Lembur: status umum tidak berubah; weekend Offsite/Tidak Hadir selalu ditolak',()=>{
 for(const status of ['WFO','WFH','Dinas','Libur']){
  const e=overtime(status);verify([e]);
  assert.equal(e.records[date].statusDocumentNeedsReview,false,status);
  assert.equal(e.records[date].documentBlocksOvertime,false,status);
 }
 const holiday=overtime('Offsite');verify([holiday],[],[],[date]);
 assert.equal(holiday.records[date].statusDocumentNeedsReview,false);
 for(const weekendDate of ['2026-11-07','2026-11-08']){
  for(const status of ['Offsite','off-site','tidak_hadir','Tidak Hadir']){
   const weekend={id:'pegawai-1',records:{[weekendDate]:{status,overtimeHours:3,originalOvertimeHours:3,manualOverride:true}}};
   const result=verify([weekend]);
   const rec=weekend.records[weekendDate];
   assert.equal(rec.statusDocumentNeedsReview,false,status);
   assert.equal(rec.statusWeekendIneligible,true,status);
   assert.equal(rec.documentBlocksOvertime,true,status);
   assert.equal(rec.originalOvertimeHours,3,status);
   assert.equal(rec.overtimeHours,0,status);
   assert.equal(rec.manualOverride,false,status);
   assert.match(rec.documentBlockNote,/Sabtu\/Minggu/);
   assert.equal(result.held,1);
   assert.equal(result.missingStatus,0);
  }
 }
});

test('Lembur: saat tanggal merah diubah, Perlu Cek ikut dihitung ulang tanpa membiarkan nilai manual melewati pembatas',()=>{
 const e=overtime('Offsite');verify([e]);
 const r=e.records[date];
 assert.equal(r.overtimeHours,0);
 LM.refreshStatusRestriction(r,date,[date]);
 assert.equal(r.statusDocumentNeedsReview,false);
 assert.equal(r.documentBlocksOvertime,false);
 LM.refreshStatusRestriction(r,date,[]);
 assert.equal(r.statusDocumentNeedsReview,true);
 assert.equal(r.documentBlocksOvertime,true);
 assert.equal(r.overtimeHours,0);
});

test('Antarmuka: peringatan, filter, dan proteksi penyesuaian tidak ditiadakan',()=>{
 const um=fs.readFileSync(path.join(main,'uang-makan-app.js'),'utf8');
 const lm=fs.readFileSync(path.join(main,'app.js'),'utf8');
 assert.match(um,/if\(day\.statusDocumentMissing\)\{alert\(/);
 assert.match(um,/Perlu Cek - Dokumen Status Presensi/);
 assert.match(um,/d\.statusDocumentMissing\?'Perlu Cek'/);
 assert.match(lm,/if\s*\(record\.documentBlocksOvertime && value !== 0\) return alert/);
 assert.match(lm,/Perlu Cek - Dokumen Status Presensi/);
 assert.match(lm,/record\.statusDocumentNeedsReview\?'Perlu Cek'/);
 assert.match(lm,/apply\(state\.employees,state\.period,state\.holidays\)/);
});

test('Uang Makan: pemeriksaan langsung tidak menerima Surat Tugas non-accountable sebagai pengesahan Offsite',()=>{
 const day=UM.classifyDay(statusRecord('Offsite'),[],[{...task,accountable:false}]);
 assert.equal(day.statusDocumentSupported,false);
 assert.equal(day.statusDocumentMissing,true);
 assert.equal(day.automatic,0);
 assert.equal(day.value,0);
});

test('Lembur: verifikasi opsional ST non-accountable tidak dapat membuka klaim Offsite tanpa bukti sah',()=>{
 const e=overtime('Offsite');
 verify([e],[],[{...task,accountable:false}]);
 const group=LM.reviewGroups([e])[0];
 assert.ok(group);
 LM.verifySource([e],group.key,group.signature,'allow');
 const r=e.records[date];
 assert.equal(r.statusDocumentNeedsReview,true);
 assert.equal(r.documentBlocksOvertime,true);
 assert.equal(r.overtimeHours,0);
});

test('Lembur: riwayat dan ekspor memiliki pembatas klaim historis serta refresh status',()=>{
 const app=fs.readFileSync(path.join(main,'app.js'),'utf8');
 assert.match(app,/const unsafeClaim=\(run\.employees\|\|\[\]\)\.some/);
 assert.match(app,/if\(unsafeClaim\)throw new Error/);
 assert.match(app,/Pilihan? Edit Data|Pilih Edit Data/);
 assert.match(app,/if\(window\.LemburDocumentIntegration\.requiresStatusDocument\(record\.status\)\)/);
 assert.match(app,/refreshStatusRestriction\(record,date,state\.holidays\)/);
});
