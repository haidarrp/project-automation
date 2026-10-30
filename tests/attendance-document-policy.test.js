'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
const path=require('node:path');
const E=require('../assets/js/register-engine.js');
const P=require('../assets/js/attendance-document-policy.js');
const dir=path.resolve(__dirname,'../assets/js');
const leave={id:'leaveA',employeeId:'emp1',leaveYear:2026,startDate:'2026-09-14',endDate:'2026-09-15',documentNumber:'C-1',leaveType:'Cuti'};
const st={id:'taskA',assignmentYear:2026,startDate:'2026-09-14',endDate:'2026-09-15',letterNumber:'ST-1',assignmentType:'official_travel',employees:[{employeeId:'emp1'},{employeeId:'emp2'}]};
const source=E.makeSources([leave],[st]);
const conflict=E.conflicts(source)[0];
const keyType=type=>source.find(x=>x.type===type&&x.employeeId==='emp1').key;
const verified=selected=>({type:'conflict',workflow:'verified',key:conflict.key,fingerprint:conflict.fingerprint,selectedSourceKeys:selected,year:2026});
test('pilih hanya Surat Tugas mengeliminasi Cuti untuk pegawai terkait, bukan pegawai lain',()=>{
 const policy=P.build([leave],[st],[verified([keyType('surat')])]);
 assert.equal(policy.exportable('cuti',leave,'emp1'),false);
 assert.equal(policy.exportable('surat',st,'emp1'),true);
 assert.equal(policy.exportable('surat',st,'emp2'),true);
 assert.match(policy.reason('cuti',leave,'emp1'),/Dieliminasi/);
});
test('pilih semua mengikutkan kedua dokumen, tidak memaksa penyesuaian hasil konflik',()=>{
 const policy=P.build([leave],[st],[verified([keyType('cuti'),keyType('surat')])]);
 assert.equal(policy.exportable('cuti',leave,'emp1'),true);
 assert.equal(policy.exportable('surat',st,'emp1'),true);
 assert.equal(policy.excluded.size,0);
});
test('sebelum verifikasi tidak mengeliminasi bukti secara diam-diam',()=>{
 const policy=P.build([leave],[st],[{...verified([keyType('surat')]),workflow:'pending'}]);
 assert.equal(policy.exportable('cuti',leave,'emp1'),true);
 assert.equal(policy.exportable('surat',st,'emp1'),true);
});
test('dokumen yang diedit memerlukan verifikasi ulang dan keputusan lama tidak dipakai',()=>{
 const changed={...leave,endDate:'2026-09-16'};
 const policy=P.build([changed],[st],[verified([keyType('surat')])]);
 assert.equal(policy.stale.length,1);
 assert.equal(policy.exportable('cuti',changed,'emp1'),true);
});
test('ketika irisan telah hilang, pilihan historis tidak mengeliminasi dokumen yang kini berdiri sendiri',()=>{
 const moved={...leave,startDate:'2026-09-10',endDate:'2026-09-10'};
 const policy=P.build([moved],[st],[verified([keyType('surat')])]);
 assert.equal(policy.stale.length,0);
 assert.equal(policy.exportable('cuti',moved,'emp1'),true);
});
test('Surat Tugas tidak dipertanggungjawabkan selalu hanya preview, bisa diaktifkan kembali',()=>{
 const disabled={...st,accountable:false};
 const first=P.build([],[disabled],[]);
 assert.equal(first.exportable('surat',disabled,'emp1'),false);
 assert.match(first.reason('surat',disabled,'emp1'),/tidak dipertanggungjawabkan/);
 const enabled={...disabled,accountable:true};
 const next=P.build([],[enabled],[]);
 assert.equal(next.exportable('surat',enabled,'emp1'),true);
 assert.notEqual(first.signature,next.signature);
});
function attendanceTukinHarness(leaves,assignments,decisions=[]){
 const files={leaveRecords:leaves,assignmentRecords:assignments};
 const db={collection(name){return {where(){return {get:async()=>({docs:(files[name]||[]).map(d=>({id:d.id,data:()=>d}))})};}};}};
 const rules={dateKey:d=>typeof d==='string'?d.slice(0,10):d.toISOString().slice(0,10),attendancePeriod:()=>({start:new Date(2026,8,1),end:new Date(2026,8,30)})};
 const win={TukinRules:rules,MasterDataService:{findByIdentity:async()=>({id:'emp1'})},FirebaseClient:{getDb:()=>db},
  AttendanceDocumentPolicy:{fetch:async()=>P.build(leaves,assignments,decisions)}};
 vm.runInNewContext(fs.readFileSync(path.join(dir,'tukin-supporting-data-integration.js'),'utf8'),{window:win,console});
 const entry={key:'2026-09-14',needsVerification:true,adjustedPercent:null,evidence:[]};
 const emp={name:'Pegawai Contoh',nip:'111',masterEmployeeId:'emp1',records:{'2026-09-14':entry}};
 return {module:win.TukinSupportingDataIntegration,employee:emp,entry};
}
test('Tukin tidak otomatis 0% untuk Surat Tugas tanpa pertanggungjawaban dan tetap menyimpan preview',async()=>{
 const doc={id:'file1',file:{driveId:'d',itemId:'x',name:'ST.pdf',webUrl:'https://sharepoint.example/st'}};
 const h=attendanceTukinHarness([],[{...st,accountable:false,documents:[doc]}]);
 const res=await h.module.apply([h.employee],{year:2026,month:9});
 assert.equal(res.assignmentAdjustedRecords,0);
 assert.equal(h.entry.adjustedPercent,null);
 assert.equal(h.entry.evidence.length,0);
 assert.equal(h.entry.nonExportDocuments[0].fileUrl,'https://sharepoint.example/st');
});
test('Tukin konflik yang keduanya dipilih tetap memuat kedua bukti namun tanpa auto 0%',async()=>{
 const a={...leave,documents:[{id:'c1',file:{driveId:'d',itemId:'cuti',name:'Cuti.pdf'}}]};
 const b={...st,documents:[{id:'st1',file:{driveId:'d',itemId:'st',name:'ST.pdf'}}]};
 const src=E.makeSources([a],[b]);const c=E.conflicts(src)[0];
 const decision={...verified(src.filter(x=>x.employeeId==='emp1').map(x=>x.key)),key:c.key,fingerprint:c.fingerprint};
 const h=attendanceTukinHarness([a],[b],[decision]);
 const res=await h.module.apply([h.employee],{year:2026,month:9});
 assert.equal(res.conflictRecords,1);
 assert.equal(h.entry.adjustedPercent,null);
 assert.equal(h.entry.evidence.length,2);
});

test('Tukin mendeteksi presensi WFO dengan Cuti atau Surat Tugas dipertanggungjawabkan untuk verifikasi klaim lembur',async()=>{
 for(const inputs of [
  {leaves:[leave],assignments:[]},
  {leaves:[],assignments:[{...st,accountable:true}]}
 ]){
  const h=attendanceTukinHarness(inputs.leaves,inputs.assignments);
  h.entry.status='WFO';h.entry.inMinutes=480;h.entry.outMinutes=1020;
  await h.module.apply([h.employee],{year:2026,month:9});
  assert.equal(h.entry.overtimeDocumentRestriction,true);
  assert.match(h.entry.overtimeDocumentReviewNote,/tidak boleh diklaim sebagai lembur/);
 }
});

test('Surat Tugas tidak dipertanggungjawabkan tidak memicu pembatas klaim lembur di Tukin',async()=>{
 const h=attendanceTukinHarness([],[{...st,accountable:false}]);
 h.entry.status='WFO';h.entry.inMinutes=480;h.entry.outMinutes=1020;
 await h.module.apply([h.employee],{year:2026,month:9});
 assert.equal(h.entry.overtimeDocumentRestriction,false);
 assert.equal(h.entry.overtimeDocumentReviewNote,'');
});
