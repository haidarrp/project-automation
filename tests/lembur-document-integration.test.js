'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const Integration=require('../assets/js/lembur-document-integration.js');
const Policy=require('../assets/js/attendance-document-policy.js');
const dates=['2026-09-14','2026-09-15'];
const base=(id='p1',records=dates)=>({id,name:`Pegawai ${id}`,records:Object.fromEntries(records.map(date=>[date,{key:date,status:'WFO',originalOvertimeHours:3,overtimeHours:3}]))});
const leave={id:'c1',employeeId:'p1',startDate:'2026-09-14',endDate:'2026-09-14',leaveType:'Cuti Tahunan',documentNumber:'C1'};
const assignment=(accountable=true,startDate='2026-09-15',endDate='2026-09-15',employeeIds=['p1'])=>({id:'st1',employeeIds,startDate,endDate,letterNumber:'ST1',accountable});
function policy(leaves,assignments){return Policy.build(leaves,assignments.map(a=>({...a,employees:a.employeeIds.map(employeeId=>({employeeId}))})),[]);}
function sync(employees,leaves,assignments){return Integration.applyDocuments(employees,leaves,assignments,policy(leaves,assignments),e=>e.id);}
function verify(employees,key,choice){const g=Integration.reviewGroups(employees).find(x=>x.key===key);assert.ok(g);return Integration.verifySource(employees,key,g.signature,choice,{displayName:'Operator'});}

test('Cuti berlaku dan Surat Tugas dipertanggungjawabkan membatasi klaim walau WFO',()=>{
 const emp=base(),st=assignment();
 const result=sync([emp],[leave],[st]);
 assert.equal(result.held,2);assert.equal(result.pending,2);
 assert.equal(emp.records[dates[0]].overtimeHours,0);
 assert.equal(emp.records[dates[1]].overtimeHours,0);
 assert.equal(emp.records[dates[0]].originalOvertimeHours,3);
 assert.equal(emp.records[dates[1]].originalOvertimeHours,3);
 assert.equal(Integration.reviewGroups([emp]).length,2);
});

test('Surat Tugas tidak dipertanggungjawabkan memerlukan verifikasi namun awalnya tidak membatasi jam lembur',()=>{
 const emp=base(),st=assignment(false);
 const result=sync([emp],[],[st]);
 assert.equal(result.held,0);assert.equal(result.pending,1);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
 assert.equal(emp.records[dates[1]].documentSources[0].accountable,false);
});

test('Satu verifikasi per Surat Tugas berlaku bagi seluruh tanggal dan pegawai dalam proses',()=>{
 const st=assignment(false,'2026-09-14','2026-09-15',['p1','p2']);
 const e1=base('p1'),e2=base('p2');
 sync([e1,e2],[],[st]);
 const g=Integration.reviewGroups([e1,e2]);
 assert.equal(g.length,1);assert.equal(g[0].records.length,4);
 assert.equal(g[0].dates.length,2);assert.equal(g[0].employees.length,2);
 assert.equal(g[0].verified,false);
 const impacted=verify([e1,e2],'surat:st1','deny');
 assert.equal(impacted.records,4);
 for(const e of [e1,e2])for(const r of Object.values(e.records)){
  assert.equal(r.documentReviewPending,false);
  assert.equal(r.overtimeHours,0);
  assert.equal(r.originalOvertimeHours,3);
 }
 assert.equal(Integration.remaining([e1,e2]),0);
 assert.equal(Integration.reviewGroups([e1,e2])[0].verified,true);
 // Perubahan verifikasi sama berlaku pada semua tanggal dan pegawai.
 verify([e1,e2],'surat:st1','allow');
 for(const e of [e1,e2])for(const r of Object.values(e.records)){
  assert.equal(r.overtimeHours,3);
  assert.equal(r.documentReviewPending,false);
 }
});

test('Surat Tugas dipertanggungjawabkan dan Cuti valid tidak boleh disetujui untuk klaim lembur',()=>{
 const emp=base(),st=assignment(true);
 sync([emp],[leave],[st]);
 assert.throws(()=>verify([emp],'surat:st1','allow'),/tidak dapat diklaim/);
 assert.throws(()=>verify([emp],'cuti:c1','allow'),/tidak dapat diklaim/);
 assert.equal(Integration.remaining([emp]),2);
 verify([emp],'surat:st1','deny');
 assert.equal(Integration.remaining([emp]),1);
 verify([emp],'cuti:c1','deny');
 assert.equal(Integration.remaining([emp]),0);
 assert.equal(emp.records[dates[0]].overtimeHours,0);
 assert.equal(emp.records[dates[1]].overtimeHours,0);
});

test('Verifikasi tersimpan dipertahankan bila dokumen tetap sama dan ditandai ulang bila pertanggungjawaban berubah',()=>{
 const emp=base(),st=assignment();
 sync([emp],[],[st]);
 verify([emp],'surat:st1','deny');
 sync([emp],[],[st]);
 assert.equal(emp.records[dates[1]].documentReviewPending,false);
 const changed=assignment(false);
 sync([emp],[],[changed]);
 assert.equal(emp.records[dates[1]].documentReviewPending,true);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
 // Bukti verifikasi versi lama tidak berlaku untuk status dokumen baru.
 assert.equal(Integration.reviewGroups([emp])[0].verified,false);
 verify([emp],'surat:st1','allow');
 assert.equal(emp.records[dates[1]].documentReviewPending,false);
});

test('Perubahan rentang Surat Tugas membatalkan keputusan lama untuk semua tanggal baru',()=>{
 const emp=base(),st=assignment(false);
 sync([emp],[],[st]);verify([emp],'surat:st1','deny');
 sync([emp],[],[assignment(false,'2026-09-14','2026-09-15')]);
 assert.equal(emp.records[dates[0]].documentReviewPending,true);
 assert.equal(emp.records[dates[1]].documentReviewPending,true);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
});

test('Pada tanggal dengan dua dokumen, keputusan dokumen yang membatasi tetap mengunci 0 jam',()=>{
 const e=base(),st=assignment(false,'2026-09-14','2026-09-14');
 sync([e],[leave],[st]);
 verify([e],'surat:st1','allow');
 assert.equal(e.records[dates[0]].overtimeHours,0);
 assert.equal(e.records[dates[0]].documentReviewPending,true);
 verify([e],'cuti:c1','deny');
 assert.equal(e.records[dates[0]].documentReviewPending,false);
 assert.equal(e.records[dates[0]].overtimeHours,0);
});

test('Penghapusan dokumen memulihkan hasil presensi, tanpa menggunakan verifikasi lama',()=>{
 const emp=base(),st=assignment();
 sync([emp],[],[st]);verify([emp],'surat:st1','deny');
 sync([emp],[],[]);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
 assert.equal(emp.records[dates[1]].documentReviewPending,false);
 assert.equal(emp.records[dates[1]].documentBlocksOvertime,false);
 assert.equal(Integration.reviewGroups([emp]).length,0);
});

test('Verifikasi signature usang dan pilihan tidak dikenal ditolak',()=>{
 const emp=base(),st=assignment(false);
 sync([emp],[],[st]);
 assert.throws(()=>Integration.verifySource([emp],'surat:st1','signature-lama','allow'),/berubah/);
 assert.throws(()=>Integration.verifySource([emp],'surat:st1',Integration.reviewGroups([emp])[0].signature,'invalid'),/tidak valid/);
});

test('Klaim tidak bisa dihitung jika identitas tidak ditemukan',()=>{
 assert.throws(()=>Integration.applyDocuments([base()],[],[],policy([],[]),()=>''),/Identitas pegawai/);
});

test('Hasil presensi dari berkas lama tidak hilang saat pembatas dokumen kemudian dicabut',()=>{
 const emp=base(),st=assignment();
 delete emp.records[dates[1]].originalOvertimeHours;
 sync([emp],[],[st]);
 assert.equal(emp.records[dates[1]].originalOvertimeHours,3);
 assert.equal(emp.records[dates[1]].overtimeHours,0);
 sync([emp],[],[]);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
});
