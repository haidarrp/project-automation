'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const Integration=require('../assets/js/lembur-document-integration.js');
const Policy=require('../assets/js/attendance-document-policy.js');
const dates=['2026-09-14','2026-09-15'];
const base=()=>({id:'p1',name:'Pegawai Contoh',records:Object.fromEntries(dates.map(date=>[date,{key:date,status:'WFO',originalOvertimeHours:3,overtimeHours:3}]))});
const leave={id:'c1',employeeId:'p1',startDate:'2026-09-14',endDate:'2026-09-14',leaveType:'Cuti Tahunan',documentNumber:'C1'};
const assignment=(accountable=true)=>({id:'st1',employeeIds:['p1'],startDate:'2026-09-15',endDate:'2026-09-15',letterNumber:'ST1',accountable});
function policy(leaves,assignments){return Policy.build(leaves,assignments.map(a=>({...a,employees:a.employeeIds.map(employeeId=>({employeeId}))})),[]);}
test('Cuti berlaku dan Surat Tugas dipertanggungjawabkan memblokir klaim lembur walau WFO',()=>{
 const emp=base(),st=assignment();
 const result=Integration.applyDocuments([emp],[leave],[st],policy([leave],[st]),e=>e.id);
 assert.equal(result.held,2);assert.equal(result.pending,2);
 assert.equal(emp.records[dates[0]].overtimeHours,0);
 assert.equal(emp.records[dates[1]].overtimeHours,0);
 assert.equal(emp.records[dates[0]].originalOvertimeHours,3);
 assert.equal(emp.records[dates[1]].originalOvertimeHours,3);
});
test('Surat Tugas tidak dipertanggungjawabkan hanya meminta verifikasi tanpa memblokir jam lembur',()=>{
 const emp=base(),st=assignment(false);
 const result=Integration.applyDocuments([emp],[],[st],policy([],[st]),e=>e.id);
 assert.equal(result.held,0);assert.equal(result.pending,1);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
 assert.equal(emp.records[dates[1]].documentSources[0].accountable,false);
});
test('Verifikasi tersimpan dipertahankan jika sumber tetap sama, dibatalkan jika pertanggungjawaban berubah',()=>{
 const emp=base(),st=assignment();
 Integration.applyDocuments([emp],[],[st],policy([],[st]),e=>e.id);
 emp.records[dates[1]].documentVerifiedSignature=emp.records[dates[1]].documentSignature;
 Integration.applyDocuments([emp],[],[st],policy([],[st]),e=>e.id);
 assert.equal(emp.records[dates[1]].documentReviewPending,false);
 const changed=assignment(false);
 Integration.applyDocuments([emp],[],[changed],policy([],[changed]),e=>e.id);
 assert.equal(emp.records[dates[1]].documentReviewPending,true);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
});
test('Penghapusan bukti pertanggungjawaban memulihkan hasil presensi, tanpa mengizinkan verifikasi lama',()=>{
 const emp=base(),st=assignment();
 Integration.applyDocuments([emp],[],[st],policy([],[st]),e=>e.id);
 Integration.applyDocuments([emp],[],[],policy([],[]),e=>e.id);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
 assert.equal(emp.records[dates[1]].documentReviewPending,false);
 assert.equal(emp.records[dates[1]].documentBlocksOvertime,false);
});
test('Klaim tidak bisa dihitung jika identitas tidak ditemukan',()=>{
 assert.throws(()=>Integration.applyDocuments([base()],[],[],policy([],[]),()=>''),/Identitas pegawai/);
});

test('Hasil presensi dari berkas lama tidak hilang saat pembatas dokumen kemudian dicabut',()=>{
 const emp=base(),st=assignment();
 delete emp.records[dates[1]].originalOvertimeHours;
 Integration.applyDocuments([emp],[],[st],policy([],[st]),e=>e.id);
 assert.equal(emp.records[dates[1]].originalOvertimeHours,3);
 assert.equal(emp.records[dates[1]].overtimeHours,0);
 Integration.applyDocuments([emp],[],[],policy([],[]),e=>e.id);
 assert.equal(emp.records[dates[1]].overtimeHours,3);
});
