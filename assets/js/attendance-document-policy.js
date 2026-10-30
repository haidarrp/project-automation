/* Aturan dokumen yang dipilih TU; shared Register / Tukin / Uang Makan. */
(function (root, factory) {
  const engine = root?.AttendanceRegisterEngine || (typeof require === 'function' ? require('./register-engine.js') : null);
  const api = factory(engine);
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceDocumentPolicy = api;
})(typeof window !== 'undefined' ? window : null, function(E) {
 'use strict';
 const key=(type,id,employeeId)=>`${type}:${id}:${employeeId}`;
 function build(leaves=[],assignments=[],cases=[]) {
  if (!E) throw new Error('Mesin Register belum tersedia.');
  const sources=E.makeSources(leaves,assignments);
  const issues=new Map(E.conflicts(sources).map(x=>[x.key,x]));
  const excluded=new Set(),stale=[];
  const currentAssignments=new Map(assignments.map(a=>[String(a.id),a]));
  for (const item of cases) {
   if (item.type!=='conflict'||item.workflow!=='verified'||!Array.isArray(item.selectedSourceKeys)) continue;
   const current=issues.get(item.key);
   if (!current) continue; // Konflik tidak lagi berlaku setelah dokumen diedit.
   if (current.fingerprint!==item.fingerprint) {stale.push(item.key);continue;}
   const valid=new Set(current.sources.map(s=>s.key));
   const chosen=new Set(item.selectedSourceKeys.filter(s=>valid.has(s)));
   if (!chosen.size) {stale.push(item.key);continue;}
   for(const source of current.sources) if(!chosen.has(source.key)) excluded.add(source.key);
  }
  const signature=E.hash(JSON.stringify([sources.map(s=>`${s.key}/${s.fingerprint}`).sort(),[...excluded].sort()]));
  return {
   excluded,stale,sources,signature,
   allows(type,id,employeeId){return !excluded.has(key(type,id,employeeId));},
   selectionReason(type,id,employeeId){return excluded.has(key(type,id,employeeId))?'Dieliminasi melalui verifikasi Register':'';},
   isAccountable(assignment){return (currentAssignments.get(String(assignment?.id))||assignment)?.accountable!==false;},
   exportable(type,record,employeeId){return this.allows(type,record?.id,employeeId)&&(type!=='surat'||this.isAccountable(record));},
   reason(type,record,employeeId){return this.selectionReason(type,record?.id,employeeId)||(type==='surat'&&!this.isAccountable(record)?'Surat Tugas tidak dipertanggungjawabkan':'');}
  };
 }
 async function fetch(years){
  const db=window.FirebaseClient?.getDb?.();if(!db)throw new Error('Firestore belum siap untuk memeriksa keputusan Register.');
  const y=[...new Set(years.map(Number).filter(Number.isInteger))];
  const batches=await Promise.all([
   Promise.all(y.map(year=>db.collection('leaveRecords').where('leaveYear','==',year).get())),
   Promise.all(y.map(year=>db.collection('assignmentRecords').where('assignmentYear','==',year).get())),
   Promise.all(y.map(year=>db.collection('attendanceCases').where('year','==',year).get()))
  ]);
  const unpack=b=>b.flatMap(s=>s.docs.map(d=>({...d.data(),id:d.id})));
  return build(unpack(batches[0]),unpack(batches[1]),unpack(batches[2]));
 }
 return Object.freeze({build,fetch});
});
