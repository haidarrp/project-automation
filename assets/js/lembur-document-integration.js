/* Cuti / Surat Tugas sebagai pembatas klaim lembur.
 * Jam dari absensi tetap menjadi referensi; jam final menjadi 0 jika ada
 * Cuti yang diakui atau Surat Tugas yang dipertanggungjawabkan.
 */
(function(root,factory){
  const api=factory(root);
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.LemburDocumentIntegration=api;
})(typeof window!=='undefined'?window:null,function(root){
  'use strict';
  function documentMatches(record,date,employeeId,policy){
    if(!record||!record.startDate||!record.endDate||date<record.startDate||date>record.endDate)return false;
    if(record.kind==='cuti')return String(record.employeeId)===String(employeeId)&&policy.allows('cuti',record.id,employeeId);
    return (record.employeeIds||[]).includes(String(employeeId))&&policy.allows('surat',record.id,employeeId);
  }
  function describe(doc){
    return doc.kind==='cuti'
      ? `Cuti ${doc.leaveType||''} (${doc.documentNumber||doc.id})`
      : `Surat Tugas ${doc.letterNumber||doc.id}${doc.accountable===false?' (tidak dipertanggungjawabkan)':''}`;
  }
  function applyDocuments(employees,leaves,assignments,policy,resolveId){
    let held=0,pending=0,linked=0;
    for(const employee of employees){
      const id=String(resolveId(employee)||'');
      if(!id)throw new Error(`Identitas pegawai ${employee.name||employee.nip||'-'} tidak ditemukan pada Master Pegawai. Klaim lembur tidak dapat diverifikasi.`);
      const docs=[...leaves.map(d=>({...d,kind:'cuti'})),...assignments.map(d=>({...d,kind:'surat'}))];
      for(const [date,record] of Object.entries(employee.records||{})){
        const matches=docs.filter(d=>documentMatches(d,date,id,policy));
        const blocked=matches.some(d=>d.kind==='cuti'||d.accountable!==false);
        const signature=JSON.stringify(matches.map(d=>[d.kind,d.id,d.startDate,d.endDate,d.accountable!==false,d.documentNumber||d.letterNumber||'']).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))));
        const beforeBlocked=Boolean(record.documentBlocksOvertime);
        // Pertahankan jam hasil absensi ketika memblokir klaim pada data lama.
        if(record.originalOvertimeHours===undefined||record.originalOvertimeHours===null){
          record.originalOvertimeHours=Number(record.overtimeHours||0);
        }
        record.documentSources=matches.map(d=>({id:d.id,kind:d.kind,label:describe(d),accountable:d.accountable!==false}));
        record.documentBlocksOvertime=blocked;
        record.documentSignature=matches.length?signature:'';
        record.documentReviewPending=matches.length>0&&record.documentVerifiedSignature!==signature;
        if(blocked){
          if(Number(record.overtimeHours||0)!==0||!beforeBlocked){record.overtimeHours=0;record.manualOverride=false;}
          record.documentBlockNote='Klaim lembur tidak diperkenankan karena ada Cuti yang berlaku atau Surat Tugas yang dipertanggungjawabkan, meskipun presensi tercatat WFO.';
          held++;
        }else{
          if(beforeBlocked){record.overtimeHours=Number(record.originalOvertimeHours||0);record.manualOverride=false;}
          record.documentBlockNote='';
        }
        if(matches.length)linked++;
        if(record.documentReviewPending)pending++;
      }
    }
    return {held,pending,linked};
  }
  function periodYears(period){return [...new Set([period.year-1,period.year])];}
  async function fetchDocuments(period){
    const db=root?.FirebaseClient?.getDb?.();
    if(!db)throw new Error('Firestore belum tersedia untuk memeriksa Cuti dan Surat Tugas sebelum klaim lembur.');
    const years=periodYears(period);
    const [cuti,st,policy]=await Promise.all([
      Promise.all(years.map(y=>db.collection('leaveRecords').where('leaveYear','==',y).get())),
      Promise.all(years.map(y=>db.collection('assignmentRecords').where('assignmentYear','==',y).get())),
      root.AttendanceDocumentPolicy.fetch(years)
    ]);
    if(policy.stale.length)throw new Error('Keputusan Register mengenai Cuti/Surat Tugas berubah. Verifikasi ulang Register sebelum memproses lembur.');
    const leaves=new Map(),assignments=new Map();
    cuti.flatMap(s=>s.docs).forEach(d=>{const x=d.data()||{};
      leaves.set(d.id,{id:d.id,employeeId:String(x.employeeId||''),startDate:String(x.startDate||''),endDate:String(x.endDate||x.startDate||''),leaveType:String(x.leaveType||'Cuti'),documentNumber:String(x.documentNumber||'')});
    });
    st.flatMap(s=>s.docs).forEach(d=>{const x=d.data()||{};
      assignments.set(d.id,{id:d.id,employeeIds:(x.employees||[]).map(e=>String(e.employeeId||e.id||'')).filter(Boolean),startDate:String(x.startDate||''),endDate:String(x.endDate||x.startDate||''),accountable:x.accountable!==false,letterNumber:String(x.letterNumber||'')});
    });
    return {leaves:[...leaves.values()],assignments:[...assignments.values()],policy};
  }
  async function apply(employees,period){
    const documents=await fetchDocuments(period);
    const ids=new Map();
    for(const e of employees){
      if(e.masterEmployeeId){ids.set(e,String(e.masterEmployeeId));continue;}
      const found=await root?.MasterDataService?.findByIdentity?.({name:e.name||'',nip:e.nip||''});
      if(!found?.id)throw new Error(`Master Pegawai tidak ditemukan untuk ${e.name||e.nip||'-'}.`);
      e.masterEmployeeId=String(found.id);
      ids.set(e,String(found.id));
    }
    return applyDocuments(employees,documents.leaves,documents.assignments,documents.policy,e=>ids.get(e));
  }
  function remaining(employees){return (employees||[]).reduce((count,e)=>count+Object.values(e.records||{}).filter(r=>r.documentReviewPending).length,0);}
  return Object.freeze({apply,applyDocuments,remaining});
});
