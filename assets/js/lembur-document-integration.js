/* Verifikasi Cuti / Surat Tugas untuk proses Lembur.
 * Satu keputusan per dokumen berlaku untuk semua tanggal dan pegawai dalam
 * proses yang sama. Jam presensi asli tidak pernah diubah oleh keputusan ini.
 */
(function(root,factory){
  const api=factory(root);
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.LemburDocumentIntegration=api;
})(typeof window!=='undefined'?window:null,function(root){
  'use strict';
  const sourceKey=doc=>`${doc.kind}:${doc.id}`;
  function signature(doc){
    return JSON.stringify([
      doc.kind,String(doc.id),String(doc.startDate),String(doc.endDate),
      doc.kind==='cuti'?String(doc.employeeId):[...(doc.employeeIds||[])].map(String).sort(),
      doc.kind==='cuti'?String(doc.leaveType||''):doc.accountable!==false,
      String(doc.documentNumber||doc.letterNumber||'')
    ]);
  }
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
  function toSource(doc){
    return {id:String(doc.id),key:sourceKey(doc),kind:doc.kind,label:describe(doc),
      signature:signature(doc),startDate:doc.startDate,endDate:doc.endDate,
      accountable:doc.accountable!==false,blockByDefault:doc.kind==='cuti'||doc.accountable!==false,
      documents:(doc.documents||[]).map((item,index)=>({id:String(item.id||`dokumen-${index}`),
        name:String(item.name||''),type:String(item.type||''),
        file:item.file?{name:String(item.file.name||''),remoteName:String(item.file.remoteName||''),
          driveId:String(item.file.driveId||''),itemId:String(item.file.itemId||''),webUrl:String(item.file.webUrl||'')}:null}))};
  }
  function validReview(record,source){
    const review=record.documentReviews?.[source.key];
    return review&&review.signature===source.signature&&['allow','deny'].includes(review.choice)?review:null;
  }
  function recalculateRecord(record){
    const sources=record.documentSources||[];
    const previousBlock=Boolean(record.documentBlocksOvertime);
    const blocked=sources.some(source=>source.blockByDefault||validReview(record,source)?.choice==='deny');
    record.documentReviewPending=sources.some(source=>!validReview(record,source));
    record.documentBlocksOvertime=blocked;
    record.documentSignature=sources.length?JSON.stringify(sources.map(s=>[s.key,s.signature]).sort((a,b)=>a[0].localeCompare(b[0]))):'';
    if(record.originalOvertimeHours===undefined||record.originalOvertimeHours===null){
      record.originalOvertimeHours=Number(record.overtimeHours||0);
    }
    if(blocked){
      if(Number(record.overtimeHours||0)!==0||!previousBlock){record.overtimeHours=0;record.manualOverride=false;}
      record.documentBlockNote=sources.some(s=>s.blockByDefault)
        ?'Klaim lembur tidak diperkenankan karena ada Cuti yang berlaku atau Surat Tugas yang dipertanggungjawabkan, meskipun presensi tercatat WFO.'
        :'Klaim lembur tidak diberikan berdasarkan penyesuaian dokumen per periode.';
    }else{
      if(previousBlock){record.overtimeHours=Number(record.originalOvertimeHours||0);record.manualOverride=false;}
      record.documentBlockNote='';
    }
    return blocked;
  }
  function applyDocuments(employees,leaves,assignments,policy,resolveId){
    let held=0,pending=0,linked=0;
    const docs=[...leaves.map(d=>({...d,kind:'cuti'})),...assignments.map(d=>({...d,kind:'surat'}))];
    for(const employee of employees){
      const id=String(resolveId(employee)||'');
      if(!id)throw new Error(`Identitas pegawai ${employee.name||employee.nip||'-'} tidak ditemukan pada Master Pegawai. Klaim lembur tidak dapat diverifikasi.`);
      for(const [date,record] of Object.entries(employee.records||{})){
        const matches=docs.filter(d=>documentMatches(d,date,id,policy));
        record.documentSources=matches.map(toSource);
        if(recalculateRecord(record))held++;
        if(matches.length)linked++;
        if(record.documentReviewPending)pending++;
      }
    }
    return {held,pending,linked};
  }
  function reviewGroups(employees){
    const groups=new Map();
    for(const employee of employees||[]){
      for(const [date,record] of Object.entries(employee.records||{})){
        for(const source of record.documentSources||[]){
          const key=source.key;
          if(!groups.has(key))groups.set(key,{...source,dates:new Set(),employees:new Set(),records:[],reviewed:0,pending:0});
          const group=groups.get(key);
          // Signature is unique to a current source. The same source must not
          // be applied across dissimilar versions within one processing batch.
          if(group.signature!==source.signature)throw new Error(`Data ${source.label} tidak konsisten; sinkronkan ulang dokumen.`);
          group.dates.add(date);
          group.employees.add(String(employee.name||employee.nip||'-'));
          group.records.push({employee,date,record});
          if(validReview(record,source))group.reviewed++;
          else group.pending++;
        }
      }
    }
    return [...groups.values()].sort((a,b)=>a.startDate.localeCompare(b.startDate)||a.key.localeCompare(b.key))
      .map(g=>{
        const savedChoices=new Set(g.records.map(({record})=>validReview(record,g)?.choice).filter(Boolean));
        const verified=g.pending===0&&savedChoices.size===1;
        return {...g,dates:[...g.dates].sort(),employees:[...g.employees].sort(),verified,
          choice:verified?[...savedChoices][0]:(g.blockByDefault?'deny':'allow')};
      });
  }
  function verifySource(employees,key,expectedSignature,choice,actor={}){
    if(!['allow','deny'].includes(choice))throw new Error('Pilihan klaim lembur tidak valid.');
    const group=reviewGroups(employees).find(g=>g.key===key);
    if(!group||group.signature!==expectedSignature)throw new Error('Dokumen berubah. Muat ulang hasil proses sebelum memverifikasi.');
    if(choice==='allow'&&group.blockByDefault){
      throw new Error('Cuti yang berlaku atau Surat Tugas yang dipertanggungjawabkan tidak dapat diklaim untuk lembur. Ubah status dokumen pada modul sumber terlebih dahulu.');
    }
    const now=new Date().toISOString();
    const review={signature:group.signature,choice,reviewedAt:now,
      reviewedBy:String(actor.displayName||actor.email||'Operator'),reviewedByUid:String(actor.uid||'')};
    for(const {record} of group.records){
      record.documentReviews={...(record.documentReviews||{}),[key]:{...review}};
      recalculateRecord(record);
    }
    return {records:group.records.length,employees:group.employees.length,dates:group.dates.length};
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
      leaves.set(d.id,{id:d.id,employeeId:String(x.employeeId||''),startDate:String(x.startDate||''),endDate:String(x.endDate||x.startDate||''),leaveType:String(x.leaveType||'Cuti'),documentNumber:String(x.documentNumber||''),documents:Array.isArray(x.documents)?x.documents:[]});
    });
    st.flatMap(s=>s.docs).forEach(d=>{const x=d.data()||{};
      assignments.set(d.id,{id:d.id,employeeIds:(x.employees||[]).map(e=>String(e.employeeId||e.id||'')).filter(Boolean),startDate:String(x.startDate||''),endDate:String(x.endDate||x.startDate||''),accountable:x.accountable!==false,letterNumber:String(x.letterNumber||''),documents:Array.isArray(x.documents)?x.documents:[]});
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
  function remaining(employees){
    const pendingDays=(employees||[]).reduce((count,e)=>count+Object.values(e.records||{}).filter(r=>r.documentReviewPending).length,0);
    // Lindungi arsip yang berisi keputusan berbeda untuk satu dokumen yang sama.
    return Math.max(pendingDays,reviewGroups(employees).filter(group=>!group.verified).length);
  }
  return Object.freeze({apply,applyDocuments,remaining,reviewGroups,verifySource});
});
