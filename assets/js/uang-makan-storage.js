/* Riwayat Uang Makan: Firestore (metadata parent, rekap per pegawai subcollection). */
(function(){
 'use strict';
 const COLLECTION='uangMakanRuns';
 function db(){const value=window.FirebaseClient?.getDb?.();if(!value)throw new Error('Firestore tidak tersedia.');return value;}
 function actor(){const u=window.FirebaseClient?.getCurrentUser?.();if(!u?.uid)throw new Error('Masuk terlebih dahulu.');return {uid:u.uid,email:u.email||'',name:u.displayName||''};}
 function sanitizeDocuments(raw){
   return (Array.isArray(raw)?raw:[]).map(doc=>{
     const d=doc&&typeof doc==='object'?doc:{},f=d.file&&typeof d.file==='object'?d.file:{};
     return {id:String(d.id||''),name:String(d.name||''),type:String(d.type||''),
       number:String(d.number||''),note:String(d.note||''),file:{name:String(f.name||''),remoteName:String(f.remoteName||''),
         driveId:String(f.driveId||''),itemId:String(f.itemId||''),webUrl:String(f.webUrl||'')}};
   });
 }
 function sanitizeEmployee(e){
   // Simpan hanya referensi SharePoint, bukan binary/File/Blob atau URL preview sementara.
   const sourceFiles=(Array.isArray(e.sourceFiles)?e.sourceFiles:[]).map(source=>({
     name:String(source.name||''),size:Number(source.size||0),type:String(source.type||''),
     lastModified:Number(source.lastModified||0),driveId:String(source.driveId||''),
     itemId:String(source.itemId||''),webUrl:String(source.webUrl||''),remoteName:String(source.remoteName||'')
   }));
   return {id:e.id,name:e.name,nip:e.nip,anakSatker:e.anakSatker||'',jenisPegawai:e.jenisPegawai||'',order:e.order||9999,
     supportReviews:Object.fromEntries(Object.entries(e.supportReviews||{}).filter(([key,review])=>
      key&&review&&['auto',0,1].includes(review.choice)&&review.signature).map(([key,review])=>
      [key,{choice:review.choice,signature:String(review.signature),reviewedAt:String(review.reviewedAt||''),
        reviewedBy:String(review.reviewedBy||'')}])),
     files:Array.isArray(e.files)?e.files:[],sourceFiles,fileWarnings:e.fileWarnings||[],coverageAcknowledged:Boolean(e.coverageAcknowledged),coverageReason:e.coverageReason||'',missingAcknowledged:Boolean(e.missingAcknowledged),missingReason:e.missingReason||'',
     days:Object.fromEntries(Object.entries(e.days||{}).map(([key,d])=>[key,{
       key,inMinutes:d.inMinutes??null,outMinutes:d.outMinutes??null,status:d.status||'',workMode:d.workMode||'',flags:d.flags||[],duplicate:Boolean(d.duplicate),isHoliday:Boolean(d.isHoliday),
       leaves:(d.leaves||[]).map(x=>({id:x.id,leaveType:x.leaveType,documentNumber:x.documentNumber||'',documentDate:x.documentDate||'',
         startDate:x.startDate,endDate:x.endDate,note:x.note||'',documents:sanitizeDocuments(x.documents)})),
       previewLeaves:(d.previewLeaves||d.leaves||[]).map(x=>({id:x.id,employeeId:x.employeeId||e.id,leaveType:x.leaveType,documentNumber:x.documentNumber||'',documentDate:x.documentDate||'',
         startDate:x.startDate,endDate:x.endDate,note:x.note||'',documents:sanitizeDocuments(x.documents)})),
       previewAssignments:(d.previewAssignments||d.assignments||[]).map(x=>({id:x.id,accountable:x.accountable!==false,assignmentType:x.assignmentType,activity:x.activity,letterNumber:x.letterNumber||'',
         letterDate:x.letterDate||'',location:x.location||'',note:x.note||'',startDate:x.startDate,endDate:x.endDate,
         employeeIds:x.employeeIds,documents:sanitizeDocuments(x.documents)})),
       assignments:(d.assignments||[]).map(x=>({id:x.id,assignmentType:x.assignmentType,activity:x.activity,letterNumber:x.letterNumber||'',
         letterDate:x.letterDate||'',location:x.location||'',note:x.note||'',startDate:x.startDate,endDate:x.endDate,
         employeeIds:x.employeeIds,documents:sanitizeDocuments(x.documents)})),
       manualValue:d.manualValue??null,manualNote:d.manualNote||'',reviewed:Boolean(d.reviewed),reviewedAt:d.reviewedAt||'',reviewedBy:d.reviewedBy||'',
       // Catatan manual per tanggal hanya sah jika tidak ada dokumen sumber.
       dateManualValue:(d.leaves?.length||d.assignments?.length)?null:(d.dateManualValue??null),
       dateReviewed:(d.leaves?.length||d.assignments?.length)?false:Boolean(d.dateReviewed),
       dateReviewedAt:(d.leaves?.length||d.assignments?.length)?'':(d.dateReviewedAt||''),
       dateReviewedBy:(d.leaves?.length||d.assignments?.length)?'':(d.dateReviewedBy||'')
     }]))};
 }
 async function list(){
   const snap=await db().collection(COLLECTION).orderBy('createdAt','desc').limit(75).get();
   return snap.docs.map(x=>({id:x.id,...x.data()})).filter(x=>x.status==='complete');
 }
 async function save(year,month,employees,summary,settings={}){
   const a=actor(),ref=db().collection(COLLECTION).doc();
   const meta={year,month,period:`${year}-${String(month).padStart(2,'0')}`,createdAt:new Date().toISOString(),
     ownerUid:a.uid,ownerEmail:a.email,ownerName:a.name,status:'writing',summary,
     holidays:Array.isArray(settings.holidays)?settings.holidays.slice():[],
     holidayNotes:settings.holidayNotes||{},archiveKey:String(settings.archiveKey||''),schemaVersion:8};
   await ref.set(meta);
   try{
     const stored=employees.map(sanitizeEmployee);
     for(let i=0;i<stored.length;i+=100){
       const batch=db().batch();stored.slice(i,i+100).forEach(e=>batch.set(ref.collection('employees').doc(e.id),e));
       await batch.commit();
     }
     await ref.update({status:'complete'});
     return {id:ref.id,...meta,status:'complete'};
   }catch(e){throw new Error('Penyimpanan belum lengkap. Hasil belum menjadi riwayat final: '+String(e.message||e));}
 }
 async function get(id){
   const ref=db().collection(COLLECTION).doc(String(id||''));const [snapshot,children]=await Promise.all([ref.get(),ref.collection('employees').get()]);
   if(!snapshot.exists||snapshot.data()?.status!=='complete')throw new Error('Riwayat tidak tersedia atau belum lengkap.');
   const meta={id:snapshot.id,...snapshot.data()};
   const employees=children.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(Number(a.order||9999)-Number(b.order||9999))||a.name.localeCompare(b.name,'id'));
   return {meta,employees};
 }
 async function remove(id){
   const ref=db().collection(COLLECTION).doc(String(id||''));
   const result=await ref.get();if(!result.exists)throw new Error('Riwayat tidak ditemukan.');
   const a=actor(),meta=result.data();
   if(meta.ownerUid!==a.uid&&!window.FirebaseClient.isAdmin())throw new Error('Hanya pembuat atau administrator dapat menghapus riwayat.');
   const children=await ref.collection('employees').get();
   for(let i=0;i<children.docs.length;i+=200){
     const batch=db().batch();children.docs.slice(i,i+200).forEach(d=>batch.delete(d.ref));await batch.commit();
   }
   await ref.delete();
 }
 window.UangMakanStorage=Object.freeze({list,save,get,remove});
})();
