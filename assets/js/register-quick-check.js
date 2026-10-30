/* Pemeriksaan non-blocking pasca-input Cuti/Surat Tugas; tidak mengubah dokumen. */
(function () {
 'use strict';
 const E=window.AttendanceRegisterEngine;
 async function checkSource(type,id,year) {
   if(!window.UserAccess?.canView?.('register','cases') || !window.FirebaseClient?.getDb?.() || !id)return 0;
   try {
     const years=[Number(year)-1,Number(year),Number(year)+1],d=window.FirebaseClient.getDb();
     const [leaveSnap,assignmentSnap]=await Promise.all([
       d.collection('leaveRecords').where('leaveYear','in',years).get(),
       d.collection('assignmentRecords').where('assignmentYear','in',years).get()
     ]);
     const sources=E.makeSources(leaveSnap.docs.map(x=>({...x.data(),id:x.id})),assignmentSnap.docs.map(x=>({...x.data(),id:x.id})));
     return E.conflicts(sources).filter(c=>c.sources.some(s=>s.type===type&&s.id===id)).length;
   } catch(error) {
     console.warn('Pemeriksaan Register pasca-input tidak berhasil, periksa kembali dari menu Register:',error);
     return -1;
   }
 }
 window.AttendanceQuickCheck=Object.freeze({checkSource});
})();
