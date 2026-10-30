/* Uang Makan: perhitungan berdasarkan bulan kalender.
 * Verifikasi satu dokumen Cuti/Surat Tugas berlaku untuk seluruh tanggal
 * dalam rentang dokumen yang termasuk periode pengolahan. Hari libur
 * tetap bernilai nol; WFA/WFH selain Jumat diberi catatan (penyesuaian opsional).
 */
(function () {
  'use strict';
  const MONTHS=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const ABSENCE_FLAGS=new Set(['S','I','TK','D','C']);
  const REMOTE_WORK=/(?:^|[^a-z0-9])(?:wfa|wfh|work[\s_-]*from[\s_-]*(?:anywhere|home)|bekerja[\s_-]*dari[\s_-]*(?:mana[\s_-]*saja|rumah))(?![a-z0-9])/i;
  function detectWorkMode(value){
    const match=String(value||'').match(REMOTE_WORK);
    if(!match)return '';
    return /wfa|anywhere|mana[\s_-]*saja/i.test(match[0])?'WFA':'WFH';
  }
  function pad(n){return String(n).padStart(2,'0');}
  function key(year,month,day){return `${year}-${pad(month)}-${pad(day)}`;}
  function periodDays(year,month){return new Date(Number(year),Number(month),0).getDate();}
  function periodLabel(year,month){return `${MONTHS[Number(month)-1]} ${year}`;}
  function keyDate(date){return key(date.getFullYear(),date.getMonth()+1,date.getDate());}
  function weekday(dateKey){const [y,m,d]=dateKey.split('-').map(Number);return new Date(y,m-1,d).getDay();}
  function weekend(dateKey){const day=weekday(dateKey);return day===0||day===6;}
  function normalizeNip(v){return String(v??'').replace(/\D/g,'');}
  function normalizeName(v){return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');}
  function monthDates(year,month){return Array.from({length:periodDays(year,month)},(_,i)=>key(year,month,i+1));}
  function isAbsenceStatus(status,flags){
    if((flags||[]).some(flag=>ABSENCE_FLAGS.has(String(flag).toUpperCase())))return true;
    return /(?:^|\b)(cuti|sakit|izin|ijin|dinas luar|perjalanan dinas|tanpa keterangan|tidak hadir|alpa)(?:\b|$)/i.test(String(status||''));
  }
  function supportKey(kind,record){
    const id=String(record?.id||'').trim();
    // Firestore selalu memberikan id, fallback ini menjaga data impor lama.
    return `${kind}:${id||[record?.startDate||'',record?.endDate||'',record?.documentNumber||record?.letterNumber||''].join('|')}`;
  }
  function supportSignature(kind,record){
    const x=record||{};
    const documents=(Array.isArray(x.documents)?x.documents:[]).map(d=>JSON.stringify([
      String(d.id||''),String(d.name||''),String(d.type||''),String(d.number||''),String(d.note||''),
      String(d.file?.name||''),String(d.file?.driveId||''),String(d.file?.itemId||'')
    ])).sort();
    return kind==='cuti'?JSON.stringify([
      String(x.id||''),String(x.employeeId||''),String(x.startDate||''),String(x.endDate||''),
      String(x.leaveType||''),String(x.documentNumber||''),String(x.documentDate||''),String(x.note||''),documents
    ]):JSON.stringify([
      String(x.id||''),(x.employeeIds||[]).map(String).sort(),String(x.startDate||''),String(x.endDate||''),
      String(x.assignmentType||''),String(x.accountable!==false),String(x.activity||''),String(x.letterNumber||''),
      String(x.letterDate||''),String(x.location||''),String(x.note||''),documents
    ]);
  }
  function classifyDay(record,leaves=[],assignments=[]){
    const dateKey=String(record?.key||'');
    const hasIn=record?.inMinutes!==null && record?.inMinutes!==undefined;
    const hasOut=record?.outMinutes!==null && record?.outMinutes!==undefined;
    const present=hasIn||hasOut;
    // Surat Tugas non-pertanggungjawaban tidak ikut proses Uang Makan.
    const activeAssignments=assignments.filter(a=>a.accountable!==false);
    const official=activeAssignments.filter(a=>a.assignmentType==='official_travel');
    const workArrangements=activeAssignments.filter(a=>a.assignmentType!=='official_travel');
    const status=String(record?.status||'').trim();
    const workMode=detectWorkMode(record?.workMode)||detectWorkMode(status);
    const absenceStatus=isAbsenceStatus(status,record?.flags);
    const isWeekend=weekend(dateKey);
    const isHoliday=Boolean(record?.isHoliday);
    const nonWorking=isWeekend||isHoliday;
    const remoteExcluded=Boolean(workMode)&&weekday(dateKey)!==5;
    const issues=[];
    if(!nonWorking){
      if(present && !(hasIn&&hasOut))issues.push('Presensi hanya satu kali');
      if(present && leaves.length)issues.push('Presensi bertentangan dengan data cuti');
      if(present && official.length)issues.push('Presensi bertentangan dengan dinas luar');
      if(present && absenceStatus)issues.push('Presensi bertentangan dengan status absensi');
      if(record?.duplicate)issues.push('Tanggal tercatat dalam lebih dari satu file');
      if(!present && /\bhadir\b/i.test(status) && !absenceStatus)issues.push('Status hadir tetapi jam presensi kosong');
      // Data sumber belum tentu mencerminkan realisasi: cuti/tugas dapat batal.
      if(leaves.length)issues.push('Konfirmasi pelaksanaan cuti (termasuk jika dibatalkan/tidak diambil)');
      if(activeAssignments.length)issues.push('Konfirmasi pelaksanaan surat tugas (termasuk jika batal/tidak berangkat)');
      // WFH/WFA di luar Jumat adalah rekomendasi nol, bukan keputusan final.
      // Pegawai mungkin memiliki surat sakit/cuti yang belum tersinkron.
      if(remoteExcluded)issues.push('WFH/WFA selain Jumat: periksa hasil aktual dan data sakit/cuti bila perlu');
    }
    // Surat Tugas yang dipertanggungjawabkan memberi rekomendasi awal 0
    // meskipun presensi WFO tercatat. Penyesuaian manual tetap opsional.
    const automatic=(present&&!nonWorking&&!remoteExcluded&&!activeAssignments.length)?1:0;
    const hasManual=record?.manualValue===0||record?.manualValue===1;
    const lockedZero=nonWorking;
    const manualValue=lockedZero?null:(hasManual?record.manualValue:null);
    const value=lockedZero?0:(manualValue===null?automatic:manualValue);
    return {key:dateKey,inMinutes:record?.inMinutes??null,outMinutes:record?.outMinutes??null,
      status,workMode,remoteExcluded,flags:Array.isArray(record?.flags)?record.flags:[],duplicate:Boolean(record?.duplicate),
      leaves,assignments:activeAssignments,workArrangements,official,issues,hasIn,hasOut,present,isWeekend,isHoliday,nonWorking,
      automatic,value,manualValue,manualNote:nonWorking?'':String(record?.manualNote||''),
      reviewed:nonWorking?false:Boolean(record?.reviewed),reviewedAt:nonWorking?'':String(record?.reviewedAt||''),
      reviewedBy:nonWorking?'':String(record?.reviewedBy||'')};
  }
  function applySupporting(employee,leaves,assignments,year,month,holidays=[]){
    const dates=monthDates(year,month),reds=new Set(holidays||[]);
    const ownLeaves=(leaves||[]).filter(l=>l.employeeId===employee.id);
    const ownAssignments=(assignments||[]).filter(a=>a.accountable!==false&&(a.employeeIds||[]).includes(employee.id));
    // Tinjau per identitas dokumen, bukan per tanggal. Tanda tangan metadata
    // mencegah persetujuan lama tetap berlaku bila isi/rentang berubah.
    const supportReviews={};
    for(const [kind,items] of [['cuti',ownLeaves],['tugas',ownAssignments]]){
      for(const item of items){
        const k=supportKey(kind,item),oldReview=employee.supportReviews?.[k];
        if(oldReview&&['auto',0,1].includes(oldReview.choice)&&oldReview.signature===supportSignature(kind,item)){
          supportReviews[k]={choice:oldReview.choice,signature:oldReview.signature,
            reviewedAt:String(oldReview.reviewedAt||''),reviewedBy:String(oldReview.reviewedBy||'')};
        }
      }
    }
    const old=employee.days||{},daily={};
    for(const date of dates){
      const base=old[date]||{key:date,inMinutes:null,outMinutes:null,status:'',flags:[]};
      const dayLeaves=ownLeaves.filter(l=>l.startDate<=date&&l.endDate>=date);
      const dayTasks=ownAssignments.filter(a=>a.startDate<=date&&a.endDate>=date);
      const hasSources=Boolean(dayLeaves.length||dayTasks.length);
      // Koreksi manual tanggal hanya berlaku untuk hari TANPA dokumen sumber
      // (misalnya WFH di luar Jumat). Abaikan koreksi khusus tanggal lama
      // apabila hari tersebut sudah terkait Cuti/Surat Tugas. Keputusan
      // dokumen harus seragam pada seluruh rentang/per peserta.
      const rawDateValue=hasSources?null:(base.dateManualValue!==undefined?base.dateManualValue:base.manualValue);
      const rawDateReviewed=!hasSources&&(base.dateReviewed!==undefined?base.dateReviewed:base.reviewed);
      const d=classifyDay({...base,key:date,isHoliday:reds.has(date),manualValue:rawDateValue,
        reviewed:rawDateReviewed},dayLeaves,dayTasks);
      d.dateManualValue=d.manualValue;
      d.dateReviewed=!d.nonWorking&&Boolean(rawDateReviewed);
      d.dateReviewedAt=d.dateReviewed?String(base.dateReviewedAt||base.reviewedAt||''):'';
      d.dateReviewedBy=d.dateReviewed?String(base.dateReviewedBy||base.reviewedBy||''):'';
      const sources=[...dayLeaves.map(x=>supportKey('cuti',x)),...dayTasks.map(x=>supportKey('tugas',x))];
      d.sourceKeys=sources;
      d.pendingSourceKeys=sources.filter(k=>!supportReviews[k]);
      d.reviewConflict=false;
      if(hasSources&&!d.nonWorking){
        const decisions=[...new Set(sources.map(k=>supportReviews[k]?.choice).filter(v=>v===0||v===1))];
        d.reviewConflict=decisions.length>1;
        const periodValue=decisions.length===1?decisions[0]:null;
        d.manualValue=periodValue;
        d.value=periodValue===null?d.automatic:periodValue;
        // Hasil dua dokumen yang bertentangan tidak bisa diselesaikan
        // memakai pengecualian tanggal; perbaiki sumber dan sinkronkan ulang.
        d.reviewed=d.pendingSourceKeys.length===0&&!d.reviewConflict;
        if(d.reviewConflict)d.issues.push('Hasil penyesuaian dokumen yang tumpang tindih berbeda; perbaiki dokumen koreksi pada modul Cuti/Surat Tugas dan sinkronkan ulang');
        if(d.reviewed){
          const latest=sources.map(k=>supportReviews[k]).filter(Boolean).sort((a,b)=>String(b.reviewedAt||'').localeCompare(String(a.reviewedAt||'')))[0];
          d.reviewedAt=latest?.reviewedAt||'';
          d.reviewedBy=latest?.reviewedBy||'';
        }else{d.reviewedAt='';d.reviewedBy='';}
      }
      daily[date]=d;
    }
    return {...employee,supportReviews,days:daily};
  }
  function overview(employee){
    const days=Object.values(employee.days||{});
    const issues=days.filter(d=>d.issues.length&&!d.reviewed).length;
    return {days:days.filter(d=>d.value===1&&!d.nonWorking).length,issues,
      present:days.filter(d=>d.present&&!d.nonWorking&&!d.remoteExcluded).length,
      leave:days.filter(d=>(d.leaves||[]).length).length,travel:days.filter(d=>(d.official||[]).length).length,
      absence:days.filter(d=>!d.present&&isAbsenceStatus(d.status,d.flags)).length,
      manual:days.filter(d=>d.manualValue!==null&&!d.nonWorking&&d.manualValue!==d.automatic).length,
      missing:false,missingType:!employee.jenisPegawai,
      coverage:!!employee.fileWarnings?.length&&!employee.coverageAcknowledged,
      status:issues||!employee.jenisPegawai||(employee.fileWarnings?.length&&!employee.coverageAcknowledged)?'Terdapat Catatan':'Siap'};
  }
  function overall(employees){
    const s={employees:employees.length,days:0,issues:0,manual:0,missing:0,ready:0};
    for(const e of employees){const a=overview(e);s.days+=a.days;s.issues+=a.issues;s.manual+=a.manual;s.missing+=Number(a.missing);s.ready+=Number(a.status==='Siap');}
    return s;
  }
  window.UangMakanRules=Object.freeze({MONTHS,key,keyDate,weekday,periodLabel,periodDays,monthDates,weekend,
    normalizeNip,normalizeName,detectWorkMode,isAbsenceStatus,supportKey,supportSignature,
    classifyDay,applySupporting,overview,overall});
})();
