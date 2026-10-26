/* Uang Makan: aturan per bulan kalender. Tidak mengikuti offset periode Tukin. */
(function () {
  'use strict';
  const MONTHS=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const ABSENCE_FLAGS=new Set(['S','I','TK','D','C']);
  function pad(n){return String(n).padStart(2,'0');}
  function key(year,month,day){return `${year}-${pad(month)}-${pad(day)}`;}
  function periodDays(year,month){return new Date(Number(year),Number(month),0).getDate();}
  function periodLabel(year,month){return `${MONTHS[Number(month)-1]} ${year}`;}
  function keyDate(date){return key(date.getFullYear(),date.getMonth()+1,date.getDate());}
  function weekend(dateKey){const [y,m,d]=dateKey.split('-').map(Number);const day=new Date(y,m-1,d).getDay();return day===0||day===6;}
  function normalizeNip(v){return String(v??'').replace(/\D/g,'');}
  function normalizeName(v){return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');}
  function monthDates(year,month){return Array.from({length:periodDays(year,month)},(_,i)=>key(year,month,i+1));}
  function isAbsenceStatus(status,flags){
    if((flags||[]).some(flag=>ABSENCE_FLAGS.has(String(flag).toUpperCase())))return true;
    return /(?:^|\b)(cuti|sakit|izin|ijin|dinas luar|perjalanan dinas|tanpa keterangan|wfh|work from home|tidak hadir|alpa)(?:\b|$)/i.test(String(status||''));
  }
  function classifyDay(record, leaves=[], assignments=[]) {
    const hasIn=record?.inMinutes!==null && record?.inMinutes!==undefined;
    const hasOut=record?.outMinutes!==null && record?.outMinutes!==undefined;
    const present=hasIn||hasOut;
    const official=assignments.filter(a=>a.assignmentType==='official_travel');
    const workArrangements=assignments.filter(a=>a.assignmentType!=='official_travel');
    const status=String(record?.status||'').trim();
    const absenceStatus=isAbsenceStatus(status,record?.flags);
    const issues=[];
    if(present && !(hasIn&&hasOut))issues.push('Presensi hanya satu kali');
    if(present && leaves.length)issues.push('Presensi bertentangan dengan data cuti');
    if(present && official.length)issues.push('Presensi bertentangan dengan dinas luar');
    if(present && absenceStatus)issues.push('Presensi bertentangan dengan status absensi');
    if(present && weekend(record.key))issues.push('Presensi pada akhir pekan');
    if(record?.duplicate)issues.push('Tanggal tercatat dalam lebih dari satu file');
    if(!present && /\bhadir\b/i.test(status) && !absenceStatus)issues.push('Status hadir tetapi jam presensi kosong');
    // Tidak ada bukti presensi fisik berarti tidak memperoleh uang makan, walaupun surat tugas ada.
    const autoValue=present?1:0;
    const hasManual=record?.manualValue===0||record?.manualValue===1;
    const value=hasManual?record.manualValue:autoValue;
    return {key:record.key,inMinutes:record.inMinutes??null,outMinutes:record.outMinutes??null,
      status,flags:Array.isArray(record?.flags)?record.flags:[],duplicate:Boolean(record?.duplicate),
      leaves,assignments,workArrangements,official,issues,hasIn,hasOut,present,
      automatic:autoValue,value,manualValue:hasManual?record.manualValue:null,
      manualNote:String(record?.manualNote||''),reviewed:Boolean(record?.reviewed),
      reviewedAt:String(record?.reviewedAt||''),reviewedBy:String(record?.reviewedBy||'')};
  }
  function applySupporting(employee,leaves,assignments,year,month){
    const days=monthDates(year,month);
    const ownLeaves=leaves.filter(l=>l.employeeId===employee.id);
    const ownAssignments=assignments.filter(a=>a.employeeIds.includes(employee.id));
    const old=employee.days||{};
    const daily={};
    for(const date of days){
      const base=old[date]||{key:date,inMinutes:null,outMinutes:null,status:'',flags:[]};
      daily[date]=classifyDay({...base,key:date},ownLeaves.filter(l=>l.startDate<=date&&l.endDate>=date),ownAssignments.filter(a=>a.startDate<=date&&a.endDate>=date));
    }
    return {...employee,days:daily};
  }
  function overview(employee){
    const days=Object.values(employee.days||{});
    const issues=days.filter(d=>d.issues.length&&!d.reviewed).length;
    return { days:days.filter(d=>d.value===1).length,issues,
      present:days.filter(d=>d.present).length,leave:days.filter(d=>d.leaves.length).length,
      travel:days.filter(d=>d.official.length).length,absence:days.filter(d=>!d.present&&isAbsenceStatus(d.status,d.flags)).length,
      manual:days.filter(d=>d.manualValue!==null).length,
      missing:!employee.files?.length&&!employee.missingAcknowledged,
      missingType:!employee.jenisPegawai,coverage:!!employee.fileWarnings?.length&&!employee.coverageAcknowledged,
      status:issues||(!employee.files?.length&&!employee.missingAcknowledged)||!employee.jenisPegawai||(employee.fileWarnings?.length&&!employee.coverageAcknowledged)?'Perlu verifikasi':'Siap' };
  }
  function overall(employees){
    const s={employees:employees.length,days:0,issues:0,manual:0,missing:0,ready:0};
    for(const e of employees){const a=overview(e);s.days+=a.days;s.issues+=a.issues;s.manual+=a.manual;s.missing+=Number(a.missing);s.ready+=Number(a.status==='Siap');}
    return s;
  }
  window.UangMakanRules=Object.freeze({MONTHS,key,keyDate,periodLabel,periodDays,monthDates,weekend,normalizeNip,normalizeName,isAbsenceStatus,classifyDay,applySupporting,overview,overall});
})();
