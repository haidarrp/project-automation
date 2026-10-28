/* Uang Makan: aturan per bulan kalender. Tidak mengikuti offset periode Tukin.
 * Sabtu/Minggu dan tanggal merah tambahan selalu bernilai 0, termasuk ketika
 * ada presensi atau hasil verifikasi lama bernilai 1.
 */
(function () {
  'use strict';
  const MONTHS=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const ABSENCE_FLAGS=new Set(['S','I','TK','D','C']);
  // WFA/WFH merupakan lokasi kerja, bukan ketidakhadiran. Aturan uang makan:
  // WFA/WFH hanya bisa dihitung pada hari Jumat, apabila ada jam presensi.
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
  function classifyDay(record, leaves=[], assignments=[]) {
    const dateKey=String(record?.key||'');
    const hasIn=record?.inMinutes!==null && record?.inMinutes!==undefined;
    const hasOut=record?.outMinutes!==null && record?.outMinutes!==undefined;
    const present=hasIn||hasOut;
    const official=assignments.filter(a=>a.assignmentType==='official_travel');
    const workArrangements=assignments.filter(a=>a.assignmentType!=='official_travel');
    const status=String(record?.status||'').trim();
    const workMode=detectWorkMode(record?.workMode)||detectWorkMode(status);
    const absenceStatus=isAbsenceStatus(status,record?.flags);
    const isWeekend=weekend(dateKey);
    const isHoliday=Boolean(record?.isHoliday);
    const nonWorking=isWeekend||isHoliday;
    const remoteExcluded=Boolean(workMode)&&weekday(dateKey)!==5;
    const issues=[];
    // Hari nonkerja tidak berhak atas uang makan; presensi tidak perlu diverifikasi
    // untuk tujuan uang makan karena hasilnya tetap 0.
    if(!nonWorking){
      if(present && !(hasIn&&hasOut))issues.push('Presensi hanya satu kali');
      if(present && leaves.length)issues.push('Presensi bertentangan dengan data cuti');
      if(present && official.length)issues.push('Presensi bertentangan dengan dinas luar');
      if(present && absenceStatus)issues.push('Presensi bertentangan dengan status absensi');
      if(record?.duplicate)issues.push('Tanggal tercatat dalam lebih dari satu file');
      if(!present && /\bhadir\b/i.test(status) && !absenceStatus)issues.push('Status hadir tetapi jam presensi kosong');
      // Surat tugas tersinkron tidak membuktikan pegawai benar melaksanakan tugas.
      // Wajib konfirmasi walaupun presensi tidak berbenturan atau tidak tersedia,
      // misalnya ketika penugasan ternyata batal / pegawai tidak jadi berangkat.
      if(assignments.length)issues.push('Konfirmasi pelaksanaan surat tugas (termasuk jika batal/tidak berangkat)');
    }
    const automatic=(present&&!nonWorking&&!remoteExcluded)?1:0;
    const hasManual=record?.manualValue===0||record?.manualValue===1;
    // Libur dan WFA/WFH di luar Jumat wajib bernilai 0, termasuk jika ada
    // koreksi manual dari proses sebelumnya atau riwayat lama.
    const lockedZero=nonWorking||remoteExcluded;
    const manualValue=lockedZero?null:(hasManual?record.manualValue:null);
    const value=lockedZero?0:(manualValue===null?automatic:manualValue);
    return {key:dateKey,inMinutes:record?.inMinutes??null,outMinutes:record?.outMinutes??null,
      status,workMode,remoteExcluded,flags:Array.isArray(record?.flags)?record.flags:[],duplicate:Boolean(record?.duplicate),
      leaves,assignments,workArrangements,official,issues,hasIn,hasOut,present,isWeekend,isHoliday,nonWorking,
      automatic,value,manualValue,manualNote:nonWorking?'':String(record?.manualNote||''),
      reviewed:nonWorking?false:Boolean(record?.reviewed),reviewedAt:nonWorking?'':String(record?.reviewedAt||''),
      reviewedBy:nonWorking?'':String(record?.reviewedBy||'')};
  }
  function applySupporting(employee,leaves,assignments,year,month,holidays=[]){
    const dates=monthDates(year,month),reds=new Set(holidays||[]);
    const ownLeaves=leaves.filter(l=>l.employeeId===employee.id);
    const ownAssignments=assignments.filter(a=>a.employeeIds.includes(employee.id));
    const old=employee.days||{},daily={};
    for(const date of dates){
      const base=old[date]||{key:date,inMinutes:null,outMinutes:null,status:'',flags:[]};
      daily[date]=classifyDay({...base,key:date,isHoliday:reds.has(date)},
        ownLeaves.filter(l=>l.startDate<=date&&l.endDate>=date),
        ownAssignments.filter(a=>a.startDate<=date&&a.endDate>=date));
    }
    return {...employee,days:daily};
  }
  function overview(employee){
    const days=Object.values(employee.days||{});
    const issues=days.filter(d=>d.issues.length&&!d.reviewed).length;
    return {days:days.filter(d=>d.value===1&&!d.nonWorking).length,issues,
      present:days.filter(d=>d.present&&!d.nonWorking&&!d.remoteExcluded).length,
      leave:days.filter(d=>d.leaves.length).length,travel:days.filter(d=>d.official.length).length,
      absence:days.filter(d=>!d.present&&isAbsenceStatus(d.status,d.flags)).length,
      manual:days.filter(d=>d.manualValue!==null&&!d.nonWorking&&d.manualValue!==d.automatic).length,
      missing:false,
      missingType:!employee.jenisPegawai,coverage:!!employee.fileWarnings?.length&&!employee.coverageAcknowledged,
      status:issues||!employee.jenisPegawai||(employee.fileWarnings?.length&&!employee.coverageAcknowledged)?'Perlu verifikasi':'Siap'};
  }
  function overall(employees){
    const s={employees:employees.length,days:0,issues:0,manual:0,missing:0,ready:0};
    for(const e of employees){const a=overview(e);s.days+=a.days;s.issues+=a.issues;s.manual+=a.manual;s.missing+=Number(a.missing);s.ready+=Number(a.status==='Siap');}
    return s;
  }
  window.UangMakanRules=Object.freeze({MONTHS,key,keyDate,weekday,periodLabel,periodDays,monthDates,weekend,normalizeNip,normalizeName,detectWorkMode,isAbsenceStatus,classifyDay,applySupporting,overview,overall});
})();
