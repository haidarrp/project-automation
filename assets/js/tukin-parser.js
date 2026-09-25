(function () {
  'use strict';
  const cfg = window.TUKIN_CONFIG;
  const rules = window.TukinRules;

  function getCellDisplay(sheet, address) {
    const cell = sheet[address];
    if (!cell) return '';
    if (cell.w !== undefined) return String(cell.w);
    if (cell.v instanceof Date) return cell.v.toISOString();
    return cell.v === undefined || cell.v === null ? '' : String(cell.v);
  }

  function parseEmployeeIdentity(text, fallbackFileName) {
    const cleaned = String(text || '').trim();
    let name = cleaned;
    let nip = '';
    let match = cleaned.match(/^(.*?)\s+-\s+(\d{8,})\s*$/);
    if (match) { name = match[1].trim(); nip = match[2].trim(); }
    else {
      match = cleaned.match(/(\d{8,})\s*$/);
      if (match) { nip = match[1]; name = cleaned.slice(0, match.index).replace(/[\s-]+$/, '').trim(); }
    }
    if (!name) {
      name = String(fallbackFileName || '').replace(/\.(xlsx?|xls)$/i,'').replace(/^laporan_presensi_/i,'').replace(/_\d{8}_\d{8}$/i,'').replace(/_/g,' ').trim();
    }
    if (!name) throw new Error('Identitas pegawai tidak ditemukan pada file.');
    return { name, nip };
  }

  function safeDate(y,m,d) {
    const out = new Date(y,m-1,d,12,0,0);
    return out.getFullYear()===y && out.getMonth()===m-1 && out.getDate()===d ? out : null;
  }
  function parseDate(rawValue, displayValue) {
    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) return new Date(rawValue.getFullYear(),rawValue.getMonth(),rawValue.getDate(),12,0,0);
    if (typeof rawValue === 'number' && Number.isFinite(rawValue) && window.XLSX?.SSF) {
      const p = XLSX.SSF.parse_date_code(rawValue); if (p) return safeDate(p.y,p.m,p.d);
    }
    const text = String(displayValue || rawValue || '').trim();
    if (!text || text === '-') return null;
    let m = text.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)$/); if (m) return safeDate(+m[1],+m[2],+m[3]);
    m = text.match(/^([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})$/); if (m) return safeDate(+m[3],+m[2],+m[1]);
    const parsed = new Date(text); if (!Number.isNaN(parsed.getTime())) return new Date(parsed.getFullYear(),parsed.getMonth(),parsed.getDate(),12,0,0);
    return null;
  }
  function parseTimeMinutes(rawValue, displayValue) {
    const text = String(displayValue || '').trim();
    if (text && text !== '-') {
      const m = text.match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?$/);
      if (m) { const h=+m[1], min=+m[2]; if (h>=0&&h<=23&&min>=0&&min<=59) return h*60+min; }
    }
    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) return rawValue.getHours()*60+rawValue.getMinutes();
    if (typeof rawValue === 'number' && Number.isFinite(rawValue)) return Math.round((rawValue-Math.floor(rawValue))*1440)%1440;
    return null;
  }

  function findHeaderRow(rows) {
    for (let i=0;i<Math.min(rows.length,15);i++) {
      const row=(rows[i]||[]).map(v=>String(v||'').trim().toLowerCase());
      if (row.includes('tanggal') && row.includes('masuk') && row.includes('keluar')) return i;
    }
    throw new Error('Header Tanggal/Masuk/Keluar tidak ditemukan.');
  }
  function findIndex(header, label, fallback) {
    const i=header.findIndex(v=>v===label); return i>=0?i:fallback;
  }
  function detectColumns(displayRows) {
    const headerRowIndex=findHeaderRow(displayRows);
    const header=(displayRows[headerRowIndex]||[]).map(v=>String(v||'').trim().toLowerCase());
    return {
      headerRowIndex,
      dataStartRow:headerRowIndex+1,
      day:findIndex(header,'hari',1), date:findIndex(header,'tanggal',2), inTime:findIndex(header,'masuk',3), outTime:findIndex(header,'keluar',4),
      status:21,
      flags:[
        {index:7,code:'S'},{index:8,code:'I'},{index:9,code:'TK'},{index:10,code:'D'},
        {index:11,code:'TL'},{index:12,code:'TB'},{index:13,code:'C'},{index:14,code:'L'}
      ]
    };
  }

  function truthyMarker(value) {
    const text=String(value ?? '').trim(); return text && text !== '-' && text !== '0' ? text : '';
  }
  function collectFlags(row, columns) {
    return columns.flags.filter(f=>truthyMarker(row[f.index])).map(f=>f.code);
  }

  async function sha256(text) {
    const bytes=new TextEncoder().encode(String(text||''));
    const digest=await crypto.subtle.digest('SHA-256',bytes);
    return [...new Uint8Array(digest)].map(b=>b.toString(16).padStart(2,'0')).join('');
  }
  async function resolveMaster(identity) {
    const master=window.TUKIN_MASTER_HASHED || [];
    const nip=String(identity.nip||'').replace(/\D/g,'');
    const name=rules.normalizeName(identity.name);
    let item=null;
    if (nip) { const h=await sha256(nip); item=master.find(x=>x.nipHash===h)||null; }
    if (!item && name) { const h=await sha256(name); item=master.find(x=>x.nameHash===h)||null; }
    return item;
  }

  async function parseAttendanceFile(file, tukinPeriod, settings) {
    if (!window.XLSX) throw new Error('Library Excel belum termuat. Muat ulang halaman.');
    const ext=String(file.name||'').split('.').pop().toLowerCase();
    if (!['xlsx','xls'].includes(ext)) throw new Error('Format file harus .xlsx atau .xls.');
    const buffer=await file.arrayBuffer();
    let workbook;
    try { workbook=XLSX.read(buffer,{type:'array',cellDates:true,cellNF:true,cellText:true}); }
    catch (e) { throw new Error(`File Excel tidak dapat dibaca: ${e.message||e}`); }
    const first=workbook.SheetNames[0]; if (!first) throw new Error('Workbook tidak memiliki sheet.');
    const sheet=workbook.Sheets[first];
    const raw=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:''});
    const display=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:''});
    const identity=parseEmployeeIdentity(getCellDisplay(sheet,'A2') || display?.[1]?.[0], file.name);
    const columns=detectColumns(display);
    const range=rules.attendancePeriod(tukinPeriod);
    const records={};
    let sourceMin=null, sourceMax=null, rowsInPeriod=0;
    for (let ri=columns.dataStartRow;ri<raw.length;ri++) {
      const firstCell=String(display?.[ri]?.[0]||'').trim().toUpperCase(); if (firstCell==='TOTAL') break;
      const rr=raw[ri]||[], dr=display[ri]||[];
      const date=parseDate(rr[columns.date],dr[columns.date]); if (!date) continue;
      if (!sourceMin || date<sourceMin) sourceMin=date; if (!sourceMax || date>sourceMax) sourceMax=date;
      if (!rules.inRange(date,range.start,range.end)) continue;
      rowsInPeriod++;
      const inMinutes=parseTimeMinutes(rr[columns.inTime],dr[columns.inTime]);
      const outMinutes=parseTimeMinutes(rr[columns.outTime],dr[columns.outTime]);
      const status=String(dr[columns.status]||'').trim();
      const flags=collectFlags(dr,columns);
      const calc=rules.calculateDay(date,inMinutes,outMinutes,status,settings);
      const key=rules.dateKey(date);
      const record={
        key,date,inMinutes,outMinutes,status,flags,
        dayLabel:String(dr[columns.day]||rules.formatDate(date,true).split(',')[0]),
        isWeekend:rules.isWeekend(date),isHoliday:calc.holiday,
        schedule:calc.schedule,tlCategory:calc.tl.category,tlPercent:calc.tl.percent,
        pswCategory:calc.psw.category,pswPercent:calc.psw.percent,shortageMinutes:calc.psw.shortageMinutes,
        requiredEndMinutes:calc.psw.requiredEndMinutes,autoTotalPercent:calc.totalPercent,reason:calc.reason,
        adjustedPercent:null,adjustmentNote:'',evidence:[],
        needsVerification:calc.totalPercent>0 || flags.some(f=>f!=='L')
      };
      records[key]=record;
    }
    if (!rowsInPeriod) throw new Error(`Tidak ditemukan data pada periode ${rules.formatPeriodRange(range)}.`);
    const master=await resolveMaster(identity);
    const warnings=[];
    if (sourceMin && sourceMin>range.start) warnings.push(`Data mulai ${rules.formatDate(sourceMin,false)}, setelah awal periode.`);
    if (sourceMax && sourceMax<range.end) warnings.push(`Data berakhir ${rules.formatDate(sourceMax,false)}, sebelum akhir periode.`);
    if (!master) warnings.push('Data master pegawai tidak ditemukan; Anak Satker perlu dicek.');
    return {
      name:identity.name,nip:identity.nip,
      satker:master?.satker || cfg.SATKER,
      anakSatker:master?.anakSatker || '',
      tukin:Number(master?.tukin || cfg.DEFAULT_TUKIN),
      masterOrder:Number(master?.order || 9999),masterMatched:Boolean(master),
      skpDeductionPercent:0,
      records,sourceFiles:[file],warnings
    };
  }

  function mergeEmployee(target,source) {
    Object.assign(target.records,source.records);
    target.sourceFiles.push(...source.sourceFiles);
    target.warnings=[...new Set([...(target.warnings||[]),...(source.warnings||[])])];
    if (!target.nip&&source.nip) target.nip=source.nip;
  }

  async function parseFiles(files,tukinPeriod,settings,onProgress) {
    const list=Array.from(files||[]), byKey=new Map(), results=[];
    for (let i=0;i<list.length;i++) {
      const file=list[i]; onProgress?.({index:i,total:list.length,file});
      try {
        const emp=await parseAttendanceFile(file,tukinPeriod,settings);
        const key=rules.employeeKey(emp);
        if (byKey.has(key)) mergeEmployee(byKey.get(key),emp); else byKey.set(key,emp);
        results.push({fileName:file.name,ok:true,employee:emp.name,warnings:emp.warnings});
      } catch (e) { results.push({fileName:file.name,ok:false,error:e.message||String(e)}); }
    }
    const employees=[...byKey.values()].sort((a,b)=>(a.masterOrder-b.masterOrder)||a.name.localeCompare(b.name,'id'));
    return {employees,results,errors:results.filter(r=>!r.ok)};
  }

  function recalculateEmployee(employee,settings) {
    Object.values(employee.records||{}).forEach(rec=>{
      const calc=rules.calculateDay(rec.date,rec.inMinutes,rec.outMinutes,rec.status,settings);
      rec.isHoliday=calc.holiday;rec.schedule=calc.schedule;rec.tlCategory=calc.tl.category;rec.tlPercent=calc.tl.percent;
      rec.pswCategory=calc.psw.category;rec.pswPercent=calc.psw.percent;rec.shortageMinutes=calc.psw.shortageMinutes;
      rec.requiredEndMinutes=calc.psw.requiredEndMinutes;rec.autoTotalPercent=calc.totalPercent;rec.reason=calc.reason;
      rec.needsVerification=calc.totalPercent>0 || (rec.flags||[]).some(f=>f!=='L');
    });
  }

  window.TukinParser=Object.freeze({parseFiles,parseAttendanceFile,recalculateEmployee,parseDate,parseTimeMinutes});
})();
