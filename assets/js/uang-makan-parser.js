(function(){
  'use strict';
  const R=window.UangMakanRules;
  function value(cell){return cell?.w!==undefined?String(cell.w):String(cell?.v??'');}
  function identity(sheet,fileName){
    const s=value(sheet.A2).trim();
    let name=s,nip='';
    const m=s.match(/^(.*?)\s*[-–]\s*(\d{8,})\s*$/)||s.match(/^(.*?)\s+(\d{8,})\s*$/);
    if(m){name=m[1].trim();nip=m[2];}
    if(!name)name=fileName.replace(/\.xlsx?$/i,'').replace(/^laporan_presensi_/i,'').replace(/_\d{8}_\d{8}$/,'').replace(/_/g,' ').trim();
    if(!name)throw new Error('Nama pegawai tidak ditemukan pada A2/nama file.');
    return {name,nip};
  }
  function headerInfo(display){
    const ix=display.slice(0,20).findIndex(row=>{const x=row.map(v=>String(v||'').trim().toLowerCase());return x.includes('tanggal')&&x.includes('masuk')&&x.includes('keluar');});
    if(ix<0)throw new Error('Header Tanggal/Masuk/Keluar tidak ditemukan. Gunakan file absensi yang sama seperti Tukin.');
    const h=display[ix].map(v=>String(v||'').trim().toLowerCase());
    const find=(x,fallback)=>{const index=h.indexOf(x);return index<0?fallback:index;};
    const stat=h.findIndex(x=>x==='keterangan'||x==='status');
    return {start:ix+1,date:find('tanggal',2),in:find('masuk',3),out:find('keluar',4),status:stat<0?21:stat};
  }
  function flag(row,index){const v=String(row[index]??'').trim();return Boolean(v&&v!=='-'&&v!=='0');}
  function parseFile(file,year,month){
    if(!window.XLSX)throw new Error('Pustaka Excel tidak tersedia. Muat ulang halaman.');
    if(!/\.xlsx?$/i.test(file.name))throw new Error('Berkas harus Excel (.xls/.xlsx).');
    return file.arrayBuffer().then(buffer=>{
      const wb=XLSX.read(buffer,{type:'array',cellDates:true,cellNF:true,cellText:true});
      if(!wb.SheetNames.length)throw new Error('Workbook kosong.');
      const sheet=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:''});
      const display=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:''});
      const info=headerInfo(display),person=identity(sheet,file.name);
      const rows={}, seen=new Set();let min='',max='',valid=0,dupes=0;
      const parseDate=window.TukinParser.parseDate;
      const parseTime=window.TukinParser.parseTimeMinutes;
      for(let i=info.start;i<raw.length;i++){
        const rr=raw[i]||[],dr=display[i]||[];
        if(String(dr[0]||'').trim().toUpperCase()==='TOTAL')break;
        const date=parseDate(rr[info.date],dr[info.date]);if(!date)continue;
        const dateKey=R.keyDate(date);
        if(!min||dateKey<min)min=dateKey;if(!max||dateKey>max)max=dateKey;
        if(date.getFullYear()!==year||date.getMonth()+1!==month)continue;
        valid++;
        const flags=[['S',7],['I',8],['TK',9],['D',10],['TL',11],['TB',12],['C',13],['L',14]].filter(x=>flag(dr,x[1])).map(x=>x[0]);
        const inMinutes=parseTime(rr[info.in],dr[info.in]);
        const outMinutes=parseTime(rr[info.out],dr[info.out]);
        const record={key:dateKey,inMinutes,outMinutes,status:String(dr[info.status]||'').trim(),flags,duplicate:seen.has(dateKey)};
        if(record.duplicate){dupes++;const prev=rows[dateKey];record.inMinutes=record.inMinutes??prev?.inMinutes??null;record.outMinutes=record.outMinutes??prev?.outMinutes??null;record.flags=[...new Set([...(prev?.flags||[]),...record.flags])];}
        rows[dateKey]=record;seen.add(dateKey);
      }
      if(!valid)throw new Error(`Tidak ada data tanggal pada ${R.periodLabel(year,month)}.`);
      const warnings=[];
      const start=R.key(year,month,1),end=R.key(year,month,R.periodDays(year,month));
      if(min>start)warnings.push(`Tanggal pertama file ${min}, melewati awal bulan ${start}.`);
      if(max<end)warnings.push(`Tanggal terakhir file ${max}, sebelum akhir bulan ${end}.`);
      if(dupes)warnings.push(`${dupes} tanggal duplikat dalam berkas.`);
      return {fileName:file.name,identity:person,rows,warnings,coverage:{min,max}};
    });
  }
  async function parseFiles(files,year,month,onProgress){
    const results=[];
    for(const file of Array.from(files||[])){
      try{const data=await parseFile(file,year,month);results.push({ok:true,...data});}
      catch(e){results.push({ok:false,fileName:file.name,error:String(e.message||e)});}
      onProgress?.(results.length,files.length);
    }
    return results;
  }
  window.UangMakanParser=Object.freeze({parseFile,parseFiles});
})();
