/**
 * Paket ZIP Uang Makan: rekap Excel, berkas absensi asli per pegawai,
 * dan bukti dukung Cuti / Surat Tugas yang terkait dengan periode proses.
 * Binary absensi diambil dari unggahan lokal atau arsip SharePoint pada riwayat.
 */
(function(){
 'use strict';
 const R=window.UangMakanRules;

 function safeFileName(raw,fallback='Berkas'){
  let result=String(raw||'').replace(/[\\/<>:"|?*\x00-\x1f\x7f]/g,'_').replace(/\s+/g,' ').trim().replace(/^\.+/,'').replace(/[. ]+$/g,'');
  if(!result||result==='.'||result==='..')result=fallback;
  if(/^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(result))result=`_${result}`;
  if(result.length>120){
   const dot=result.lastIndexOf('.'),ext=dot>0&&result.length-dot<=10?result.slice(dot):'';
   result=result.slice(0,120-ext.length)+ext;
  }
  return result.replace(/[. ]+$/g,'')||fallback;
 }
 function uniqueName(used,raw,fallback){
  const name=safeFileName(raw,fallback);
  let candidate=name,index=2;
  while(used.has(candidate.toLocaleLowerCase('id-ID'))){
   const dot=name.lastIndexOf('.');
   candidate=dot>0?`${name.slice(0,dot)} (${index++})${name.slice(dot)}`:`${name} (${index++})`;
  }
  used.add(candidate.toLocaleLowerCase('id-ID'));
  return candidate;
 }
 function zipFileName(year,month){return `Uang_Makan_PUSDATIN_${R.MONTHS[month-1]}_${year}.zip`;}
 function attachmentLabel(item,fallback){return String(item?.name||item?.remoteName||item?.file?.name||fallback||'Berkas');}
 async function zipBinary(blob){return blob&&typeof blob.arrayBuffer==='function'?blob.arrayBuffer():blob;}

 async function resolveBinary(item,cache,label){
  if(item?.file && (typeof item.file.arrayBuffer==='function'||typeof item.file.stream==='function'))return item.file;
  if(typeof Blob!=='undefined'&&item instanceof Blob)return item;
  if(item?.driveId&&item?.itemId){
   if(!window.SharePointStorage?.downloadAttachment)throw new Error('Layanan download SharePoint tidak tersedia.');
   const key=`${item.driveId}:${item.itemId}`;
   if(!cache.has(key))cache.set(key,Promise.resolve().then(()=>window.SharePointStorage.downloadAttachment({...item,file:null},true)));
   try{
    const binary=await cache.get(key);
    if(binary)return binary;
   }catch(error){throw new Error(`Gagal mengambil ${label} (${attachmentLabel(item)}): ${error.message||error}`);}
  }
  throw new Error(`Berkas ${label} tidak tersedia: ${attachmentLabel(item)}. Periksa unggahan atau akses SharePoint.`);
 }
 function supportingDocuments(employee,policy){
  const documents=[],missing=[],seen=new Set(),seenMissing=new Set();
  for(const day of Object.values(employee.days||{})){
   for(const [kind,records] of [['Cuti',day.previewLeaves||day.leaves||[]],['Surat_Tugas',day.previewAssignments||day.assignments||[]]]){
    for(const record of records){
     const kindKey=kind==='Cuti'?'cuti':'surat';
     if(!policy.exportable(kindKey,record,employee.id))continue;
     for(const doc of record.documents||[]){
      const file=doc?.file||{};
      // Metadata dokumen tanpa berkas digital dicatat sebagai kekurangan,
      // bukan diam-diam dianggap berkas yang telah disertakan ke ZIP.
      if(!file.driveId&&!file.itemId&&!file.file){
       const key=`${kind}:${record.id||''}:${doc.id||''}:${doc.name||''}`;
       if(!seenMissing.has(key)){
        missing.push(`[${kind==='Cuti'?'Cuti':'Surat Tugas'}] ${doc.name||file.name||doc.type||'Dokumen tanpa nama'}${doc.number?' - '+doc.number:''}`);
        seenMissing.add(key);
       }
       continue;
      }
      const key=file.driveId&&file.itemId?`remote:${file.driveId}:${file.itemId}`:
        `${kind}:${record.id||''}:${doc.id||''}:${file.name||''}`;
      if(seen.has(key))continue;
      seen.add(key);
      documents.push({kind,file:{...file,name:file.name||file.remoteName||doc.name||'Bukti Dukung'}});
     }
    }
   }
  }
  return {documents,missing};
 }

 async function generate(year,month,employees,options={}){
  if(!window.JSZip)throw new Error('Pustaka ZIP belum tersedia.');
  if(!window.UangMakanExport?.generate)throw new Error('Generator Excel Uang Makan belum tersedia.');
  if(!Array.isArray(employees)||!employees.length)throw new Error('Belum ada data pegawai untuk diunduh.');
  const policy=await window.AttendanceDocumentPolicy.fetch([year-1,year]);
  if(policy.stale.length)throw new Error('Ada verifikasi Register yang perlu diperbarui sebelum mengekspor ZIP Uang Makan.');
  const recap=await window.UangMakanExport.generate(year,month,employees,{holidays:options.holidays||[],download:false});
  const zip=new window.JSZip();
  zip.file(safeFileName(recap.filename,'Rekap_Uang_Makan.xlsx'),await zipBinary(recap.blob));
  const pegawai=zip.folder('Pegawai'),usedEmployees=new Set(),cache=new Map();
  let attendanceCount=0,evidenceCount=0,missingEvidenceCount=0;
  for(const employee of employees){
   const folderName=uniqueName(usedEmployees,employee.name||employee.nip||'Pegawai','Pegawai');
   const folder=pegawai.folder(folderName);
   // Semua berkas milik pegawai ditempatkan langsung di folder ini.
   // Gunakan satu penampung nama untuk menghindari tabrakan lintas jenis dokumen.
   const usedNames=new Set();
   const sources=Array.isArray(employee.sourceFiles)?employee.sourceFiles:[];
   const expected=Array.isArray(employee.files)?employee.files.length:0;
   if(!sources.length||expected>sources.length){
    throw new Error(`File absensi asli ${employee.name} tidak tersedia pada riwayat ini. Riwayat lama hanya menyimpan nama file; unggah ulang absensi dan simpan proses baru untuk mengaktifkan ZIP lengkap.`);
   }
   for(const source of sources){
    const binary=await resolveBinary(source,cache,`absensi ${employee.name}`);
    folder.file(uniqueName(usedNames,attachmentLabel(source,'Absensi.xlsx'),'Absensi.xlsx'),await zipBinary(binary));
    attendanceCount++;
    options.onProgress?.({employee:employee.name,attendanceCount,evidenceCount});
   }
   const supporting=supportingDocuments(employee,policy);
   for(const doc of supporting.documents){
    const binary=await resolveBinary(doc.file,cache,`bukti dukung ${employee.name}`);
    folder.file(uniqueName(usedNames,attachmentLabel(doc.file,'Bukti Dukung'),'Bukti Dukung'),await zipBinary(binary));
    evidenceCount++;
    options.onProgress?.({employee:employee.name,attendanceCount,evidenceCount});
   }
   if(supporting.missing.length){
    missingEvidenceCount+=supporting.missing.length;
    folder.file(uniqueName(usedNames,'BERKAS_BELUM_TERSEDIA.txt','BERKAS_BELUM_TERSEDIA.txt'),
     `Dokumen berikut tercatat pada Cuti/Surat Tugas, tetapi berkas digital belum tersedia di data sumber:\n\n${supporting.missing.join('\n')}\n`);
   }
  }
  const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
  return {name:zipFileName(year,month),blob,attendanceCount,evidenceCount,missingEvidenceCount};
 }
 function download(result){
  if(!result?.blob||!result?.name)throw new Error('Paket ZIP belum tersedia.');
  const url=URL.createObjectURL(result.blob);
  const a=document.createElement('a');
  a.href=url;a.download=result.name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
 }
 window.UangMakanZip=Object.freeze({generate,download,zipFileName});
})();
