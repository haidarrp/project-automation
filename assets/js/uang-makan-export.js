/**
 * Ekspor Uang Makan PUSDATIN - workbook dibuat dari nol dari konfigurasi JavaScript.
 * Tidak perlu menyimpan, mengambil, atau mengunggah template .xlsx.
 * XML format dibangkitkan berdasarkan UangMakanFormat (font, border, warna,
 * ukuran baris/kolom, merged cells, pengaturan cetak, dan rumus).
 */
(function () {
  'use strict';
  const R=window.UangMakanRules;
  const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const XML='http://www.w3.org/XML/1998/namespace';
  const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  function make(doc,tag,attrs,parent){
    const item=doc.createElementNS(NS,tag);
    for(const [key,value] of Object.entries(attrs||{}))item.setAttribute(key,String(value));
    if(parent)parent.appendChild(item);
    return item;
  }
  function fromSettings(doc,setting,parent){
    const item=make(doc,setting.tag,setting.attrs,parent);
    for(const child of setting.children||[])fromSettings(doc,child,item);
    return item;
  }
  function colName(index){
    let col='';
    while(index>0){index--;col=String.fromCharCode(65+(index%26))+col;index=Math.floor(index/26);}
    return col;
  }
  function colNum(coord){
    let n=0;
    for(const letter of (coord.match(/^[A-Z]+/)||[''])[0])n=n*26+(letter.charCodeAt(0)-64);
    return n;
  }
  function getCell(doc,row,coord){
    const existing=Array.from(row.children).find(c=>c.localName==='c'&&c.getAttribute('r')===coord);
    if(existing)return existing;
    const target=colNum(coord),cell=make(doc,'c',{r:coord});
    const next=Array.from(row.children).find(c=>c.localName==='c'&&colNum(c.getAttribute('r'))>target);
    row.insertBefore(cell,next||null);
    return cell;
  }
  function clearValue(cell){
    for(const child of Array.from(cell.children))if(['f','v','is'].includes(child.localName))cell.removeChild(child);
    cell.removeAttribute('t');
  }
  function setString(doc,row,coord,value){
    const cell=getCell(doc,row,coord);clearValue(cell);
    const text=String(value??'');
    if(!text)return;
    cell.setAttribute('t','inlineStr');
    const inner=make(doc,'is',{},cell);
    const node=make(doc,'t',{},inner);
    node.setAttributeNS(XML,'xml:space','preserve');
    node.textContent=text;
  }
  function setNumber(doc,row,coord,value){
    const cell=getCell(doc,row,coord);clearValue(cell);
    if(value===null||value===undefined||value==='')return;
    const v=make(doc,'v',{},cell);v.textContent=String(value);
  }
  function setFormula(doc,row,coord,formula,cached){
    const cell=getCell(doc,row,coord);clearValue(cell);
    const f=make(doc,'f',{},cell);f.textContent=String(formula).replace(/^=/,'');
    const v=make(doc,'v',{},cell);v.textContent=String(cached??0);
  }
  function getStyles(row,rowNumber){
    const weekday=Array.from(row.children).find(c=>c.getAttribute('r')===`D${rowNumber}`);
    const saturday=Array.from(row.children).find(c=>c.getAttribute('r')===`H${rowNumber}`);
    if(!weekday||!saturday||weekday.getAttribute('s')===saturday.getAttribute('s'))return null;
    return {weekday:weekday.getAttribute('s'),weekend:saturday.getAttribute('s')};
  }
  function styleWeekends(row,year,month,rowNo,styles){
    if(!styles)return;
    const map=new Map(Array.from(row.children).filter(c=>c.localName==='c').map(c=>[c.getAttribute('r'),c]));
    for(let day=1;day<=31;day++){
      const cell=map.get(`${colName(day+3)}${rowNo}`);
      if(!cell)continue;
      const isWeekend=day<=R.periodDays(year,month)&&R.weekend(R.key(year,month,day));
      cell.setAttribute('s',isWeekend?styles.weekend:styles.weekday);
    }
  }
  function createWorksheet(year,month,employees,settings){
    const finalRow=Math.max(settings.baseEndRow,employees.length+4);
    const doc=document.implementation.createDocument(NS,'worksheet',null);
    const root=doc.documentElement;
    const layout=settings.layout;
    fromSettings(doc,layout.sheetPr,root);
    make(doc,'dimension',{ref:`A1:AV${finalRow}`},root);
    fromSettings(doc,layout.sheetViews,root);
    fromSettings(doc,layout.sheetFormatPr,root);
    fromSettings(doc,layout.cols,root);
    const data=make(doc,'sheetData',{},root);
    const rows=new Map();
    for(let rowNo=1;rowNo<=finalRow;rowNo++){
      const index=String(rowNo),base=index in settings.rowCells?index:String(settings.baseEndRow);
      const attrs={r:String(rowNo),...settings.rowAttrs[base]};
      const row=make(doc,'row',attrs,data);
      for(const [col,s] of settings.rowCells[base])make(doc,'c',{r:`${colName(col)}${rowNo}`,s},row);
      rows.set(rowNo,row);
    }
    for(const [coord,type,value] of settings.headingValues){
      const rowNo=Number(coord.match(/\d+$/)[0]);
      if(type==='s')setString(doc,rows.get(rowNo),coord,value);
      else setNumber(doc,rows.get(rowNo),coord,value);
    }
    setNumber(doc,rows.get(2),'B2',year);
    setString(doc,rows.get(2),'C2',`Uang_Makan_${R.MONTHS[month-1].toUpperCase()}_${year}`);
    setString(doc,rows.get(3),'B3',String(month).padStart(2,'0'));
    styleWeekends(rows.get(4),year,month,4,getStyles(rows.get(4),4));
    for(let index=0;index<employees.length;index++){
      const rowNo=index+5,employee=employees[index],row=rows.get(rowNo);
      styleWeekends(row,year,month,rowNo,getStyles(row,rowNo));
      setNumber(doc,row,`A${rowNo}`,index+1);
      setString(doc,row,`B${rowNo}`,R.normalizeNip(employee.nip));
      setString(doc,row,`C${rowNo}`,employee.name);
      let total=0;
      for(let day=1;day<=R.periodDays(year,month);day++){
        const attendance=employee.days?.[R.key(year,month,day)];
        if(attendance?.value===1){setNumber(doc,row,`${colName(day+3)}${rowNo}`,1);total++;}
      }
      setFormula(doc,row,`AI${rowNo}`,`SUM(D${rowNo}:AH${rowNo})`,total);
      setFormula(doc,row,`AJ${rowNo}`,`COUNTA(D${rowNo}:AH${rowNo})`,total);
      setFormula(doc,row,`AK${rowNo}`,`AI${rowNo}-AJ${rowNo}`,0);
      setString(doc,row,`AL${rowNo}`,'Pusat Data dan Informasi');
      setString(doc,row,`AM${rowNo}`,String(employee.anakSatker||''));
      setString(doc,row,`AN${rowNo}`,String(employee.jenisPegawai||''));
      // AP-AU sengaja kosong: kolom rekonsiliasi GPP memerlukan data sumber lain.
    }
    const filter=fromSettings(doc,layout.autoFilter,root);
    filter.setAttribute('ref',`A4:AV${finalRow}`);
    fromSettings(doc,layout.mergeCells,root);
    fromSettings(doc,layout.pageMargins,root);
    return new XMLSerializer().serializeToString(doc);
  }
  function encodeXml(text){return String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function createPackage(year,month,employees,settings){
    const zip=new JSZip(),finalRow=Math.max(settings.baseEndRow,employees.length+4);
    const now=new Date().toISOString().replace(/\.\d{3}Z$/,'Z');
    const ns='http://schemas.openxmlformats.org';
    zip.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`);
    zip.file('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${ns}/package/2006/relationships"><Relationship Type="${ns}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" Id="rId1"/><Relationship Type="${ns}/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" Id="rId2"/><Relationship Type="${ns}/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" Id="rId3"/></Relationships>`);
    zip.file('docProps/app.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="${ns}/officeDocument/2006/extended-properties"><Application>PUSDATIN Uang Makan</Application><AppVersion>1.0</AppVersion></Properties>`);
    zip.file('docProps/core.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${ns}/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Pusdatin PKP</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    const viewAttrs=Object.entries(settings.workbookView||{}).map(([k,v])=>`${k}="${encodeXml(v)}"`).join(' ');
    zip.file('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${ns}/officeDocument/2006/relationships"><workbookPr/><bookViews><workbookView ${viewAttrs}/></bookViews><sheets><sheet name="${encodeXml(settings.sheetName)}" sheetId="1" state="visible" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'PUSDATIN'!$A$4:$AV$${finalRow}</definedName></definedNames><calcPr calcId="191028" fullCalcOnLoad="1"/></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${ns}/package/2006/relationships"><Relationship Type="${ns}/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml" Id="rId1"/><Relationship Type="${ns}/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId2"/><Relationship Type="${ns}/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" Id="rId3"/></Relationships>`);
    zip.file('xl/styles.xml',settings.stylesXml);
    zip.file('xl/theme/theme1.xml',settings.themeXml);
    zip.file('xl/worksheets/sheet1.xml',createWorksheet(year,month,employees,settings));
    return zip;
  }
  function verifyInput(year,month,employees){
    if(!Number.isInteger(year)||year<2000||year>2100||!Number.isInteger(month)||month<1||month>12)throw new Error('Periode tidak valid.');
    if(!Array.isArray(employees)||!employees.length)throw new Error('Daftar pegawai kosong.');
    if(employees.some(e=>!e.name||!e.nip))throw new Error('Nama/NIP pegawai belum lengkap.');
  }
  async function generate(year,month,employees,options={}){
    verifyInput(year,month,employees);
    if(!window.JSZip)throw new Error('Pustaka pembuat Excel belum tersedia.');
    const format=window.UangMakanFormat;
    if(!format?.rowCells||!format?.stylesXml||!format?.themeXml)throw new Error('Konfigurasi format PUSDATIN tidak tersedia.');
    const blob=await createPackage(year,month,employees,format).generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
    const filename=`Uang_Makan_PUSDATIN_${R.MONTHS[month-1]}_${year}.xlsx`;
    if(options.download!==false){
      const url=URL.createObjectURL(new Blob([blob],{type:MIME}));
      const a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),30000);
    }
    return {blob,filename};
  }
  window.UangMakanExport=Object.freeze({generate});
})();
