'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const service = require('../assets/js/surat-tugas-permohonan-service.js');
const root = path.resolve(__dirname,'..');

const mockRows = [
  ['','','','','','','','','',''],
  ['Id','Start time','Completion time','Email','Name','Nama','Jenis','Nomor Surat/ Perihal','Deksripsi','Status'],
  [4,45601.1,'','','','','','','',''],
  [7,46335.6,46335.65,'anonymous','','Muhammad Hatta Rahmatul Kholiq','Perubahan','285/ST/Sj.08/2026','Tambah daftar pelaksana','Selesai'],
  [8,46335.7,46335.74,'anonymous','','Rinaldy Dwi Istanto','Perubahan','286/ST/Sj.08/2026','Perubahan tanggal','Selesai'],
  [9,46339.4,46339.5,'anonymous','','Khusnul Arifin','Perubahan','286/ST/Sj.08/2026','Perubahan tanggal penugasan',null]
];

test('parse Excel: header row is detected even with blank top rows, unfinished responses ignored', () => {
  const records = service.parseWorksheet({values:mockRows});
  assert.equal(records.length,3);
  assert.equal(records.find(r=>r.id==='9').name,'Khusnul Arifin');
  assert.equal(records.find(r=>r.id==='9').status,'Belum ditangani');
  assert.equal(records.find(r=>r.id==='7').status,'Selesai');
  assert.equal(records.find(r=>r.id==='7').email,'');
  assert.equal(records.filter(r=>r.letter==='286/ST/Sj.08/2026').length,2);
});

test('duplicate rows are suppressed, alternate spelling Deskripsi supported', () => {
  const variant = mockRows.map(row=>row.slice());
  variant[1][8]='Deskripsi';
  variant.push(variant[5].slice());
  assert.equal(service.parseWorksheet({values:variant}).length,3);
  assert.equal(service.parseWorksheet({values:[['NOT','MATCH']]}),null);
});

test('service uses signed in delegated Graph token and only GET endpoints', async () => {
  const calls=[];
  const serviceCode = fs.readFileSync(path.join(root,'assets/js/surat-tugas-permohonan-service.js'),'utf8');
  const fakeWindow = { MICROSOFT_CONFIG: {graphBaseUrl:'https://graph.microsoft.com/v1.0'},
    ST_CHANGE_REQUESTS_CONFIG: {shareUrl:'https://contoso.sharepoint.com/file?token=abc'},
    MicrosoftAuth:{async acquireToken(interactive){assert.equal(interactive,true);return 'test-token';}}};
  const context={window:fakeWindow,TextEncoder, btoa, fetch:async(url,options)=>{
    calls.push({url,options});
    if(url.includes('/shares/')) return {ok:true,async json(){return {id:'item1',name:'Responses.xlsx',webUrl:'https://example.com',parentReference:{driveId:'drive1'}};}};
    if(url.endsWith('/worksheets?$select=id,name')) return {ok:true,async json(){return {value:[{id:'sheet1',name:'Sheet1'}]};}};
    if(url.includes('usedRange')) return {ok:true,async json(){return {values:mockRows};}};
    throw new Error('Unexpected '+url);
  },module:{exports:{}}};
  vm.runInNewContext(serviceCode,context,{filename:'surat-tugas-permohonan-service.js'});
  const result=await context.module.exports.fetchRequests();
  assert.equal(result.rows.length,3);
  assert.equal(calls.length,3);
  for(const c of calls){assert.equal(c.options.method,'GET');assert.equal(c.options.headers.Authorization,'Bearer test-token');}
  assert.match(calls[0].url,/\/shares\/u!/);
  assert.match(calls[2].url,/usedRange\(valuesOnly=true\)/);
});

test('new view integrated into navigation and permission models', () => {
  const menu=fs.readFileSync(path.join(root,'assets/js/ui-shell.js'),'utf8');
  const access=fs.readFileSync(path.join(root,'assets/js/user-access.js'),'utf8');
  const assignment=fs.readFileSync(path.join(root,'assets/js/surat-tugas.js'),'utf8');
  assert.match(menu,/view: 'permohonan'/);
  assert.match(access,/id:'permohonan'/);
  assert.match(assignment,/function refreshRequests\(/);
  assert.match(assignment,/function requestDetailModal\(/);
  assert.match(assignment,/st-request-refresh/);
  assert.match(assignment,/saveRequestMarkers/);
});

async function mountRequests() {
  const handlers = {};
  const storage = new Map();
  const app = {innerHTML:'',addEventListener(type,fn){handlers[type]=fn;}};
  let resultRows = service.parseWorksheet({values:mockRows});
  let writes = [];
  let shouldFail = false;
  const browserWindow = {
    confirm(){return true;},
    ST_CHANGE_REQUESTS_CONFIG:{shareUrl:'https://example.com/forms.xlsx'},
    SharePointStorage:{},
    AppShell:{render({content='',overlays=''}){return content+overlays;}},
    UserAccess:{async requireModule(){return true;},canView(){return true;},firstView(){return 'permohonan';}},
    MasterDataService:{async getBundle(){return {employees:[],directory:new Map()};}},
    STChangeRequests:{async fetchRequests(){return {rows:resultRows,fileName:'Responses.xlsx',worksheet:'Sheet1'};},
      async completeRequest(id){writes.push(id);if(shouldFail)throw new Error('Microsoft Graph 403 (accessDenied): Perubahan ditolak');resultRows=resultRows.map(row=>row.id===id?{...row,status:'Selesai'}:row);return {id,status:'Selesai',alreadyDone:false};}},
    FirebaseClient:{getDb(){return {collection(){return {async get(){return {docs:[]};}};}};},getCurrentUser(){return {uid:'staff-1'};}},
    addEventListener(){},requestAnimationFrame(){},
  };
  const context={window:browserWindow,document:{getElementById(id){return id==='app'?app:null;},body:{classList:{toggle(){}}}},
    location:{hash:'#permohonan',search:''},history:{replaceState(){}},
    localStorage:{getItem(k){return storage.get(k)||null;},setItem(k,v){storage.set(k,v);}},
    console,URLSearchParams,setTimeout,clearTimeout,
    Intl,Date,Map,Set};
  vm.runInNewContext(fs.readFileSync(path.join(root,'assets/js/surat-tugas.js'),'utf8'),context);
  await new Promise(resolve=>setImmediate(resolve));
  function click(action, requestId) {
    const el={dataset:{action,requestId},classList:{contains(){return false;}}};
    handlers.click({target:{closest(selector){return selector==='[data-action]'?el:null;}}});
  }
  return {app,click,window:browserWindow,handlers,setRows(rows){resultRows=rows;},storage,getWrites(){return writes.slice();},setFail(value){shouldFail=value;}};
}

test('completeRequest PATCH writes only the Status cell for the ID re-read from latest usedRange',async()=>{
  const calls=[];
  const serviceCode=fs.readFileSync(path.join(root,'assets/js/surat-tugas-permohonan-service.js'),'utf8');
  const fakeWindow={MICROSOFT_CONFIG:{graphBaseUrl:'https://graph.microsoft.com/v1.0'},
    ST_CHANGE_REQUESTS_CONFIG:{shareUrl:'https://contoso.sharepoint.com/file?token=abc'},
    MicrosoftAuth:{async acquireToken(interactive){assert.equal(interactive,true);return 'valid-token';}}};
  const shiftedRange={ rowIndex:2, columnIndex:2, values:mockRows.slice(1) };
  const context={window:fakeWindow,TextEncoder,btoa,module:{exports:{}},fetch:async(url,options)=>{
    calls.push({url,options});
    if(url.includes('/shares/'))return {ok:true,async json(){return {id:'item1',parentReference:{driveId:'drive1'}};}};
    if(url.includes('/worksheets?$select'))return {ok:true,async json(){return {value:[{id:'sheet1',name:'Responses'}]};}};
    if(url.includes('usedRange'))return {ok:true,async json(){return shiftedRange;}};
    if(options.method==='PATCH')return {ok:true,status:200,async json(){return {values:[['Selesai']]};}};
    throw Error('Unexpected Graph URL: '+url);
  }};
  vm.runInNewContext(serviceCode,context);
  const updated=await context.module.exports.completeRequest('9');
  assert.equal(updated.status,'Selesai');
  assert.equal(updated.address,'L7'); // usedRange starts at row 3; the final mocked ID is at relative index 4
  assert.equal(calls.filter(c=>c.options.method==='PATCH').length,1);
  const patch=calls.find(c=>c.options.method==='PATCH');
  assert.match(patch.url,/\/range\(address='L7'\)$/);
  assert.equal(patch.options.headers['Content-Type'],'application/json');
  assert.deepEqual(JSON.parse(patch.options.body),{values:[['Selesai']]});
  assert.equal(patch.options.headers.Authorization,'Bearer valid-token');
});

test('already completed, missing Status, missing location and duplicated IDs never cause a write', async()=>{
  const existing=service.findStatusCell({values:mockRows,rowIndex:0,columnIndex:0},'9');
  assert.equal(existing.address,'J6');
  assert.throws(()=>service.findStatusCell({values:mockRows},'9'),/Posisi baris/);
  assert.throws(()=>service.findStatusCell({values:mockRows.map(r=>r.slice(0,9)),rowIndex:0,columnIndex:0},'9'),/Kolom Status/);
  const dup=[...mockRows,mockRows[5].slice()];
  assert.throws(()=>service.findStatusCell({values:dup,rowIndex:0,columnIndex:0},'9'),/duplikasi ID/);
  assert.throws(()=>service.findStatusCell({values:mockRows,rowIndex:0,columnIndex:0},'999'),/tidak ditemukan/);
  assert.equal(service.findStatusCell({values:mockRows,address:'Sheet1!$C$3:$L$8'},'9').address,'L8');
});

test('UI refresh reads rows, completion updates SharePoint, then badge/status; failed write preserves ticket',async()=>{
  const ui=await mountRequests();
  assert.match(ui.app.innerHTML,/Klik Refresh untuk membaca data terbaru/);
  ui.click('st-request-refresh');
  await new Promise(resolve=>setImmediate(resolve));
  assert.match(ui.app.innerHTML,/Khusnul Arifin/);
  assert.equal(ui.window.STRequestMenuBadgeCount,1);
  ui.click('st-request-detail','9');
  assert.match(ui.app.innerHTML,/Tandai Sudah Dibaca & Selesai/);
  ui.setFail(true);
  ui.click('st-request-read','9');
  await new Promise(resolve=>setImmediate(resolve));
  assert.match(ui.app.innerHTML,/Gagal menyimpan status/);
  assert.match(ui.app.innerHTML,/Belum ditangani/);
  assert.equal(ui.window.STRequestMenuBadgeCount,1);
  assert.equal(ui.getWrites().length,1);
  ui.setFail(false);
  ui.click('st-request-read','9');
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(ui.window.STRequestMenuBadgeCount,0);
  assert.match(ui.app.innerHTML,/berhasil disimpan sebagai Selesai/);
  assert.equal(ui.getWrites().length,2);
  assert.doesNotMatch(ui.app.innerHTML,/Tandai semua sudah dibaca/);
  const added=[...service.parseWorksheet({values:mockRows}),{
    id:'10',name:'Pegawai Baru',letter:'287/ST/Sj.08/2026',type:'Perubahan',description:'Update',status:'Belum ditangani',completed:'13/11/2026',completedSort:2000000000000
  }];
  ui.setRows(added);
  ui.click('st-request-refresh');
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(ui.window.STRequestMenuBadgeCount,1);
  assert.match(ui.app.innerHTML,/Pegawai Baru/);
});
