const test=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
const E=require('../assets/js/register-engine.js');
function snapshot(rows){return {docs:rows.map(x=>({id:x.id,data:()=>x}))};}
async function boot(){
 const app={innerHTML:'',listeners:{},addEventListener(name,fn){this.listeners[name]=fn;}};
 const items={
   leaveRecords:[{id:'c1',employeeId:'p1',startDate:'2026-10-12',endDate:'2026-10-14',leaveYear:2026,documentNumber:'CT/1',leaveType:'Cuti Tahunan'}],
   assignmentRecords:[{id:'st1',startDate:'2026-10-13',endDate:'2026-10-16',assignmentYear:2026,letterNumber:'ST/1',employees:[{employeeId:'p1'}]}],
   attendanceCases:[],attendanceReconciliations:[],attendanceCaseLogs:[]
 };
 const mockDb={collection:name=>({where:()=>({get:async()=>snapshot(items[name]||[])})})};
 const window={AttendanceRegisterEngine:E,
   UserAccess:{can:()=>true,canView:()=>true,requireModule:async()=>true},
   FirebaseClient:{getDb:()=>mockDb,getCurrentUser:()=>({uid:'p1',email:'tu@pkp.go.id'})},
   MasterDataService:{getBundle:async()=>({employees:[{id:'p1'}],directory:new Map([['p1',{name:'Pegawai Contoh'}]])})},
   AppShell:{render:({content,overlays})=>content+(overlays||'')},
   addEventListener(){},
   firebase:{firestore:{FieldValue:{serverTimestamp:()=>new Date()}}}
 };
 const context={window,document:{getElementById:()=>app,querySelector:()=>null},location:{hash:'#dashboard'},Date,TextEncoder,crypto:require('node:crypto').webcrypto,
   console,alert(){throw Error('Unexpected alert')},setTimeout,URL,Blob,Map,Set,Promise};
 vm.runInNewContext(fs.readFileSync('assets/js/register.js','utf8'),context);
 for(let i=0;i<10&&!app.innerHTML.includes('Pegawai Contoh');i++)await new Promise(r=>setTimeout(r,5));
 return app;
}
test('UI Register memuat data existing dan menampilkan potensi konflik yang nyata',async()=>{
 const app=await boot();
 assert.match(app.innerHTML,/Register Administrasi Kehadiran/);
 assert.match(app.innerHTML,/Pegawai Contoh/);
 assert.match(app.innerHTML,/Perlu tindak lanjut/);
 assert.match(app.innerHTML,/Cuti dan Surat Tugas beririsan/);
 assert.doesNotMatch(app.innerHTML,/Data register gagal dimuat/);
});

test('UI menampilkan formulir klarifikasi dan laporan realisasi tanpa duplikasi modul lama',async()=>{
 const app=await boot();
 const match=app.innerHTML.match(/data-open="([^"]+)"/);
 assert.ok(match,'tombol kasus terdeteksi');
 const target={closest(selector){return selector==='[data-open]'?{dataset:{open:match[1]}}:null;}};
 app.listeners.click({target});
 assert.match(app.innerHTML,/Tindak Lanjut Konflik/);
 assert.match(app.innerHTML,/Simpan &amp; verifikasi TU|Simpan & verifikasi TU/);
 const newTarget={closest(selector){return selector==='[data-new-realization]'?{}:null;}};
 app.listeners.click({target:newTarget});
 assert.match(app.innerHTML,/Status Realisasi Dokumen/);
 assert.match(app.innerHTML,/name="actualDates"/);
});
