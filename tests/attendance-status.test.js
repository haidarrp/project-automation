const test = require('node:test');
const assert = require('node:assert/strict');
const T = require('../assets/js/attendance-status-storage.js');
const I = require('../assets/js/attendance-import.js');

test('ringkasan Tukin hanya menyimpan ID pegawai, tanggal dan status (bukan nilai Tukin)', () => {
  const result = T.fromTukin({id:'run1',employees:[
    {masterEmployeeId:'p1',tukin:8000000,records:{
      '2026-09-30':{status:'WFO',autoTotalPercent:1.5},
      '2026-10-01':{status:'offsite',adjustedPercent:0},
      '2026-10-02':{status:''}}},
    {masterEmployeeId:'',records:{'2026-10-01':{status:'tidak_hadir'}}}
  ]});
  assert.equal(result.unmatched,1);
  assert.deepEqual(result.records.map(r=>r.id),['2026-09_p1_tukin_run1','2026-10_p1_tukin_run1']);
  assert.deepEqual(result.records[1].days,{'2026-10-01':{status:'offsite'}});
  assert.doesNotMatch(JSON.stringify(result.records),/8000000|adjustedPercent|autoTotalPercent/);
});

test('ringkasan Uang Makan hanya menyalin status presensi', () => {
  const records = T.fromUangMakan({id:'um1',ownerUid:'u1'},[{id:'pegawai1',days:{
    '2026-10-10':{status:'Libur',value:0},
    '2026-10-11':{status:'tidak_hadir',value:0}
  }}]);
  assert.equal(records.records.length,1);
  assert.equal(records.records[0].source,'uangMakan');
  assert.equal(Object.keys(records.records[0].days).length,2);
  assert.doesNotMatch(JSON.stringify(records),/"value"/);
});

test('sinkronisasi presensi hanya menulis koleksi ringkasan, tanpa mengubah dokumen sumber', async () => {
  const previous = global.window;
  const writes=[];
  global.window={FirebaseClient:{getCurrentUser:()=>({uid:'u1'}),getDb:()=>({
    collection(name) { assert.equal(name,'attendanceStatusRecords');return {doc:id=>({id})}; },
    batch:()=>({set:(ref, data)=>writes.push({ref,data}),commit:async()=>{}})
  })}};
  try {
    const result=await T.saveFromUangMakan({id:'um1',ownerUid:'u1'},[{id:'p1',days:{'2026-10-12':{status:'Offsite'}}}]);
    assert.equal(result.written,1);
    assert.equal(writes[0].ref.id,'2026-10_p1_uangMakan_um1');
    assert.equal(writes[0].data.updatedByUid,'u1');
    assert.deepEqual(writes[0].data.days,{'2026-10-12':{status:'Offsite'}});
    assert.equal(writes[0].data.sourceRunId,'um1');
  } finally {global.window=previous;}
});

test('impor Excel membaca kolom Status dan identitas master tanpa menyimpan file', async () => {
  const previous = global.window;
  const header=['No','Hari','Tanggal','Masuk','Keluar','Status'];
  const display=[['Laporan'],['Pegawai Uji - 12345678'],header,
    ['1','Selasa','2026-10-13','08:00','17:00','Offsite'],
    ['2','Rabu','14/10/2026','','','tidak_hadir'],
    ['3','Kamis','2026-09-17','','','WFO'],
    ['TOTAL']];
  const sheet={A2:{w:'Pegawai Uji - 12345678'},display};
  global.window={XLSX:{read:()=>({Sheets:{Sheet1:sheet},SheetNames:['Sheet1']}),
    utils:{sheet_to_json:(_sheet,opts)=>display}},
    MasterDataService:{findByIdentity:async(person)=>{
      assert.equal(person.nip,'12345678');return {id:'master1'};
    }}};
  try {
    const file={name:'presensi.xlsx',arrayBuffer:async()=>new ArrayBuffer(0)};
    const rows=await I.parseFile(file,2026);
    assert.equal(rows.length,3);
    assert.deepEqual(rows.map(x=>x.status),['Offsite','tidak_hadir','WFO']);
    assert.equal(rows[1].date,'2026-10-14');
    assert.ok(rows.every(x=>x.employeeId==='master1'));
    assert.equal(I.dateKey('2026-02-30','2026-02-30'),'');
  } finally {global.window=previous;}
});

test('penghapusan riwayat hanya membersihkan snapshot dari sumber yang cocok',async()=>{
 const previous=global.window;
 const deleted=[];
 const docs=[
  {id:'a',ref:{id:'a'},data:()=>({source:'tukin'})},
  {id:'b',ref:{id:'b'},data:()=>({source:'uangMakan'})}
 ];
 global.window={FirebaseClient:{getDb:()=>({
  collection:name=>({where:(field,operator,value)=>{
   assert.equal(name,'attendanceStatusRecords');assert.equal(field,'sourceRunId');
   assert.equal(value,'run1');return {get:async()=>({docs})};
  }}),
  batch:()=>({delete:ref=>deleted.push(ref.id),commit:async()=>{}})
 })}};
 try {
  assert.equal(await T.removeForRun('tukin','run1'),1);
  assert.deepEqual(deleted,['a']);
 }finally{global.window=previous;}
});

test('ringkasan presensi memisahkan sumber dan run agar periode yang berbeda tidak saling menimpa',()=>{
 const e=[{id:'p1',days:{'2026-10-01':{status:'WFO'}}}];
 const a=T.fromUangMakan({id:'old',ownerUid:'u1',year:2026,month:10},e);
 const b=T.fromUangMakan({id:'new',ownerUid:'u1',year:2026,month:10},e);
 assert.notEqual(a.records[0].id,b.records[0].id);
 assert.equal(a.records[0].ownerUid,'u1');
 assert.equal(a.records[0].days['2026-10-01'].status,'WFO');
});
