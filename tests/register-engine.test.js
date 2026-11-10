const test=require('node:test');
const assert=require('node:assert/strict');
const E=require('../assets/js/register-engine.js');
const leave={id:'cuti1',employeeId:'pegawai1',startDate:'2026-10-12',endDate:'2026-10-14',documentNumber:'C1'};
const st={id:'st1',startDate:'2026-10-13',endDate:'2026-10-16',letterNumber:'ST1',employees:[{employeeId:'pegawai1'},{employeeId:'pegawai2'}]};
test('S11 irisan cuti dan surat tugas per pegawai',()=>{
 const src=E.makeSources([leave],[st]);const issues=E.conflicts(src);
 assert.equal(src.length,3);assert.equal(issues.length,1);assert.equal(issues[0].kind,'cuti_surat');assert.equal(issues[0].startDate,'2026-10-13');assert.equal(issues[0].endDate,'2026-10-14');
});
test('S06 beberapa surat pada tanggal sama ditandai potensi konflik',()=>{
 const issues=E.conflicts(E.makeSources([], [st,{...st,id:'st2',startDate:'2026-10-14',employees:[{employeeId:'pegawai2'}]}]));
 assert.equal(issues.length,1);assert.equal(issues[0].kind,'surat_ganda');
});
test('periode lintas bulan/tahun dan tanggal tidak valid',()=>{
 const rows=E.makeSources([{...leave,startDate:'2026-12-31',endDate:'2027-01-02'}],[]);
 assert.equal(E.inPeriod(rows[0],2027,1),true);
 assert.equal(E.inPeriod(rows[0],2026,11),false);
 assert.equal(E.iso('2026-02-30'),false);
});
test('perubahan sumber mengubah fingerprint rekonsiliasi',()=>{
 const s1=E.makeSources([leave],[st]);const s2=E.makeSources([{...leave,endDate:'2026-10-13'}],[st]);
 assert.notEqual(E.reconciliationFingerprint(s1,[],2026,10),E.reconciliationFingerprint(s2,[],2026,10));
 assert.deepEqual(E.conflicts(s1).map(c=>c.key),E.conflicts(s2).map(c=>c.key));
});
test('bukan konflik jika tanggal tidak bertabrakan atau pegawai beda',()=>{
 const a=E.makeSources([{...leave,startDate:'2026-10-01',endDate:'2026-10-02'}],[st]);
 assert.equal(E.conflicts(a).length,0);
});
test('keenam status presensi dinormalisasi termasuk variasi ejaan',()=>{
 const statuses={WFO:'WFO',WFH:'WFH',Libur:'Libur',Dinas:'Dinas',OFFSITE:'Offsite',off_site:'Offsite',tidak_hadir:'tidak_hadir','Tidak Hadir':'tidak_hadir'};
 for(const [input,want] of Object.entries(statuses))assert.equal(E.normalizeStatus(input),want,input);
 assert.equal(E.normalizeStatus('lainnya'),'');
});
test('Offsite/tidak_hadir perlu dokumen pada pegawai dan tanggal yang sama',()=>{
 const sources=E.makeSources([
  {id:'c1',employeeId:'p1',startDate:'2026-10-03',endDate:'2026-10-04'}
 ],[
  {id:'st1',startDate:'2026-10-06',endDate:'2026-10-07',employees:[{employeeId:'p1'}]}
 ]);
 const rows=E.makeAttendance([
  {employeeId:'p1',days:{'2026-10-02':{status:'Offsite'},'2026-10-03':{status:'tidak_hadir'},
    '2026-10-06':{status:'Offsite'},'2026-10-08':{status:'Tidak Hadir'},
    '2026-10-09':{status:'Libur'},'2026-10-10':{status:'Dinas'},'2026-10-11':{status:'WFH'}}},
  {employeeId:'p2',days:{'2026-10-06':{status:'Offsite'}}}
 ]);
 const issues=E.attendanceIssues(rows,sources);
 assert.deepEqual(issues.map(r=>`${r.employeeId}:${r.startDate}`).sort(),[
  'p1:2026-10-02','p1:2026-10-08','p2:2026-10-06']);
 assert.equal(issues.find(r=>r.startDate==='2026-10-02').kind,'offsite_tanpa_dokumen');
 assert.equal(issues.find(r=>r.startDate==='2026-10-08').kind,'tidak_hadir_tanpa_dokumen');
 assert.deepEqual(E.statusCounts(rows,2026,10),{WFO:0,WFH:1,Libur:1,Dinas:1,Offsite:3,tidak_hadir:2});
});
test('file impor lokal menimpa status snapshot lama pada tanggal yang sama',()=>{
 const stored=[{employeeId:'p1',updatedAt:'2026-10-01',days:{'2026-10-03':{status:'WFH'}}}];
 const imported=[{employeeId:'p1',date:'2026-10-03',status:'Offsite',source:'Impor Excel'}];
 const rows=E.makeAttendance(stored,imported);
 assert.equal(rows.length,1);
 assert.equal(rows[0].normalizedStatus,'Offsite');
});
