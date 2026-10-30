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
