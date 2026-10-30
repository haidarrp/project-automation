'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const root=path.join(__dirname,'..');
const app=fs.readFileSync(path.join(root,'assets/js/app.js'),'utf8');
const css=fs.readFileSync(path.join(root,'assets/css/lembur-calendar.css'),'utf8');
const html=fs.readFileSync(path.join(root,'lembur.html'),'utf8');
test('Form checkbox verifikasi harian dihilangkan dan panel verifikasi periode ditampilkan pada review',()=>{
 assert.doesNotMatch(app,/id="lm-verify-document"/);
 assert.match(app,/Verifikasi Cuti \/ Surat Tugas per Periode/);
 assert.match(app,/renderDocumentPeriodReview\(\)/);
 assert.match(app,/data-lm-document-save/);
});
test('Koreksi untuk data bersurat diarahkan ke panel per periode, tidak ke pop-up verifikasi',()=>{
 assert.match(app,/if\(record\.documentSources\?\.length\)\{scrollToDocumentPeriodReview\(\);return;\}/);
 assert.match(app,/data-action="open-period-document-review"/);
 assert.match(app,/saveDocumentPeriodReview/);
});
test('Kalender pegawai dan kalender tanggal menggunakan penanda berwarna untuk perlu verifikasi',()=>{
 assert.match(app,/lm-needs-review/);
 assert.match(app,/info\.pending \? 'lm-needs-review'/);
 assert.match(app,/Perlu Verifikasi/);
 assert.match(css,/\.lm-cal-day\.lm-needs-review/);
 assert.match(css,/\.lm-dot-orange/);
});
test('Berkas HTML mengarah ke versi JavaScript dan CSS yang baru',()=>{
 assert.match(html,/lembur-calendar\.css\?v=2/);
 assert.match(html,/lembur-document-integration\.js\?v=2/);
 assert.match(html,/assets\/js\/app\.js\?v=14/);
});
