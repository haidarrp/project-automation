(function(){
'use strict';
const root=document.getElementById('app');
const R=window.UangMakanRules;
const prev=new Date();prev.setDate(1);prev.setMonth(prev.getMonth()-1);
const state={view:'process',step:1,year:prev.getFullYear(),month:prev.getMonth()+1,master:[],
  files:[],parsed:[],employees:[],leaves:[],assignments:[],supportLoaded:false,history:[],
  busy:false,error:'',message:'',filter:'all',search:'',modalId:'',modalDate:'',savedRun:null,dirty:false,
  holidays:[],holidayNotes:{}};
const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
const permission=key=>window.UserAccess.assert('uangMakan',key);
const fmtTime=n=>n===null||n===undefined?'-':`${String(Math.floor(n/60)).padStart(2,'0')}:${String(n%60).padStart(2,'0')}`;
const dateLabel=k=>{const [y,m,d]=k.split('-').map(Number);return new Intl.DateTimeFormat('id-ID',{weekday:'short',day:'numeric',month:'short',year:'numeric'}).format(new Date(y,m-1,d));};
const WEEKDAYS=['Min','Sen','Sel','Rab','Kam','Jum','Sab'];
const WEEKDAYS_MONDAY=['Sen','Sel','Rab','Kam','Jum','Sab','Min'];
function calendarOffset(k){return (R.weekday(k)+6)%7;}
function holidayStoreKey(year,month){return `pkp-um-libur-${year}-${String(month).padStart(2,'0')}`;}
function cleanHolidaySettings(settings,year,month){
 const valid=new Set(R.monthDates(year,month).filter(k=>!R.weekend(k)));
 const holidays=[...new Set(Array.isArray(settings?.holidays)?settings.holidays:[])].filter(k=>valid.has(k)).sort();
 const holidayNotes={};
 for(const k of holidays)holidayNotes[k]=String(settings?.holidayNotes?.[k]||'').trim().slice(0,120);
 return {holidays,holidayNotes};
}
function readHolidaySettings(year,month){
 try{return cleanHolidaySettings(JSON.parse(localStorage.getItem(holidayStoreKey(year,month))||'{}'),year,month);}
 catch(_){return {holidays:[],holidayNotes:{}};}
}
function persistHolidaySettings(){
 const cleaned=cleanHolidaySettings(state,state.year,state.month);
 state.holidays=cleaned.holidays;state.holidayNotes=cleaned.holidayNotes;
 try{localStorage.setItem(holidayStoreKey(state.year,state.month),JSON.stringify(cleaned));}catch(_){}
}
Object.assign(state,readHolidaySettings(state.year,state.month));
function holidayCount(){return state.holidays.length;}
function calendarNonWorkingDescription(day){
 return day.isHoliday?'Tanggal merah':day.isWeekend?'Akhir pekan':'Hari kerja';
}

function clearNotice(){state.error='';state.message='';}
function fail(error){state.error=String(error?.message||error);state.message='';state.busy=false;render();console.error('[Uang Makan]',error);}
function notify(s){state.error='';state.message=String(s||'');render();}
function summary(){return R.overall(state.employees);}
function pendingCount(){return state.employees.filter(e=>R.overview(e).status!=='Siap').length;}
function makeEmptyEmployee(e){return {...e,files:[],fileWarnings:[],coverageAcknowledged:false,coverageReason:'',days:{}};}
function resetProcess(){state.files=[];state.parsed=[];state.employees=[];state.leaves=[];state.assignments=[];state.supportLoaded=false;state.modalId='';state.modalDate='';state.savedRun=null;state.dirty=false;state.search='';state.filter='all';state.step=1;clearNotice();}
function pageTitle(view){return {process:'Proses Uang Makan',dashboard:'Dashboard Uang Makan',history:'Riwayat Uang Makan'}[view]||'Uang Makan';}
function metric(label,value,description){return `<div class="um-panel um-kpi"><small>${esc(label)}</small><strong>${esc(value)}</strong><span>${esc(description||'')}</span></div>`;}
function showStatus(){return `${state.error?`<div class="um-note error" role="alert" style="margin-bottom:13px">${esc(state.error)}</div>`:''}${state.message?`<div class="um-note good" role="status" style="margin-bottom:13px">${esc(state.message)}</div>`:''}`;}
function header(title,description,actions=''){
 return `<div class="um-header"><div><h1>${esc(title)}</h1><p class="um-desc">${esc(description)}</p></div><div class="um-row">${actions}</div></div>`;
}
function stepper(){
 const steps=['Periode','Unggah Absensi','Verifikasi','Pratinjau & Ekspor'];
 return `<div class="um-steps">${steps.map((text,i)=>`${i?'<div class="um-step-rule"></div>':''}<button type="button" class="um-step ${state.step===i+1?'active':state.step>i+1?'passed':''}" data-um-action="step" data-step="${i+1}" ${i+1>3&&state.employees.length===0?'disabled':''}><strong>${state.step>i+1?'✓':i+1}</strong>${text}</button>`).join('')}</div>`;
}
function holidayPanel(){
 const dates=R.monthDates(state.year,state.month),start=dates[0];
 const weekends=dates.filter(k=>R.weekend(k)).length;
 const cells=Array.from({length:calendarOffset(start)},()=>'<span class="um-calendar-blank" aria-hidden="true"></span>').join('')+
   dates.map(k=>{
     const day=Number(k.slice(-2)),weekend=R.weekend(k),selected=state.holidays.includes(k);
     const label=weekend?'Akhir pekan':selected?'Tanggal merah':'Hari kerja';
     return `<button type="button" class="um-holiday-day ${weekend?'weekend':''} ${selected?'holiday':''}" data-um-action="toggle-holiday" data-date="${k}" ${weekend?'disabled':''} aria-label="${esc(dateLabel(k))} - ${label}" aria-pressed="${selected}" title="${weekend?'Sabtu/Minggu otomatis tidak dihitung':selected?'Klik untuk menghapus tanggal merah':'Klik untuk menetapkan tanggal merah'}"><span>${day}</span><small>${weekend?'Libur':selected?'Tanggal merah':''}</small></button>`;
   }).join('');
 const selected=state.holidays.map(k=>`<div class="um-holiday-item"><div class="um-holiday-item-date"><strong>${esc(dateLabel(k))}</strong><small>Tidak dihitung meskipun ada presensi</small></div><input type="text" data-um-holiday-note="${esc(k)}" aria-label="Keterangan tanggal merah ${esc(dateLabel(k))}" maxlength="120" placeholder="Keterangan (opsional)" value="${esc(state.holidayNotes[k]||'')}"><button type="button" class="um-btn small" data-um-action="remove-holiday" data-date="${k}">Hapus</button></div>`).join('');
 return `<section class="um-holiday-settings" aria-label="Pengaturan tanggal merah"><div class="um-panel-title"><div><h3>Tanggal Merah</h3><p>Pilih tanggal libur di luar Sabtu–Minggu. Klik tanggal pada kalender untuk menandai atau membatalkan tanggal merah.</p></div><span class="um-pill">${holidayCount()} tanggal dipilih</span></div>
 <div class="um-holiday-layout"><div><div class="um-calendar-weekdays">${WEEKDAYS_MONDAY.map(n=>`<span>${n}</span>`).join('')}</div><div class="um-holiday-calendar">${cells}</div><div class="um-legend"><span><i class="um-legend-dot work"></i>Hari kerja</span><span><i class="um-legend-dot weekend"></i>Sabtu–Minggu</span><span><i class="um-legend-dot holiday"></i>Tanggal merah</span></div></div>
 <div class="um-holiday-list"><div class="um-holiday-list-head">Tanggal merah yang dipilih</div>${selected||'<div class="um-holiday-empty">Belum ada tanggal merah tambahan. Sabtu dan Minggu tetap otomatis tidak memperoleh uang makan.</div>'}</div></div>
 <div class="um-holiday-footer">${weekends} hari akhir pekan + ${holidayCount()} tanggal merah tambahan tidak dihitung. Pilihan tersimpan pada browser ini dan ikut tercatat saat hasil disimpan ke Riwayat.</div></section>`;
}
function periodPanel(){
 const years=Array.from({length:9},(_,i)=>new Date().getFullYear()-4+i);
 return `<section class="um-panel"><div class="um-panel-title"><div><h2>Tentukan Periode Uang Makan</h2><p>Penghitungan mengikuti tanggal 1 hingga akhir bulan, dengan Sabtu–Minggu dan tanggal merah tidak dihitung.</p></div></div><div class="um-grid">
 <div class="um-field"><label for="um-year">Tahun</label><select id="um-year">${years.map(y=>`<option value="${y}" ${y===state.year?'selected':''}>${y}</option>`).join('')}</select></div>
 <div class="um-field"><label for="um-month">Bulan</label><select id="um-month">${R.MONTHS.map((name,i)=>`<option value="${i+1}" ${state.month===i+1?'selected':''}>${name}</option>`).join('')}</select></div>
 <div class="um-field"><label>Rentang Absensi</label><input type="text" readonly value="01–${R.periodDays(state.year,state.month)} ${R.MONTHS[state.month-1]} ${state.year}"></div>
 </div>${holidayPanel()}<div class="um-actions"><span class="um-fine">Presensi pada Sabtu, Minggu, dan tanggal merah selalu bernilai 0 hari uang makan.</span><button class="um-btn primary" type="button" data-um-action="period-next">Lanjut: Unggah Absensi →</button></div></section>`;
}
function uploadPanel(){
 return `<section class="um-panel"><div class="um-panel-title"><div><h2>Unggah Absensi — ${R.periodLabel(state.year,state.month)}</h2><p>Pilih file absensi per pegawai seperti yang digunakan pada modul Tukin. Dapat memilih banyak file sekaligus.</p></div><span class="um-pill">${state.files.length} file</span></div>
 <div class="um-note" style="margin-bottom:14px">Hanya pegawai yang memiliki file absensi diunggah yang dihitung dan masuk Excel. Pegawai di master tanpa file tidak disertakan. Hari Sabtu–Minggu dan ${holidayCount()} tanggal merah tambahan tetap tidak dihitung.</div>
 <div id="um-drop" class="um-drop" role="region" aria-label="Area unggah absensi"><div class="um-drop-icon">⇧</div><strong>Tarik file Excel absensi ke sini</strong><p>Format .xlsx atau .xls, satu atau beberapa file per pegawai.</p><button class="um-btn" type="button" data-um-action="upload">Pilih File</button><input id="um-files" type="file" accept=".xls,.xlsx" multiple hidden></div>
 ${state.files.length?`<div style="margin-top:13px">${state.files.map((f,i)=>`<div class="um-file"><span>📄</span><strong title="${esc(f.name)}">${esc(f.name)}</strong><small>${(f.size/1024).toFixed(1)} KB</small><button class="um-btn small" type="button" data-um-action="remove-file" data-index="${i}" aria-label="Hapus file ${esc(f.name)}">Hapus</button></div>`).join('')}</div>`:''}
 <div class="um-actions"><button class="um-btn" type="button" data-um-action="back">← Kembali</button><div class="um-row"><button class="um-btn gold" type="button" data-um-action="clear-files" ${state.files.length?'':'disabled'}>Bersihkan</button><button class="um-btn primary" type="button" data-um-action="calculate" ${!state.files.length||state.busy?'disabled':''}>${state.busy?'Memproses...':'Hitung & Sinkronkan →'}</button></div></div></section>`;
}
function visibleEmployees(){
 return state.employees.filter(e=>{
  const q=state.search.trim().toLowerCase();
  if(q&&![e.name,e.nip,e.anakSatker,e.jenisPegawai].some(x=>String(x||'').toLowerCase().includes(q)))return false;
  const s=R.overview(e);return state.filter==='all'||(state.filter==='pending'&&s.status!=='Siap')||(state.filter==='ready'&&s.status==='Siap');
 });
}
function verifyPanel(){
 const s=summary();
 return `${kpis(s)}<section class="um-panel"><div class="um-panel-title"><div><h2>Validasi & Verifikasi Uang Makan</h2><p>Rekap kehadiran fisik pegawai untuk sheet PUSDATIN. Klik Detail untuk memeriksa setiap tanggal.</p></div><div class="um-row"><button class="um-btn small" data-um-action="sync" type="button" ${state.busy?'disabled':''}>⟳ Sinkronkan Cuti & Surat Tugas</button></div></div>
 ${pendingCount()?`<div class="um-note warn">Terdapat ${pendingCount()} pegawai dari berkas yang diunggah yang perlu pemeriksaan, konfirmasi cakupan tanggal, atau kelengkapan jenis pegawai. Ekspor final tersedia setelah pegawai yang diproses siap.</div>`:''}
 <div class="um-note" style="margin-bottom:12px">Periode ${esc(R.periodLabel(state.year,state.month))}: Sabtu–Minggu dan ${holidayCount()} tanggal merah tambahan otomatis tidak memperoleh uang makan.</div>
 <div class="um-toolbar"><input class="um-search" id="um-search" placeholder="Cari nama pegawai atau NIP..." value="${esc(state.search)}"><div class="um-row"><select id="um-filter"><option value="all" ${state.filter==='all'?'selected':''}>Semua pegawai</option><option value="pending" ${state.filter==='pending'?'selected':''}>Perlu Verifikasi</option><option value="ready" ${state.filter==='ready'?'selected':''}>Siap</option></select><span class="um-pill good">Hitung berdasarkan kehadiran fisik</span></div></div>
 <div class="um-table-outer"><table class="um-table"><thead><tr><th>No.</th><th>Nama / NIP</th><th>Anak Satker</th><th>Jenis</th><th>Hari Hadir</th><th>Cuti</th><th>Dinas Luar</th><th>Hari Uang Makan</th><th>Status</th><th>Aksi</th></tr></thead><tbody>
 ${visibleEmployees().length?visibleEmployees().map((e,i)=>{const a=R.overview(e);return `<tr><td>${i+1}</td><td><div class="name">${esc(e.name)}</div><div class="sub">${esc(e.nip)}</div></td><td>${esc(e.anakSatker||'-')}</td><td>${esc(e.jenisPegawai||'-')}</td><td class="number">${a.present}</td><td class="number">${a.leave}</td><td class="number">${a.travel}</td><td class="number">${a.days}</td><td><span class="um-pill ${a.status==='Siap'?'good':'warn'}">${esc(a.status)}</span>${a.issues?`<div class="sub">${a.issues} tanggal perlu cek</div>`:''}</td><td><button type="button" class="um-btn small" data-um-action="verify" data-id="${esc(e.id)}">Detail →</button></td></tr>`;}).join(''):`<tr><td colspan="10" class="um-empty">Tidak ada pegawai dari file unggahan yang sesuai filter.</td></tr>`}
 </tbody></table></div><div class="um-summary"><span>${visibleEmployees().length} dari ${state.employees.length} pegawai hasil unggah</span><span><b>${s.days} hari</b> uang makan · ${s.issues} tanggal belum diperiksa</span></div>
 <div class="um-actions"><button type="button" class="um-btn" data-um-action="back">← Kembali ke Unggah</button><div class="um-row"><button type="button" class="um-btn" data-um-action="restart">Ubah Periode</button><button type="button" class="um-btn primary" data-um-action="preview">Pratinjau Output Excel →</button></div></div></section>`;
}
function kpis(s){return `<div class="um-kpis">${metric('Pegawai Diproses',s.employees,'Dari file absensi yang diunggah')}${metric('Hari Uang Makan',s.days,'Total pegawai yang diproses')}${metric('Tanggal Perlu Cek',s.issues,'Belum diverifikasi')}${metric('Koreksi Manual',s.manual,'Perubahan bernotasi')}</div>`;}
function previewPanel(){
 const s=summary(),pending=pendingCount(),ready=pending===0;
 const shown=state.employees.slice(0,Math.min(12,state.employees.length));
 return `${kpis(s)}<section class="um-panel"><div class="um-panel-title"><div><h2>Pratinjau Sheet PUSDATIN</h2><p>Kolom tanggal diisi angka 1 apabila memperoleh uang makan, dan kosong untuk tanggal lainnya.</p></div><span class="um-pill ${ready?'good':'warn'}">${ready?'Siap Ekspor':`${pending} pegawai belum siap`}</span></div>
 ${!ready?`<div class="um-note warn" style="margin-bottom:13px">Selesaikan verifikasi dan kelengkapan data pegawai yang tercantum pada file absensi. Pegawai master yang tidak diunggah tidak memengaruhi ekspor.</div>`:''}
 <div class="um-table-outer um-preview-scroll"><table class="um-table export"><thead><tr><th>No.</th><th>NIP</th><th>Nama</th>${Array.from({length:R.periodDays(state.year,state.month)},(_,i)=>`<th class="um-preview-day">${i+1}</th>`).join('')}<th>Jumlah</th><th>Unit Kerja</th><th>Bank</th><th>Jenis Pegawai</th></tr></thead><tbody>${shown.map((e,i)=>`<tr><td>${i+1}</td><td>${esc(e.nip)}</td><td class="name">${esc(e.name)}</td>${R.monthDates(state.year,state.month).map(k=>`<td class="number um-preview-day">${e.days[k]?.value===1?'1':''}</td>`).join('')}<td class="number">${R.overview(e).days}</td><td>Pusat Data dan Informasi</td><td>${esc(e.anakSatker||'')}</td><td>${esc(e.jenisPegawai||'')}</td></tr>`).join('')}</tbody></table></div>
 ${state.employees.length>shown.length?`<p class="um-fine" style="margin-top:10px">Menampilkan ${shown.length} pegawai pertama dari ${state.employees.length}. File hasil hanya mencakup pegawai yang diunggah.</p>`:''}
 <div class="um-summary um-preview-period"><span>Periode ${esc(R.periodLabel(state.year,state.month))}</span></div>
 <div class="um-actions"><button class="um-btn" type="button" data-um-action="back">← Kembali ke Verifikasi</button><div class="um-row"><button class="um-btn" type="button" data-um-action="save" ${!ready||state.busy?'disabled':''}>${state.busy?'Menyimpan...':'Simpan ke Riwayat'}</button><button class="um-btn primary" type="button" data-um-action="export" ${!ready||state.busy?'disabled':''}>Unduh Excel PUSDATIN ↓</button></div></div></section>`;
}
function periodCard(){return `<section class="um-panel" style="margin-bottom:16px"><h2>Periode Saat Ini</h2><p>${esc(R.periodLabel(state.year,state.month))}</p><div style="margin-top:14px"><button type="button" class="um-btn primary" data-um-action="goto-process">Proses Uang Makan →</button></div></section>`;}
function historyTable(embedded=false){
 return `<div class="um-table-outer"><table class="um-table"><thead><tr><th>Periode</th><th>Diproses</th><th>Jumlah Pegawai</th><th>Hari Uang Makan</th><th>Operator</th><th>Aksi</th></tr></thead><tbody>${state.history.length?state.history.slice(0,embedded?7:75).map(h=>`<tr><td class="name">${esc(R.periodLabel(h.year,h.month))}</td><td>${esc(new Date(h.createdAt).toLocaleString('id-ID'))}</td><td>${Number(h.summary?.employees||0)}</td><td>${Number(h.summary?.days||0)}</td><td>${esc(h.ownerName||h.ownerEmail||'-')}</td><td><div class="um-row"><button class="um-btn small" type="button" data-um-action="history-open" data-id="${esc(h.id)}">Lihat</button><button class="um-btn small" type="button" data-um-action="history-download" data-id="${esc(h.id)}">Excel ↓</button>${!embedded&&((h.ownerUid===window.FirebaseClient?.getCurrentUser?.()?.uid)||window.FirebaseClient?.isAdmin?.())?`<button class="um-btn small danger" type="button" data-um-action="history-delete" data-id="${esc(h.id)}">Hapus</button>`:''}</div></td></tr>`).join(''):`<tr><td class="um-empty" colspan="6">Belum ada riwayat uang makan tersimpan.</td></tr>`}</tbody></table></div>`;
}
function dashboard(){
 const total=state.history.reduce((s,h)=>s+Number(h.summary?.days||0),0);
 return `${header('Dashboard Uang Makan','Ringkasan proses uang makan berdasarkan periode kalender.',`<button class="um-btn primary" type="button" data-um-action="goto-process">+ Proses Uang Makan</button>`)}${showStatus()}
 <div class="um-kpis">${metric('Total Proses',state.history.length,'Riwayat tersimpan')}${metric('Akumulasi Hari',total,'Dari riwayat terbaru')}${metric('Periode Terakhir',state.history[0]?R.periodLabel(state.history[0].year,state.history[0].month):'-','Sudah diproses')}${metric('Status','Aktif','Siap unggah absensi')}</div>
 <section class="um-panel"><div class="um-panel-title"><div><h2>Riwayat Terbaru</h2><p>Rekap proses yang telah disimpan.</p></div><button type="button" class="um-btn small" data-um-action="goto-history">Lihat Semua</button></div>${historyTable(true)}</section>`;
}
function history(){return `${header('Riwayat Uang Makan','Daftar rekap yang telah disimpan dan dapat diunduh ulang.',`<button type="button" class="um-btn primary" data-um-action="goto-process">+ Proses Baru</button>`)}${showStatus()}<section class="um-panel">${historyTable()}</section>`;}
function process(){
 return `${header('Proses Uang Makan','Unggah data absensi, periksa hasil sinkronisasi cuti dan surat tugas, kemudian hasilkan Excel PUSDATIN.',`<button class="um-btn" type="button" data-um-action="goto-history">⟲ Riwayat Proses</button>`)}${showStatus()}${stepper()}
 ${state.step===1?periodPanel():state.step===2?uploadPanel():state.step===3?verifyPanel():previewPanel()}`;
}
function findEmployee(id){return state.employees.find(e=>e.id===id);}
function modal(){
 const e=findEmployee(state.modalId);if(!e)return '';
 const a=R.overview(e),keys=R.monthDates(state.year,state.month);
 const day=e.days?.[state.modalDate]||e.days?.[keys[0]];
 if(!day)return '';
 const isProblem=day.issues.length>0;
 return `<div class="um-modal-bg" data-um-action="close-modal"><section class="um-modal" role="dialog" aria-modal="true" aria-label="Detail Uang Makan">
 <div class="um-modal-head"><div><h2>${esc(e.name)}</h2><p>${esc(e.nip)} · ${esc(e.jenisPegawai||'Jenis belum dilengkapi')} · Anak Satker ${esc(e.anakSatker||'-')} · ${esc(R.periodLabel(state.year,state.month))}</p></div><button class="um-btn small" type="button" data-um-action="close-modal">Tutup ×</button></div>
 <div class="um-modal-body"><div class="um-row"><span class="um-pill good">${a.days} hari uang makan</span><span class="um-pill ${a.issues?'warn':'good'}">${a.issues} tanggal perlu cek</span><span class="um-pill">${e.files?.length||0} file absensi</span></div>
 ${!e.jenisPegawai?'<div class="um-note warn" style="margin-top:14px">Jenis pegawai belum terisi pada master. Lengkapi melalui menu Administrasi → Master Data.</div>':''}
 ${e.fileWarnings?.length?`<div class="um-note warn" style="margin-top:13px">${e.fileWarnings.map(v=>esc(v)).join(' · ')}${e.coverageAcknowledged?`<div>Sudah ditinjau: ${esc(e.coverageReason||'-')}</div>`:`<div class="um-row" style="margin-top:9px"><input class="um-search" style="max-width:390px" id="um-coverage-reason" placeholder="Catatan verifikasi cakupan absensi"><button type="button" class="um-btn small" data-um-action="acknowledge" data-kind="coverage">Konfirmasi Cakupan</button></div>`}</div>`:''}
 <div class="um-weekday-row">${WEEKDAYS_MONDAY.map(day=>`<span>${day}</span>`).join('')}</div>
 <div class="um-day-grid">${Array.from({length:calendarOffset(keys[0])},()=>'<span class="um-day-placeholder" aria-hidden="true"></span>').join('')}${keys.map(k=>{const d=e.days[k],warn=d.issues.length&&!d.reviewed,cls=d.value===1?'earned':'';
   const status=d.isHoliday?'Tanggal merah':d.isWeekend?'Akhir pekan':warn?'Perlu verifikasi':d.value===1?'1 hari':'Tidak dibayar';
   return `<button class="um-day ${cls} ${warn?'warn':''} ${state.modalDate===k?'selected':''} ${d.isWeekend?'weekend':''} ${d.isHoliday?'holiday':''}" type="button" data-um-action="select-day" data-date="${k}" aria-pressed="${state.modalDate===k}" aria-label="${esc(dateLabel(k))}, ${esc(status)}" title="${esc(d.issues.join('; ')||status)}"><span class="um-day-week">${WEEKDAYS[R.weekday(k)]}</span><strong>${Number(k.slice(-2))}</strong><small>${d.nonWorking?'Libur':d.value===1?'1 hari':warn?'Periksa':'–'}</small></button>`;}).join('')}</div>
 <div class="um-detail"><h3>Verifikasi ${esc(dateLabel(day.key))}</h3><div class="um-detail-grid"><div><div class="um-fine"><b>Jam masuk:</b> ${esc(fmtTime(day.inMinutes))} · <b>Jam pulang:</b> ${esc(fmtTime(day.outMinutes))}<br><b>Status absensi:</b> ${esc(day.status||'—')}<br><b>Kode:</b> ${esc(day.flags.join(', ')||'—')}<br><b>Cuti:</b> ${day.leaves.length?esc(day.leaves.map(x=>x.leaveType||'Cuti').join(', ')):'—'}<br><b>Surat tugas:</b> ${day.assignments.length?esc(day.assignments.map(x=>x.activity||x.letterNumber||'Surat Tugas').join(', ')):'—'}<br><b>Hasil otomatis:</b> ${day.automatic===1?'1 hari':'Tidak dibayar'}</div>
  ${day.nonWorking?'<div class="um-note" style="margin-top:11px">Tanggal tidak termasuk hari yang memperoleh uang makan.</div>':isProblem?`<div class="um-note warn" style="margin-top:11px">${day.issues.map(x=>esc(x)).join(' · ')}</div>`:'<div class="um-note good" style="margin-top:11px">Tidak ditemukan konflik otomatis.</div>'}</div>
 <div>${day.nonWorking?`<div class="um-note um-note-nonworking"><b>${day.isHoliday?'Tanggal merah':'Sabtu/Minggu'}</b><p>Hari ini otomatis bernilai 0. Presensi atau verifikasi manual tidak dapat mengubahnya menjadi hari uang makan.</p>${day.isHoliday&&state.holidayNotes[day.key]?`<p>Keterangan: ${esc(state.holidayNotes[day.key])}</p>`:''}</div>`:`<div class="um-field"><label for="um-manual">Hasil Setelah Verifikasi</label><select id="um-manual"><option value="" ${day.manualValue===null?'selected':''}>Ikuti otomatis (${day.automatic} hari)</option><option value="1" ${day.manualValue===1?'selected':''}>1 — Memperoleh Uang Makan</option><option value="0" ${day.manualValue===0?'selected':''}>0 — Tidak Memperoleh Uang Makan</option></select></div><div class="um-field" style="margin-top:11px"><label for="um-note">Catatan Verifikasi / Alasan Koreksi</label><textarea id="um-note" placeholder="Wajib diisi jika ada konflik atau koreksi manual">${esc(day.manualNote||'')}</textarea></div><div class="um-row" style="margin-top:12px"><button class="um-btn primary small" type="button" data-um-action="save-day">Simpan Verifikasi Tanggal</button>${day.reviewed?'<span class="um-pill good">Sudah diverifikasi</span>':''}</div>`}</div></div></div>
 </div><div class="um-modal-footer"><span class="um-fine">Perubahan tercatat bersama catatan, operator, dan waktu verifikasi.</span><button class="um-btn" type="button" data-um-action="close-modal">Selesai</button></div>
 </section></div>`;
}
function render(){
 if(!window.AppShell||!window.FirebaseClient?.getCurrentUser?.())return;
 const content=`<div class="um-page ${state.busy?'um-busy':''}">${state.view==='dashboard'?dashboard():state.view==='history'?history():process()}</div>`;
 root.innerHTML=window.AppShell.render({module:'uangMakan',view:state.view,viewLabel:pageTitle(state.view),content,overlays:modal()});
 window.UserAccess?.scheduleDecoration?.();
}
async function loadMaster(){
 const bundle=await window.MasterDataService.getBundle(true);
 state.master=bundle.employees.filter(x=>x.active!==false).map(x=>{
   const identity=bundle.directory.get(x.id)||{};
   return {id:x.id,name:String(identity.name||'').trim(),nip:R.normalizeNip(identity.nip),
     anakSatker:String(x.anakSatker||'').trim(),jenisPegawai:String(x.jenisPegawai||'').toUpperCase(),
     unit:'Pusat Data dan Informasi',order:Number(x.order||9999)};
 }).sort((a,b)=>a.order-b.order||a.name.localeCompare(b.name,'id'));
 // Data master tidak divalidasi secara global. Hanya pegawai dari berkas yang diunggah
 // yang perlu dicocokkan dan divalidasi saat proses perhitungan.
}
async function getSupportingData(){
 const year=state.year,month=state.month;
 const from=R.key(year,month,1),to=R.key(year,month,R.periodDays(year,month));
 const years=[year-1,year];
 const [leaveSnaps,assignmentSnaps]=await Promise.all([
   Promise.all(years.map(y=>window.FirebaseClient.getDb().collection('leaveRecords').where('leaveYear','==',y).get())),
   Promise.all(years.map(y=>window.FirebaseClient.getDb().collection('assignmentRecords').where('assignmentYear','==',y).get()))
 ]);
 const ls=new Map(),as=new Map();
 leaveSnaps.flatMap(s=>s.docs).forEach(doc=>{
   const x=doc.data()||{},startDate=String(x.startDate||''),endDate=String(x.endDate||x.startDate||'');
   if(x.employeeId&&startDate<=to&&endDate>=from){ls.set(doc.id,{id:doc.id,employeeId:String(x.employeeId),startDate,endDate,leaveType:String(x.leaveType||'Cuti'),documentNumber:String(x.documentNumber||'')});}
 });
 assignmentSnaps.flatMap(s=>s.docs).forEach(doc=>{
   const x=doc.data()||{},startDate=String(x.startDate||''),endDate=String(x.endDate||x.startDate||'');
   const type=['official_travel','work_arrangement','other'].includes(x.assignmentType)?x.assignmentType:'official_travel';
   const ids=(x.employees||[]).map(e=>String(e.employeeId||e.id||'')).filter(Boolean);
   if(ids.length&&startDate<=to&&endDate>=from)as.set(doc.id,{id:doc.id,employeeIds:ids,startDate,endDate,
     assignmentType:type,activity:String(x.activity||''),letterNumber:String(x.letterNumber||'')});
 });
 return {leaves:Array.from(ls.values()),assignments:Array.from(as.values())};
}
function computeEmployees(list,leaves,assignments){
 return list.map(e=>{
   const prior=e.days||{};
   const updated=R.applySupporting(e,leaves,assignments,state.year,state.month,state.holidays);
   for(const [key,d] of Object.entries(updated.days)){
     const old=prior[key];
     if(old?.reviewed&&(old.issues||[]).join('|')!==d.issues.join('|')){
       d.reviewed=false;d.reviewedAt='';d.reviewedBy='';
     }
   }
   return updated;
 });
}
async function calculate(){
 if(!permission('process.calculate')||state.busy)return;
 if(!state.files.length){notify('Unggah minimal satu berkas absensi.');return;}
 state.busy=true;clearNotice();render();
 try{
  await window.ResourceLoader.ensure('xlsx');
  const parsed=await window.UangMakanParser.parseFiles(state.files,state.year,state.month);
  const faulty=parsed.filter(x=>!x.ok);
  if(faulty.length)throw new Error('Beberapa file gagal dibaca:\n'+faulty.map(x=>`${x.fileName}: ${x.error}`).join('\n'));
  // Berkas unggahan adalah satu-satunya sumber daftar pegawai yang diproses.
  // Master hanya berfungsi sebagai referensi identitas, Anak Satker, dan jenis pegawai.
  const employees=[],selectedById=new Map(),errors=[],list=[];
  for(const item of parsed){
    const fileNip=R.normalizeNip(item.identity.nip);
    const normalizedName=R.normalizeName(item.identity.name);
    // Utamakan NIP, karena nama bisa sama atau berbeda penulisan pada laporan.
    const candidates=fileNip
      ?state.master.filter(e=>e.nip&&e.nip===fileNip)
      :state.master.filter(e=>e.name&&R.normalizeName(e.name)===normalizedName);
    let master=candidates.length===1?candidates[0]:null;
    // Lookup hash tetap mendukung variasi nama/alias yang terdaftar di master.
    if(!master&&candidates.length===0){
      const found=await window.MasterDataService.findByIdentity(item.identity);
      master=state.master.find(e=>e.id===found?.id)||null;
      // Jangan mencocokkan berdasarkan nama jika NIP di berkas bertentangan.
      if(master&&fileNip&&master.nip&&master.nip!==fileNip)master=null;
    }
    if(!master){
      errors.push(`${item.fileName}: ${candidates.length>1?'identitas ganda pada master':'identitas tidak ditemukan atau NIP tidak cocok dengan master aktif'} (${item.identity.name} / ${fileNip||'-'}).`);
      continue;
    }
    if(!selectedById.has(master.id)){
      const e=makeEmptyEmployee({...master,
        // Izinkan data direktori yang belum lengkap hanya jika ada identitas
        // pengganti pada berkas absensi yang sama.
        name:master.name||String(item.identity.name||'').trim(),
        nip:master.nip||fileNip,
        order:employees.length+1});
      if(!e.name||!e.nip){
        errors.push(`${item.fileName}: nama atau NIP pegawai yang diunggah belum lengkap. Lengkapi identitas pada Master Pegawai / file absensi.`);
        continue;
      }
      selectedById.set(master.id,e);
      employees.push(e);
    }
    const e=selectedById.get(master.id);
    if(!e)continue;
    // Hindari dua berkas dengan NIP berbeda digabung hanya karena nama sama.
    if(fileNip&&e.nip&&fileNip!==e.nip){
      errors.push(`${item.fileName}: NIP ${fileNip} berbeda dari NIP berkas sebelumnya (${e.nip}) untuk pegawai ${e.name}.`);
      continue;
    }
    e.files.push(item.fileName);
    list.push({file:item.fileName,employee:e.name});
    const prior=e.days;
    for(const [key,record] of Object.entries(item.rows)){
      if(prior[key]){
        const old=prior[key];
        prior[key]={...record,inMinutes:record.inMinutes??old.inMinutes,outMinutes:record.outMinutes??old.outMinutes,
          flags:[...new Set([...(old.flags||[]),...(record.flags||[])])],status:record.status||old.status,duplicate:true};
      }else prior[key]=record;
    }
  }
  if(errors.length)throw new Error('Periksa identitas hanya pada berkas yang diunggah:\n'+errors.join('\n'));
  if(!employees.length)throw new Error('Tidak ditemukan pegawai valid dari file absensi yang diunggah.');
  for(const e of employees){
    if(!e.files.length)continue;
    const observed=Object.keys(e.days).sort(),start=R.key(state.year,state.month,1),end=R.key(state.year,state.month,R.periodDays(state.year,state.month));
    if(observed[0]>start)e.fileWarnings.push(`Tanggal pertama yang tersedia ${observed[0]} (seharusnya ${start})`);
    if(observed[observed.length-1]<end)e.fileWarnings.push(`Tanggal terakhir yang tersedia ${observed[observed.length-1]} (seharusnya ${end})`);
  }
  const supports=await getSupportingData();
  state.leaves=supports.leaves;state.assignments=supports.assignments;state.supportLoaded=true;
  state.employees=computeEmployees(employees,state.leaves,state.assignments);
  state.parsed=list;state.step=3;state.savedRun=null;state.dirty=true;
  state.message=`Berhasil mengolah ${parsed.length} berkas untuk ${employees.length} pegawai yang diunggah. Pegawai master lainnya tidak ikut dihitung; cuti dan surat tugas telah disinkronkan.`;
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
async function sync(){
 if(!permission('process.calculate')||state.busy)return;
 state.busy=true;clearNotice();render();
 try{
   const supports=await getSupportingData();
   state.leaves=supports.leaves;state.assignments=supports.assignments;
   state.employees=computeEmployees(state.employees,supports.leaves,supports.assignments);state.dirty=true;
   state.message=`Sinkronisasi selesai: ${supports.leaves.length} data cuti, ${supports.assignments.length} surat tugas.`;
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
function saveDay(){
 if(!permission('process.verify'))return;
 const e=findEmployee(state.modalId),d=e?.days[state.modalDate];if(!d)return;
 const raw=document.getElementById('um-manual')?.value||'';
 const manual=raw===''?null:Number(raw),reason=String(document.getElementById('um-note')?.value||'').trim();
 if(d.nonWorking){alert('Sabtu/Minggu atau tanggal merah selalu bernilai 0 dan tidak dapat dikoreksi menjadi 1.');return;}
 if((manual!==null||d.issues.length)&&!reason){alert('Catatan verifikasi wajib diisi untuk tanggal yang memiliki konflik atau dikoreksi.');return;}
 if(manual!==null&&manual!==0&&manual!==1){alert('Pilihan hasil verifikasi tidak valid.');return;}
 d.manualValue=manual;d.manualNote=reason;d.value=manual===null?d.automatic:manual;d.reviewed=true;
 d.reviewedAt=new Date().toISOString();d.reviewedBy=window.FirebaseClient.getCurrentUser()?.email||'';
 state.dirty=true;state.savedRun=null;notify('Verifikasi tanggal berhasil dicatat.');
}
function acknowledge(kind){
 if(!permission('process.verify')||kind!=='coverage')return;
 const e=findEmployee(state.modalId);if(!e)return;
 const reason=String(document.getElementById('um-coverage-reason')?.value||'').trim();
 if(reason.length<5){alert('Isi alasan atau catatan minimal 5 karakter.');return;}
 e.coverageAcknowledged=true;e.coverageReason=reason;
 state.dirty=true;state.savedRun=null;notify('Konfirmasi cakupan absensi sudah dicatat.');
}
function allowFinal(){
 if(!state.employees.length)throw new Error('Belum ada hasil perhitungan.');
 if(pendingCount())throw new Error(`${pendingCount()} pegawai belum siap. Periksa status verifikasi terlebih dahulu.`);
}
async function exportData(){
 if(!permission('process.export'))return;
 try{allowFinal();await window.ResourceLoader.ensure('jszip');const result=await window.UangMakanExport.generate(state.year,state.month,state.employees,{holidays:state.holidays});notify(`File ${result.filename} berhasil dibuat.`);}
 catch(e){fail(e);}
}
async function saveRun(){
 if(!permission('process.save')||state.busy)return;
 try{allowFinal();}catch(e){fail(e);return;}
 if(!state.dirty&&state.savedRun){notify('Hasil ini sudah tersimpan pada riwayat.');return;}
 state.busy=true;clearNotice();render();
 try{const run=await window.UangMakanStorage.save(state.year,state.month,state.employees,summary(),{holidays:state.holidays,holidayNotes:state.holidayNotes});state.savedRun=run;state.dirty=false;await refreshHistory();state.message='Hasil periode '+R.periodLabel(state.year,state.month)+' tersimpan di Riwayat.';}
 catch(e){fail(e);return;}
 state.busy=false;render();
}
async function refreshHistory(){state.history=await window.UangMakanStorage.list();}
async function openHistory(id,download=false){
 if(!permission(download?'history.download':'history.view')||state.busy)return;
 state.busy=true;clearNotice();render();
 try{
   const {meta,employees}=await window.UangMakanStorage.get(id);
   const runSettings=cleanHolidaySettings({holidays:meta.holidays||[],holidayNotes:meta.holidayNotes||{}},Number(meta.year),Number(meta.month));
   const holidays=new Set(runSettings.holidays);
   const hydrated=employees.map(e=>({...e,days:Object.fromEntries(Object.entries(e.days||{}).map(([key,d])=>[key,R.classifyDay({...d,isHoliday:d.isHoliday||holidays.has(key)},d.leaves||[],d.assignments||[])]))}));
   if(download){await window.ResourceLoader.ensure('jszip');await window.UangMakanExport.generate(meta.year,meta.month,hydrated,{holidays:runSettings.holidays});state.message='File Excel riwayat berhasil dibuat.';}
   else {state.year=Number(meta.year);state.month=Number(meta.month);state.holidays=runSettings.holidays;state.holidayNotes=runSettings.holidayNotes;state.employees=hydrated;state.files=[];state.savedRun=meta;state.dirty=false;state.step=4;state.view='process';location.hash='process';state.message='Riwayat dimuat. Untuk mengolah ulang, pilih Ubah Periode dan unggah berkas baru.';}
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
async function deleteHistory(id){
 if(!permission('history.delete')||state.busy)return;
 if(!confirm('Hapus riwayat Uang Makan ini beserta rincian pegawainya? Tindakan tidak dapat dibatalkan.'))return;
 state.busy=true;clearNotice();render();
 try{await window.UangMakanStorage.remove(id);await refreshHistory();state.message='Riwayat dihapus.';}
 catch(e){fail(e);return;}
 state.busy=false;render();
}
function changeView(view){
 const allowed=['process','dashboard','history'].includes(view)?view:'process';
 if(!window.UserAccess.canView('uangMakan',allowed))return;
 state.view=allowed;state.error='';state.message='';render();
}
root.addEventListener('click',async event=>{
 const button=event.target.closest('[data-um-action]');if(!button)return;
 const action=button.dataset.umAction;
 if(action==='close-modal' && button.classList.contains('um-modal-bg') && event.target!==button)return;
 switch(action){
  case 'goto-process':location.hash='process';break;
  case 'goto-history':location.hash='history';break;
  case 'step':{
   const step=Number(button.dataset.step);if(step===1){if(state.employees.length&&!confirm('Mengubah periode akan menghapus hasil proses yang belum disimpan. Lanjutkan?'))return;resetProcess();}
   else if(step===2){if(!state.employees.length||confirm('Kembali ke unggah akan mengulang proses perhitungan. Lanjutkan?')){state.employees=[];state.step=2;state.savedRun=null;}}
   else if(step===3&&state.employees.length)state.step=3;
   else if(step===4&&state.employees.length)state.step=4;
   render();break;
  }
  case 'toggle-holiday':
  case 'remove-holiday':{
   if(!permission('process.calculate')||state.busy)return;
   const date=button.dataset.date;
   if(!R.monthDates(state.year,state.month).includes(date)||R.weekend(date))return;
   const wasSelected=state.holidays.includes(date);
   if(action==='remove-holiday'||wasSelected){state.holidays=state.holidays.filter(k=>k!==date);delete state.holidayNotes[date];}
   else {state.holidays.push(date);state.holidayNotes[date]='';}
   persistHolidaySettings();render();break;
  }
  case 'period-next':persistHolidaySettings();state.step=2;clearNotice();render();break;
  case 'upload':if(permission('process.upload'))root.querySelector('#um-files')?.click();break;
  case 'remove-file':if(permission('process.upload')){state.files.splice(Number(button.dataset.index),1);state.employees=[];render();}break;
  case 'clear-files':if(permission('process.upload')){state.files=[];state.employees=[];render();}break;
  case 'calculate':await calculate();break;
  case 'sync':await sync();break;
  case 'verify':if(permission('process.verify')){state.modalId=button.dataset.id;const e=findEmployee(state.modalId);state.modalDate=Object.keys(e?.days||{}).find(k=>e.days[k].issues.length&&!e.days[k].reviewed)||R.key(state.year,state.month,1);render();}break;
  case 'select-day':state.modalDate=button.dataset.date;render();break;
  case 'save-day':saveDay();break;
  case 'acknowledge':acknowledge(button.dataset.kind);break;
  case 'close-modal':state.modalId='';state.modalDate='';render();break;
  case 'back':state.step=Math.max(1,state.step-1);render();break;
  case 'restart':if(state.employees.length&&!confirm('Proses saat ini akan dikosongkan. Periode baru membutuhkan unggah absensi. Lanjutkan?'))break;resetProcess();render();break;
  case 'preview':state.step=4;render();break;
  case 'export':await exportData();break;
  case 'save':await saveRun();break;
  case 'history-open':await openHistory(button.dataset.id);break;
  case 'history-download':await openHistory(button.dataset.id,true);break;
  case 'history-delete':await deleteHistory(button.dataset.id);break;
 }
});
root.addEventListener('change',event=>{
 const input=event.target;
 if(input.matches('[data-um-holiday-note]')){
  const date=input.dataset.umHolidayNote;
  if(state.holidays.includes(date)){
   state.holidayNotes[date]=String(input.value||'').trim().slice(0,120);
   persistHolidaySettings();
  }
  return;
 }
 if(input.id==='um-files'){
  if(!permission('process.upload'))return;
  const files=Array.from(input.files||[]).filter(x=>/\.xlsx?$/i.test(x.name));
  state.files.push(...files);state.files=Array.from(new Map(state.files.map(f=>[`${f.name}::${f.size}::${f.lastModified}`,f])).values());
  state.employees=[];state.parsed=[];clearNotice();render();
 }else if(input.id==='um-filter'){state.filter=input.value;render();}
 else if(input.id==='um-year'||input.id==='um-month'){
  const year=Number(document.getElementById('um-year')?.value||state.year),month=Number(document.getElementById('um-month')?.value||state.month);
  if(year!==state.year||month!==state.month){resetProcess();state.year=year;state.month=month;Object.assign(state,readHolidaySettings(year,month));render();}
 }
});
root.addEventListener('input',event=>{
 if(event.target?.id==='um-search'){
  const start=event.target.selectionStart;state.search=event.target.value;render();const next=document.getElementById('um-search');if(next){next.focus();try{next.setSelectionRange(start,start);}catch(_){}}
 }
});
root.addEventListener('dragover',event=>{const drop=event.target.closest('#um-drop');if(drop){event.preventDefault();drop.classList.add('is-over');}});
root.addEventListener('dragleave',event=>{const drop=event.target.closest('#um-drop');if(drop)drop.classList.remove('is-over');});
root.addEventListener('drop',event=>{
 const drop=event.target.closest('#um-drop');if(!drop)return;
 event.preventDefault();if(!permission('process.upload'))return;
 const files=Array.from(event.dataTransfer?.files||[]).filter(f=>/\.xlsx?$/i.test(f.name));
 state.files.push(...files);state.files=Array.from(new Map(state.files.map(f=>[`${f.name}::${f.size}::${f.lastModified}`,f])).values());
 state.employees=[];clearNotice();render();
});
window.addEventListener('hashchange',()=>changeView(location.hash.replace('#','')));
async function init(){
 try{
  if(!await window.UserAccess.requireModule('uangMakan'))return;
  const requested=location.hash.replace('#','');
  state.view=window.UserAccess.canView('uangMakan',requested)?requested:window.UserAccess.firstView('uangMakan');
  state.busy=true;render();
  await loadMaster();
  await refreshHistory();
  state.busy=false;render();
  window.ResourceLoader?.warm?.(['xlsx','jszip']);
 }catch(e){state.busy=false;state.error=e.message||String(e);render();}
}
init();
})();
