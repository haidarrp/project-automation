(function(){
'use strict';
const root=document.getElementById('app');
const R=window.UangMakanRules;
const prev=new Date();prev.setDate(1);prev.setMonth(prev.getMonth()-1);
const state={view:'process',step:1,year:prev.getFullYear(),month:prev.getMonth()+1,master:[],
  files:[],parsed:[],employees:[],leaves:[],assignments:[],supportLoaded:false,documentPolicy:null,history:[],
  busy:false,error:'',message:'',filter:'all',search:'',modalId:'',modalDate:'',savedRun:null,archiveKey:'',dirty:false,
  holidays:[],holidayNotes:{},documentPreviews:{},
  sharePointChecking:false,sharePointBusy:false,sharePointReady:false,sharePointError:''};
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
function makeEmptyEmployee(e){return {...e,files:[],sourceFiles:[],fileWarnings:[],coverageAcknowledged:false,coverageReason:'',days:{}};}
function resetProcess(){state.files=[];state.parsed=[];state.employees=[];state.leaves=[];state.assignments=[];state.documentPolicy=null;state.supportLoaded=false;state.modalId='';state.modalDate='';state.savedRun=null;state.archiveKey='';state.dirty=false;state.search='';state.filter='all';state.documentPreviews={};state.step=1;clearNotice();}
function pageTitle(view){return {process:'Proses Uang Makan',dashboard:'Dashboard Uang Makan',history:'Riwayat Uang Makan'}[view]||'Uang Makan';}
function metric(label,value,description){return `<div class="um-panel um-kpi"><small>${esc(label)}</small><strong>${esc(value)}</strong><span>${esc(description||'')}</span></div>`;}
function showStatus(){return `${state.error?`<div class="um-note error" role="alert" style="margin-bottom:13px">${esc(state.error)}</div>`:''}${state.message?`<div class="um-note good" role="status" style="margin-bottom:13px">${esc(state.message)}</div>`:''}`;}
function sharePointPanel(){
 const status=window.MicrosoftAuth?.getStatus?.()||{configured:false,libraryLoaded:false,connected:false};
 const busy=state.sharePointBusy||state.sharePointChecking;
 if(!status.configured||!status.libraryLoaded){
  const message=!status.configured?'Konfigurasi Microsoft/SharePoint belum lengkap. Periksa microsoft-config.js.':'Library autentikasi Microsoft belum tersedia. Jalankan proses build MSAL sebelum menggunakan SharePoint.';
  return `<div class="um-sp-panel warning" role="status"><div class="um-sp-info"><strong>SharePoint belum tersedia</strong><span>${esc(message)}</span></div><button class="um-btn small" type="button" disabled>Hubungkan SharePoint</button></div>`;
 }
 if(state.sharePointChecking){
  return '<div class="um-sp-panel checking" role="status"><div class="um-sp-info"><strong>Memeriksa koneksi SharePoint...</strong><span>Mengecek akun Microsoft 365 dan akses penyimpanan.</span></div></div>';
 }
 if(status.connected&&state.sharePointReady){
  const account=status.name||status.email||'Akun Microsoft 365';
  return `<div class="um-sp-panel connected" role="status"><div class="um-sp-info"><strong>SharePoint terhubung</strong><span>${esc(account)}${status.email&&account!==status.email?` · ${esc(status.email)}`:''}</span></div><span class="um-sp-indicator">Terhubung</span></div>`;
 }
 const message=state.sharePointError||'Hubungkan akun Microsoft 365 Kementerian PKP untuk menyimpan arsip absensi dan mengakses bukti dukung.';
 return `<div class="um-sp-panel warning" role="status"><div class="um-sp-info"><strong>${status.connected?'Koneksi SharePoint perlu diperiksa':'Hubungkan SharePoint'}</strong><span>${esc(message)}</span></div><button class="um-btn small" type="button" data-um-action="connect-sharepoint" ${busy?'disabled':''}>${state.sharePointBusy?'Menghubungkan...':status.connected?'Periksa Koneksi SharePoint':'Hubungkan SharePoint'}</button></div>`;
}
function refreshSharePointPanel(){
 // Tidak menggambar ulang formulir atau modal yang sedang diedit saat
 // pemeriksaan koneksi Microsoft Graph selesai secara asinkron.
 const panel=root.querySelector('.um-sp-panel');
 if(panel)panel.outerHTML=sharePointPanel();
 else render();
}
async function checkSharePoint(){
 const auth=window.MicrosoftAuth;
 if(!auth?.isConfigured?.())return;
 state.sharePointChecking=true;refreshSharePointPanel();
 try{
  await auth.initialize();
  if(auth.getStatus().connected){
   await window.SharePointStorage.ensureReady(false);
   state.sharePointReady=true;
  }else state.sharePointReady=false;
  state.sharePointError='';
 }catch(e){state.sharePointReady=false;state.sharePointError=`Tidak dapat mengakses penyimpanan SharePoint: ${String(e?.message||e)}`;}
 finally{state.sharePointChecking=false;refreshSharePointPanel();}
}
async function connectSharePoint(){
 if(state.sharePointBusy||state.sharePointChecking)return;
 state.sharePointBusy=true;state.sharePointError='';refreshSharePointPanel();
 try{
  await window.MicrosoftAuth.connect();
  await window.SharePointStorage.ensureReady(true);
  state.sharePointReady=true;
 }catch(e){state.sharePointReady=false;state.sharePointError=`Gagal menghubungkan SharePoint: ${String(e?.message||e)}`;}
 finally{state.sharePointBusy=false;refreshSharePointPanel();}
}
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
 return `${kpis(s)}<section class="um-panel"><div class="um-panel-title"><div><h2>Validasi & Verifikasi Uang Makan</h2><p>Rekap uang makan pegawai untuk sheet PUSDATIN. Satu verifikasi berlaku untuk seluruh rentang tanggal dan semua pegawai pada Surat Tugas yang sama.</p></div><div class="um-row"><button class="um-btn small" data-um-action="sync" type="button" ${state.busy?'disabled':''}>⟳ Sinkronkan Cuti & Surat Tugas</button></div></div>
 ${pendingCount()?`<div class="um-note warn">Terdapat ${pendingCount()} pegawai dari berkas yang diunggah yang perlu pemeriksaan, konfirmasi cakupan tanggal, atau kelengkapan jenis pegawai. Ekspor final tersedia setelah pegawai yang diproses siap.</div>`:''}
 <div class="um-note" style="margin-bottom:12px">Periode ${esc(R.periodLabel(state.year,state.month))}: Sabtu–Minggu dan ${holidayCount()} tanggal merah tambahan otomatis tidak memperoleh uang makan. Setiap cuti dan Surat Tugas yang dipertanggungjawabkan wajib dikonfirmasi pada hari kerja. Surat Tugas yang tidak dipertanggungjawabkan tidak disertakan pada perhitungan, verifikasi, maupun ZIP bukti dukung Uang Makan. Satu keputusan per dokumen diterapkan ke seluruh tanggal terkait pada bulan ini.</div>
 <div class="um-toolbar"><input class="um-search" id="um-search" placeholder="Cari nama pegawai atau NIP..." value="${esc(state.search)}"><div class="um-row"><select id="um-filter"><option value="all" ${state.filter==='all'?'selected':''}>Semua pegawai</option><option value="pending" ${state.filter==='pending'?'selected':''}>Perlu Verifikasi</option><option value="ready" ${state.filter==='ready'?'selected':''}>Siap</option></select><span class="um-pill good">Rekomendasi berdasarkan presensi dan dokumen</span></div></div>
 <div class="um-table-outer"><table class="um-table"><thead><tr><th>No.</th><th>Nama / NIP</th><th>Anak Satker</th><th>Jenis</th><th>Hari Hadir</th><th>Cuti</th><th>Dinas Luar</th><th>Hari Uang Makan</th><th>Status</th><th>Aksi</th></tr></thead><tbody>
 ${visibleEmployees().length?visibleEmployees().map((e,i)=>{const a=R.overview(e);return `<tr><td>${i+1}</td><td><div class="name">${esc(e.name)}</div><div class="sub">${esc(e.nip)}</div></td><td>${esc(e.anakSatker||'-')}</td><td>${esc(e.jenisPegawai||'-')}</td><td class="number">${a.present}</td><td class="number">${a.leave}</td><td class="number">${a.travel}</td><td class="number">${a.days}</td><td><span class="um-pill ${a.status==='Siap'?'good':'warn'}">${esc(a.status)}</span>${a.issues?`<div class="sub">${a.issues} tanggal perlu cek</div>`:''}</td><td><button type="button" class="um-btn small" data-um-action="verify" data-id="${esc(e.id)}">Detail →</button></td></tr>`;}).join(''):`<tr><td colspan="10" class="um-empty">Tidak ada pegawai dari file unggahan yang sesuai filter.</td></tr>`}
 </tbody></table></div><div class="um-summary"><span>${visibleEmployees().length} dari ${state.employees.length} pegawai hasil unggah</span><span><b>${s.days} hari</b> uang makan · ${s.issues} tanggal belum diperiksa</span></div>
 <div class="um-actions"><button type="button" class="um-btn" data-um-action="back">← Kembali ke Unggah</button><div class="um-row"><button type="button" class="um-btn" data-um-action="restart">Ubah Periode</button><button type="button" class="um-btn primary" data-um-action="preview">Pratinjau Output Excel →</button></div></div></section>`;
}
function kpis(s){return `<div class="um-kpis">${metric('Pegawai Diproses',s.employees,'Dari file absensi yang diunggah')}${metric('Hari Uang Makan',s.days,'Total pegawai yang diproses')}${metric('Tanggal Perlu Cek',s.issues,'Belum diverifikasi')}${metric('Koreksi Manual',s.manual,'Perbedaan dari rekomendasi sistem')}</div>`;}
function previewPanel(){
 const s=summary(),pending=pendingCount(),ready=pending===0;
 const shown=state.employees;
 return `${kpis(s)}<section class="um-panel"><div class="um-panel-title"><div><h2>Pratinjau Sheet PUSDATIN</h2><p>Kolom tanggal diisi angka 1 apabila memperoleh uang makan, dan kosong untuk tanggal lainnya.</p></div><span class="um-pill ${ready?'good':'warn'}">${ready?'Siap Ekspor':`${pending} pegawai belum siap`}</span></div>
 ${!ready?`<div class="um-note warn" style="margin-bottom:13px">Selesaikan verifikasi dan kelengkapan data pegawai yang tercantum pada file absensi. Pegawai master yang tidak diunggah tidak memengaruhi ekspor.</div>`:''}
 <div class="um-table-outer um-preview-scroll"><table class="um-table export"><thead><tr><th>No.</th><th>NIP</th><th>Nama</th>${Array.from({length:R.periodDays(state.year,state.month)},(_,i)=>`<th class="um-preview-day">${i+1}</th>`).join('')}<th>Jumlah</th><th>Unit Kerja</th><th>Bank</th><th>Jenis Pegawai</th></tr></thead><tbody>${shown.map((e,i)=>`<tr><td>${i+1}</td><td>${esc(e.nip)}</td><td class="name">${esc(e.name)}</td>${R.monthDates(state.year,state.month).map(k=>`<td class="number um-preview-day">${e.days[k]?.value===1?'1':''}</td>`).join('')}<td class="number">${R.overview(e).days}</td><td>Pusat Data dan Informasi</td><td>${esc(e.anakSatker||'')}</td><td>${esc(e.jenisPegawai||'')}</td></tr>`).join('')}</tbody></table></div>
 <p class="um-fine" style="margin-top:10px">Menampilkan seluruh ${shown.length} pegawai dari berkas yang diunggah. Geser tabel ke samping untuk melihat semua tanggal. Penyimpanan riwayat akan mengarsipkan absensi asli ke SharePoint agar ZIP dapat diunduh ulang.</p>
 <div class="um-summary um-preview-period"><span>Periode ${esc(R.periodLabel(state.year,state.month))}</span></div>
 <div class="um-actions"><button class="um-btn" type="button" data-um-action="back">← Kembali ke Verifikasi</button><div class="um-row"><button class="um-btn" type="button" data-um-action="save" ${!ready||state.busy?'disabled':''}>${state.busy?'Menyimpan...':'Simpan ke Riwayat'}</button><button class="um-btn" type="button" data-um-action="export" ${!ready||state.busy?'disabled':''}>Unduh Excel PUSDATIN ↓</button><button class="um-btn primary" type="button" data-um-action="export-zip" ${!ready||state.busy?'disabled':''}>${state.busy?'Menyiapkan...':'Unduh ZIP Lengkap ↓'}</button></div></div></section>`;
}
function periodCard(){return `<section class="um-panel" style="margin-bottom:16px"><h2>Periode Saat Ini</h2><p>${esc(R.periodLabel(state.year,state.month))}</p><div style="margin-top:14px"><button type="button" class="um-btn primary" data-um-action="goto-process">Proses Uang Makan →</button></div></section>`;}
function historyTable(embedded=false){
 return `<div class="um-table-outer"><table class="um-table"><thead><tr><th>Periode</th><th>Diproses</th><th>Jumlah Pegawai</th><th>Hari Uang Makan</th><th>Operator</th><th>Aksi</th></tr></thead><tbody>${state.history.length?state.history.slice(0,embedded?7:75).map(h=>`<tr><td class="name">${esc(R.periodLabel(h.year,h.month))}</td><td>${esc(new Date(h.createdAt).toLocaleString('id-ID'))}</td><td>${Number(h.summary?.employees||0)}</td><td>${Number(h.summary?.days||0)}</td><td>${esc(h.ownerName||h.ownerEmail||'-')}</td><td><div class="um-row"><button class="um-btn small" type="button" data-um-action="history-open" data-id="${esc(h.id)}">Lihat</button><button class="um-btn small" type="button" data-um-action="history-download" data-id="${esc(h.id)}">Excel ↓</button><button class="um-btn small" type="button" data-um-action="history-download-zip" data-id="${esc(h.id)}">ZIP ↓</button>${!embedded&&((h.ownerUid===window.FirebaseClient?.getCurrentUser?.()?.uid)||window.FirebaseClient?.isAdmin?.())?`<button class="um-btn small danger" type="button" data-um-action="history-delete" data-id="${esc(h.id)}">Hapus</button>`:''}</div></td></tr>`).join(''):`<tr><td class="um-empty" colspan="6">Belum ada riwayat uang makan tersimpan.</td></tr>`}</tbody></table></div>`;
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

// Metadata berkas mengikuti model dokumen pada modul Cuti dan Surat Tugas.
// Hanya metadata tersimpan yang digunakan; URL pratinjau diambil sementara
// dari Microsoft Graph, tidak disimpan pada riwayat atau Firestore.
function supportDocuments(raw){
 return (Array.isArray(raw)?raw:[]).map((doc,index)=>{
  const item=doc&&typeof doc==='object'?doc:{};
  const f=item.file&&typeof item.file==='object'?item.file:{};
  return {id:String(item.id||`dokumen-${index+1}`),name:String(item.name||''),type:String(item.type||''),
    number:String(item.number||''),note:String(item.note||''),
    file:{name:String(f.name||f.remoteName||''),remoteName:String(f.remoteName||''),
      driveId:String(f.driveId||''),itemId:String(f.itemId||''),webUrl:String(f.webUrl||'')}};
 });
}
function assignmentTypeLabel(type){
 return {official_travel:'Dinas Luar / Perjalanan Dinas',work_arrangement:'Pengaturan Kerja',other:'Penugasan Lainnya'}[type]||'Penugasan';
}
// Hanya dokumen yang diikutkan pada proses Uang Makan ditampilkan.
// Surat Tugas yang tidak dipertanggungjawabkan dikeluarkan sepenuhnya.
function dailySupporting(day){
 const matches=[];
 (day.previewLeaves||day.leaves||[]).forEach((item,index)=>matches.push({kind:'cuti',id:String(item.id||`cuti-${index+1}`),
   title:`Cuti: ${item.leaveType||'Cuti'}`,type:'Cuti',subject:item.leaveType||'Cuti',
   excluded:!(day.leaves||[]).some(d=>String(d.id)===String(item.id)),excludeReason:!(day.leaves||[]).some(d=>String(d.id)===String(item.id))?'Dieliminasi pada Register':'',
   reference:item.documentNumber||'',referenceLabel:'Nomor dokumen',referenceDate:item.documentDate||'',
   startDate:item.startDate||'',endDate:item.endDate||'',note:item.note||'',
   documents:Array.isArray(item.documents)?item.documents:[]}));
 (day.previewAssignments||day.assignments||[]).filter(item=>item.accountable!==false).forEach((item,index)=>matches.push({kind:'tugas',id:String(item.id||`tugas-${index+1}`),
   title:`Surat Tugas: ${assignmentTypeLabel(item.assignmentType)}`,type:'Surat Tugas',
   excluded:!(day.assignments||[]).some(d=>String(d.id)===String(item.id)),excludeReason:'Dieliminasi pada Register',
   subject:item.activity||assignmentTypeLabel(item.assignmentType),
   assignmentType:assignmentTypeLabel(item.assignmentType),
   reference:item.letterNumber||'',referenceLabel:'Nomor surat',referenceDate:item.letterDate||'',
   startDate:item.startDate||'',endDate:item.endDate||'',location:item.location||'',note:item.note||'',
   documents:Array.isArray(item.documents)?item.documents:[]}));
 return matches;
}
function supportDate(raw){
 const match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(raw||''));
 if(!match)return '—';
 const date=new Date(Number(match[1]),Number(match[2])-1,Number(match[3]));
 if(Number.isNaN(date.getTime())||date.getFullYear()!==Number(match[1])||date.getMonth()!==Number(match[2])-1||date.getDate()!==Number(match[3]))return '—';
 return new Intl.DateTimeFormat('id-ID',{day:'numeric',month:'short',year:'numeric'}).format(date);
}
function supportPeriod(group){
 const start=supportDate(group.startDate),end=supportDate(group.endDate||group.startDate);
 return start===end?start:`${start} – ${end}`;
}
function supportingRecordsSection(day){
 const groups=dailySupporting(day);
 if(!groups.length)return '<div class="um-source-empty">Tidak ada data cuti atau surat tugas yang tercatat pada tanggal ini.</div>';
 return `<section class="um-source-records" aria-label="Data cuti dan surat tugas terintegrasi">
  <div class="um-source-head"><h4>Data Cuti &amp; Surat Tugas Terintegrasi</h4><span>${groups.length} catatan dari data sumber</span></div>
  <div class="um-source-record-list">${groups.map(group=>`<article class="um-source-record ${group.kind}">
   <div class="um-source-record-top"><span class="um-source-type ${group.kind}">${esc(group.type)}</span><strong>${esc(group.subject)}</strong>${group.excluded?'<span class="um-pill">Tidak diikutkan</span>':''}</div>
   ${group.excluded?`<div class="um-note" style="color:#9b5b15">Hanya pratinjau — ${esc(group.excludeReason)}; tidak memengaruhi hitungan/ZIP.</div>`:''}<dl><div><dt>${esc(group.referenceLabel)}</dt><dd>${esc(group.reference||'—')}</dd></div>
   <div><dt>Periode</dt><dd>${esc(supportPeriod(group))}</dd></div>
   ${group.referenceDate?`<div><dt>Tanggal dokumen</dt><dd>${esc(supportDate(group.referenceDate))}</dd></div>`:''}
   ${group.assignmentType?`<div><dt>Jenis penugasan</dt><dd>${esc(group.assignmentType)}</dd></div>`:''}
   ${group.location?`<div><dt>Lokasi / tujuan</dt><dd>${esc(group.location)}</dd></div>`:''}
   ${group.note?`<div class="um-source-wide"><dt>Keterangan</dt><dd>${esc(group.note)}</dd></div>`:''}
   </dl><small>${group.documents.length} dokumen pendukung tercatat</small>
  </article>`).join('')}</div>
 </section>`;
}
function documentKey(group,doc){
 return [group.kind,group.id,doc.id,doc.file?.driveId||'',doc.file?.itemId||''].join('::');
}
function safeHttps(value){
 try{const url=new URL(String(value||''));return url.protocol==='https:'?url.href:'';}catch(_){return '';}
}
function previewPostFields(raw){
 return [...new URLSearchParams(String(raw||'')).entries()].map(([name,value])=>`<input type="hidden" name="${esc(name)}" value="${esc(value)}">`).join('');
}
function supportFilePreview(group,doc,index){
 const file=doc.file||{},key=documentKey(group,doc),hasFile=Boolean(file.driveId&&file.itemId);
 const ctx=state.documentPreviews[key]||{loading:true};
 const fileName=file.name||doc.name||file.remoteName||'Berkas pendukung';
 const sharePointUrl=safeHttps(file.webUrl);
 const actions=sharePointUrl?`<a class="um-btn small" href="${esc(sharePointUrl)}" target="_blank" rel="noopener noreferrer">Buka SharePoint ↗</a>`:'';
 let body='';
 if(!hasFile)body='<div class="um-document-status">Berkas digital belum terunggah pada data sumber.</div>';
 else if(ctx.loading)body='<div class="um-document-status">Menyiapkan pratinjau dari Microsoft 365...</div>';
 else if(ctx.error)body=`<div class="um-document-status"><strong>Pratinjau belum tersedia</strong><p>${esc(ctx.error)}</p><button class="um-btn small" type="button" data-um-action="retry-support-preview" data-doc-key="${esc(key)}">Coba Lagi</button></div>`;
 // Jangan buat iframe sebelum pengguna benar-benar mencapai area preview.
 // Iframe Microsoft 365 (khususnya form POST) dapat memindahkan posisi scroll modal saat dimuat.
 else if(safeHttps(ctx.getUrl))body=`<div class="um-document-placeholder" data-um-preview-url="${esc(safeHttps(ctx.getUrl))}" data-um-preview-title="Pratinjau ${esc(fileName)}">Pratinjau dimuat saat bagian ini ditampilkan.</div>`;
 else if(safeHttps(ctx.postUrl)){
  const target=`um-support-${index}`;
  body=`<div class="um-document-placeholder" data-um-preview-target="${target}" data-um-preview-title="Pratinjau ${esc(fileName)}">Pratinjau dimuat saat bagian ini ditampilkan.</div><form data-um-preview-post method="post" action="${esc(safeHttps(ctx.postUrl))}" target="${target}">${previewPostFields(ctx.postParameters)}</form>`;
 }else body='<div class="um-document-status">Microsoft 365 tidak memberikan tautan pratinjau yang valid.</div>';
 return `<article class="um-document-card"><div class="um-document-heading"><div><strong>${esc(fileName)}${group.excluded?' (hanya pratinjau — tidak diekspor)':''}</strong><small>${esc(group.title)}${group.reference?' · '+esc(group.reference):''}${doc.type?' · '+esc(doc.type):''}${doc.number?' · '+esc(doc.number):''}</small></div>${actions}</div>${body}</article>`;
}
function supportingFilesSection(day){
 const groups=dailySupporting(day);
 if(!groups.length)return '';
 let index=0;
 const cards=groups.flatMap(group=>group.documents.map(doc=>supportFilePreview(group,doc,++index)));
 const missing=groups.filter(group=>!group.documents.length);
 return `<section class="um-supporting-documents" id="um-supporting-documents" aria-label="Berkas pendukung cuti dan surat tugas">
  <div class="um-supporting-head"><h4>Berkas Pendukung Cuti / Surat Tugas</h4><p>Dokumen dari modul Cuti dan Surat Tugas untuk ${esc(dateLabel(day.key))}. Tetap tersedia meskipun tidak ada konflik presensi.</p></div>
  ${cards.length?`<div class="um-document-list">${cards.join('')}</div>`:''}
  ${missing.length?`<div class="um-note um-supporting-missing">Belum terdapat berkas digital pada ${missing.map(group=>esc(group.title)).join(' · ')}. Periksa dokumen di modul Cuti atau Surat Tugas.</div>`:''}
 </section>`;
}
let documentPreviewObserver=null;
let documentPreviewFallback=null;
function stopDocumentPreviewTracking(){
 documentPreviewObserver?.disconnect();
 documentPreviewObserver=null;
 if(documentPreviewFallback){
  documentPreviewFallback.modal.removeEventListener('scroll',documentPreviewFallback.onScroll);
  documentPreviewFallback=null;
 }
}
function activateDocumentPreview(placeholder){
 if(!placeholder.isConnected)return;
 const url=safeHttps(placeholder.dataset.umPreviewUrl);
 const target=placeholder.dataset.umPreviewTarget||'';
 if(!url&&!target)return;
 const frame=document.createElement('iframe');
 frame.className='um-document-frame';
 frame.title=placeholder.dataset.umPreviewTitle||'Pratinjau dokumen';
 frame.loading='lazy';
 frame.referrerPolicy='no-referrer';
 frame.setAttribute('allow','fullscreen');
 if(target)frame.name=target;
 placeholder.replaceWith(frame);
 if(url)frame.src=url;
 else{
  const form=frame.closest('.um-document-card')?.querySelector('[data-um-preview-post]');
  if(form&&form.dataset.submitted!=='1'){
   form.dataset.submitted='1';
   try{form.submit();}catch(error){console.warn('[Uang Makan] Pratinjau POST gagal',error);}
  }
 }
}
function observeDocumentPreviews(){
 stopDocumentPreviewTracking();
 const modal=root.querySelector('.um-modal');
 if(!modal)return;
 const placeholders=Array.from(modal.querySelectorAll('.um-document-placeholder[data-um-preview-url], .um-document-placeholder[data-um-preview-target]'));
 if(!placeholders.length)return;
 if(typeof window.IntersectionObserver==='function'){
  documentPreviewObserver=new window.IntersectionObserver((entries,observer)=>{
   for(const entry of entries){
    if(!entry.isIntersecting)continue;
    observer.unobserve(entry.target);
    activateDocumentPreview(entry.target);
   }
  },{root:modal,rootMargin:'0px',threshold:0.01});
  placeholders.forEach(node=>documentPreviewObserver.observe(node));
 }else{
  // Fallback tanpa IntersectionObserver: aktifkan hanya preview yang terlihat.
  const onScroll=()=>{
   const bounds=modal.getBoundingClientRect();
   placeholders.forEach(node=>{
    if(!node.isConnected)return;
    const rect=node.getBoundingClientRect();
    if(rect.top<bounds.bottom&&rect.bottom>bounds.top)activateDocumentPreview(node);
   });
  };
  documentPreviewFallback={modal,onScroll};
  modal.addEventListener('scroll',onScroll,{passive:true});
  onScroll();
 }
}
function refreshSupportingFiles(){
 const day=findEmployee(state.modalId)?.days?.[state.modalDate];
 const node=root.querySelector('#um-supporting-documents');
 if(day&&node){node.outerHTML=supportingFilesSection(day);observeDocumentPreviews();}
}
async function loadSupportingPreviews(){
 if(!window.UserAccess?.can?.('uangMakan','process.verify'))return;
 const employeeId=state.modalId,date=state.modalDate;
 const day=findEmployee(employeeId)?.days?.[date];
 if(!day)return;
 for(const group of dailySupporting(day)){
  for(const doc of group.documents){
   if(state.modalId!==employeeId||state.modalDate!==date)return;
   if(!doc.file?.driveId||!doc.file?.itemId)continue;
   const key=documentKey(group,doc);
   if(state.documentPreviews[key])continue;
   state.documentPreviews[key]={loading:true};
   refreshSupportingFiles();
   try{
    if(!window.SharePointStorage?.getPreviewInfo)throw new Error('Layanan pratinjau SharePoint tidak tersedia.');
    const preview=await window.SharePointStorage.getPreviewInfo({...doc.file,file:null},true);
    state.documentPreviews[key]={loading:false,error:'',getUrl:preview.getUrl||'',
      postUrl:preview.postUrl||'',postParameters:preview.postParameters||''};
   }catch(error){
    state.documentPreviews[key]={loading:false,error:String(error?.message||error),getUrl:'',postUrl:'',postParameters:''};
   }
   // Pratinjau yang sama dapat dipakai pada beberapa tanggal. Ketika pengguna
   // berpindah tanggal saat request masih berjalan, segarkan panel yang aktif.
   const activeDay=findEmployee(state.modalId)?.days?.[state.modalDate];
   if(activeDay&&dailySupporting(activeDay).some(g=>g.documents.some(d=>documentKey(g,d)===key)))refreshSupportingFiles();
  }
 }
}
// Sebuah ID Surat Tugas dari Firestore dapat mencakup beberapa pegawai.
// Hanya ID + tanda tangan metadata yang identik boleh berbagi verifikasi.
function reviewParticipants(kind,source){
 const key=R.supportKey(kind,source),signature=R.supportSignature(kind,source);
 return state.employees.map(employee=>({employee,dates:Object.values(employee.days||{}).filter(day=>
  !day.nonWorking&&(kind==='cuti'?day.leaves||[]:day.assignments||[]).some(item=>
   R.supportKey(kind,item)===key&&R.supportSignature(kind,item)===signature))
 })).filter(item=>item.dates.length);
}
function periodRecommendation(participants){
 const days=participants.flatMap(item=>item.dates);
 // Rekomendasi konservatif: memberi hanya jika SELURUH tanggal/pegawai
 // pada dokumen direkomendasikan memperoleh oleh mesin presensi.
 return days.length&&days.every(day=>day.automatic===1)?1:0;
}
function sourceReviewPanel(employee,day){
 const groups=dailySupporting(day);
 if(!groups.length)return '';
 return `<section class="um-period-review" aria-label="Verifikasi periode cuti dan surat tugas">
  <div class="um-period-review-title"><h4>Verifikasi Cuti / Surat Tugas per Periode</h4>
   <p><b>Cukup satu kali verifikasi untuk setiap dokumen.</b> Hasil berlaku pada seluruh tanggal kerja yang tercakup. Untuk satu Surat Tugas yang memuat beberapa pegawai, keputusan juga diterapkan kepada seluruh pegawai terkait yang sudah diunggah dalam proses ini.</p></div>
  <div class="um-period-review-items">${groups.map((group,index)=>{
   const source=(group.kind==='cuti'?day.leaves:day.assignments).find(x=>String(x.id)===group.id);
   const id=R.supportKey(group.kind,source||group),review=employee.supportReviews?.[id];
   const participants=reviewParticipants(group.kind,source||group);
   const dates=participants.find(x=>x.employee.id===employee.id)?.dates||[];
   const allDays=participants.flatMap(x=>x.dates);
   const recommended=periodRecommendation(participants);
   const mixed=allDays.some(x=>x.automatic!==recommended);
   const selected=review?.choice===0?0:review?.choice===1?1:recommended;
   const status=review?'Sudah diverifikasi':allDays.length?'Perlu verifikasi':'Tidak perlu verifikasi (hari libur)';
   const crossEmployee=group.kind==='tugas'&&participants.length>1;
   return `<article class="um-period-review-item ${group.kind}">
    <div class="um-period-review-top"><div class="um-period-review-head"><span class="um-source-type ${group.kind}">${esc(group.type)}</span><strong>${esc(group.subject)}</strong></div><span class="um-pill ${review||!allDays.length?'good':'warn'}">${status}</span></div>
    <div class="um-period-review-meta"><span><b>Periode:</b> ${esc(supportPeriod(group))}</span><span><b>Nomor:</b> ${esc(group.reference||'-')}</span><span><b>Cakupan:</b> ${dates.length} tanggal kerja pada ${esc(R.periodLabel(state.year,state.month))}</span>
     ${group.kind==='tugas'?`<span><b>Penerapan bersama:</b> ${participants.length} pegawai, ${allDays.length} tanggal-pegawai pada proses ini</span>`:''}</div>
    <div class="um-period-review-controls"><div class="um-field"><label for="um-source-choice-${index}">Hasil Setelah Verifikasi</label>
      <select id="um-source-choice-${index}"><option value="1" ${selected===1?'selected':''}>Memperoleh uang makan</option><option value="0" ${selected===0?'selected':''}>Tidak memperoleh uang makan</option></select></div>
      <button type="button" class="um-btn primary small" data-um-action="save-support" data-kind="${group.kind}" data-source-id="${esc(group.id)}" data-index="${index}" ${!allDays.length?'disabled':''}>${!allDays.length?'Tidak ada hari kerja':review?'Perbarui Verifikasi Periode':'Simpan Verifikasi Periode'}</button></div>
    ${group.kind==='tugas'?'<div class="um-note">Surat Tugas ini berstatus <b>dipertanggungjawabkan</b>, sehingga rekomendasi awal uang makan reguler adalah <b>Tidak memperoleh uang makan</b> meskipun terdapat presensi WFO. Pilihan verifikasi tidak mengubah status pertanggungjawaban; status tersebut dikelola melalui Data Surat Tugas.</div>':''}
    <small>Rekomendasi awal sistem: <b>${recommended===1?'Memperoleh uang makan':'Tidak memperoleh uang makan'}</b>. ${mixed?'Hasil otomatis antar tanggal atau pegawai tidak seragam; periksa realisasi sebelum menyimpan. ':''}${crossEmployee?'Keputusan berlaku bagi semua peserta Surat Tugas dalam proses ini. Perbedaan tanggal harus didukung dokumen koreksi pada modul sumber. ':''}Sabtu/Minggu dan tanggal merah tetap 0. WFH/WFA selain Jumat dapat dikoreksi setelah verifikasi. ${review?`Diverifikasi oleh ${esc(review.reviewedBy||'operator')} pada ${esc(new Date(review.reviewedAt).toLocaleString('id-ID'))}.`:'Keputusan belum disimpan.'}</small>
   </article>`;
  }).join('')}</div></section>`;
}
function dayReviewPanel(day){
 if(day.nonWorking)return `<div class="um-note um-note-nonworking"><b>${day.isHoliday?'Tanggal merah':'Sabtu/Minggu'}</b><p>Hari ini otomatis bernilai 0. Verifikasi periode tetap dapat dilakukan untuk tanggal kerja lain yang termasuk dokumen yang sama.</p>${day.isHoliday&&state.holidayNotes[day.key]?`<p>Keterangan: ${esc(state.holidayNotes[day.key])}</p>`:''}</div>`;
 if(day.sourceKeys?.length)return `<div class="um-note um-review-hint"><b>Verifikasi menurut dokumen</b><p>Pilih hasil aktual di bagian <b>Verifikasi Cuti / Surat Tugas per Periode</b> di bawah. Tidak perlu mengulang untuk setiap tanggal yang tercakup. Apabila realisasi berbeda pada satu tanggal, perbarui atau tambahkan dokumen koreksi di modul sumber dan sinkronkan ulang.</p>${day.reviewConflict?'<div class="um-note warn" style="margin-top:10px">Dokumen yang tumpang tindih memiliki keputusan berbeda. Perbaiki dokumen Cuti/Surat Tugas yang menjadi dasar, lalu klik Sinkronkan Ulang. Tidak tersedia koreksi manual per tanggal untuk dokumen tersebut.</div>':''}${day.reviewed?'<span class="um-pill good">Seluruh dokumen sudah dikonfirmasi</span>':''}</div>`;
 const selected=day.manualValue===0?0:day.manualValue===1?1:day.automatic;
 return `<div class="um-field"><label for="um-manual">Hasil Setelah Verifikasi</label>
   <select id="um-manual"><option value="1" ${selected===1?'selected':''}>Memperoleh uang makan</option><option value="0" ${selected===0?'selected':''}>Tidak memperoleh uang makan</option></select>
   <small>Rekomendasi awal sistem: <b>${day.automatic===1?'Memperoleh uang makan':'Tidak memperoleh uang makan'}</b>. ${day.remoteExcluded?'WFH/WFA di luar Jumat memerlukan pemeriksaan data sakit/cuti yang mungkin belum tercatat; hasil akhir tidak dikunci.':'Pilih hasil berdasarkan kondisi aktual dan bukti pendukung.'}</small></div>
   <div class="um-row" style="margin-top:12px"><button class="um-btn primary small" type="button" data-um-action="save-day">Simpan Verifikasi Tanggal</button>${day.reviewed?'<span class="um-pill good">Sudah diverifikasi</span>':''}</div>`;
}
function modal(){
 const e=findEmployee(state.modalId);if(!e)return '';
 const a=R.overview(e),keys=R.monthDates(state.year,state.month);
 const day=e.days?.[state.modalDate]||e.days?.[keys[0]];
 if(!day)return '';
 const isProblem=day.issues.length>0&&!day.reviewed;
 const leaveDays=keys.filter(k=>e.days[k]?.leaves?.length).length;
 const assignmentDays=keys.filter(k=>e.days[k]?.assignments?.length).length;
 return `<div class="um-modal-bg" data-um-action="close-modal"><section class="um-modal" role="dialog" aria-modal="true" aria-label="Detail Uang Makan">
 <div class="um-modal-head"><div><h2>${esc(e.name)}</h2><p>${esc(e.nip)} · ${esc(e.jenisPegawai||'Jenis belum dilengkapi')} · Anak Satker ${esc(e.anakSatker||'-')} · ${esc(R.periodLabel(state.year,state.month))}</p></div><button class="um-btn small" type="button" data-um-action="close-modal">Tutup ×</button></div>
 <div class="um-modal-body"><div class="um-row"><span class="um-pill good">${a.days} hari uang makan</span><span class="um-pill ${a.issues?'warn':'good'}">${a.issues} tanggal perlu cek</span><span class="um-pill">${e.files?.length||0} file absensi</span><span class="um-pill um-pill-source cuti">${leaveDays} tanggal cuti</span><span class="um-pill um-pill-source tugas">${assignmentDays} tanggal surat tugas</span></div>
 ${!e.jenisPegawai?'<div class="um-note warn" style="margin-top:14px">Jenis pegawai belum terisi pada master. Lengkapi melalui menu Administrasi → Master Data.</div>':''}
 ${e.fileWarnings?.length?`<div class="um-note warn" style="margin-top:13px">${e.fileWarnings.map(v=>esc(v)).join(' · ')}${e.coverageAcknowledged?`<div>Sudah ditinjau: ${esc(e.coverageReason||'-')}</div>`:`<div class="um-row" style="margin-top:9px"><input class="um-search" style="max-width:390px" id="um-coverage-reason" placeholder="Catatan verifikasi cakupan absensi"><button type="button" class="um-btn small" data-um-action="acknowledge" data-kind="coverage">Konfirmasi Cakupan</button></div>`}</div>`:''}
 <div class="um-weekday-row">${WEEKDAYS_MONDAY.map(day=>`<span>${day}</span>`).join('')}</div>
 <div class="um-day-grid">${Array.from({length:calendarOffset(keys[0])},()=>'<span class="um-day-placeholder" aria-hidden="true"></span>').join('')}${keys.map(k=>{const d=e.days[k],warn=d.issues.length&&!d.reviewed,cls=d.value===1?'earned':'';
   const status=d.isHoliday?'Tanggal merah':d.isWeekend?'Akhir pekan':warn?'Perlu verifikasi':d.value===1?'1 hari':d.remoteExcluded?'WFH/WFA: 0 hari':'Tidak dibayar';
   const leaves=d.leaves?.length||0,assignments=d.assignments?.length||0;
   const sourceLabels=[leaves?`Cuti (${leaves})`:'',assignments?`Surat Tugas (${assignments})`:''].filter(Boolean);
   const fullStatus=[status,...sourceLabels].join(' · ');
   const title=[d.issues.join('; ')||status,...sourceLabels].join(' · ');
   return `<button class="um-day ${cls} ${warn?'warn':''} ${state.modalDate===k?'selected':''} ${d.isWeekend?'weekend':''} ${d.isHoliday?'holiday':''} ${d.remoteExcluded&&!d.nonWorking?'remote':''} ${sourceLabels.length?'has-support':''}" type="button" data-um-action="select-day" data-date="${k}" aria-pressed="${state.modalDate===k}" aria-label="${esc(dateLabel(k))}, ${esc(fullStatus)}" title="${esc(title)}"><span class="um-day-week">${WEEKDAYS[R.weekday(k)]}</span><strong>${Number(k.slice(-2))}</strong><span class="um-day-sources">${leaves?`<span class="um-day-source cuti">Cuti${leaves>1?' '+leaves:''}</span>`:''}${assignments?`<span class="um-day-source tugas">Tugas${assignments>1?' '+assignments:''}</span>`:''}</span><small class="um-day-status">${d.nonWorking?'Libur':warn?'Periksa':d.value===1?'1 hari':d.remoteExcluded?d.workMode:'–'}</small></button>`;}).join('')}</div>
 <div class="um-source-legend" aria-label="Legenda integrasi kalender"><span><i class="um-legend-source cuti"></i>Cuti</span><span><i class="um-legend-source tugas"></i>Surat Tugas</span><small>Cuti dan surat tugas pada hari kerja memerlukan konfirmasi. Satu verifikasi berlaku untuk semua tanggal dalam satu dokumen; klik tanggal untuk melihat periode dan berkas.</small></div>
 <div class="um-detail"><h3>Verifikasi ${esc(dateLabel(day.key))}</h3><div class="um-detail-grid"><div><div class="um-fine"><b>Jam masuk:</b> ${esc(fmtTime(day.inMinutes))} · <b>Jam pulang:</b> ${esc(fmtTime(day.outMinutes))}<br><b>Status absensi:</b> ${esc(day.status||'—')}<br><b>Pola kerja:</b> ${esc(day.workMode||'—')}<br><b>Kode:</b> ${esc(day.flags.join(', ')||'—')}<br><b>Cuti:</b> ${day.leaves.length?esc(day.leaves.map(x=>x.leaveType||'Cuti').join(', ')):'—'}<br><b>Surat tugas:</b> ${day.assignments.length?esc(day.assignments.map(x=>x.activity||x.letterNumber||'Surat Tugas').join(', ')):'—'}<br><b>Hasil otomatis:</b> ${day.automatic===1?'Memperoleh uang makan':'Tidak memperoleh uang makan'}<br><b>Hasil setelah verifikasi:</b> ${day.value===1?'Memperoleh uang makan':'Tidak memperoleh uang makan'}${day.issues.length&&!day.reviewed?' (menunggu konfirmasi)':''}</div>
  ${day.remoteExcluded&&!day.nonWorking?`<div class="um-note um-remote-policy" style="margin-top:11px">${esc(day.workMode)} di luar hari Jumat: rekomendasi awal <b>tidak memperoleh uang makan</b>. Hasil dapat diubah setelah verifikasi, termasuk jika data sakit/cuti belum tersinkron.</div>`:''}
  ${day.nonWorking?'<div class="um-note" style="margin-top:11px">Tanggal tidak termasuk hari yang memperoleh uang makan.</div>':day.reviewed&&day.issues.length?`<div class="um-note good" style="margin-top:11px">Pemeriksaan selesai. ${day.issues.map(x=>esc(x)).join(' · ')}</div>`:isProblem?`<div class="um-note warn" style="margin-top:11px">${day.issues.map(x=>esc(x)).join(' · ')}</div>`:`<div class="um-note good" style="margin-top:11px">Tidak ditemukan konflik otomatis.${day.leaves.length||day.assignments.length?' Data cuti/surat tugas terkait tetap ditampilkan di bawah.':''}</div>`}</div>
 ${dayReviewPanel(day)}</div></div>${sourceReviewPanel(e,day)}${supportingRecordsSection(day)}${supportingFilesSection(day)}</div>
 <div class="um-modal-footer"><span class="um-fine">Hasil verifikasi tercatat bersama operator dan waktu pemeriksaan.</span><button class="um-btn" type="button" data-um-action="close-modal">Selesai</button></div>
 </section></div>`;
}
function render(){
 if(!window.AppShell||!window.FirebaseClient?.getCurrentUser?.())return;
 const content=`<div class="um-page ${state.busy?'um-busy':''}">${sharePointPanel()}${state.view==='dashboard'?dashboard():state.view==='history'?history():process()}</div>`;
 const previousModal=root.querySelector('.um-modal'),previousScroll=previousModal?.scrollTop||0;
 stopDocumentPreviewTracking();
 root.innerHTML=window.AppShell.render({module:'uangMakan',view:state.view,viewLabel:pageTitle(state.view),content,overlays:modal()});
 const currentModal=root.querySelector('.um-modal');if(currentModal)currentModal.scrollTop=previousScroll;
 window.UserAccess?.scheduleDecoration?.();
 observeDocumentPreviews();
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
   if(x.employeeId&&startDate<=to&&endDate>=from){ls.set(doc.id,{id:doc.id,employeeId:String(x.employeeId),startDate,endDate,leaveType:String(x.leaveType||'Cuti'),documentNumber:String(x.documentNumber||''),documentDate:String(x.documentDate||''),note:String(x.note||''),documents:supportDocuments(x.documents)});}
 });
 assignmentSnaps.flatMap(s=>s.docs).forEach(doc=>{
   const x=doc.data()||{},startDate=String(x.startDate||''),endDate=String(x.endDate||x.startDate||'');
   const type=['official_travel','work_arrangement','other'].includes(x.assignmentType)?x.assignmentType:'official_travel';
   const ids=(x.employees||[]).map(e=>String(e.employeeId||e.id||'')).filter(Boolean);
   if(ids.length&&startDate<=to&&endDate>=from)as.set(doc.id,{id:doc.id,employeeIds:ids,startDate,endDate,
     assignmentType:type,accountable:x.accountable!==false,activity:String(x.activity||''),letterNumber:String(x.letterNumber||''),letterDate:String(x.letterDate||''),location:String(x.location||''),note:String(x.note||''),documents:supportDocuments(x.documents)});
 });
 const policy=await window.AttendanceDocumentPolicy.fetch(years);
 if(policy.stale.length)throw new Error('Ada keputusan Register yang perlu diverifikasi ulang. Selesaikan dahulu sebelum memproses Uang Makan.');
 return {leaves:Array.from(ls.values()),assignments:Array.from(as.values()),policy};
}
// Saat data sumber diubah/dibatalkan, verifikasi yang lama tidak boleh
// dianggap masih berlaku hanya karena jumlah/jenis isu tetap sama.
function supportSignature(day){
 const leaves=(day?.leaves||[]).map(x=>R.supportSignature('cuti',x)).sort();
 const tasks=(day?.assignments||[]).map(x=>R.supportSignature('tugas',x)).sort();
 return JSON.stringify([leaves,tasks]);
}
function computeEmployees(list,leaves,assignments){
 return list.map(employee=>{
   const policy=state.documentPolicy;
   if(!policy)throw new Error('Keputusan Register belum dimuat.');
   const goodLeaves=leaves.filter(item=>String(item.employeeId)===String(employee.id)&&policy.exportable('cuti',item,employee.id));
   // Surat Tugas yang tidak dipertanggungjawabkan dikecualikan dari proses
   // uang makan, termasuk perhitungan, verifikasi, pratinjau dan ZIP.
   const goodAssignments=assignments.filter(item=>item.employeeIds.includes(String(employee.id))&&
     item.accountable!==false&&
     (policy.allows?policy.allows('surat',item.id,employee.id):policy.exportable('surat',item,employee.id)));
   let updated=R.applySupporting(employee,goodLeaves,goodAssignments,state.year,state.month,state.holidays);
   const revised={...updated.days};
   let changed=false;
   for(const [date,day] of Object.entries(updated.days)){
     const old=employee.days?.[date];
     if(old&&(supportSignature(old)!==supportSignature(day)||
       (old.issues||[]).filter(x=>!String(x).startsWith('Hasil verifikasi dokumen yang tumpang tindih')).join('|')!==
        day.issues.filter(x=>!String(x).startsWith('Hasil verifikasi dokumen yang tumpang tindih')).join('|'))){
       // Perubahan sumber membatalkan verifikasi harian lama, tetapi dokumen lain yang
       // masih identik tetap menyimpan keputusan verifikasi periodenya sendiri.
       revised[date]={...day,dateReviewed:false,dateReviewedAt:'',dateReviewedBy:'',
         dateManualValue:null,manualValue:null,manualNote:'',reviewed:false};
       changed=true;
     }
   }
   if(changed)updated=R.applySupporting({...updated,days:revised},goodLeaves,goodAssignments,
     state.year,state.month,state.holidays);
   for (const [date,day] of Object.entries(updated.days||{})) {
     day.previewLeaves=leaves.filter(x=>String(x.employeeId)===String(employee.id)&&x.startDate<=date&&x.endDate>=date);
     day.previewAssignments=assignments.filter(x=>x.accountable!==false&&x.employeeIds.includes(String(employee.id))&&x.startDate<=date&&x.endDate>=date);
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
  for(const [fileIndex,item] of parsed.entries()){
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
    const originalFile=state.files[fileIndex];
    if(!originalFile||originalFile.name!==item.fileName)throw new Error(`Berkas absensi ${item.fileName} tidak dapat ditautkan dengan hasil parsing.`);
    e.sourceFiles.push({name:originalFile.name,size:originalFile.size,type:originalFile.type||'',
      lastModified:originalFile.lastModified||0,file:originalFile});
    list.push({file:item.fileName,employee:e.name});
    const prior=e.days;
    for(const [key,record] of Object.entries(item.rows)){
      if(prior[key]){
        const old=prior[key];
        prior[key]={...record,inMinutes:record.inMinutes??old.inMinutes,outMinutes:record.outMinutes??old.outMinutes,
          flags:[...new Set([...(old.flags||[]),...(record.flags||[])])],status:record.status||old.status,
          workMode:record.workMode||old.workMode||'',duplicate:true};
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
  state.leaves=supports.leaves;state.assignments=supports.assignments;state.documentPolicy=supports.policy;state.supportLoaded=true;state.documentPreviews={};
  state.employees=computeEmployees(employees,state.leaves,state.assignments);
  state.parsed=list;state.step=3;state.savedRun=null;state.archiveKey='';state.dirty=true;
  state.message=`Berhasil mengolah ${parsed.length} berkas untuk ${employees.length} pegawai yang diunggah. Pegawai master lainnya tidak ikut dihitung; cuti dan surat tugas telah disinkronkan.`;
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
async function sync(){
 if(!permission('process.calculate')||state.busy)return;
 state.busy=true;clearNotice();render();
 try{
   const supports=await getSupportingData();
   state.leaves=supports.leaves;state.assignments=supports.assignments;state.documentPolicy=supports.policy;state.documentPreviews={};
   state.employees=computeEmployees(state.employees,supports.leaves,supports.assignments);state.dirty=true;
   state.message=`Sinkronisasi selesai: ${supports.leaves.length} data cuti, ${supports.assignments.filter(item=>item.accountable!==false).length} Surat Tugas yang dipertanggungjawabkan. ${supports.assignments.filter(item=>item.accountable===false).length} Surat Tugas tidak dipertanggungjawabkan dikecualikan dari proses Uang Makan.`;
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
function recalculateEmployee(employee){
 const revised=computeEmployees([employee],state.leaves,state.assignments)[0];
 state.employees=state.employees.map(e=>e.id===employee.id?revised:e);
 state.dirty=true;state.savedRun=null;
 return revised;
}
async function saveSupport(kind,sourceId,index){
 if(!permission('process.verify')||state.busy)return;
 if(!['cuti','tugas'].includes(kind))return;
 const e=findEmployee(state.modalId),day=e?.days?.[state.modalDate];
 if(!day)return;
 const source=(kind==='cuti'?day.leaves:day.assignments).find(x=>String(x.id)===sourceId);
 if(!source)return;
 const raw=document.getElementById('um-source-choice-'+index)?.value;
 if(!['0','1'].includes(raw)){
  alert('Pilih Memperoleh atau Tidak memperoleh uang makan.');return;
 }
 const choice=Number(raw),id=R.supportKey(kind,source),signature=R.supportSignature(kind,source);
 const participants=reviewParticipants(kind,source);
 if(!participants.length){alert('Tidak ada pegawai atau tanggal terkait dalam proses ini. Sinkronkan ulang data sumber.');return;}
 const employeeIds=new Set(participants.map(p=>p.employee.id));
 const reviewedAt=new Date().toISOString(),reviewedBy=window.FirebaseClient?.getCurrentUser?.()?.email||'';
 state.busy=true;
 try{
  const decision={choice,signature,reviewedAt,reviewedBy};
  state.employees=state.employees.map(employee=>{
   if(!employeeIds.has(employee.id))return employee;
   return computeEmployees([{...employee,supportReviews:{...(employee.supportReviews||{}),[id]:{...decision}}}],state.leaves,state.assignments)[0];
  });
  state.dirty=true;state.savedRun=null;
  const count=participants.reduce((n,p)=>n+p.dates.length,0);
  notify(`Verifikasi ${kind==='cuti'?'cuti':'surat tugas'} diterapkan ke ${count} tanggal-pegawai (${participants.length} pegawai) pada ${R.periodLabel(state.year,state.month)}.${state.documentPolicy?.stale?.length?' Keputusan konflik Register perlu diperiksa ulang sebelum ekspor.':''}`);
 }catch(error){fail(error);}
 finally{state.busy=false;render();}
}
function saveDay(){
 if(!permission('process.verify'))return;
 const e=findEmployee(state.modalId),day=e?.days?.[state.modalDate];if(!day)return;
 if(day.nonWorking){alert('Sabtu/Minggu atau tanggal merah selalu bernilai 0.');return;}
 // Hari yang memiliki sumber Cuti/Surat Tugas hanya dapat ditetapkan
 // melalui verifikasi dokumen. Perubahan per tanggal harus dibuktikan
 // dengan dokumen koreksi yang diunggah melalui modul sumber.
 if(day.sourceKeys?.length){
  alert('Tanggal dengan Cuti/Surat Tugas tidak dapat dikoreksi secara manual. Perbaiki atau tambahkan dokumen koreksi pada modul sumber, kemudian Sinkronkan Ulang.');return;
 }
 const raw=document.getElementById('um-manual')?.value;
 if(!['0','1'].includes(raw)){alert('Pilih Memperoleh atau Tidak memperoleh uang makan.');return;}
 const manual=Number(raw);
 const now=new Date().toISOString(),actor=window.FirebaseClient?.getCurrentUser?.()?.email||'';
 day.dateManualValue=manual;
 day.dateReviewed=true;day.dateReviewedAt=now;day.dateReviewedBy=actor;
 day.manualNote='';
 recalculateEmployee(e);
 notify('Verifikasi tanggal tanpa dokumen Cuti/Surat Tugas berhasil dicatat.');
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
async function assertDocumentPolicyFresh(){
 if(!state.documentPolicy)throw new Error('Keputusan Register belum dimuat. Sinkronkan ulang data Cuti/Surat Tugas.');
 const latest=await window.AttendanceDocumentPolicy.fetch([state.year-1,state.year]);
 if(latest.stale.length||latest.signature!==state.documentPolicy.signature)throw new Error('Dokumen/keputusan Register berubah setelah perhitungan. Klik Sinkronkan Ulang dan periksa kembali sebelum menyimpan atau mengekspor.');
}
async function exportData(){
 if(!permission('process.export')||state.busy)return;
 try{allowFinal();await assertDocumentPolicyFresh();await window.ResourceLoader.ensure('jszip');const result=await window.UangMakanExport.generate(state.year,state.month,state.employees,{holidays:state.holidays});notify(`File ${result.filename} berhasil dibuat.`);}
 catch(e){fail(e);}
}
async function exportZip(){
 if(!permission('process.export')||state.busy)return;
 try{allowFinal();}catch(e){fail(e);return;}
 state.busy=true;clearNotice();render();
 try{
  await assertDocumentPolicyFresh();
  await window.ResourceLoader.ensure('jszip');
  const result=await window.UangMakanZip.generate(state.year,state.month,state.employees,{holidays:state.holidays});
  window.UangMakanZip.download(result);
  state.message=`ZIP berhasil dibuat: ${result.attendanceCount} berkas absensi dan ${result.evidenceCount} bukti dukung.${result.missingEvidenceCount?' '+result.missingEvidenceCount+' dokumen belum memiliki berkas digital (lihat catatan dalam ZIP).':''}`;
 }catch(e){fail(e);return;}
 state.busy=false;render();
}
// Berkas absensi harus diarsipkan ke SharePoint sebelum menyimpan metadata
// riwayat. Referensi file dipertahankan agar download ZIP di perangkat lain
// tidak bergantung pada file input browser yang sudah hilang setelah reload.
async function archiveAbsences(){
 const toUpload=state.employees.some(e=>(e.sourceFiles||[]).some(f=>!f.driveId||!f.itemId));
 if(!state.employees.every(e=>(e.sourceFiles||[]).length&&(e.sourceFiles||[]).length>=(e.files||[]).length)){
  throw new Error('Ada absensi asli yang belum tersedia. Unggah dan hitung ulang berkas sebelum menyimpan riwayat.');
 }
 if(!toUpload)return;
 const sharePoint=window.SharePointStorage;
 if(!sharePoint?.uploadAttachment)throw new Error('Penyimpanan SharePoint belum tersedia untuk arsip absensi Uang Makan.');
 await sharePoint.ensureReady(true);
 if(!state.archiveKey){
  const unique=window.crypto?.randomUUID?.()||`${Date.now()}-${Math.random().toString(36).slice(2)}`;
  state.archiveKey=String(unique).replace(/[^a-zA-Z0-9-]/g,'').slice(0,45);
 }
 for(const employee of state.employees){
  for(const [index,source] of (employee.sourceFiles||[]).entries()){
   if(source.driveId&&source.itemId)continue;
   if(!source.file)throw new Error(`File absensi ${employee.name} tidak tersedia untuk diarsipkan.`);
   const stored=await sharePoint.uploadAttachment(source,{
    kind:'absensi',
    path:['Uang Makan',String(state.year),`${String(state.month).padStart(2,'0')}_${R.MONTHS[state.month-1]}`,`Run_${state.archiveKey}`,'Pegawai',`${employee.nip}_${employee.name}`,'Absensi'],
    remoteName:`${index+1}_${source.name||'Absensi.xlsx'}`
   },true);
   Object.assign(source,stored);
  }
 }
}
async function saveRun(){
 if(!permission('process.save')||state.busy)return;
 try{allowFinal();}catch(e){fail(e);return;}
 if(!state.dirty&&state.savedRun){notify('Hasil ini sudah tersimpan pada riwayat.');return;}
 state.busy=true;clearNotice();render();
 try{
  await assertDocumentPolicyFresh();
  await archiveAbsences();
  const run=await window.UangMakanStorage.save(state.year,state.month,state.employees,summary(),{
   holidays:state.holidays,holidayNotes:state.holidayNotes,archiveKey:state.archiveKey
  });
  state.savedRun=run;state.dirty=false;await refreshHistory();
  state.message='Hasil periode '+R.periodLabel(state.year,state.month)+' tersimpan di Riwayat bersama referensi arsip absensi SharePoint.';
 }
 catch(e){fail(e);return;}
 state.busy=false;render();
}
async function refreshHistory(){state.history=await window.UangMakanStorage.list();}
// Riwayat skema baru merekonstruksi hasil per tanggal dari keputusan per dokumen.
// Data lama tetap dibaca menggunakan mekanisme verifikasi harian yang asli.
function storedSupporting(employees){
 const leaves=new Map(),assignments=new Map();
 for(const e of employees){
  for(const d of Object.values(e.days||{})){
   for(const x of d.leaves||[])if(x.id)leaves.set(x.id,{...x,employeeId:e.id});
   for(const x of d.assignments||[])if(x.id)assignments.set(x.id,{...x,
     employeeIds:Array.isArray(x.employeeIds)?x.employeeIds:[e.id]});
  }
 }
 return {leaves:[...leaves.values()],assignments:[...assignments.values()]};
}
// Riwayat skema 5 dapat berisi pengecualian tanggal yang dulu diizinkan.
// Hasil arsip merupakan bukti perhitungan historis dan harus tetap sama;
// pengecualian historis ini TIDAK dipakai untuk pengolahan/sinkronisasi baru.
function hydrateArchivedEmployee(employee,supportData,year,month,holidays,schemaVersion){
 const calculated=R.applySupporting(employee,supportData.leaves,supportData.assignments,year,month,holidays);
 for(const [date,day] of Object.entries(calculated.days||{})){
   const source=employee.days?.[date]||{};
   day.previewLeaves=source.previewLeaves||source.leaves||[];
   day.previewAssignments=source.previewAssignments||source.assignments||[];
 }
 if(schemaVersion>=8)return calculated;
 // Rekap historis versi <=7 menggunakan rekomendasi berbasis presensi.
 // Jangan menghitung ulang arsip tersebut memakai aturan ST baru.
 const days={...calculated.days};
 for(const [date,original] of Object.entries(employee.days||{})){
  const d=days[date];if(!d)continue;
  const oldAutomatic=d.present&&!d.nonWorking&&!d.remoteExcluded?1:0;
  const sourceChoices=(d.sourceKeys||[]).map(k=>employee.supportReviews?.[k]?.choice)
   .filter(x=>x===0||x===1);
  let oldValue=sourceChoices.length&&sourceChoices.every(x=>x===sourceChoices[0])
   ?sourceChoices[0]:(original.manualValue===0||original.manualValue===1?original.manualValue:oldAutomatic);
  if(schemaVersion<6&&original.dateReviewed&&[0,1].includes(original.dateManualValue)&&
    (original.leaves?.length||original.assignments?.length))oldValue=original.dateManualValue;
  days[date]={...d,automatic:oldAutomatic,value:d.nonWorking?0:oldValue,
   manualValue:sourceChoices.length?sourceChoices[0]:(original.manualValue??null),
   reviewed:Boolean(original.reviewed),reviewedAt:String(original.reviewedAt||''),
   reviewedBy:String(original.reviewedBy||'')};
 }
 return {...calculated,days};
}
async function openHistory(id,mode='view'){
 if(!permission(mode==='view'?'history.view':'history.download')||state.busy)return;
 state.busy=true;clearNotice();render();
 try{
   const {meta,employees}=await window.UangMakanStorage.get(id);
   const year=Number(meta.year),month=Number(meta.month);
   const runSettings=cleanHolidaySettings({holidays:meta.holidays||[],holidayNotes:meta.holidayNotes||{}},year,month);
   const holidays=new Set(runSettings.holidays);
   const supportData=storedSupporting(employees);
   const hydrated=Number(meta.schemaVersion||0)>=5
    ?employees.map(e=>hydrateArchivedEmployee(e,supportData,year,month,runSettings.holidays,Number(meta.schemaVersion)))
    :employees.map(e=>({...e,days:Object.fromEntries(Object.entries(e.days||{}).map(([key,d])=>{
       const oldDay=R.classifyDay({...d,isHoliday:d.isHoliday||holidays.has(key)},d.leaves||[],d.assignments||[]);
       oldDay.issues=oldDay.issues.filter(issue=>!issue.startsWith('Konfirmasi pelaksanaan cuti'));
       return [key,oldDay];
      }))}));
   if(mode==='excel'){
    await window.ResourceLoader.ensure('jszip');
    await window.UangMakanExport.generate(year,month,hydrated,{holidays:runSettings.holidays});
    state.message='File Excel riwayat berhasil dibuat.';
   }else if(mode==='zip'){
    await window.ResourceLoader.ensure('jszip');
    const result=await window.UangMakanZip.generate(year,month,hydrated,{holidays:runSettings.holidays});
    window.UangMakanZip.download(result);
    state.message=`ZIP riwayat berhasil dibuat: ${result.attendanceCount} absensi dan ${result.evidenceCount} bukti dukung.${result.missingEvidenceCount?' '+result.missingEvidenceCount+' dokumen belum memiliki berkas digital (lihat catatan dalam ZIP).':''}`;
   }else{
    state.year=year;state.month=month;state.holidays=runSettings.holidays;state.holidayNotes=runSettings.holidayNotes;
    state.employees=hydrated;state.leaves=supportData.leaves;state.assignments=supportData.assignments;state.files=[];state.savedRun=meta;state.archiveKey=meta.archiveKey||'';
    state.dirty=false;state.step=4;state.view='process';location.hash='process';
    state.message='Riwayat dimuat. Untuk mengolah ulang, pilih Ubah Periode dan unggah berkas baru.';
   }
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
  case 'connect-sharepoint':await connectSharePoint();break;
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
  case 'verify':if(permission('process.verify')){state.modalId=button.dataset.id;const e=findEmployee(state.modalId);state.modalDate=Object.keys(e?.days||{}).find(k=>e.days[k].issues.length&&!e.days[k].reviewed)||R.key(state.year,state.month,1);render();void loadSupportingPreviews();}break;
  case 'select-day':state.modalDate=button.dataset.date;render();void loadSupportingPreviews();break;
  case 'retry-support-preview':{
   if(!permission('process.verify'))break;
   delete state.documentPreviews[button.dataset.docKey];
   refreshSupportingFiles();void loadSupportingPreviews();break;
  }
  case 'save-day':saveDay();break;
  case 'save-support':await saveSupport(button.dataset.kind,button.dataset.sourceId,button.dataset.index);break;
  case 'acknowledge':acknowledge(button.dataset.kind);break;
  case 'close-modal':state.modalId='';state.modalDate='';render();break;
  case 'back':state.step=Math.max(1,state.step-1);render();break;
  case 'restart':if(state.employees.length&&!confirm('Proses saat ini akan dikosongkan. Periode baru membutuhkan unggah absensi. Lanjutkan?'))break;resetProcess();render();break;
  case 'preview':state.step=4;render();break;
  case 'export':await exportData();break;
  case 'export-zip':await exportZip();break;
  case 'save':await saveRun();break;
  case 'history-open':await openHistory(button.dataset.id);break;
  case 'history-download':await openHistory(button.dataset.id,'excel');break;
  case 'history-download-zip':await openHistory(button.dataset.id,'zip');break;
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
  // Memuat status dari akun Microsoft yang masih tersimpan di sesi browser;
  // hanya memeriksa, tidak memunculkan login tanpa tindakan pengguna.
  // Pemeriksaan SharePoint berjalan terpisah agar keterlambatan Microsoft Graph
  // tidak menghalangi pengolahan absensi dan riwayat Uang Makan.
  void checkSharePoint();
  await loadMaster();
  await refreshHistory();
  state.busy=false;render();
  window.ResourceLoader?.warm?.(['xlsx','jszip']);
 }catch(e){state.busy=false;state.error=e.message||String(e);render();}
}
init();
})();
