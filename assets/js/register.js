/* Register Status Administrasi Kehadiran Pegawai - pengembangan modul eksisting. */
(function () {
 'use strict';
 const app=document.getElementById('app'), E=window.AttendanceRegisterEngine;
 const now=new Date();
 const state={year:now.getFullYear(),month:now.getMonth()+1,view:'dashboard',loading:true,busy:false,error:'',notice:'',employees:new Map(),sources:[],records:[],recons:[],logs:[],modal:null,search:'',filter:'all'};
 const VIEWS={dashboard:'Dashboard',cases:'Konflik & Realisasi',reconcile:'Rekonsiliasi',history:'Riwayat Klarifikasi'};
 const STATUS={pending:'Belum Verifikasi',awaiting:'Belum Verifikasi',verified:'Sudah Verifikasi',closed:'Sudah Verifikasi'};
 const REAL={unknown:'Belum dikonfirmasi',done:'Terlaksana',cancelled:'Tidak terlaksana/dilaporkan batal',partial:'Terlaksana sebagian',rescheduled:'Jadwal berubah',leave_not_taken:'Cuti tidak diambil'};
 const KIND={cuti_surat:'Cuti dan Surat Tugas beririsan',surat_ganda:'Beberapa Surat Tugas beririsan',cuti_ganda:'Beberapa Cuti beririsan',realization:'Pelaporan realisasi'};
 const $=s=>document.querySelector(s);
 const esc=v=>String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
 const db=()=>window.FirebaseClient.getDb();
 const actor=()=>window.FirebaseClient.getCurrentUser();
 const ts=()=>window.firebase.firestore.FieldValue.serverTimestamp();
 const can=(key)=>window.UserAccess.can('register',key);
 const formatDate=s=>E.iso(s)?new Date(`${s}T00:00:00Z`).toLocaleDateString('id-ID',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'}):'-';
 const fmt=(s)=>s?new Date(s).toLocaleString('id-ID'):'-';
 const period=()=>`${state.year}-${String(state.month).padStart(2,'0')}`;
 const inPeriod=s=>E.inPeriod(s,state.year,state.month);
 const empName=id=>state.employees.get(String(id))?.name || state.employees.get(String(id))?.fullName || String(id);
 const recordMap=()=>new Map(state.records.map(r=>[r.key,r]));
 const sourceMap=()=>new Map(state.sources.map(s=>[s.key,s]));
 
 function dateYear(s){return Number(String(s||'').slice(0,4))||state.year;}
 function sourceLabel(s){return `${s.type==='cuti'?'Cuti':'Surat Tugas'} · ${s.number || s.id} · ${empName(s.employeeId)} · ${formatDate(s.startDate)}–${formatDate(s.endDate)}`;}
 function allRows(){
   const saved=recordMap(), sources=sourceMap();
   const automatic=E.conflicts(state.sources).filter(inPeriod).map(c=>({ ...c, type:'conflict', record:saved.get(c.key)||null }));
   const keys=new Set(automatic.map(c=>c.key));
   const entries=state.records.filter(r=>r.type==='realization' || (r.type==='conflict'&&!keys.has(r.key)))
     .filter(r=>inPeriod({startDate:r.startDate,endDate:r.endDate||r.startDate})).map(r=>{
       const s=r.type==='realization'?sources.get(r.sourceKey):null;
       return {key:r.key,type:r.type,kind:r.kind||'realization',employeeId:r.employeeId,
         startDate:r.startDate,endDate:r.endDate, sources:s?[s]:[],fingerprint:s?.fingerprint||'',record:r,missing:!s&&r.type==='realization'||r.type==='conflict'};
     });
   return [...automatic,...entries].map(r=>{
     const stale=Boolean(r.record && (r.missing || !r.fingerprint || r.record.fingerprint!==r.fingerprint));
     return {...r,stale,workflow:stale?'stale':r.record?.workflow||'pending'};
   }).sort((a,b)=>b.startDate.localeCompare(a.startDate)||a.key.localeCompare(b.key));
 }
 function counts(){const rows=allRows();return {total:rows.length,open:rows.filter(r=>!['verified','closed'].includes(r.workflow)).length,
   conflict:rows.filter(r=>r.type==='conflict'&&!['verified','closed'].includes(r.workflow)).length,
   stale:rows.filter(r=>r.stale).length,verified:rows.filter(r=>['verified','closed'].includes(r.workflow)).length};}
 function reconcileState(){
   const rec=state.recons.find(r=>r.id===period());
   const rows=allRows();
   const fp=E.reconciliationFingerprint(state.sources,rows.map(r=>({key:r.key,startDate:r.startDate,endDate:r.endDate,fingerprint:r.fingerprint,workflow:r.workflow,realization:r.record?.realization||'',selection:(r.record?.selectedSourceKeys||[]).join('|'),note:r.record?.note||'',at:r.record?.updatedAt?.seconds||''})),state.year,state.month);
   const changed=!!rec&&rec.fingerprint!==fp;
   return {rec,fp,changed,rows,open:rows.filter(r=>!['verified','closed'].includes(r.workflow)).length,
     sources:state.sources.filter(inPeriod).length};
 }
 function banner(){return `<div class="rg-lead"><div><span class="rg-overline">PUSDATIN · TATA USAHA</span><h1>Register Administrasi Kehadiran</h1><p>Rekonsiliasi Cuti, Surat Tugas, dan laporan realisasi. Data sumber tetap dikelola pada modul masing-masing.</p></div><div class="rg-period"><label for="rg-year">Tahun pemeriksaan</label><select id="rg-year">${[state.year-2,state.year-1,state.year,state.year+1].map(y=>`<option value="${y}" ${y===state.year?'selected':''}>${y}</option>`).join('')}</select></div></div>`;}
 function toolbar(){return `<div class="rg-tabs" role="tablist" aria-label="Menu Register">${Object.entries(VIEWS).filter(([k])=>can(`${k}.view`)).map(([k,label])=>`<a role="tab" aria-selected="${state.view===k}" class="${state.view===k?'active':''}" href="#${k}">${esc(label)}</a>`).join('')}</div>`;}
 function metric(n,label,sub,flag){return `<div class="rg-metric ${flag||''}"><strong>${n}</strong><span>${label}</span><small>${sub}</small></div>`;}
 function dashboard(){const n=counts(),recon=reconcileState();const recent=allRows().slice(0,6);
   return `<div class="rg-metrics">${metric(n.open,'Perlu tindak lanjut','Konflik dan laporan realisasi','warn')}${metric(n.conflict,'Potensi konflik','Belum diverifikasi','danger')}${metric(n.stale,'Perlu pemeriksaan ulang','Sumber berubah atau tidak ditemukan','warn')}${metric(n.verified,'Terverifikasi','Kasus dalam periode','')}</div>
   <div class="rg-grid"><section class="card rg-card"><div class="rg-card-head"><h2>Perhatian TU</h2><a href="#cases">Lihat semua kasus</a></div><p class="rg-muted">Kasus yang belum diselesaikan ditampilkan terlebih dahulu. Dua dokumen beririsan merupakan indikasi, bukan keputusan bahwa dokumen tidak sah.</p>${recent.length?recent.map(rowCompact).join(''):'<div class="rg-empty">Belum ada konflik atau laporan realisasi pada periode ini.</div>'}</section>
   <section class="card rg-card"><h2>Rekonsiliasi ${esc(period())}</h2><div class="rg-status ${recon.changed?'rg-warn':recon.rec?.status==='complete'?'rg-ok':'rg-warn'}">${recon.changed?'Perlu ditinjau ulang — data berubah':recon.rec?.status==='complete'?'Selesai direkonsiliasi':'Belum selesai direkonsiliasi'}</div><dl class="rg-dl"><dt>Entri dokumen per pegawai</dt><dd>${recon.sources}</dd><dt>Kasus belum selesai</dt><dd>${recon.open}</dd><dt>Rekonsiliasi terakhir</dt><dd>${esc(recon.rec?.reviewedAt?fmt(recon.rec.reviewedAt):'-')}</dd></dl><p class="rg-muted">Status rekonsiliasi mencatat hasil pemeriksaan TU, bukan jaminan bahwa seluruh surat resmi sudah masuk ke aplikasi.</p><a class="btn btn-primary" href="#reconcile">Buka Rekonsiliasi</a></section></div>`;
 }
 function rowCompact(row){return `<button class="rg-compact" data-open="${esc(row.key)}" type="button"><span><b>${esc(empName(row.employeeId))}</b><small>${esc(KIND[row.kind]||row.kind)} · ${formatDate(row.startDate)}</small></span>${statusPill(row)}</button>`;}
 function statusPill(row){return `<span class="rg-pill ${row.workflow==='stale'?'rg-pill-warn':['verified','closed'].includes(row.workflow)?'rg-pill-ok':row.workflow==='awaiting'?'rg-pill-warn':'rg-pill-gray'}">${esc(row.workflow==='stale'?'Belum Verifikasi':STATUS[row.workflow]||STATUS.pending)}</span>`;}
 function cases(){const rows=allRows().filter(row=>{
   if(state.filter==='open'&&['verified','closed'].includes(row.workflow))return false;
   if(state.filter==='conflict'&&row.type!=='conflict')return false;
   if(state.filter==='stale'&&!row.stale)return false;
   const search=state.search.trim().toLowerCase();return !search || [empName(row.employeeId),KIND[row.kind],row.startDate,...row.sources.map(s=>`${s.number} ${s.activity}`)].join(' ').toLowerCase().includes(search);
 });
 return `<section class="card rg-card"><div class="rg-card-head"><div><h2>Konflik dan Realisasi</h2><p class="rg-muted">Mendeteksi irisan otomatis dan mencatat laporan perubahan per pegawai/dokumen.</p></div>${can('cases.create')?'<button class="btn btn-primary" type="button" data-new-realization>+ Catat realisasi</button>':''}</div>
 <div class="rg-controls"><input id="rg-search" value="${esc(state.search)}" placeholder="Cari nama, nomor surat, atau jenis kasus" aria-label="Cari kasus"><select id="rg-filter"><option value="all">Semua status</option><option value="open" ${state.filter==='open'?'selected':''}>Perlu tindak lanjut</option><option value="conflict" ${state.filter==='conflict'?'selected':''}>Potensi konflik</option><option value="stale" ${state.filter==='stale'?'selected':''}>Perlu tinjau ulang</option></select><button class="btn btn-secondary" type="button" data-export>Ekspor CSV</button></div>
 <div class="rg-table-wrap"><table class="rg-table"><thead><tr><th>Pegawai</th><th>Jenis temuan</th><th>Tanggal irisan / kegiatan</th><th>Dokumen terkait</th><th>Status</th><th>Tindakan</th></tr></thead><tbody>${rows.length?rows.map(r=>`<tr><td><b>${esc(empName(r.employeeId))}</b></td><td>${esc(KIND[r.kind]||r.kind)}</td><td>${formatDate(r.startDate)}${r.endDate&&r.endDate!==r.startDate?' – '+formatDate(r.endDate):''}</td><td>${r.sources.length?r.sources.map(s=>`<span class="rg-doc">${esc(s.label)}</span>`).join(''):'<span class="rg-muted">Sumber berubah / tidak ditemukan</span>'}</td><td>${statusPill(r)}</td><td><button class="btn btn-secondary btn-sm" data-open="${esc(r.key)}" type="button">${r.record?'Detail':'Tindak lanjut'}</button></td></tr>`).join(''):'<tr><td colspan="6" class="rg-empty">Tidak ditemukan kasus sesuai filter. Ini belum berarti seluruh dokumen resmi sudah tercatat.</td></tr>'}</tbody></table></div><p class="rg-muted rg-footnote">Total ${rows.length} temuan dalam filter. Surat tugas dengan dua peserta diproses per pegawai, bukan digabung menjadi satu hasil verifikasi.</p></section>`;}
 function reconcile(){const r=reconcileState(),rec=r.rec||{};return `<section class="card rg-card"><div class="rg-card-head"><div><h2>Rekonsiliasi Dokumen</h2><p class="rg-muted">Pemeriksaan sebelum finalisasi dan pemeriksaan ulang ketika ada dokumen terlambat.</p></div><label class="rg-select"><span>Periode</span><input id="rg-month" type="month" value="${esc(period())}"></label></div>
 <div class="rg-check-grid">${metric(r.sources,'Entri dokumen pegawai','Periode terpilih')}${metric(r.rows.length,'Total kasus','Potensi konflik dan perubahan')}${metric(r.open,'Belum selesai','Perlu tindak lanjut',r.open?'warn':'')}${metric(r.changed?'Ya':'Tidak','Perubahan sejak checkpoint','Pencatatan sebelumnya',r.changed?'danger':'')}</div>
 ${r.changed?'<div class="rg-alert rg-alert-warn"><b>Pemeriksaan ulang diperlukan.</b> Dokumen, realisasi, atau hasil verifikasi berubah setelah rekonsiliasi terakhir. Tinjau dampaknya pada hasil Tukin/Uang Makan/Lembur yang mungkin sudah diproses. Tidak ada hasil keuangan yang diubah otomatis.</div>':''}
 ${r.open?`<div class="rg-alert rg-alert-warn">Ada ${r.open} kasus belum selesai. Status rekonsiliasi tidak dapat ditandai lengkap sebelum kasus tersebut ditindaklanjuti.</div>`:''}
 <form id="rg-reconcile-form" class="rg-form"><div class="rg-fields"><label><span>Hasil pemeriksaan TU</span><select name="status"><option value="pending" ${rec.status!=='complete'?'selected':''}>Masih ada yang perlu diperiksa</option><option value="complete" ${rec.status==='complete'?'selected':''}>Selesai direkonsiliasi</option></select></label><label><span>Konfirmasi sumber dokumen (wajib jika selesai)</span><input name="sourceConfirmation" maxlength="240" placeholder="Misalnya: daftar surat dari tiap bidang telah dicocokkan" value="${esc(rec.sourceConfirmation||'')}"></label></div><label><span>Catatan pemeriksaan</span><textarea name="notes" rows="3" maxlength="1500" placeholder="Dokumen masuk, tindak lanjut, rujukan, pengecualian">${esc(rec.notes||'')}</textarea></label><div class="rg-actions"><button class="btn btn-primary" type="submit" ${!can('reconcile.save')||Boolean(state.error)?'disabled':''}>Simpan hasil rekonsiliasi</button><a class="btn btn-secondary" href="#cases">Tinjau konflik</a></div></form><p class="rg-muted">Rekonsiliasi menyimpan fingerprint data pada saat pemeriksaan. Perubahan sumber setelahnya akan ditandai ketika register dibuka kembali.</p></section>`;}
 function history(){const rows=state.logs.filter(l=>l.year>=state.year-1&&l.year<=state.year+1).sort((a,b)=>String(b.createdAt?.seconds||b.at||'').localeCompare(String(a.createdAt?.seconds||a.at||''))).slice(0,150);
 return `<section class="card rg-card"><h2>Riwayat Perubahan dan Klarifikasi</h2><p class="rg-muted">Catatan tindakan bersifat append-only. Riwayat tidak dapat diedit atau dihapus melalui aplikasi.</p><div class="rg-table-wrap"><table class="rg-table"><thead><tr><th>Waktu</th><th>Petugas</th><th>Tindakan</th><th>Pegawai / kasus</th><th>Keterangan</th></tr></thead><tbody>${rows.length?rows.map(r=>`<tr><td>${esc(r.at?fmt(r.at):r.createdAt?.toDate?.().toLocaleString('id-ID')||'-')}</td><td>${esc(r.actorEmail||'-')}</td><td>${esc(r.action||'-')}</td><td>${esc(empName(r.employeeId))}<small class="rg-doc">${esc(r.kind||'')}</small></td><td>${esc(r.detail||'')}</td></tr>`).join(''):'<tr><td colspan="5" class="rg-empty">Belum ada riwayat pada rentang tahun ini.</td></tr>'}</tbody></table></div></section>`;}
 function modalMarkup(){if(!state.modal)return '';
   const m=state.modal, r=m.record||{}, isConflict=m.type==='conflict';
   const sourceOptions=state.sources.filter(s=>s.startDate.slice(0,4)===String(state.year)||inPeriod(s)).sort((a,b)=>a.startDate.localeCompare(b.startDate));
   const selected=state.sources.find(s=>s.key===(m.sourceKey||r.sourceKey));
   const mayWrite=can(r.key?'cases.edit':'cases.create');
   const mayVerify=can('cases.verify') && (Boolean(r.key) || can('cases.create'));
   return `<div class="rg-overlay" data-close-modal><section class="rg-modal" role="dialog" aria-modal="true" aria-label="Tindak lanjut administrasi"><div class="rg-modal-head"><div><span class="rg-overline">REGISTER KEHADIRAN</span><h2>${isConflict?'Tindak Lanjut Konflik':'Status Realisasi Dokumen'}</h2></div><button class="rg-close" type="button" data-close-modal aria-label="Tutup">×</button></div><div class="rg-modal-body"><div class="rg-context"><b>${esc(empName(m.employeeId||r.employeeId||selected?.employeeId))}</b><small>${esc(isConflict?(KIND[m.kind]||m.kind):(selected?sourceLabel(selected):'Pilih dokumen dan pegawai'))}</small>${isConflict?m.sources.map(s=>`<small>${esc(sourceLabel(s))}</small>`).join(''):''}</div>
 ${m.stale?'<div class="rg-alert rg-alert-warn">Dokumen sumber telah berubah atau tidak ditemukan. Verifikasi sebelumnya tidak lagi dianggap berlaku. Cocokkan kembali dengan sumber resmi.</div>':''}
 <form id="rg-case-form"><input type="hidden" name="caseKey" value="${esc(m.key||'')}"><div class="rg-fields">
 ${!isConflict?`<label class="rg-full"><span>Dokumen dan pegawai *</span><select name="sourceKey" required ${r.key?'disabled':''}><option value="">Pilih dokumen</option>${sourceOptions.map(s=>`<option value="${esc(s.key)}" ${s.key===(m.sourceKey||r.sourceKey)?'selected':''}>${esc(sourceLabel(s))}</option>`).join('')}</select></label><label><span>Realisasi *</span><select name="realization">${Object.entries(REAL).map(([key,label])=>`<option value="${key}" ${r.realization===key?'selected':''}>${esc(label)}</option>`).join('')}</select></label><label><span>Status tindak lanjut</span><input value="${esc(STATUS[r.workflow]||STATUS.pending)}" disabled></label><label><span>Tanggal realisasi mulai</span><input type="date" name="actualStart" value="${esc(r.actualStart||'')}"></label><label><span>Tanggal realisasi akhir</span><input type="date" name="actualEnd" value="${esc(r.actualEnd||'')}"></label><label class="rg-full"><span>Tanggal pelaksanaan tertentu (opsional, untuk realisasi tidak berurutan)</span><input name="actualDates" placeholder="2026-10-12, 2026-10-14" value="${esc((r.actualDates||[]).join(', '))}"></label><label><span>Jam mulai (opsional)</span><input type="time" name="actualStartTime" value="${esc(r.actualStartTime||'')}"></label><label><span>Jam selesai (opsional)</span><input type="time" name="actualEndTime" value="${esc(r.actualEndTime||'')}"></label>`:
 `<div class="rg-full"><strong>Dokumen yang ditetapkan</strong><p class="rg-muted">Centang satu atau beberapa dokumen yang berlaku. Dokumen tidak dicentang tetap tersimpan dan dapat dipilih kembali, tetapi tidak ikut perhitungan/ekspor setelah verifikasi.</p><label class="rg-select-all"><input type="checkbox" data-select-all ${(!Array.isArray(r.selectedSourceKeys)||m.sources.every(s=>r.selectedSourceKeys.includes(s.key)))?'checked':''}> Pilih semua</label>${m.sources.map(s=>`<label class="rg-select-doc"><input type="checkbox" name="selectedSourceKeys" value="${esc(s.key)}" ${!Array.isArray(r.selectedSourceKeys)||r.selectedSourceKeys.includes(s.key)?'checked':''}>${esc(sourceLabel(s))}</label>`).join('')}</div>`}
 <label class="rg-full"><span>Catatan (opsional)</span><textarea name="note" rows="3" maxlength="2200" placeholder="Catatan tambahan apabila diperlukan">${esc(r.note||'')}</textarea></label></div>
 <div class="rg-actions"><button type="button" class="btn btn-secondary" data-close-modal>Tutup</button>${mayWrite?'<button type="submit" class="btn btn-secondary" name="intent" value="save">Simpan — Belum Verifikasi</button>':''}${mayVerify?'<button type="submit" class="btn btn-primary" name="intent" value="verify">Simpan — Sudah Verifikasi</button>':''}</div></form><p class="rg-muted">Verifikasi TU merupakan pencatatan administratif, bukan pembatalan formal keputusan pejabat atau penetapan otomatis hak keuangan.</p></div></section></div>`;
 }
 function render(){
   if(!app)return;const content=state.loading?'<div class="card rg-card">Memuat data dokumen dan pemeriksaan register...</div>':
    `${state.error?`<div class="rg-alert rg-alert-warn">${esc(state.error)}</div>`:''}${state.notice?`<div class="rg-alert rg-alert-ok">${esc(state.notice)}</div>`:''}${banner()}${toolbar()}${state.view==='dashboard'?dashboard():state.view==='cases'?cases():state.view==='reconcile'?reconcile():history()}`;
   app.innerHTML=window.AppShell.render({module:'register',view:state.view,viewLabel:VIEWS[state.view],content:`<div class="rg">${content}</div>`,overlays:modalMarkup()});
 }
 async function load(){state.loading=true;state.error='';state.notice='';render();
   try {
    const years=[state.year-1,state.year,state.year+1];
    const d=db();
    const [bundle,leaveSnap,stSnap,caseSnap,reconSnap,logSnap]=await Promise.all([
      window.MasterDataService.getBundle(false),
      d.collection('leaveRecords').where('leaveYear','in',years).get(),
      d.collection('assignmentRecords').where('assignmentYear','in',years).get(),
      d.collection('attendanceCases').where('year','in',years).get(),
      d.collection('attendanceReconciliations').where('year','==',state.year).get(),
      can('history.view')?d.collection('attendanceCaseLogs').where('year','in',years).get():Promise.resolve({docs:[]})
    ]);
    const employees=[...bundle.employees];
    state.employees=new Map(employees.map(e=>[String(e.id), {...e,...(bundle.directory.get(String(e.id))||{}),id:e.id}]));
    state.sources=E.makeSources(leaveSnap.docs.map(x=>({...x.data(),id:x.id})),stSnap.docs.map(x=>({...x.data(),id:x.id})));
    state.records=caseSnap.docs.map(x=>({id:x.id,...x.data()}));
    state.recons=reconSnap.docs.map(x=>({id:x.id,...x.data()}));
    state.logs=logSnap.docs.map(x=>({id:x.id,...x.data()}));
   } catch(e){state.error=`Data register gagal dimuat: ${e.message||e}. Periksa hak akses, koneksi, dan Firestore Rules terbaru.`;state.sources=[];state.records=[];state.logs=[];state.recons=[];}
   finally{state.loading=false;render();}
 }
 function openCase(key){const row=allRows().find(x=>x.key===key);if(!row)return;state.modal={...row};render();}
 function newRealization(){state.modal={type:'realization',kind:'realization',key:'',employeeId:'',record:null};render();}
 async function caseId(key){const data=new TextEncoder().encode(key);const bytes=await crypto.subtle.digest('SHA-256',data);return 'rg-'+[...new Uint8Array(bytes)].map(x=>x.toString(16).padStart(2,'0')).join('');}
 async function saveCase(form,intent){
   if(!state.modal) return;
   if(state.error)throw new Error('Kasus tidak dapat disimpan selama data belum berhasil dimuat.');
   const m=state.modal, isConflict=m.type==='conflict', old=m.record||{}, values=Object.fromEntries(new FormData(form));
   const src=isConflict?null:state.sources.find(s=>s.key===(m.sourceKey||old.sourceKey||values.sourceKey));
   if(!isConflict&&!src)throw new Error('Pilih dokumen sumber yang masih tersedia.');
   const key=isConflict?m.key:`realization:${src.key}`;
   const verify=intent==='verify';
   const needsUpdate=Boolean(old.key);
   if(verify&&!can('cases.verify'))throw new Error('Tidak memiliki izin verifikasi TU.');
   if(!verify&&!can(needsUpdate?'cases.edit':'cases.create'))throw new Error('Tidak memiliki izin menyimpan tindak lanjut.');
   const note=String(values.note||'').trim();
   const selectedSourceKeys=isConflict?[...new FormData(form).getAll('selectedSourceKeys')]:[];
   if(isConflict&&verify&&!selectedSourceKeys.length)throw new Error('Pilih sekurang-kurangnya satu dokumen untuk diverifikasi.');
   const actualDates=String(values.actualDates||'').split(',').map(v=>v.trim()).filter(Boolean);
   if(actualDates.some(v=>!E.iso(v)))throw new Error('Daftar tanggal realisasi harus berformat YYYY-MM-DD dan dipisahkan koma.');
   const uniqueDates=[...new Set(actualDates)].sort();
   if(uniqueDates.length&&(!values.actualStart||!values.actualEnd)){
     values.actualStart=uniqueDates[0];values.actualEnd=uniqueDates[uniqueDates.length-1];
   }
   if(values.actualStart&&!E.iso(values.actualStart)||values.actualEnd&&!E.iso(values.actualEnd))throw new Error('Tanggal realisasi tidak valid.');
   if(values.actualStart&&values.actualEnd&&values.actualEnd<values.actualStart)throw new Error('Tanggal akhir realisasi tidak boleh sebelum tanggal mulai.');
   if(!isConflict&&['done','partial','rescheduled'].includes(values.realization)&&(!values.actualStart||!values.actualEnd))throw new Error('Isi tanggal realisasi mulai dan akhir untuk kegiatan yang terlaksana atau berubah jadwal.');
   if(!isConflict&&values.actualStartTime&&values.actualEndTime&&values.actualStartTime>values.actualEndTime&&values.actualStart===values.actualEnd)throw new Error('Jam selesai tidak boleh lebih awal untuk tanggal yang sama.');
   const nowIso=new Date().toISOString(), sourceList=isConflict?m.sources:[src];
   const payload={key,type:m.type,kind:m.kind||'realization',year:dateYear(m.startDate||src?.startDate),employeeId:m.employeeId||src.employeeId,
     startDate:m.startDate||src.startDate,endDate:m.endDate||src.endDate,
     sourceKey:isConflict?'':src.key, sourceRefs:sourceList.map(s=>({type:s.type,id:s.id,employeeId:s.employeeId,number:s.number,startDate:s.startDate,endDate:s.endDate})),
     fingerprint:isConflict?m.fingerprint:src.fingerprint,
     workflow:verify?'verified':'pending',selectedSourceKeys:isConflict?selectedSourceKeys:[],
     realization:isConflict?'':String(values.realization||'unknown'),
     actualStart:isConflict?'':String(values.actualStart||''),actualEnd:isConflict?'':String(values.actualEnd||''),
     actualStartTime:isConflict?'':String(values.actualStartTime||''),actualEndTime:isConflict?'':String(values.actualEndTime||''),actualDates:isConflict?[]:uniqueDates,
     clarifiedBy:'',evidence:'',note,
     updatedAt:ts(),updatedByUid:actor().uid,updatedByEmail:actor().email||'',
     verifiedAt:verify?nowIso:'',verifiedByUid:verify?actor().uid:'',verifiedByEmail:verify?(actor().email||''):''};
   const ref=db().collection('attendanceCases').doc(await caseId(key)),batch=db().batch();
   if(!needsUpdate){payload.createdAt=ts();payload.createdByUid=actor().uid;payload.createdByEmail=actor().email||'';}
   batch.set(ref,payload,{merge:needsUpdate});
   batch.set(db().collection('attendanceCaseLogs').doc(),{
     year:payload.year,caseKey:key,kind:payload.kind,employeeId:payload.employeeId,action:verify?'Verifikasi TU':needsUpdate?'Perbarui tindak lanjut':'Catat kasus/realisasi',
     detail:`${note.slice(0,650)}${isConflict?' | Dokumen dipilih: '+selectedSourceKeys.join(', ').slice(0,300):''}`, 
     actorUid:actor().uid,actorEmail:actor().email||'',at:nowIso,createdAt:ts()});
   await batch.commit();state.modal=null;await load();state.notice='Tindak lanjut dan riwayat berhasil disimpan.';render();
 }
 async function saveRecon(form){
  if(state.error)throw new Error('Rekonsiliasi tidak dapat disimpan selama data belum berhasil dimuat.');
  if(!can('reconcile.save'))throw new Error('Anda tidak memiliki izin untuk menyimpan rekonsiliasi.');
  const v=Object.fromEntries(new FormData(form)),r=reconcileState();
  if(v.status==='complete'&&(r.open>0||!String(v.sourceConfirmation||'').trim()))throw new Error('Selesaikan kasus terbuka dan isi konfirmasi pemeriksaan sumber sebelum menandai rekonsiliasi selesai.');
  const payload={year:state.year,month:state.month,status:v.status==='complete'?'complete':'pending',fingerprint:r.fp,
    totalSources:r.sources,openCases:r.open,sourceConfirmation:String(v.sourceConfirmation||'').trim().slice(0,240),notes:String(v.notes||'').trim().slice(0,1500),
    reviewedAt:new Date().toISOString(),reviewedByUid:actor().uid,reviewedByEmail:actor().email||'',updatedAt:ts()};
  await db().collection('attendanceReconciliations').doc(period()).set(payload,{merge:true});
  await load();state.notice='Checkpoint rekonsiliasi tersimpan. Perubahan sumber selanjutnya akan ditandai.';render();
 }
 function exportCsv(){const rows=allRows().map(r=>({pegawai:empName(r.employeeId),jenis:KIND[r.kind]||r.kind,mulai:r.startDate,selesai:r.endDate, dokumen:r.sources.map(s=>s.number||s.id).join(' / '),status:STATUS[r.workflow]||'Perlu pemeriksaan ulang',realisasi:REAL[r.record?.realization]||'',klarifikasi:r.record?.note||''}));
  const safe=v=>{let s=String(v??'');if(/^[\s]*[=+@-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';};
  const headers=['Pegawai','Jenis','Mulai','Selesai','Dokumen','Status','Realisasi','Klarifikasi'];const csv='\uFEFF'+[headers.map(safe).join(';'),...rows.map(r=>Object.values(r).map(safe).join(';'))].join('\r\n');
  const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=`register-kehadiran-${period()}.csv`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 }
 function nav(){const v=String(location.hash||'').replace('#','').split('?')[0];state.view=VIEWS[v]&&can(`${v}.view`)?v:(Object.keys(VIEWS).find(k=>can(`${k}.view`))||'dashboard');state.modal=null;render();}
 app.addEventListener('click',e=>{
  const close=e.target.closest('[data-close-modal]');if(close&&(e.target===close||close.tagName==='BUTTON')){state.modal=null;render();return;}
  const all=e.target.closest('[data-select-all]');if(all){app.querySelectorAll('[name=selectedSourceKeys]').forEach(x=>{x.checked=all.checked;});return;}
  const open=e.target.closest('[data-open]');if(open){openCase(open.dataset.open);return;}
  if(e.target.closest('[data-new-realization]')){newRealization();return;}
  if(e.target.closest('[data-export]')){exportCsv();return;}
 });
 app.addEventListener('change',e=>{
  if(e.target.id==='rg-year'){state.year=Number(e.target.value);state.modal=null;load();}
  if(e.target.id==='rg-month'){const [y,m]=e.target.value.split('-').map(Number);if(y!==state.year){state.year=y;state.month=m;load();}else{state.month=m;render();}}
  if(e.target.id==='rg-filter'){state.filter=e.target.value;render();}
  if(e.target.name==='selectedSourceKeys'){const boxes=[...app.querySelectorAll('[name=selectedSourceKeys]')];const all=app.querySelector('[data-select-all]');if(all)all.checked=boxes.every(x=>x.checked);}
 });
 app.addEventListener('input',e=>{
  if(e.target.id==='rg-search'){state.search=e.target.value;const pos=e.target.selectionStart;render();const field=$('#rg-search');if(field){field.focus();field.setSelectionRange(pos,pos);}}
 });
 app.addEventListener('submit',async e=>{
  if(!['rg-case-form','rg-reconcile-form'].includes(e.target.id))return;e.preventDefault();
  if(state.busy)return;
  state.busy=true;
  const button=e.submitter;const intent=button?.value||'save';if(button)button.disabled=true;
  try{if(e.target.id==='rg-case-form')await saveCase(e.target,intent);else await saveRecon(e.target);}
  catch(err){alert(err.message||String(err));}
  finally{state.busy=false;if(button&&button.isConnected)button.disabled=false;}
 });
 window.addEventListener('hashchange',nav);
 (async()=>{
  try{if(!await window.UserAccess.requireModule('register'))return;nav();await load();}
  catch(e){state.loading=false;state.error=e.message||String(e);render();}
 })();
})();
