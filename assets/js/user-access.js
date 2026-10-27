/* Hak akses rinci Pusdatin PKP. Kunci izin harus sama dengan firebase/firestore.rules. */
(function () {
  'use strict';
  const MODULES = [
    { id: 'cuti', label: 'Cuti', sections: [
      { id:'dashboard',label:'Dashboard Cuti',actions:[] },
      { id:'data',label:'Data Cuti',actions:[['create','Input Data Cuti'],['edit','Edit Data Cuti'],['delete','Hapus Data Cuti']] },
      { id:'kalender',label:'Kalender Cuti',actions:[] },
      { id:'pegawai',label:'Data Pegawai',actions:[] },
      { id:'saldo',label:'Saldo Cuti',actions:[['edit_balance','Edit Saldo Cuti Tahunan']] },
      { id:'laporan',label:'Laporan Cuti',actions:[['print','Cetak Laporan']] },
      { id:'documents',label:'Bukti Dukung Cuti',parent:'data',actions:[['download','Unduh Dokumen'],['create','Unggah Dokumen'],['edit','Edit Dokumen'],['delete','Hapus Dokumen']] }
    ]},
    { id:'surat',label:'Surat Tugas',sections:[
      { id:'dashboard',label:'Dashboard Surat Tugas',actions:[] },
      { id:'data',label:'Data Surat Tugas',actions:[['create','Input Surat Tugas'],['edit','Edit Surat Tugas'],['delete','Hapus Surat Tugas']] },
      { id:'documents',label:'Berkas Surat Tugas',parent:'data',actions:[['upload','Unggah/Ganti Berkas'],['download','Unduh Berkas']] },
      { id:'kalender',label:'Kalender Surat Tugas',actions:[] }
    ]},
    { id:'pelatihan',label:'Pelatihan Pegawai',sections:[
      { id:'dashboard',label:'Dashboard Pelatihan',actions:[] },
      { id:'pegawai',label:'Data Pegawai',actions:[] },
      { id:'data',label:'Data Pelatihan',actions:[['create','Input Pelatihan'],['edit','Edit Data Pelatihan'],['delete','Hapus Pelatihan']] },
      { id:'participants',label:'Status & Catatan Peserta',parent:'data',actions:[['edit','Edit Status/Catatan Peserta']] },
      { id:'documents',label:'Dokumen Pelatihan',parent:'data',actions:[['create','Tambah Metadata'],['edit','Edit Metadata'],['delete','Hapus Metadata']] }
    ]},
    { id:'tukin',label:'Tunjangan Kinerja',sections:[
      { id:'dashboard',label:'Dashboard Tukin',actions:[] },
      { id:'process',label:'Proses Tukin',actions:[['upload','Unggah Data Absensi'],['calculate','Jalankan Perhitungan'],['verify','Verifikasi/Koreksi Potongan'],['evidence','Kelola Bukti Dukung'],['save','Simpan Hasil Proses']] },
      { id:'parameters',label:'Parameter Perhitungan',parent:'process',actions:[['holidays','Ubah Tanggal Merah'],['ramadan','Ubah Pengaturan Ramadan']] },
      { id:'history',label:'Riwayat Tukin',actions:[['download','Unduh Hasil'],['edit','Edit Riwayat'],['delete','Hapus Riwayat']] }
    ]},
    { id:'uangMakan',label:'Uang Makan',sections:[
      {id:'dashboard',label:'Dashboard Uang Makan',actions:[]},
      {id:'process',label:'Proses Uang Makan',actions:[['upload','Unggah Absensi'],['calculate','Hitung dan Sinkronkan'],['verify','Verifikasi/Koreksi Harian'],['save','Simpan Hasil'],['export','Ekspor Excel / ZIP']]},
      {id:'history',label:'Riwayat Uang Makan',actions:[['download','Unduh Excel / ZIP'],['delete','Hapus Riwayat']]}
    ]},
    { id:'lembur',label:'Lembur',sections:[
      { id:'dashboard',label:'Dashboard Lembur',actions:[] },
      { id:'process',label:'Proses Lembur',actions:[['upload','Unggah Data Absensi'],['calculate','Jalankan Proses'],['correct','Koreksi Hasil'],['holiday','Ubah Tanggal Merah'],['generate','Buat Dokumen']] },
      { id:'history',label:'Riwayat Lembur',actions:[['download','Unduh Hasil'],['edit','Edit Riwayat'],['delete','Hapus Riwayat']] }
    ]}
  ];
  const MOD = Object.fromEntries(MODULES.map(m => [m.id,m]));
  const PAGE_MODULE = {leave:'cuti',assignment:'surat',training:'pelatihan',tukin:'tukin',lembur:'lembur',uangMakan:'uangMakan'};
  const legacy = Object.fromEntries(['tukin','lembur','uangMakan'].map(id => [id,Object.fromEntries(
    MOD[id].sections.flatMap(s => [`${s.id}.view`,...s.actions.map(a => `${s.id}.${a[0]}`)]).map(key=>[key,true]))]));
  let grants = null;
  let active = false;
  let loaded = false;
  let unsubscribe = null;
  let currentUid = '';
  let loadPromise = null;
  let refreshScheduled = false;
  const escape = v => String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  function normalizeGrants(input) {
    const result = {};
    for (const m of MODULES) {
      const raw = input?.[m.id];
      const allowed = new Set(m.sections.flatMap(s=>[`${s.id}.view`, ...s.actions.map(a=>`${s.id}.${a[0]}`)]));
      result[m.id] = {};
      if (raw && typeof raw === 'object') {
        for (const [key,value] of Object.entries(raw)) {
          if (allowed.has(key) && value === true) result[m.id][key] = true;
        }
      }
      for (const s of m.sections) {
        if (s.parent && !result[m.id][`${s.parent}.view`]) {
          for (const key of Object.keys(result[m.id])) if (key.startsWith(`${s.id}.`)) delete result[m.id][key];
        }
        if (!result[m.id][`${s.id}.view`]) {
          for (const key of Object.keys(result[m.id])) if (key.startsWith(`${s.id}.`) && key !== `${s.id}.view`) delete result[m.id][key];
        }
      }
    }
    return result;
  }

  function defaultGrants() { return normalizeGrants(legacy); }
  function ownIsAdmin() { return window.FirebaseClient?.isAdmin?.() === true; }
  function allowed(module, key) {
    if (!loaded && !ownIsAdmin()) return false;
    if (ownIsAdmin()) return true;
    if (!active) return false;
    const m = MOD[module];
    const sectionId = String(key || '').split('.')[0];
    const section = m?.sections.find(s=>s.id===sectionId);
    if (!section) return false;
    if (!grants?.[module]?.[key]) return false;
    if (key !== `${sectionId}.view` && !grants[module][`${sectionId}.view`]) return false;
    if (section.parent && !grants[module][`${section.parent}.view`]) return false;
    return true;
  }
  function can(module,key) { return allowed(module,key); }
  function canView(module,view) {
    const normalized=module==='pelatihan'&&view==='pelatihan'?'data':view;
    const section=MOD[module]?.sections.find(s=>s.id===normalized);
    return Boolean(section && can(module,`${normalized}.view`) && (!section.parent || can(module,`${section.parent}.view`)));
  }
  function someView(module) { return Boolean(MOD[module]?.sections.some(s=>canView(module,s.id))); }
  function firstView(module) { const v=MOD[module]?.sections.find(s=>!s.parent&&canView(module,s.id))?.id || ''; return module==='pelatihan'&&v==='data'?'pelatihan':v; }
  function loadedAccess() { return loaded || ownIsAdmin(); }
  function deny(message) {
    if (message !== false) alert('Anda tidak memiliki izin untuk tindakan ini. Hubungi administrator untuk meminta akses.');
    return false;
  }
  function assert(module,key,message) { return can(module,key) || deny(message); }
  function permissionDeniedMarkup(module) {
    const destination=homeHref();
    const action=destination==='index.html'
      ? '<button type="button" class="btn btn-secondary" onclick="window.FirebaseClient.signOut()">Keluar dari akun</button>'
      : `<a class="btn btn-secondary" href="${destination}">Buka menu yang tersedia</a>`;
    return `<div class="card permission-denied"><div class="permission-denied-icon">&#128274;</div><h2>Akses tidak tersedia</h2><p>Akun ini belum memperoleh izin untuk membuka ${escape(MOD[module]?.label || 'halaman')} atau fitur yang dipilih. Hubungi administrator aplikasi.</p>${action}</div>`;
  }
  function homeHref() {
    const links = {tukin:'tukin.html',lembur:'lembur.html',uangMakan:'uang-makan.html',pelatihan:'pelatihan.html',cuti:'cuti.html',surat:'surat-tugas.html'};
    for (const mod of MODULES) if (firstView(mod.id)) return `${links[mod.id]}#${firstView(mod.id)}`;
    return 'index.html';
  }
  function emit() {
    window.dispatchEvent(new CustomEvent('permissionschanged'));
    scheduleDecoration();
  }
  function useSnapshot(snapshot) {
    const previous = loaded ? JSON.stringify({active,grants}) : null;
    const data = snapshot?.exists ? snapshot.data() : null;
    grants = data ? normalizeGrants(data.grants || {}) : defaultGrants();
    active = data ? data.active !== false : true;
    loaded = true;
    const changed = previous !== null && previous !== JSON.stringify({active,grants});
    if (changed && document.readyState === 'complete') { location.reload(); return; }
    emit();
  }
  async function load(force=false) {
    const client = window.FirebaseClient;
    const user = client?.getCurrentUser?.();
    if (!user) throw new Error('Masuk terlebih dahulu.');
    if (currentUid === user.uid && loaded && !force) return;
    if (loadPromise && currentUid === user.uid && !force) return loadPromise;
    unsubscribe?.(); unsubscribe = null;
    currentUid = user.uid; loaded = false; active = false; grants = null;
    if (ownIsAdmin()) { grants = normalizeGrants({}); active = true; loaded = true; emit(); return; }
    const ref = client.getDb().collection('userPermissions').doc(user.uid);
    loadPromise = (async () => {
      try {
        const snap = await ref.get({source:'server'});
        useSnapshot(snap);
        unsubscribe = ref.onSnapshot(useSnapshot, (error) => {
          console.warn('Sinkronisasi hak akses gagal:',error);
          // Biarkan status server yang sudah dibaca, tetapi aturan Firestore tetap memverifikasi setiap operasi.
        });
      } catch (error) {
        console.error('Gagal memverifikasi hak akses', error);
        loaded = true; active = false; grants = {}; emit();
        throw new Error('Hak akses gagal diverifikasi. Periksa koneksi Firestore dan deploy Security Rules terbaru.');
      }
    })();
    try { await loadPromise; } finally { loadPromise=null; }
  }
  async function requireModule(module) {
    await window.FirebaseClient.requireAuth();
    try { await load(); } catch (error) { document.getElementById('app').innerHTML = permissionDeniedMarkup(module) + `<p class="permission-error">${escape(error.message || error)}</p>`; return false; }
    if (!someView(module)) {
      document.getElementById('app').innerHTML = permissionDeniedMarkup(module);
      return false;
    }
    return true;
  }
  function grantedCount(module) {
    const m=MOD[module]; if(!m) return 0;
    return m.sections.reduce((sum,s)=>sum+(can(module,`${s.id}.view`)?1:0)+s.actions.filter(a=>can(module,`${s.id}.${a[0]}`)).length,0);
  }
  function totalCount(module) {return MOD[module]?.sections.reduce((sum,s)=>sum+1+s.actions.length,0)||0;}

  // Tombol yang tidak diizinkan disembunyikan, dan event dalam capture phase ditolak.
  const MATCH = {
    cuti: [
      ['[data-action="add-leave"]','data.create'],['[data-action="save-leave"]',node=>node.closest('.leave-modal')?.querySelector('h3')?.textContent.includes('Edit')?'data.edit':'data.create'],
      ['[data-leave-edit]','data.edit'],['[data-leave-delete]','data.delete'],
      ['[data-balance-edit], [data-action="save-balance"]','saldo.edit_balance'],
      ['[data-action="print-report"]','laporan.print'],
      ['[data-add-leave-document]','documents.create'],['[data-edit-leave-document]','documents.edit'],['[data-delete-leave-document]','documents.delete'],['[data-download-leave-document]','documents.download'],
      ['[data-action="save-leave-document"]',node=>node.closest('.leave-modal')?.querySelector('h3')?.textContent.includes('Edit')?'documents.edit':'documents.create']
    ],
    surat:[
      ['[data-action="add-assignment"]','data.create'],['#assignment-file','documents.upload'],['[data-assignment-edit]','data.edit'],['[data-assignment-delete]','data.delete'],
      ['[data-action="save-assignment"]',node=>node.closest('.assignment-modal')?.querySelector('h3')?.textContent.includes('Edit')?'data.edit':'data.create'],
      ['[data-download-assignment-document]','documents.download'],
      ['[data-action="assignment-file"], [data-assignment-file]','documents.upload']
    ],
    pelatihan:[
      ['[data-action="add-training"]','data.create'],['[data-training-edit]','data.edit'],['[data-training-delete]','data.delete'],
      ['[data-action="save-training"]',node=>node.closest('.training-modal')?.querySelector('h3')?.textContent.includes('Edit')?'data.edit':'data.create'],
      ['[data-add-training-document]','documents.create'],['[data-edit-training-document]','documents.edit'],['[data-delete-training-document]','documents.delete'],
      ['[data-action="save-training-document"]',node=>node.closest('.training-modal')?.querySelector('h3')?.textContent.includes('Edit')?'documents.edit':'documents.create'],
      ['[data-participant-status-select], [data-participant-note], [data-participants-edit],[data-action="save-training-participants"]','participants.edit']
    ],
    tukin:[
      ['[data-action="add-holiday"],[data-remove-holiday]','parameters.holidays'],
      ['#ramadan-enabled,#ramadan-start,#ramadan-end','parameters.ramadan'],
      ['[data-action="choose-files"],#tukin-files,[data-remove-file]','process.upload'],
      ['[data-action="calculate"]','process.calculate'],
      ['[data-edit-record],[data-action="save-edit"],[data-action="reset-auto"]','process.verify'],
      ['[data-edit-employee],[data-action="save-employee"]','process.verify'],
      ['#edit-evidence,[data-remove-evidence]','process.evidence'],
      ['[data-action="save-run"]',()=>can('tukin','process.save')?'process.save':'history.edit'],['[data-action="new-process"],[data-action="new-process-from-history"]','process.view'],
      ['[data-edit-tukin-history]','history.edit'],['[data-delete-tukin-history]','history.delete'],['[data-download-tukin-history]','history.download']
    ],
    uangMakan:[
      ['[data-um-action="upload"],#um-files','process.upload'],
      ['[data-um-action="calculate"]','process.calculate'],
      ['[data-um-action="verify"],[data-um-action="save-day"],[data-um-action="acknowledge"]','process.verify'],
      ['[data-um-action="save"]','process.save'],
      ['[data-um-action="export"],[data-um-action="export-zip"]','process.export'],
      ['[data-um-action="history-download"],[data-um-action="history-download-zip"]','history.download'],
      ['[data-um-action="history-delete"]','history.delete']
    ],
    lembur:[
      ['[data-action="start-process"],[data-action="new-period"]','process.view'],
      ['[data-action="add-holiday"],[data-remove-holiday],[data-action="confirm-holiday-preview"]',()=>document.querySelector('.lm-editor')?.dataset.historyEdit==='true'?'history.edit':'process.holiday'],
      ['[data-action="choose-files"],#file-input,[data-remove-file]','process.upload'],
      ['[data-action="validate"],[data-action="process-data"]','process.calculate'],
      ['[data-review-edit],[data-action="save-review-edit"],[data-overtime-key]',()=>document.querySelector('.lm-editor')?.dataset.historyEdit==='true'?'history.edit':'process.correct'],
      ['[data-action="generate"]',()=>can('lembur','process.generate')?'process.generate':'history.edit'],
      ['[data-action="save-history-edit"],[data-edit-history-id],[data-action="edit-current-run"]','history.edit'],
      ['[data-delete-history-id]','history.delete'],
      ['[data-download]','history.download']
    ]
  };
  function requiredForElement(node, module) {
    for (const [selector,key] of MATCH[module] || []) {
      if (node.matches?.(selector)) return typeof key === 'function'?key(node):key;
    }
    return '';
  }
  function decorate() {
    const el=document.getElementById('app'); if (!el || !loadedAccess()) return;
    const page=PAGE_MODULE[document.body.dataset.permissionModule] || document.body.dataset.permissionModule;
    if(!page) return;
    if (page === 'cuti' && !can('cuti','documents.view')) el.querySelectorAll('.leave-document-head,.leave-document-list-preview').forEach(node => node.classList.add('permission-hidden'));
    if (page === 'pelatihan' && !can('pelatihan','documents.view')) el.querySelectorAll('.training-document-head,.training-document-row').forEach(node=>node.classList.add('permission-hidden'));
    if (page === 'surat' && !can('surat','documents.view')) el.querySelectorAll('.assignment-document-card,.assignment-file-preview').forEach(node=>node.classList.add('permission-hidden'));
    for (const [selector,key] of MATCH[page] || []) {
      el.querySelectorAll(selector).forEach(node => {
        const requested=typeof key==='function'?key(node):key;
        node.classList.toggle('permission-hidden',!can(page,requested));
        if ('disabled' in node) node.disabled = !can(page,requested);
      });
    }
  }
  function scheduleDecoration() {
    if (refreshScheduled) return;
    refreshScheduled=true;
    queueMicrotask(()=>{refreshScheduled=false;decorate();});
  }
  const observer=new MutationObserver(scheduleDecoration);
  function startGuard() {
    if (document.getElementById('app')) observer.observe(document.getElementById('app'),{subtree:true,childList:true});
    document.addEventListener('click', event => {
      const page=PAGE_MODULE[document.body.dataset.permissionModule] || document.body.dataset.permissionModule;
      const node=event.target.closest?.('[data-action],button,[data-leave-edit],[data-leave-delete],[data-balance-edit],[data-add-leave-document],[data-edit-leave-document],[data-delete-leave-document],[data-download-leave-document],[data-assignment-edit],[data-assignment-delete],[data-download-assignment-document],[data-training-edit],[data-training-delete],[data-add-training-document],[data-edit-training-document],[data-delete-training-document],[data-edit-record],[data-edit-employee],[data-edit-tukin-history],[data-delete-tukin-history],[data-download-tukin-history],[data-edit-history-id],[data-delete-history-id],[data-download],[data-um-action]');
      const key=page&&node?requiredForElement(node,page):'';
      if (key&&!can(page,key)) {event.preventDefault();event.stopImmediatePropagation();deny();}
    },true);
    document.addEventListener('change',event=>{
      const page=PAGE_MODULE[document.body.dataset.permissionModule] || document.body.dataset.permissionModule;
      const node=event.target;
      const key=page?requiredForElement(node,page):'';
      if (key&&!can(page,key)) {event.preventDefault();event.stopImmediatePropagation();node.checked=false;deny();}
    },true);
    document.addEventListener('click',event=>{
      const anchor=event.target.closest?.('a[href]');
      if(!anchor) return;
      const path=anchor.getAttribute('href')||'';
      const match=path.match(/^(cuti|surat-tugas|pelatihan|tukin|lembur|uang-makan)\.html(?:#([a-z]+))?$/i);
      if (!match) return;
      const module={cuti:'cuti','surat-tugas':'surat',pelatihan:'pelatihan',tukin:'tukin',lembur:'lembur','uang-makan':'uangMakan'}[match[1].toLowerCase()];
      if(!canView(module,match[2]||firstView(module))) {event.preventDefault();event.stopImmediatePropagation();deny();}
    },true);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',startGuard); else startGuard();
  window.UserAccess=Object.freeze({MODULES,MOD,normalizeGrants,defaultGrants,can,canView,someView,firstView,loadedAccess,load,requireModule,assert,deny,permissionDeniedMarkup,homeHref,grantedCount,totalCount,decorate,scheduleDecoration});
})();
