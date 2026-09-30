(function () {
  'use strict';

  const app = document.getElementById('app');
  const state = {
    tab: 'master',
    loading: true,
    busy: false,
    error: '',
    toast: '',
    search: '',
    employees: [],
    directory: new Map(),
    settings: { lembur: {}, tukin: {} },
    audits: [],
    trainings: [],
    leaveRecords: [],
    leaveBalances: [],
    modal: null,
    importOpen: false,
    importFile: null
  };

  const esc = (value) => String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');

  const money = (value) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(Number(value || 0));

  function normalizeAnakSatker(value) {
    const code = String(value == null ? '' : value).trim();
    // Pertahankan kompatibilitas kode numerik lama (mis. 1 -> 01),
    // tetapi jangan menghapus huruf pada kode alfanumerik seperti P1/P2/P3.
    return /^\d$/.test(code) ? code.padStart(2, '0') : code;
  }

  function db() {
    return window.FirebaseClient.getDb();
  }

  function user() {
    return window.FirebaseClient.getCurrentUser();
  }

  function serverTimestamp() {
    return firebase.firestore.FieldValue.serverTimestamp();
  }

  function dateTime(value) {
    try {
      const d = value?.toDate ? value.toDate() : new Date(value || 0);
      if (!d || Number.isNaN(d.getTime())) return '-';
      return new Intl.DateTimeFormat('id-ID', { dateStyle: 'medium', timeStyle: 'short' }).format(d);
    } catch (_) {
      return '-';
    }
  }

  function docLabel(item) {
    const privateData = state.directory.get(item.id) || {};
    const name = String(privateData.name || '').trim();
    const nip = String(privateData.nip || '').replace(/\D/g, '');
    if (name) return { title: name, sub: nip ? `NIP ${nip}` : 'NIP belum tersedia', resolved: true };
    return { title: 'Nama pegawai belum tersedia', sub: 'Lengkapi identitas tampilan pada master pegawai', resolved: false };
  }

  function auditDisplay(item) {
    const target = String(item?.target || '');
    if (state.directory.has(target)) return docLabel({ id: target }).title;
    const trainingId = target.replace(/^trainings\//, '');
    const training = state.trainings.find((row) => row.id === trainingId);
    if (training) return String(training.name || 'Data pelatihan');
    if (target === 'appSettings/lembur') return 'Pengaturan Lembur';
    if (target === 'appSettings/tukin') return 'Pengaturan Tunjangan Kinerja';
    const detail = String(item?.detail || '').trim();
    const looksLikeHash = /^[a-f0-9]{32,}$/i.test(detail) || /[a-f0-9]{48,}/i.test(detail);
    return detail && !looksLikeHash ? detail : 'Perubahan administrasi';
  }

  async function audit(action, target, detail) {
    const ref = db().collection('adminAudit').doc();
    await ref.set({
      action: String(action || ''),
      target: String(target || ''),
      detail: String(detail || ''),
      actorUid: user()?.uid || '',
      actorEmail: user()?.email || '',
      createdAt: serverTimestamp()
    });
  }

  function showToast(message) {
    state.toast = String(message || '');
    render();
    clearTimeout(window.__adminToastTimer);
    window.__adminToastTimer = setTimeout(() => {
      state.toast = '';
      render();
    }, 3500);
  }

  async function loadData() {
    state.loading = true;
    state.error = '';
    render();
    try {
      const [masterSnap, directorySnap, lemburSnap, tukinSnap, auditSnap, trainingSnap, leaveSnap, balanceSnap] = await Promise.all([
        db().collection('masterEmployees').get(),
        db().collection('masterDirectory').get(),
        db().collection('appSettings').doc('lembur').get(),
        db().collection('appSettings').doc('tukin').get(),
        db().collection('adminAudit').orderBy('createdAt', 'desc').limit(20).get(),
        db().collection('trainings').get(),
        db().collection('leaveRecords').get(),
        db().collection('leaveBalances').get()
      ]);
      state.employees = masterSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.directory = new Map(directorySnap.docs.map((doc) => [doc.id, { id: doc.id, ...doc.data() }]));
      state.settings = {
        lembur: lemburSnap.exists ? lemburSnap.data() : {},
        tukin: tukinSnap.exists ? tukinSnap.data() : {}
      };
      state.audits = auditSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.trainings = trainingSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.leaveRecords = leaveSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.leaveBalances = balanceSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.employees.sort((a, b) => Number(a.order || 9999) - Number(b.order || 9999) || docLabel(a).title.localeCompare(docLabel(b).title, 'id'));
    } catch (error) {
      state.error = error.message || String(error);
    } finally {
      state.loading = false;
      window.MasterDataService?.clearCache?.();
      render();
    }
  }

  function metric(label, value, note) {
    return `<div class="card admin-metric"><div class="admin-metric-label">${esc(label)}</div><div class="admin-metric-value">${esc(value)}</div><div class="card-subtitle">${esc(note || '')}</div></div>`;
  }

  function filteredEmployees() {
    const q = state.search.trim().toLowerCase();
    if (!q) return state.employees;
    return state.employees.filter((item) => {
      const label = docLabel(item);
      const fields = [
        label.title, label.sub,
        item.lemburSatkerCode, item.satker, item.anakSatker, item.tukin, item.order
      ].map((x) => String(x || '').toLowerCase());
      return fields.some((x) => x.includes(q));
    });
  }

  function masterTable() {
    const rows = filteredEmployees();
    return `<div class="card">
      <div class="admin-card-head"><div class="admin-card-head-copy"><h2>Master Pegawai</h2><p>Satu sumber master digunakan bersama oleh modul Lembur dan Tunjangan Kinerja.</p></div><button class="btn btn-primary btn-sm" type="button" data-action="add-master">+ Tambah Master</button></div>
      <div class="admin-toolbar"><div class="search"><input id="admin-search" placeholder="Cari nama, NIP, satker, anak satker..." value="${esc(state.search)}"></div><span class="card-subtitle">${rows.length} dari ${state.employees.length} data</span></div>
      <div class="admin-table-wrap"><table class="data-table admin-table"><thead><tr><th>Identitas</th><th>Satker Lembur</th><th>Satker Tukin</th><th>Anak Satker</th><th>Besaran Tukin</th><th>Urutan</th><th>Status</th><th>Aksi</th></tr></thead><tbody>
        ${rows.length ? rows.map((item) => {
          const label = docLabel(item);
          const unresolved = !label.resolved;
          return `<tr>
            <td class="admin-identity"><strong>${esc(label.title)}</strong><span>${esc(label.sub)}</span>${unresolved ? '<span class="admin-badge legacy" style="margin-top:6px">Identitas belum lengkap</span>' : ''}</td>
            <td>${esc(item.lemburSatkerCode ?? '-')}</td>
            <td class="admin-code">${esc(item.satker || '-')}</td>
            <td>${esc(item.anakSatker || '-')}</td>
            <td class="admin-money">${money(item.tukin)}</td>
            <td>${esc(item.order ?? '-')}</td>
            <td><span class="admin-badge ${item.active === false ? 'off' : ''}">${item.active === false ? 'Nonaktif' : 'Aktif'}</span></td>
            <td><div class="admin-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-edit-master="${esc(item.id)}">Edit</button><button class="btn ${item.active === false ? 'btn-secondary' : 'btn-danger'} btn-sm" type="button" data-toggle-master="${esc(item.id)}">${item.active === false ? 'Aktifkan' : 'Nonaktifkan'}</button></div></td>
          </tr>`;
        }).join('') : '<tr><td colspan="8"><div class="empty-state"><strong>Master belum tersedia</strong>Tambahkan manual atau impor file master awal.</div></td></tr>'}
      </tbody></table></div>
    </div>`;
  }

  const UNIT_OPTIONS = [
    'Pemeliharaan Infrastruktur Teknologi Informasi',
    'Manajemen Data dan Pengembangan Sistem Informasi',
    'Umum dan Tata Usaha'
  ];

  function field(label, id, value, opts) {
    const o = opts || {};
    const cls = o.span2 ? 'field span-2' : 'field';
    const type = o.type || 'text';
    const attrs = [o.min != null ? `min="${o.min}"` : '', o.step != null ? `step="${o.step}"` : '', o.maxlength ? `maxlength="${o.maxlength}"` : ''].filter(Boolean).join(' ');
    return `<div class="${cls}"><label for="${id}">${esc(label)}</label><input id="${id}" type="${type}" value="${esc(value ?? '')}" ${attrs}></div>`;
  }

  function unitField(value) {
    const selected = String(value || '').trim();
    return `<div class="field span-2"><label for="m-unit">Unit/Bidang/Bagian</label><select id="m-unit" required><option value="">Pilih Unit/Bidang/Bagian</option>${UNIT_OPTIONS.map((unit) => `<option value="${esc(unit)}" ${selected === unit ? 'selected' : ''}>${esc(unit)}</option>`).join('')}</select></div>`;
  }

  function settingsView() {
    const l = state.settings.lembur || {};
    const t = state.settings.tukin || {};
    return `<div class="admin-settings-grid">
      <section class="card admin-settings-card"><h2>Pengaturan Lembur</h2><p>Data penandatangan dan identitas unit yang digunakan pada dokumen keluaran.</p>
        <div class="admin-form">
          ${field('Unit Organisasi', 'set-l-org', l.ORGANIZATION_UNIT || window.APP_CONFIG.ORGANIZATION_UNIT, { span2: true })}
          ${field('Unit Kerja', 'set-l-work', l.WORK_UNIT || window.APP_CONFIG.WORK_UNIT)}
          ${field('Kota', 'set-l-city', l.CITY || window.APP_CONFIG.CITY)}
          ${field('Judul Penanggung Jawab', 'set-l-title1', l.RESPONSIBLE_TITLE_1 || window.APP_CONFIG.RESPONSIBLE_TITLE_1, { span2: true })}
          ${field('Jabatan Penanggung Jawab', 'set-l-title2', l.RESPONSIBLE_TITLE_2 || window.APP_CONFIG.RESPONSIBLE_TITLE_2, { span2: true })}
          ${field('Nama Penanggung Jawab', 'set-l-name', l.RESPONSIBLE_NAME || '', { span2: true })}
          ${field('NIP Penanggung Jawab', 'set-l-nip', l.RESPONSIBLE_NIP || '', { span2: true, maxlength: 30 })}
          <div class="field span-2"><label for="set-l-address">Alamat SPKL</label><textarea id="set-l-address">${esc(l.SPKL_ADDRESS || '')}</textarea></div>
        </div>
        <div class="admin-settings-actions"><button class="btn btn-primary" type="button" data-action="save-lembur-settings">Simpan Pengaturan Lembur</button></div>
      </section>
      <section class="card admin-settings-card"><h2>Pengaturan Tunjangan Kinerja</h2><p>Parameter identitas satker dan nilai default yang dipakai ketika data master pegawai belum cocok.</p>
        <div class="admin-form">
          ${field('Kode Satker', 'set-t-satker', t.SATKER || '')}
          ${field('Tukin Default', 'set-t-default', t.DEFAULT_TUKIN || 0, { type: 'number', min: 0, step: 1000 })}
          ${field('Unit Kerja', 'set-t-work', t.UNIT_WORK || window.TUKIN_CONFIG.UNIT_WORK, { span2: true })}
          ${field('Unit Organisasi', 'set-t-org', t.UNIT_ORGANIZATION || window.TUKIN_CONFIG.UNIT_ORGANIZATION, { span2: true })}
          ${field('Label Unit', 'set-t-label', t.UNIT_LABEL || window.TUKIN_CONFIG.UNIT_LABEL, { span2: true })}
          ${field('Nilai SKP', 'set-t-skp', t.SKP_SCORE ?? window.TUKIN_CONFIG.SKP_SCORE, { type: 'number', min: 0 })}
          ${field('Potongan SKP Default (%)', 'set-t-skp-deduction', t.DEFAULT_SKP_DEDUCTION ?? window.TUKIN_CONFIG.DEFAULT_SKP_DEDUCTION, { type: 'number', min: 0, step: 0.01 })}
        </div>
        <div class="admin-settings-actions"><button class="btn btn-primary" type="button" data-action="save-tukin-settings">Simpan Pengaturan Tukin</button></div>
      </section>
      <section class="card admin-settings-card admin-audit" style="grid-column:1/-1"><h2>Aktivitas Administrasi Terakhir</h2><p>Catatan perubahan master dan pengaturan yang dilakukan melalui halaman ini.</p>
        ${state.audits.length ? state.audits.map((item) => `<div class="admin-audit-row"><span>${esc(dateTime(item.createdAt))}</span><span><strong>${esc(item.action || '-')}</strong><br>${esc(auditDisplay(item))}</span><span>${esc(item.actorEmail || '-')}</span></div>`).join('') : '<div class="empty-state"><strong>Belum ada aktivitas</strong>Log akan muncul setelah admin melakukan perubahan.</div>'}
      </section>
    </div>`;
  }

  function importPanel() {
    return `<div class="admin-import-box ${state.importOpen ? 'open' : ''}"><div class="admin-import-row"><input id="seed-file" type="file" accept="application/json,.json"><button class="btn btn-primary btn-sm" type="button" data-action="run-import" ${state.importFile ? '' : 'disabled'}>Impor ke Firestore</button><span class="card-subtitle">${state.importFile ? `Dipilih: <strong>${esc(state.importFile.name)}</strong>. ` : ''}Pilih <strong>master-data-seed.json</strong> dari folder <code>firebase-setup-private</code>. File tersebut jangan dipublikasikan.</span></div></div>`;
  }

  function masterModal() {
    if (!state.modal) return '';
    const existing = state.modal.id ? state.employees.find((x) => x.id === state.modal.id) : null;
    const privateData = existing ? (state.directory.get(existing.id) || {}) : {};
    const title = existing ? 'Edit Master Pegawai' : 'Tambah Master Pegawai';
    const help = existing && !state.directory.has(existing.id)
      ? 'Lengkapi Nama dan NIP agar identitas pegawai dapat ditampilkan dengan benar pada seluruh menu aplikasi.'
      : 'Nama, NIP, dan Unit/Bidang/Bagian pada bagian ini digunakan sebagai identitas tampilan pada seluruh menu aplikasi.';
    return `<div class="admin-modal-backdrop" data-action="close-modal"><div class="admin-modal" role="dialog" aria-modal="true" onclick="event.stopPropagation()">
      <div class="admin-modal-head"><div><h3>${title}</h3><p>${esc(help)}</p></div><button class="icon-btn" type="button" data-action="close-modal">×</button></div>
      <div class="admin-form">
        ${field('Nama Pegawai', 'm-name', privateData.name || '', { span2: true })}
        ${field('NIP', 'm-nip', privateData.nip || '', { span2: true, maxlength: 30 })}
        ${unitField(privateData.unit || '')}
        ${field('Kode Satker Lembur', 'm-lembur-code', existing?.lemburSatkerCode ?? '', { maxlength: 20 })}
        ${field('Kode Satker Tukin', 'm-satker', existing?.satker || window.TUKIN_CONFIG.SATKER || '')}
        ${field('Anak Satker', 'm-anak', existing?.anakSatker || '', { maxlength: 20 })}
        ${field('Besaran Tukin', 'm-tukin', existing?.tukin || window.TUKIN_CONFIG.DEFAULT_TUKIN || 0, { type: 'number', min: 0, step: 1000 })}
        ${field('Urutan', 'm-order', existing?.order || state.employees.length + 1, { type: 'number', min: 1 })}
        <div class="field"><label>Status</label><label class="admin-check"><input id="m-active" type="checkbox" ${existing?.active === false ? '' : 'checked'}><span>Aktif digunakan aplikasi</span></label></div>
      </div>
      ${existing ? `<div class="admin-private-note">Identitas yang ditampilkan pada aplikasi: <strong>${esc(privateData.name || 'Nama belum tersedia')}</strong>${privateData.nip ? `<br>NIP ${esc(privateData.nip)}` : ''}</div>` : '<div class="admin-private-note">Nama dan NIP akan digunakan sebagai identitas tampilan pegawai pada seluruh menu.</div>'}
      <div class="actions"><span></span><div class="actions-right"><button class="btn btn-secondary" type="button" data-action="close-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-master" ${state.busy ? 'disabled' : ''}>Simpan</button></div></div>
    </div></div>`;
  }

  function render() {
    if (!window.FirebaseClient?.getCurrentUser?.()) return;
    const active = state.employees.filter((x) => x.active !== false).length;
    const labeled = state.employees.filter((x) => String(state.directory.get(x.id)?.name || '').trim()).length;
    const content = `<div class="admin-page">
      <div class="admin-head"><div><h1>Administrasi Master Data</h1><p>Kelola master pegawai dan pengaturan operasional yang digunakan oleh proses Tunjangan Kinerja dan Lembur. Perubahan disimpan langsung di Cloud Firestore.</p></div><div class="admin-head-actions"><button class="btn btn-secondary" type="button" data-action="toggle-import">Impor Master Awal</button><button class="btn btn-secondary" type="button" data-action="export-backup">Ekspor Backup</button></div></div>
      ${state.error ? `<div class="alert alert-danger"><div class="alert-title">Data administrasi gagal dimuat</div>${esc(state.error)}</div>` : ''}
      ${importPanel()}
      <div class="admin-metrics">
        ${metric('Total Master', state.employees.length, 'Seluruh dokumen master')}
        ${metric('Master Aktif', active, 'Digunakan untuk pencocokan')}
        ${metric('Identitas Terpetakan', labeled, 'Memiliki Nama/NIP untuk tampilan')}
        ${metric('Pengaturan', (state.settings.lembur && Object.keys(state.settings.lembur).length ? 1 : 0) + (state.settings.tukin && Object.keys(state.settings.tukin).length ? 1 : 0), 'Dari 2 modul')}
      </div>
      <div class="admin-tabs"><button class="admin-tab ${state.tab === 'master' ? 'active' : ''}" data-tab="master" type="button">Master Pegawai</button><button class="admin-tab ${state.tab === 'settings' ? 'active' : ''}" data-tab="settings" type="button">Pengaturan & Audit</button></div>
      ${state.loading ? '<div class="card admin-loading">Memuat data Firestore...</div>' : (state.tab === 'master' ? masterTable() : settingsView())}
    </div>`;
    app.innerHTML = window.AppShell.render({ module: 'admin', view: state.tab, viewLabel: state.tab === 'master' ? 'Master Pegawai' : 'Pengaturan', content, overlays: masterModal() }) + (state.toast ? `<div class="admin-toast">${esc(state.toast)}</div>` : '');
    bindEvents();
  }

  function bindEvents() {
    document.querySelectorAll('[data-tab]').forEach((button) => button.addEventListener('click', () => { state.tab = button.dataset.tab; render(); }));
    document.querySelector('[data-action="toggle-import"]')?.addEventListener('click', () => { state.importOpen = !state.importOpen; render(); });
    document.getElementById('seed-file')?.addEventListener('change', (event) => { state.importFile = event.target.files?.[0] || null; render(); });
    document.querySelector('[data-action="run-import"]')?.addEventListener('click', importSeed);
    document.querySelector('[data-action="export-backup"]')?.addEventListener('click', exportBackup);
    document.querySelector('[data-action="add-master"]')?.addEventListener('click', () => { state.modal = { id: null }; render(); });
    document.querySelectorAll('[data-edit-master]').forEach((button) => button.addEventListener('click', () => { state.modal = { id: button.dataset.editMaster }; render(); }));
    document.querySelectorAll('[data-toggle-master]').forEach((button) => button.addEventListener('click', () => toggleMaster(button.dataset.toggleMaster)));
    document.querySelectorAll('[data-action="close-modal"]').forEach((button) => button.addEventListener('click', () => { state.modal = null; render(); }));
    document.querySelector('[data-action="save-master"]')?.addEventListener('click', saveMaster);
    document.getElementById('admin-search')?.addEventListener('input', (event) => {
      state.search = event.target.value;
      clearTimeout(window.__adminSearchTimer);
      window.__adminSearchTimer = setTimeout(render, 140);
    });
    document.querySelector('[data-action="save-lembur-settings"]')?.addEventListener('click', saveLemburSettings);
    document.querySelector('[data-action="save-tukin-settings"]')?.addEventListener('click', saveTukinSettings);
  }

  async function saveMaster() {
    if (state.busy) return;
    state.busy = true;
    try {
      const existing = state.modal?.id ? state.employees.find((x) => x.id === state.modal.id) : null;
      const name = String(document.getElementById('m-name')?.value || '').trim();
      const nip = String(document.getElementById('m-nip')?.value || '').replace(/\D/g, '');
      const unit = String(document.getElementById('m-unit')?.value || '').trim();
      if (!UNIT_OPTIONS.includes(unit)) {
        throw new Error('Pilih Unit/Bidang/Bagian dari daftar yang tersedia.');
      }
      const hashes = await window.MasterDataService.hashIdentity({ name, nip });
      let id = existing?.id || hashes.nipHash || hashes.nameHash;
      if (!id) throw new Error('Isi minimal Nama Pegawai atau NIP.');

      if (existing?.nipHash && nip && hashes.nipHash !== existing.nipHash) {
        throw new Error('NIP tidak sesuai dengan identitas master yang tersimpan. Periksa kembali NIP pegawai.');
      }

      const existingNameHashes = window.MasterDataService.nameHashes(existing || {});
      if (hashes.nameHash) existingNameHashes.push(hashes.nameHash);
      const operational = {
        nipHash: existing?.nipHash || hashes.nipHash || '',
        nameHashes: [...new Set(existingNameHashes.filter(Boolean))],
        lemburSatkerCode: String(document.getElementById('m-lembur-code')?.value || '').trim(),
        satker: String(document.getElementById('m-satker')?.value || '').trim(),
        anakSatker: normalizeAnakSatker(document.getElementById('m-anak')?.value),
        tukin: Number(document.getElementById('m-tukin')?.value || 0),
        order: Number(document.getElementById('m-order')?.value || 9999),
        active: Boolean(document.getElementById('m-active')?.checked),
        updatedAt: serverTimestamp(),
        updatedBy: user()?.email || '',
        schemaVersion: 4
      };
      if (!operational.lemburSatkerCode || !operational.satker || !operational.anakSatker || !Number.isFinite(operational.tukin) || operational.tukin < 0) {
        throw new Error('Kode Satker Lembur, Satker Tukin, Anak Satker, dan Besaran Tukin harus valid.');
      }
      if (!Number.isFinite(operational.order) || operational.order < 1) operational.order = 9999;

      const batch = db().batch();
      batch.set(db().collection('masterEmployees').doc(id), operational, { merge: true });
      if (name || nip) {
        batch.set(db().collection('masterDirectory').doc(id), {
          name,
          nip,
          unit,
          updatedAt: serverTimestamp(),
          updatedBy: user()?.email || ''
        }, { merge: true });
      }
      const auditRef = db().collection('adminAudit').doc();
      batch.set(auditRef, {
        action: existing ? 'UPDATE_MASTER' : 'CREATE_MASTER',
        target: id,
        detail: `${name || 'Master pegawai'} · Lembur ${operational.lemburSatkerCode} · Anak Satker ${operational.anakSatker}`,
        actorUid: user()?.uid || '', actorEmail: user()?.email || '', createdAt: serverTimestamp()
      });
      await batch.commit();
      state.modal = null;
      window.MasterDataService.clearCache();
      await loadData();
      showToast(existing ? 'Master pegawai berhasil diperbarui.' : 'Master pegawai berhasil ditambahkan.');
    } catch (error) {
      alert(`Master gagal disimpan: ${error.message || error}`);
    } finally {
      state.busy = false;
    }
  }

  async function toggleMaster(id) {
    const item = state.employees.find((x) => x.id === id);
    if (!item) return;
    const next = item.active === false;
    const label = docLabel(item).title;
    const ok = window.confirm(`${next ? 'Aktifkan' : 'Nonaktifkan'} master ${label}?`);
    if (!ok) return;
    try {
      const batch = db().batch();
      batch.update(db().collection('masterEmployees').doc(id), { active: next, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' });
      batch.set(db().collection('adminAudit').doc(), {
        action: next ? 'ACTIVATE_MASTER' : 'DEACTIVATE_MASTER', target: id, detail: label,
        actorUid: user()?.uid || '', actorEmail: user()?.email || '', createdAt: serverTimestamp()
      });
      await batch.commit();
      window.MasterDataService.clearCache();
      await loadData();
      showToast(`Master ${next ? 'diaktifkan' : 'dinonaktifkan'}.`);
    } catch (error) {
      alert(`Status master gagal diubah: ${error.message || error}`);
    }
  }

  function read(id) {
    return String(document.getElementById(id)?.value || '').trim();
  }

  async function saveLemburSettings() {
    const data = {
      ORGANIZATION_UNIT: read('set-l-org'), WORK_UNIT: read('set-l-work'), CITY: read('set-l-city'),
      RESPONSIBLE_TITLE_1: read('set-l-title1'), RESPONSIBLE_TITLE_2: read('set-l-title2'),
      RESPONSIBLE_NAME: read('set-l-name'), RESPONSIBLE_NIP: read('set-l-nip').replace(/\D/g, ''),
      SPKL_ADDRESS: read('set-l-address'), updatedAt: serverTimestamp(), updatedBy: user()?.email || '', schemaVersion: 3
    };
    if (!data.RESPONSIBLE_NAME || !data.RESPONSIBLE_NIP || !data.SPKL_ADDRESS) {
      alert('Nama, NIP penanggung jawab, dan alamat SPKL wajib diisi.');
      return;
    }
    try {
      const batch = db().batch();
      batch.set(db().collection('appSettings').doc('lembur'), data, { merge: true });
      batch.set(db().collection('adminAudit').doc(), { action: 'UPDATE_SETTINGS', target: 'appSettings/lembur', detail: 'Pengaturan modul Lembur diperbarui.', actorUid: user()?.uid || '', actorEmail: user()?.email || '', createdAt: serverTimestamp() });
      await batch.commit();
      await window.AppSettingsService.loadAndApply(true);
      await loadData();
      showToast('Pengaturan Lembur berhasil disimpan.');
    } catch (error) { alert(`Pengaturan gagal disimpan: ${error.message || error}`); }
  }

  async function saveTukinSettings() {
    const data = {
      SATKER: read('set-t-satker'), DEFAULT_TUKIN: Number(read('set-t-default') || 0),
      UNIT_WORK: read('set-t-work'), UNIT_ORGANIZATION: read('set-t-org'), UNIT_LABEL: read('set-t-label'),
      SKP_SCORE: Number(read('set-t-skp') || 0), DEFAULT_SKP_DEDUCTION: Number(read('set-t-skp-deduction') || 0),
      updatedAt: serverTimestamp(), updatedBy: user()?.email || '', schemaVersion: 3
    };
    if (!data.SATKER || !Number.isFinite(data.DEFAULT_TUKIN) || data.DEFAULT_TUKIN < 0) {
      alert('Kode Satker dan Tukin Default harus valid.');
      return;
    }
    try {
      const batch = db().batch();
      batch.set(db().collection('appSettings').doc('tukin'), data, { merge: true });
      batch.set(db().collection('adminAudit').doc(), { action: 'UPDATE_SETTINGS', target: 'appSettings/tukin', detail: 'Pengaturan modul Tunjangan Kinerja diperbarui.', actorUid: user()?.uid || '', actorEmail: user()?.email || '', createdAt: serverTimestamp() });
      await batch.commit();
      await window.AppSettingsService.loadAndApply(true);
      await loadData();
      showToast('Pengaturan Tunjangan Kinerja berhasil disimpan.');
    } catch (error) { alert(`Pengaturan gagal disimpan: ${error.message || error}`); }
  }

  async function importSeed() {
    if (!state.importFile) return;
    const ok = window.confirm('Impor master awal akan menulis/menimpa data dengan ID yang sama di Firestore. Lanjutkan?');
    if (!ok) return;
    try {
      const payload = JSON.parse(await state.importFile.text());
      if (!Array.isArray(payload.masterEmployees)) throw new Error('File tidak memiliki array masterEmployees.');
      const writes = payload.masterEmployees.map((item) => ({ type: 'master', item }));
      if (payload.masterDirectory && typeof payload.masterDirectory === 'object') {
        Object.entries(payload.masterDirectory).forEach(([id, item]) => writes.push({ type: 'directory', id, item }));
      }
      if (Array.isArray(payload.trainings)) payload.trainings.forEach((item) => writes.push({ type: 'training', item }));
      if (Array.isArray(payload.leaveRecords)) payload.leaveRecords.forEach((item) => writes.push({ type: 'leave', item }));
      if (Array.isArray(payload.leaveBalances)) payload.leaveBalances.forEach((item) => writes.push({ type: 'leaveBalance', item }));
      if (payload.appSettings?.lembur) writes.push({ type: 'setting', id: 'lembur', item: payload.appSettings.lembur });
      if (payload.appSettings?.tukin) writes.push({ type: 'setting', id: 'tukin', item: payload.appSettings.tukin });

      for (let i = 0; i < writes.length; i += 400) {
        const batch = db().batch();
        writes.slice(i, i + 400).forEach((entry) => {
          if (entry.type === 'master') {
            const item = entry.item || {};
            const id = String(item.id || item.nipHash || item.nameHashes?.[0] || '').trim();
            if (!id) return;
            const clean = {
              nipHash: String(item.nipHash || ''),
              nameHashes: [...new Set((item.nameHashes || (item.nameHash ? [item.nameHash] : [])).map(String).filter(Boolean))],
              lemburSatkerCode: String(item.lemburSatkerCode ?? item.code ?? '').trim(),
              satker: String(item.satker || ''), anakSatker: normalizeAnakSatker(item.anakSatker), tukin: Number(item.tukin || 0),
              order: Number(item.order || 9999), active: item.active !== false,
              schemaVersion: 4, updatedAt: serverTimestamp(), updatedBy: user()?.email || ''
            };
            batch.set(db().collection('masterEmployees').doc(id), clean, { merge: true });
          } else if (entry.type === 'directory') {
            batch.set(db().collection('masterDirectory').doc(String(entry.id)), {
              name: String(entry.item?.name || ''),
              nip: String(entry.item?.nip || '').replace(/\D/g, ''),
              unit: String(entry.item?.unit || ''),
              updatedAt: serverTimestamp(), updatedBy: user()?.email || ''
            }, { merge: true });
          } else if (entry.type === 'training') {
            const item = entry.item || {};
            const id = String(item.id || '').trim();
            if (!id) return;
            const participants = Array.isArray(item.participants) ? item.participants.map((row) => ({
              employeeId: String(row?.employeeId || ''),
              status: String(row?.status || 'Diusulkan'),
              note: String(row?.note || ''),
              documents: Array.isArray(row?.documents) ? row.documents : []
            })).filter((row) => row.employeeId) : [];
            const participantIds = [...new Set([...(Array.isArray(item.participantIds) ? item.participantIds : []), ...participants.map((row) => row.employeeId)].map(String).filter(Boolean))];
            batch.set(db().collection('trainings').doc(id), {
              name: String(item.name || ''),
              organizer: String(item.organizer || ''),
              startDate: String(item.startDate || ''),
              endDate: String(item.endDate || ''),
              participantIds,
              participants,
              schemaVersion: participants.length ? 2 : 1,
              updatedAt: serverTimestamp(),
              updatedBy: user()?.email || ''
            }, { merge: true });
          } else if (entry.type === 'leave') {
            const item = entry.item || {};
            const id = String(item.id || '').trim();
            if (!id) return;
            const clean = { ...item };
            delete clean.id;
            batch.set(db().collection('leaveRecords').doc(id), { ...clean, schemaVersion: 1, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });
          } else if (entry.type === 'leaveBalance') {
            const item = entry.item || {};
            const id = String(item.id || '').trim();
            if (!id) return;
            const clean = { ...item };
            delete clean.id;
            batch.set(db().collection('leaveBalances').doc(id), { ...clean, schemaVersion: 1, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });
          } else {
            batch.set(db().collection('appSettings').doc(entry.id), { ...entry.item, schemaVersion: 3, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });
          }
        });
        await batch.commit();
      }
      await audit('IMPORT_SEED', 'masterEmployees', `${payload.masterEmployees.length} master pegawai dan pengaturan awal diimpor.`);
      state.importFile = null;
      state.importOpen = false;
      window.MasterDataService.clearCache();
      await window.AppSettingsService.loadAndApply(true);
      await loadData();
      showToast('Master awal berhasil diimpor ke Firestore.');
    } catch (error) {
      alert(`Impor gagal: ${error.message || error}`);
    }
  }

  async function exportBackup() {
    try {
      const payload = {
        exportedAt: new Date().toISOString(),
        schemaVersion: 4,
        masterEmployees: state.employees.map((item) => ({ ...item, updatedAt: undefined })),
        masterDirectory: Object.fromEntries([...state.directory.entries()].map(([id, value]) => [id, { name: value.name || '', nip: value.nip || '', unit: value.unit || '' }])),
        trainings: state.trainings.map((item) => ({ id: item.id, name: item.name || '', organizer: item.organizer || '', startDate: item.startDate || '', endDate: item.endDate || '', participantIds: Array.isArray(item.participantIds) ? item.participantIds : [], participants: Array.isArray(item.participants) ? item.participants : [] })),
        leaveRecords: state.leaveRecords.map((item) => ({ ...item, updatedAt: undefined, createdAt: undefined })),
        leaveBalances: state.leaveBalances.map((item) => ({ ...item, updatedAt: undefined, createdAt: undefined })),
        appSettings: { lembur: state.settings.lembur || {}, tukin: state.settings.tukin || {} }
      };
      const replacer = (key, value) => {
        if (value && typeof value.toDate === 'function') return value.toDate().toISOString();
        return value;
      };
      const blob = new Blob([JSON.stringify(payload, replacer, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `backup-master-pusdatin-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      await audit('EXPORT_BACKUP', 'masterEmployees', 'Admin mengekspor backup master dan pengaturan.');
      showToast('Backup berhasil diunduh. Simpan sebagai file terbatas.');
    } catch (error) { alert(`Backup gagal dibuat: ${error.message || error}`); }
  }

  async function init() {
    await window.FirebaseClient.requireAdmin();
    await window.AppSettingsService.loadAndApply(true);
    await loadData();
  }

  init();
})();
