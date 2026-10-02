(function () {
  'use strict';

  const app = document.getElementById('app');
  const sharePoint = window.SharePointStorage;
  const now = new Date();
  const CURRENT_YEAR = now.getFullYear();
  const CURRENT_MONTH = now.getMonth() + 1;
  const TODAY = toIsoDate(now);
  const VIEW_LABELS = Object.freeze({
    dashboard: 'Dashboard',
    data: 'Data Surat Tugas',
    kalender: 'Kalender'
  });

  const state = {
    loading: true,
    error: '',
    toast: '',
    view: 'dashboard',
    employees: [],
    directory: new Map(),
    assignments: [],
    filters: {
      year: CURRENT_YEAR,
      search: '',
      unit: '',
      status: ''
    },
    calendar: {
      year: CURRENT_YEAR,
      month: CURRENT_MONTH
    },
    modal: null,
    detailId: '',
    previewModal: null,
    inlinePreviews: {}
  };

  const esc = (value) => String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');

  function db() {
    return window.FirebaseClient.getDb();
  }

  function user() {
    return window.FirebaseClient.getCurrentUser();
  }

  function serverTimestamp() {
    return firebase.firestore.FieldValue.serverTimestamp();
  }

  function toIsoDate(value) {
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function parseDate(value) {
    const raw = String(value || '');
    const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return null;
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function formatDate(value) {
    const date = parseDate(value);
    if (!date) return '—';
    return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
  }

  function formatDateRange(item) {
    if (!item?.startDate) return '—';
    if (!item.endDate || item.endDate === item.startDate) return formatDate(item.startDate);
    return `${formatDate(item.startDate)} – ${formatDate(item.endDate)}`;
  }

  function monthName(month) {
    return new Intl.DateTimeFormat('id-ID', { month: 'long' }).format(new Date(2026, Number(month || 1) - 1, 1));
  }

  function normalizeView(value) {
    const key = String(value || '').toLowerCase();
    return Object.prototype.hasOwnProperty.call(VIEW_LABELS, key) ? key : 'dashboard';
  }

  function viewFromHash() {
    return normalizeView(String(location.hash || '').replace(/^#/, ''));
  }

  function employeeIdentity(id, fallback) {
    const row = state.directory.get(String(id || '')) || fallback || {};
    return {
      id: String(id || ''),
      name: String(row.name || '').trim(),
      nip: String(row.nip || '').replace(/\D/g, ''),
      unit: String(row.unit || '').trim()
    };
  }

  function activeEmployees() {
    return state.employees
      .filter((item) => item.active !== false)
      .map((item) => ({ ...item, ...employeeIdentity(item.id) }))
      .filter((item) => item.name)
      .sort((a, b) => Number(a.order || 9999) - Number(b.order || 9999) || a.name.localeCompare(b.name, 'id'));
  }

  function uniqueUnits() {
    return [...new Set(activeEmployees().map((item) => item.unit).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'id'));
  }

  function normalizeDocument(raw, index) {
    const row = raw && typeof raw === 'object' ? raw : {};
    const file = row.file && typeof row.file === 'object' ? { ...row.file } : null;
    if (file) delete file.file;
    return {
      id: String(row.id || `doc-${index + 1}`),
      type: String(row.type || 'Surat Tugas'),
      name: String(row.name || ''),
      number: String(row.number || ''),
      note: String(row.note || ''),
      file
    };
  }

  function normalizeAssignment(data, id) {
    const row = data || {};
    const employees = (Array.isArray(row.employees) ? row.employees : [])
      .map((employee) => ({
        employeeId: String(employee?.employeeId || employee?.id || ''),
        name: String(employee?.name || ''),
        nip: String(employee?.nip || '').replace(/\D/g, ''),
        unit: String(employee?.unit || '')
      }))
      .filter((employee) => employee.employeeId);
    return {
      ...row,
      id,
      letterNumber: String(row.letterNumber || ''),
      letterDate: String(row.letterDate || ''),
      startDate: String(row.startDate || ''),
      endDate: String(row.endDate || row.startDate || ''),
      assignmentYear: Number(row.assignmentYear || String(row.startDate || '').slice(0, 4) || CURRENT_YEAR),
      employees,
      activity: String(row.activity || ''),
      location: String(row.location || ''),
      note: String(row.note || ''),
      adjustmentType: String(row.adjustmentType || 'official_duty'),
      adjustmentPercent: Number.isFinite(Number(row.adjustmentPercent)) ? Number(row.adjustmentPercent) : 0,
      documents: (Array.isArray(row.documents) ? row.documents : []).map(normalizeDocument)
    };
  }

  function assignmentPeople(item) {
    return (item?.employees || []).map((employee) => employeeIdentity(employee.employeeId, employee));
  }

  function statusOf(item) {
    if (!item?.startDate || !item?.endDate) return { key: 'done', label: 'Tidak lengkap' };
    if (TODAY < item.startDate) return { key: 'upcoming', label: 'Akan datang' };
    if (TODAY > item.endDate) return { key: 'done', label: 'Selesai' };
    return { key: 'active', label: 'Berjalan' };
  }

  function formatFileSize(value) {
    const bytes = Number(value || 0);
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(bytes < 10240 ? 1 : 0)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0)} MB`;
  }

  function primaryDocument(item) {
    return (item?.documents || [])[0] || null;
  }

  function assignmentSharePointPath(item) {
    const year = Number(item?.assignmentYear || String(item?.startDate || '').slice(0, 4) || CURRENT_YEAR);
    const number = String(item?.letterNumber || '').replace(/[^A-Za-z0-9_-]/g, '-').slice(0, 48);
    const folder = sharePoint?.safeName?.(`ST_${number || item?.id || 'surat-tugas'}`, item?.id || 'surat-tugas') || `ST_${number || item?.id || 'surat-tugas'}`;
    return [String(year), folder];
  }

  function assignmentRemoteFileName(item, file) {
    const original = String(file?.name || 'Surat_Tugas.pdf');
    const dot = original.lastIndexOf('.');
    const ext = dot > 0 && original.length - dot <= 10 ? original.slice(dot) : '';
    const number = String(item?.letterNumber || item?.id || 'Surat_Tugas').replace(/[^A-Za-z0-9_-]/g, '-').slice(0, 70);
    return sharePoint?.safeName?.(`Surat_Tugas_${number}${ext}`, `Surat_Tugas${ext}`) || `Surat_Tugas_${number}${ext}`;
  }

  function metric(label, value, note) {
    return `<section class="card assignment-metric"><span>${esc(label)}</span><strong>${esc(value)}</strong><small>${esc(note || '')}</small></section>`;
  }

  function yearOptions(selected) {
    const years = new Set([CURRENT_YEAR - 2, CURRENT_YEAR - 1, CURRENT_YEAR, CURRENT_YEAR + 1]);
    state.assignments.forEach((item) => years.add(Number(item.assignmentYear || CURRENT_YEAR)));
    return [...years].sort((a, b) => b - a).map((year) => `<option value="${year}" ${Number(selected) === year ? 'selected' : ''}>${year}</option>`).join('');
  }

  function unitOptions(selected, allowAll) {
    return `<select id="assignment-filter-unit">${allowAll ? '<option value="">Semua Unit/Bidang</option>' : ''}${uniqueUnits().map((unit) => `<option value="${esc(unit)}" ${selected === unit ? 'selected' : ''}>${esc(unit)}</option>`).join('')}</select>`;
  }

  function filteredAssignments() {
    const f = state.filters;
    const query = f.search.trim().toLowerCase();
    return state.assignments
      .filter((item) => Number(item.assignmentYear) === Number(f.year))
      .filter((item) => {
        const people = assignmentPeople(item);
        if (f.unit && !people.some((person) => person.unit === f.unit)) return false;
        if (f.status && statusOf(item).key !== f.status) return false;
        if (!query) return true;
        const haystack = [
          item.letterNumber, item.activity, item.location,
          ...people.flatMap((person) => [person.name, person.nip, person.unit])
        ].join(' ').toLowerCase();
        return haystack.includes(query);
      })
      .sort((a, b) => b.startDate.localeCompare(a.startDate) || b.letterDate.localeCompare(a.letterDate));
  }

  function peopleCompact(item) {
    const people = assignmentPeople(item);
    if (!people.length) return '<span class="assignment-sub">Belum ada pegawai</span>';
    const first = people[0];
    return `<div class="assignment-person">${esc(first.name || 'Pegawai')}</div><div class="assignment-person-sub">${first.nip ? `NIP ${esc(first.nip)} · ` : ''}${esc(first.unit || 'Unit belum diisi')}${people.length > 1 ? ` · +${people.length - 1} pegawai` : ''}</div>`;
  }

  function dashboardView() {
    const yearRows = state.assignments.filter((item) => Number(item.assignmentYear) === CURRENT_YEAR);
    const active = yearRows.filter((item) => statusOf(item).key === 'active').length;
    const upcoming = yearRows.filter((item) => statusOf(item).key === 'upcoming').length;
    const uniqueEmployeeIds = new Set(yearRows.flatMap((item) => (item.employees || []).map((employee) => employee.employeeId)));
    const upcomingRows = yearRows
      .filter((item) => item.endDate >= TODAY)
      .sort((a, b) => a.startDate.localeCompare(b.startDate))
      .slice(0, 6);

    return `<div class="assignment-page">
      <div class="assignment-head"><div><h1>Surat Tugas</h1><p>Arsip penugasan/dinas pegawai yang terhubung otomatis dengan proses Tunjangan Kinerja.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <div class="assignment-metrics">${metric('Surat Tugas', yearRows.length, `Tahun ${CURRENT_YEAR}`)}${metric('Sedang Berjalan', active, 'Berdasarkan tanggal hari ini')}${metric('Akan Datang', upcoming, 'Penugasan yang belum dimulai')}${metric('Pegawai Ditugaskan', uniqueEmployeeIds.size, 'Pegawai unik pada tahun berjalan')}</div>
      <section class="card assignment-card"><div class="assignment-card-head"><div><h2>Penugasan Terdekat</h2><p>Surat Tugas berjalan dan akan datang.</p></div><a class="btn btn-secondary btn-sm" href="surat-tugas.html#data">Lihat Semua</a></div>
        ${upcomingRows.length ? `<div class="assignment-upcoming-list">${upcomingRows.map((item) => `<article class="assignment-upcoming"><strong>${esc(item.activity || item.letterNumber || 'Surat Tugas')}</strong><span>${esc(formatDateRange(item))} · ${esc(item.location || 'Lokasi belum diisi')}</span><span>${esc(assignmentPeople(item).map((person) => person.name).filter(Boolean).slice(0, 3).join(', '))}${assignmentPeople(item).length > 3 ? ` +${assignmentPeople(item).length - 3}` : ''}</span><div style="margin-top:8px"><button class="btn btn-secondary btn-sm" type="button" data-assignment-detail="${esc(item.id)}">Detail</button></div></article>`).join('')}</div>` : '<div class="assignment-empty"><strong>Belum ada penugasan</strong>Tambahkan Surat Tugas untuk mulai membangun arsip penugasan.</div>'}
      </section>
    </div>`;
  }

  function dataView() {
    const rows = filteredAssignments();
    return `<div class="assignment-page">
      <div class="assignment-head"><div><h1>Data Surat Tugas</h1><p>Satu record dapat mencakup beberapa pegawai dan satu rentang tanggal penugasan.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <section class="card assignment-filter-card"><div class="assignment-filter-grid">
        <div class="search"><input id="assignment-filter-search" type="search" value="${esc(state.filters.search)}" placeholder="Cari nomor ST, kegiatan, pegawai, NIP, lokasi"></div>
        <div class="field"><label>Tahun</label><select id="assignment-filter-year">${yearOptions(state.filters.year)}</select></div>
        <div class="field"><label>Unit/Bidang</label>${unitOptions(state.filters.unit, true)}</div>
        <div class="field"><label>Status</label><select id="assignment-filter-status"><option value="">Semua Status</option><option value="upcoming" ${state.filters.status === 'upcoming' ? 'selected' : ''}>Akan datang</option><option value="active" ${state.filters.status === 'active' ? 'selected' : ''}>Berjalan</option><option value="done" ${state.filters.status === 'done' ? 'selected' : ''}>Selesai</option></select></div>
        <button class="btn btn-secondary" type="button" data-action="reset-filter">Reset</button>
      </div></section>
      <section class="card assignment-card"><div class="assignment-table-wrap"><table class="data-table assignment-table"><thead><tr><th>No</th><th>Pegawai</th><th>Nomor Surat</th><th>Tanggal Surat</th><th>Periode Tugas</th><th>Kegiatan</th><th>Lokasi</th><th>Status</th><th>Dokumen</th><th>Aksi</th></tr></thead><tbody>${rows.length ? rows.map((item, index) => {
        const status = statusOf(item);
        const doc = primaryDocument(item);
        return `<tr><td>${index + 1}</td><td>${peopleCompact(item)}</td><td><strong>${esc(item.letterNumber || '—')}</strong></td><td class="nowrap">${esc(formatDate(item.letterDate))}</td><td class="nowrap">${esc(formatDateRange(item))}</td><td><strong>${esc(item.activity || '—')}</strong><div class="assignment-sub">${esc(item.note || '')}</div></td><td>${esc(item.location || '—')}</td><td><span class="assignment-pill ${status.key}">${esc(status.label)}</span></td><td>${doc?.file?.itemId ? `<span class="assignment-pill">Tersimpan</span><div class="assignment-sub">${esc(doc.file.name || doc.file.remoteName || 'Surat Tugas')}</div>` : '<span class="assignment-sub">Belum ada file</span>'}</td><td><div class="assignment-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-assignment-detail="${esc(item.id)}">Detail</button><button class="btn btn-secondary btn-sm" type="button" data-assignment-edit="${esc(item.id)}">Edit</button></div></td></tr>`;
      }).join('') : '<tr><td colspan="10"><div class="assignment-empty"><strong>Tidak ada data</strong>Tidak ada Surat Tugas yang sesuai filter.</div></td></tr>'}</tbody></table></div></section>
    </div>`;
  }

  function calendarRows() {
    const year = Number(state.calendar.year);
    const month = Number(state.calendar.month);
    const start = `${year}-${String(month).padStart(2, '0')}-01`;
    const end = `${year}-${String(month).padStart(2, '0')}-${String(new Date(year, month, 0).getDate()).padStart(2, '0')}`;
    return state.assignments.filter((item) => item.startDate <= end && item.endDate >= start).sort((a, b) => a.startDate.localeCompare(b.startDate));
  }

  function calendarView() {
    const rows = calendarRows();
    return `<div class="assignment-page">
      <div class="assignment-head"><div><h1>Kalender Surat Tugas</h1><p>Daftar penugasan berdasarkan rentang tanggal dinas.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <section class="card assignment-card"><div class="assignment-card-head assignment-calendar-toolbar"><div><h2>${esc(monthName(state.calendar.month))} ${esc(state.calendar.year)}</h2><p>${rows.length} Surat Tugas beririsan dengan bulan ini.</p></div><div class="assignment-calendar-nav"><button class="btn btn-secondary btn-sm" type="button" data-action="calendar-prev">←</button><strong>${esc(monthName(state.calendar.month))}</strong><button class="btn btn-secondary btn-sm" type="button" data-action="calendar-next">→</button></div></div>
        <div class="assignment-calendar-list">${rows.length ? rows.map((item) => `<button class="assignment-calendar-row" type="button" data-assignment-detail="${esc(item.id)}" style="width:100%;border-left:0;border-right:0;border-top:0;background:#fff;text-align:left"><div class="date">${esc(formatDateRange(item))}</div><div class="activity"><strong>${esc(item.activity || item.letterNumber || 'Surat Tugas')}</strong><span>${esc(item.letterNumber || 'Tanpa nomor')}</span></div><div>${esc(assignmentPeople(item).map((person) => person.name).filter(Boolean).slice(0, 2).join(', '))}${assignmentPeople(item).length > 2 ? ` +${assignmentPeople(item).length - 2}` : ''}</div><div><span class="assignment-pill ${statusOf(item).key}">${esc(statusOf(item).label)}</span></div></button>`).join('') : '<div class="assignment-empty"><strong>Tidak ada penugasan</strong>Belum ada Surat Tugas pada bulan ini.</div>'}</div>
      </section>
    </div>`;
  }

  function employeePickerMarkup() {
    const selected = state.modal?.selectedEmployeeIds || new Set();
    return `<div class="assignment-employee-picker"><div class="assignment-employee-tools"><input id="assignment-employee-search" type="search" placeholder="Cari pegawai, NIP, atau unit"><button class="btn btn-secondary btn-sm" type="button" data-action="select-all-employees">Pilih Semua</button><button class="btn btn-secondary btn-sm" type="button" data-action="clear-employees">Kosongkan</button><span class="assignment-employee-summary" id="assignment-employee-summary">${selected.size} pegawai dipilih</span></div><div class="assignment-employee-list" id="assignment-employee-list">${activeEmployees().map((employee) => `<label class="assignment-employee-option" data-employee-search-text="${esc(`${employee.name} ${employee.nip} ${employee.unit}`.toLowerCase())}"><input type="checkbox" data-assignment-employee="${esc(employee.id)}" ${selected.has(employee.id) ? 'checked' : ''}><span><strong>${esc(employee.name)}</strong><span>${employee.nip ? `NIP ${esc(employee.nip)} · ` : ''}${esc(employee.unit || 'Unit/Bidang belum diisi')}</span></span></label>`).join('')}</div></div>`;
  }

  function assignmentModal() {
    if (!state.modal) return '';
    const item = state.modal.record || {};
    const doc = primaryDocument(item);
    const currentFile = doc?.file || null;
    const currentFileMarkup = currentFile?.itemId ? `<div class="assignment-current-file"><span>File saat ini</span><strong>${esc(currentFile.name || currentFile.remoteName || doc.name || 'Surat Tugas')}</strong>${currentFile.size ? `<small>${esc(formatFileSize(currentFile.size))}</small>` : '<small>SharePoint</small>'}</div>` : '';
    return `<div class="assignment-modal-backdrop" data-action="close-assignment-modal"><div class="assignment-modal" role="dialog" aria-modal="true" data-assignment-modal-panel>
      <div class="assignment-modal-head"><div><h3>${item.id ? 'Edit Surat Tugas' : 'Tambah Surat Tugas'}</h3><p>Upload surat, pilih satu atau lebih pegawai, kemudian isi rentang tanggal penugasan. Data ini akan menjadi sumber penyesuaian otomatis pada proses Tukin.</p></div><button class="icon-btn" type="button" data-action="close-assignment-modal">×</button></div>
      <div class="assignment-form-grid">
        <div class="field"><label for="assignment-letter-number">Nomor Surat Tugas</label><input id="assignment-letter-number" type="text" value="${esc(item.letterNumber || '')}" placeholder="Contoh: ST/123/Pusdatin/2026"></div>
        <div class="field"><label for="assignment-letter-date">Tanggal Surat</label><input id="assignment-letter-date" type="date" value="${esc(item.letterDate || '')}"></div>
        <div class="field"><label for="assignment-start">Tanggal Mulai Tugas</label><input id="assignment-start" type="date" value="${esc(item.startDate || '')}"></div>
        <div class="field"><label for="assignment-end">Tanggal Selesai Tugas</label><input id="assignment-end" type="date" value="${esc(item.endDate || '')}"></div>
        <div class="field span-2"><label>Pegawai yang Ditugaskan</label>${employeePickerMarkup()}<div class="assignment-field-note">Satu Surat Tugas cukup dicatat satu kali walaupun memuat banyak pegawai.</div></div>
        <div class="field span-2"><label for="assignment-activity">Kegiatan / Keperluan</label><input id="assignment-activity" type="text" value="${esc(item.activity || '')}" placeholder="Contoh: Rapat Koordinasi ..."></div>
        <div class="field"><label for="assignment-location">Lokasi / Tujuan</label><input id="assignment-location" type="text" value="${esc(item.location || '')}" placeholder="Opsional"></div>
        <div class="field"><label>Perlakuan Tukin</label><input type="text" value="Dinas / Surat Tugas — penyesuaian 0%" readonly><div class="assignment-field-note">Hanya diterapkan pada tanggal yang berstatus perlu verifikasi. Nilai otomatis asal tetap ditampilkan.</div></div>
        <div class="field span-2"><label for="assignment-note">Keterangan</label><textarea id="assignment-note" rows="3" placeholder="Opsional">${esc(item.note || '')}</textarea></div>
        <div class="field span-2"><label for="assignment-file">Berkas Surat Tugas${currentFile?.itemId ? ' <span class="assignment-sub">(pilih file baru untuk mengganti)</span>' : ''}</label><input class="assignment-file-input" id="assignment-file" type="file" accept=".pdf,.doc,.docx,.jpg,.jpeg,.png">${currentFileMarkup}<div class="assignment-field-note">File disimpan pada folder SURAT_TUGAS di SharePoint. PDF dan gambar dapat dipreview langsung sebelum disimpan.</div><div id="assignment-local-preview"></div></div>
      </div>
      <div class="assignment-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-assignment-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-assignment">Simpan</button></div>
    </div></div>`;
  }

  function previewPostFields(raw) {
    const params = new URLSearchParams(String(raw || ''));
    return [...params.entries()].map(([name, value]) => `<input type="hidden" name="${esc(name)}" value="${esc(value)}">`).join('');
  }

  function inlinePreviewKey(assignmentId, documentId) {
    return `${String(assignmentId || '')}::${String(documentId || '')}`;
  }

  function inlinePreviewMarkup(item, doc) {
    const file = doc?.file || null;
    if (!file?.driveId || !file?.itemId) return '';
    const key = inlinePreviewKey(item.id, doc.id);
    const ctx = state.inlinePreviews[key] || { loading: true, error: '', getUrl: '', postUrl: '', postParameters: '' };
    const fileName = file.name || file.remoteName || doc.name || 'Surat Tugas';
    let body = '';
    if (ctx.loading) {
      body = '<div class="assignment-inline-preview-status"><strong>Menyiapkan preview dokumen</strong><span>Mengambil tampilan sementara dari Microsoft 365...</span></div>';
    } else if (ctx.error) {
      body = `<div class="assignment-inline-preview-status assignment-inline-preview-error"><strong>Preview tidak tersedia</strong><span>${esc(ctx.error)}</span><small>Gunakan tombol Unduh atau SharePoint untuk membuka dokumen.</small></div>`;
    } else if (ctx.getUrl) {
      body = `<iframe class="assignment-inline-preview-frame" src="${esc(ctx.getUrl)}" title="Preview ${esc(fileName)}" allow="fullscreen" referrerpolicy="no-referrer"></iframe>`;
    } else if (ctx.postUrl) {
      const target = `assignment-inline-${String(item.id).replace(/[^A-Za-z0-9_-]/g, '')}`;
      body = `<iframe class="assignment-inline-preview-frame" name="${esc(target)}" title="Preview ${esc(fileName)}" allow="fullscreen" referrerpolicy="no-referrer"></iframe><form class="assignment-preview-post-form" data-assignment-inline-preview-post method="post" action="${esc(ctx.postUrl)}" target="${esc(target)}">${previewPostFields(ctx.postParameters)}</form>`;
    } else {
      body = '<div class="assignment-inline-preview-status assignment-inline-preview-error"><strong>Preview tidak tersedia</strong><span>Microsoft 365 tidak memberikan URL preview untuk dokumen ini.</span></div>';
    }
    return `<div class="assignment-inline-preview"><div class="assignment-inline-preview-label"><span>Preview Surat Tugas</span><small>Dokumen SharePoint dimuat otomatis</small></div><div class="assignment-inline-preview-body">${body}</div></div>`;
  }

  function assignmentDetail() {
    if (!state.detailId) return '';
    const item = state.assignments.find((row) => row.id === state.detailId);
    if (!item) return '';
    const people = assignmentPeople(item);
    const doc = primaryDocument(item);
    const file = doc?.file || null;
    const status = statusOf(item);
    return `<div class="assignment-detail-backdrop" data-action="close-assignment-detail"></div><section class="assignment-detail-modal" role="dialog" aria-modal="true" aria-label="Detail Surat Tugas">
      <div class="assignment-modal-head assignment-detail-head"><div><h3>${esc(item.letterNumber || 'Surat Tugas')}</h3><p>${esc(item.activity || 'Detail penugasan/dinas pegawai.')}</p></div><button class="icon-btn" type="button" data-action="close-assignment-detail">×</button></div>
      <div class="assignment-section"><div class="assignment-section-title">Informasi Surat Tugas</div><div class="assignment-detail-grid"><div><span>Nomor Surat</span><strong>${esc(item.letterNumber || '—')}</strong></div><div><span>Tanggal Surat</span><strong>${esc(formatDate(item.letterDate))}</strong></div><div><span>Status</span><strong>${esc(status.label)}</strong></div><div><span>Perlakuan Tukin</span><strong>Dinas · 0%</strong></div><div><span>Tanggal Mulai</span><strong>${esc(formatDate(item.startDate))}</strong></div><div><span>Tanggal Selesai</span><strong>${esc(formatDate(item.endDate))}</strong></div><div class="span-2"><span>Lokasi / Tujuan</span><strong>${esc(item.location || '—')}</strong></div><div class="span-all"><span>Kegiatan / Keperluan</span><strong>${esc(item.activity || '—')}</strong></div><div class="span-all"><span>Keterangan</span><strong>${esc(item.note || '—')}</strong></div></div></div>
      <div class="assignment-section"><div class="assignment-section-title">Pegawai yang Ditugaskan (${people.length})</div><div class="assignment-people-list">${people.length ? people.map((person) => `<div class="assignment-people-item"><strong>${esc(person.name || 'Pegawai')}</strong><span>${person.nip ? `NIP ${esc(person.nip)} · ` : ''}${esc(person.unit || 'Unit/Bidang belum diisi')}</span></div>`).join('') : '<div class="assignment-empty"><strong>Belum ada pegawai</strong>Record ini belum memiliki pegawai.</div>'}</div></div>
      <div class="assignment-section"><div class="assignment-section-title">Berkas Surat Tugas</div>${doc ? `<article class="assignment-document-card"><div class="assignment-document-row"><div><strong>${esc(file?.name || doc.name || 'Surat Tugas')}</strong><span>${esc(doc.type || 'Surat Tugas')}${file?.size ? ` · ${esc(formatFileSize(file.size))}` : ''}</span></div><div class="assignment-row-actions">${file?.itemId ? `<button class="btn btn-secondary btn-sm" type="button" data-download-assignment-document="${esc(doc.id)}" data-assignment-id="${esc(item.id)}">Unduh</button><button class="btn btn-secondary btn-sm" type="button" data-preview-assignment-document="${esc(doc.id)}" data-assignment-id="${esc(item.id)}">Preview Besar</button>${file.webUrl ? `<a class="btn btn-secondary btn-sm" href="${esc(file.webUrl)}" target="_blank" rel="noopener noreferrer">SharePoint ↗</a>` : ''}` : ''}</div></div>${inlinePreviewMarkup(item, doc)}</article>` : '<div class="assignment-empty"><strong>Berkas belum tersedia</strong>Edit Surat Tugas untuk mengunggah file.</div>'}</div>
      <div class="assignment-modal-actions" style="justify-content:space-between"><button class="btn btn-danger" type="button" data-assignment-delete="${esc(item.id)}">Hapus</button><div class="assignment-row-actions"><button class="btn btn-secondary" type="button" data-action="close-assignment-detail">Tutup</button><button class="btn btn-primary" type="button" data-assignment-edit="${esc(item.id)}">Edit Data</button></div></div>
    </section>`;
  }

  function previewModal() {
    if (!state.previewModal) return '';
    const ctx = state.previewModal;
    const sharePointAction = ctx.webUrl ? `<a class="btn btn-secondary btn-sm" href="${esc(ctx.webUrl)}" target="_blank" rel="noopener noreferrer">Buka SharePoint ↗</a>` : '';
    let body = '';
    if (ctx.loading) body = '<div class="assignment-preview-status"><strong>Menyiapkan preview dokumen</strong><span>Mengambil tampilan sementara dari Microsoft 365...</span></div>';
    else if (ctx.error) body = `<div class="assignment-preview-status"><strong>Preview tidak tersedia</strong><span>${esc(ctx.error)}</span></div>`;
    else if (ctx.getUrl) body = `<iframe class="assignment-preview-frame" src="${esc(ctx.getUrl)}" title="Preview ${esc(ctx.fileName)}" allow="fullscreen" referrerpolicy="no-referrer"></iframe>`;
    else if (ctx.postUrl) {
      const target = `assignment-preview-${String(ctx.assignmentId || 'assignment').replace(/[^A-Za-z0-9_-]/g, '')}`;
      body = `<iframe class="assignment-preview-frame" name="${esc(target)}" title="Preview ${esc(ctx.fileName)}" allow="fullscreen" referrerpolicy="no-referrer"></iframe><form id="assignment-preview-post-form" class="assignment-preview-post-form" method="post" action="${esc(ctx.postUrl)}" target="${esc(target)}">${previewPostFields(ctx.postParameters)}</form>`;
    }
    return `<div class="assignment-modal-backdrop assignment-preview-backdrop" data-action="close-preview-modal"><div class="assignment-modal assignment-preview-modal" role="dialog" aria-modal="true"><div class="assignment-modal-head assignment-preview-head"><div><h3>${esc(ctx.fileName || 'Surat Tugas')}</h3><p>Preview ditampilkan langsung dari file SharePoint.</p></div><div class="assignment-row-actions">${sharePointAction}<button class="icon-btn" type="button" data-action="close-preview-modal">×</button></div></div><div class="assignment-preview-body">${body}</div></div></div>`;
  }

  function overlays() {
    return `${assignmentDetail()}${assignmentModal()}${previewModal()}${state.toast ? `<div class="assignment-toast">${esc(state.toast)}</div>` : ''}`;
  }

  function viewMarkup() {
    if (state.view === 'data') return dataView();
    if (state.view === 'kalender') return calendarView();
    return dashboardView();
  }

  function render() {
    if (state.loading) {
      app.innerHTML = window.AppShell.render({ module: 'assignment', view: state.view, viewLabel: VIEW_LABELS[state.view], content: '<div class="card assignment-loading"><strong>Memuat modul Surat Tugas</strong>Membaca master pegawai dan arsip penugasan dari Cloud Firestore...</div>' });
      return;
    }
    const content = `${state.error ? `<div class="alert alert-warning"><div class="alert-title">Data belum dapat dimuat sempurna</div>${esc(state.error)}</div>` : ''}${viewMarkup()}`;
    app.innerHTML = window.AppShell.render({ module: 'assignment', view: state.view, viewLabel: VIEW_LABELS[state.view], content, overlays: overlays() });
    if (state.detailId) window.requestAnimationFrame(submitInlinePreviewPostForms);
  }

  async function audit(action, target, detail) {
    try {
      await db().collection('adminAudit').doc().set({
        action: String(action || ''), target: String(target || ''), detail: String(detail || ''),
        actorUid: user()?.uid || '', actorEmail: user()?.email || '', createdAt: serverTimestamp()
      });
    } catch (error) {
      console.warn('Audit Surat Tugas gagal ditulis:', error);
    }
  }

  async function loadData() {
    state.loading = true;
    state.error = '';
    render();
    try {
      const [masterBundle, snapshot] = await Promise.all([
        window.MasterDataService.getBundle(false),
        db().collection('assignmentRecords').get()
      ]);
      state.employees = [...masterBundle.employees];
      state.directory = new Map(masterBundle.directory);
      state.assignments = snapshot.docs.map((doc) => normalizeAssignment(doc.data(), doc.id));
    } catch (error) {
      state.error = error.message || String(error);
    } finally {
      state.loading = false;
      render();
    }
  }

  function showToast(message) {
    state.toast = String(message || '');
    render();
    clearTimeout(window.__assignmentToastTimer);
    window.__assignmentToastTimer = setTimeout(() => { state.toast = ''; render(); }, 3200);
  }

  function openAssignmentModal(id) {
    const existing = id ? state.assignments.find((item) => item.id === id) : null;
    const selected = new Set((existing?.employees || []).map((employee) => employee.employeeId));
    state.detailId = '';
    state.previewModal = null;
    state.modal = {
      record: existing ? { ...existing, employees: (existing.employees || []).map((employee) => ({ ...employee })), documents: (existing.documents || []).map((doc) => ({ ...doc, file: doc.file ? { ...doc.file } : null })) } : { letterNumber: '', letterDate: TODAY, startDate: '', endDate: '', activity: '', location: '', note: '', documents: [] },
      selectedEmployeeIds: selected
    };
    render();
  }

  function selectedEmployeeIdsFromForm() {
    return new Set([...document.querySelectorAll('[data-assignment-employee]:checked')].map((input) => String(input.dataset.assignmentEmployee || '')).filter(Boolean));
  }

  function updateEmployeeSummary() {
    if (!state.modal) return;
    state.modal.selectedEmployeeIds = selectedEmployeeIdsFromForm();
    const target = document.getElementById('assignment-employee-summary');
    if (target) target.textContent = `${state.modal.selectedEmployeeIds.size} pegawai dipilih`;
  }

  function filterEmployeePicker(query) {
    const needle = String(query || '').trim().toLowerCase();
    document.querySelectorAll('[data-employee-search-text]').forEach((row) => {
      row.style.display = !needle || String(row.dataset.employeeSearchText || '').includes(needle) ? '' : 'none';
    });
  }

  function renderLocalFilePreview(file) {
    const target = document.getElementById('assignment-local-preview');
    if (!target) return;
    if (window.__assignmentLocalPreviewUrl) {
      URL.revokeObjectURL(window.__assignmentLocalPreviewUrl);
      window.__assignmentLocalPreviewUrl = '';
    }
    if (!file) { target.innerHTML = ''; return; }
    const url = URL.createObjectURL(file);
    window.__assignmentLocalPreviewUrl = url;
    const lower = String(file.name || '').toLowerCase();
    let body = '';
    if (file.type === 'application/pdf' || lower.endsWith('.pdf')) {
      body = `<iframe src="${esc(url)}" title="Preview ${esc(file.name)}"></iframe>`;
    } else if ((file.type || '').startsWith('image/') || /\.(png|jpe?g|webp)$/i.test(lower)) {
      body = `<img src="${esc(url)}" alt="Preview ${esc(file.name)}">`;
    } else {
      body = `<div class="assignment-local-preview-status"><strong>${esc(file.name)}</strong><br>Preview lokal belum tersedia untuk format ini. Setelah disimpan, aplikasi akan mencoba preview melalui Microsoft 365.</div>`;
    }
    target.innerHTML = `<div class="assignment-local-preview"><div class="assignment-local-preview-head"><span>Preview File yang Dipilih</span><span>${esc(formatFileSize(file.size))}</span></div><div class="assignment-local-preview-body">${body}</div></div>`;
  }

  async function saveAssignment(button) {
    if (!state.modal) return;
    const selectedIds = selectedEmployeeIdsFromForm();
    const letterNumber = String(document.getElementById('assignment-letter-number')?.value || '').trim();
    const letterDate = String(document.getElementById('assignment-letter-date')?.value || '');
    const startDate = String(document.getElementById('assignment-start')?.value || '');
    const endDate = String(document.getElementById('assignment-end')?.value || '');
    const activity = String(document.getElementById('assignment-activity')?.value || '').trim();
    const locationText = String(document.getElementById('assignment-location')?.value || '').trim();
    const note = String(document.getElementById('assignment-note')?.value || '').trim();
    const selectedFile = document.getElementById('assignment-file')?.files?.[0] || null;
    const existing = state.modal.record || {};
    const existingDoc = primaryDocument(existing);
    const existingFile = existingDoc?.file || null;

    if (!letterNumber) { alert('Nomor Surat Tugas wajib diisi.'); return; }
    if (!letterDate) { alert('Tanggal Surat wajib diisi.'); return; }
    if (!startDate) { alert('Tanggal mulai tugas wajib diisi.'); return; }
    if (!endDate) { alert('Tanggal selesai tugas wajib diisi.'); return; }
    if (endDate < startDate) { alert('Tanggal selesai tugas tidak boleh lebih awal dari tanggal mulai.'); return; }
    if (!selectedIds.size) { alert('Pilih sekurang-kurangnya satu pegawai yang ditugaskan.'); return; }
    if (!activity) { alert('Kegiatan/keperluan wajib diisi.'); return; }
    if (!selectedFile && !existingFile?.itemId) { alert('Upload berkas Surat Tugas terlebih dahulu.'); return; }

    const docRef = existing.id ? db().collection('assignmentRecords').doc(existing.id) : db().collection('assignmentRecords').doc();
    const id = docRef.id;
    const assignmentYear = Number(startDate.slice(0, 4)) || CURRENT_YEAR;
    const employees = [...selectedIds].map((employeeId) => {
      const identity = employeeIdentity(employeeId);
      return { employeeId, name: identity.name, nip: identity.nip, unit: identity.unit };
    });
    const draft = { id, letterNumber, letterDate, startDate, endDate, assignmentYear, employees, activity, location: locationText, note };
    let storedFile = existingFile ? { ...existingFile } : null;
    let uploadedFile = null;
    let firestoreSaved = false;

    button.disabled = true;
    button.textContent = selectedFile ? 'Mengunggah...' : 'Menyimpan...';
    try {
      if (selectedFile) {
        await sharePoint.ensureReady(true);
        const localAttachment = sharePoint.normalizeAttachment(selectedFile, 'assignment');
        localAttachment.id = `assignment-evidence-${id}`;
        uploadedFile = await sharePoint.uploadAttachment(localAttachment, {
          kind: 'assignment',
          path: assignmentSharePointPath(draft),
          remoteName: assignmentRemoteFileName(draft, selectedFile)
        }, true);
        storedFile = sharePoint.serializableAttachment(uploadedFile);
      }

      const documents = [{
        id: 'surat-tugas',
        type: 'Surat Tugas',
        name: storedFile?.name || existingDoc?.name || selectedFile?.name || `Surat Tugas ${letterNumber}`,
        number: letterNumber,
        note: activity,
        file: storedFile
      }];

      const payload = {
        letterNumber, letterDate, startDate, endDate, assignmentYear,
        employees, activity, location: locationText, note,
        adjustmentType: 'official_duty', adjustmentPercent: 0,
        documents,
        schemaVersion: 1,
        updatedAt: serverTimestamp(), updatedBy: user()?.email || ''
      };
      if (!existing.id) {
        payload.createdAt = serverTimestamp();
        payload.createdBy = user()?.email || '';
      }

      button.textContent = 'Menyimpan...';
      await docRef.set(payload, { merge: true });
      firestoreSaved = true;

      if (selectedFile && existingFile?.driveId && existingFile?.itemId && existingFile.itemId !== storedFile?.itemId) {
        try { await sharePoint.deleteAttachment(existingFile, true); }
        catch (cleanupError) { console.warn('File Surat Tugas lama tidak berhasil dihapus:', cleanupError); }
      }

      await audit(existing.id ? 'UPDATE_ASSIGNMENT' : 'CREATE_ASSIGNMENT', `assignmentRecords/${id}`, `${letterNumber} · ${activity} · ${employees.length} pegawai`);
      state.modal = null;
      await loadData();
      showToast(existing.id ? 'Surat Tugas berhasil diperbarui.' : 'Surat Tugas berhasil disimpan dan file telah diunggah ke SharePoint.');
    } catch (error) {
      if (uploadedFile?.driveId && uploadedFile?.itemId && !firestoreSaved) {
        try { await sharePoint.deleteAttachment(uploadedFile, true); } catch (cleanupError) { console.warn(cleanupError); }
      }
      button.disabled = false;
      button.textContent = 'Simpan';
      alert(`Surat Tugas gagal disimpan: ${error.message || error}`);
    }
  }

  async function deleteAssignment(id) {
    const item = state.assignments.find((row) => row.id === id);
    if (!item) return;
    if (!window.confirm(`Hapus Surat Tugas ${item.letterNumber || ''}?\n\nData ini tidak akan lagi digunakan pada sinkronisasi Tukin.`)) return;
    const file = primaryDocument(item)?.file || null;
    try {
      await db().collection('assignmentRecords').doc(id).delete();
      let cleanupFailed = false;
      if (file?.driveId && file?.itemId) {
        try { await sharePoint.deleteAttachment(file, true); }
        catch (error) { cleanupFailed = true; console.warn('File SharePoint tidak berhasil dihapus:', error); }
      }
      await audit('DELETE_ASSIGNMENT', `assignmentRecords/${id}`, `${item.letterNumber} · ${item.activity}`);
      state.detailId = '';
      await loadData();
      showToast(cleanupFailed ? 'Data Surat Tugas dihapus, tetapi file SharePoint belum berhasil dibersihkan.' : 'Surat Tugas dan file pendukung berhasil dihapus.');
    } catch (error) {
      alert(`Surat Tugas gagal dihapus: ${error.message || error}`);
    }
  }

  function submitInlinePreviewPostForms() {
    document.querySelectorAll('[data-assignment-inline-preview-post]').forEach((form) => {
      try { form.submit(); } catch (error) { console.warn('Preview POST Microsoft 365 gagal dikirim:', error); }
    });
  }

  async function loadAssignmentPreview(id) {
    const item = state.assignments.find((row) => row.id === id);
    const doc = primaryDocument(item);
    if (!doc?.file?.driveId || !doc?.file?.itemId) return;
    const key = inlinePreviewKey(id, doc.id);
    state.inlinePreviews[key] = { loading: true, error: '', getUrl: '', postUrl: '', postParameters: '' };
    render();
    try {
      const preview = await sharePoint.getPreviewInfo({ ...doc.file, file: null }, true);
      if (state.detailId !== id) return;
      state.inlinePreviews[key] = { loading: false, error: '', getUrl: preview.getUrl || '', postUrl: preview.postUrl || '', postParameters: preview.postParameters || '' };
    } catch (error) {
      if (state.detailId !== id) return;
      state.inlinePreviews[key] = { loading: false, error: error.message || String(error), getUrl: '', postUrl: '', postParameters: '' };
    }
    if (state.detailId === id) {
      render();
      window.requestAnimationFrame(submitInlinePreviewPostForms);
    }
  }

  function openAssignmentDetail(id) {
    state.modal = null;
    state.previewModal = null;
    state.inlinePreviews = {};
    state.detailId = id;
    render();
    loadAssignmentPreview(id);
  }

  async function previewAssignmentDocument(assignmentId, documentId) {
    const item = state.assignments.find((row) => row.id === assignmentId);
    const doc = (item?.documents || []).find((row) => row.id === documentId);
    if (!doc?.file?.driveId || !doc?.file?.itemId) { alert('File SharePoint untuk Surat Tugas ini tidak tersedia.'); return; }
    state.previewModal = {
      assignmentId, documentId,
      fileName: doc.file.name || doc.file.remoteName || doc.name || 'Surat Tugas',
      webUrl: doc.file.webUrl || '',
      loading: true, error: '', getUrl: '', postUrl: '', postParameters: ''
    };
    render();
    try {
      const preview = await sharePoint.getPreviewInfo({ ...doc.file, file: null }, true);
      if (!state.previewModal || state.previewModal.assignmentId !== assignmentId) return;
      state.previewModal = { ...state.previewModal, loading: false, getUrl: preview.getUrl || '', postUrl: preview.postUrl || '', postParameters: preview.postParameters || '' };
      render();
      if (!preview.getUrl && preview.postUrl) {
        window.requestAnimationFrame(() => document.getElementById('assignment-preview-post-form')?.submit());
      }
    } catch (error) {
      if (!state.previewModal || state.previewModal.assignmentId !== assignmentId) return;
      state.previewModal = { ...state.previewModal, loading: false, error: error.message || String(error) };
      render();
    }
  }

  async function downloadAssignmentDocument(assignmentId, documentId, button) {
    const item = state.assignments.find((row) => row.id === assignmentId);
    const doc = (item?.documents || []).find((row) => row.id === documentId);
    if (!doc?.file?.driveId || !doc?.file?.itemId) { alert('File SharePoint untuk Surat Tugas ini tidak tersedia.'); return; }
    const original = button?.textContent || 'Unduh';
    if (button) { button.disabled = true; button.textContent = 'Mengunduh...'; }
    try {
      const file = await sharePoint.downloadAttachment({ ...doc.file, file: null }, true);
      if (!file) throw new Error('File tidak berhasil diunduh.');
      const url = URL.createObjectURL(file);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = doc.file.name || doc.file.remoteName || doc.name || 'Surat_Tugas';
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      alert(`File Surat Tugas gagal diunduh: ${error.message || error}`);
    } finally {
      if (button) { button.disabled = false; button.textContent = original; }
    }
  }

  function changeCalendarMonth(delta) {
    let year = Number(state.calendar.year);
    let month = Number(state.calendar.month) + Number(delta || 0);
    if (month < 1) { month = 12; year -= 1; }
    if (month > 12) { month = 1; year += 1; }
    state.calendar = { year, month };
    render();
  }

  app.addEventListener('change', (event) => {
    const target = event.target;
    if (target.matches?.('[data-assignment-employee]')) { updateEmployeeSummary(); return; }
    if (target.id === 'assignment-file') { renderLocalFilePreview(target.files?.[0] || null); return; }
    if (target.id === 'assignment-filter-year') { state.filters.year = Number(target.value || CURRENT_YEAR); render(); return; }
    if (target.id === 'assignment-filter-unit') { state.filters.unit = String(target.value || ''); render(); return; }
    if (target.id === 'assignment-filter-status') { state.filters.status = String(target.value || ''); render(); }
  });

  app.addEventListener('input', (event) => {
    const target = event.target;
    if (target.id === 'assignment-filter-search') {
      state.filters.search = String(target.value || '');
      const cursor = target.selectionStart;
      render();
      requestAnimationFrame(() => {
        const input = document.getElementById('assignment-filter-search');
        if (input) { input.focus(); try { input.setSelectionRange(cursor, cursor); } catch (_) {} }
      });
      return;
    }
    if (target.id === 'assignment-employee-search') filterEmployeePicker(target.value);
  });

  app.addEventListener('click', (event) => {
    const actionEl = event.target.closest?.('[data-action]');
    const action = actionEl?.dataset.action;
    if (action === 'close-assignment-modal' && actionEl.classList.contains('assignment-modal-backdrop') && event.target !== actionEl) return;
    if (action === 'close-assignment-detail' && actionEl.classList.contains('assignment-detail-backdrop') && event.target !== actionEl) return;
    if (action === 'close-preview-modal' && actionEl.classList.contains('assignment-modal-backdrop') && event.target !== actionEl) return;

    if (action === 'add-assignment') { openAssignmentModal(''); return; }
    if (action === 'close-assignment-modal') { state.modal = null; render(); return; }
    if (action === 'close-assignment-detail') { state.detailId = ''; state.inlinePreviews = {}; render(); return; }
    if (action === 'close-preview-modal') { state.previewModal = null; render(); return; }
    if (action === 'save-assignment') { saveAssignment(actionEl); return; }
    if (action === 'reset-filter') { state.filters = { year: CURRENT_YEAR, search: '', unit: '', status: '' }; render(); return; }
    if (action === 'calendar-prev') { changeCalendarMonth(-1); return; }
    if (action === 'calendar-next') { changeCalendarMonth(1); return; }
    if (action === 'select-all-employees') {
      document.querySelectorAll('[data-assignment-employee]').forEach((input) => { if (input.closest('[data-employee-search-text]')?.style.display !== 'none') input.checked = true; });
      updateEmployeeSummary();
      return;
    }
    if (action === 'clear-employees') {
      document.querySelectorAll('[data-assignment-employee]').forEach((input) => { input.checked = false; });
      updateEmployeeSummary();
      return;
    }

    const detailId = event.target.closest?.('[data-assignment-detail]')?.dataset.assignmentDetail;
    if (detailId) { openAssignmentDetail(detailId); return; }
    const editId = event.target.closest?.('[data-assignment-edit]')?.dataset.assignmentEdit;
    if (editId) { openAssignmentModal(editId); return; }
    const deleteId = event.target.closest?.('[data-assignment-delete]')?.dataset.assignmentDelete;
    if (deleteId) { deleteAssignment(deleteId); return; }
    const download = event.target.closest?.('[data-download-assignment-document]');
    if (download) { downloadAssignmentDocument(download.dataset.assignmentId, download.dataset.downloadAssignmentDocument, download); return; }
    const preview = event.target.closest?.('[data-preview-assignment-document]');
    if (preview) { previewAssignmentDocument(preview.dataset.assignmentId, preview.dataset.previewAssignmentDocument); }
  });

  window.addEventListener('hashchange', () => {
    state.view = viewFromHash();
    state.modal = null;
    state.detailId = '';
    state.previewModal = null;
    state.inlinePreviews = {};
    render();
  });

  window.addEventListener('beforeunload', () => {
    if (window.__assignmentLocalPreviewUrl) URL.revokeObjectURL(window.__assignmentLocalPreviewUrl);
  });

  async function init() {
    await window.FirebaseClient.requireAdmin();
    state.view = viewFromHash();
    if (!location.hash) history.replaceState(null, '', '#dashboard');
    await loadData();
  }

  init();
})();
