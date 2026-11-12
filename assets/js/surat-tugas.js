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
    kalender: 'Kalender',
    permohonan: 'Permohonan Perubahan ST'
  });
  const ASSIGNMENT_TYPES = Object.freeze({
    official_travel: 'Dinas Luar / Perjalanan Dinas',
    work_arrangement: 'Pengaturan Kerja',
    other: 'Lainnya'
  });

  function normalizeAssignmentType(value, legacyAdjustmentType) {
    const key = String(value || '').trim();
    if (Object.prototype.hasOwnProperty.call(ASSIGNMENT_TYPES, key)) return key;
    // Data Surat Tugas lama belum memiliki assignmentType dan seluruhnya dahulu
    // diperlakukan sebagai dinas yang menghasilkan penyesuaian otomatis 0%.
    if (!key && String(legacyAdjustmentType || '') === 'official_duty') return 'official_travel';
    return 'official_travel';
  }

  function assignmentTypeLabel(value) {
    return ASSIGNMENT_TYPES[normalizeAssignmentType(value)] || ASSIGNMENT_TYPES.official_travel;
  }

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
      month: CURRENT_MONTH,
      unit: '',
      type: ''
    },
    calendarMode: 'calendar',
    expandedCalendarDays: new Set(),
    modal: null,
    detailId: '',
    previewModal: null,
    inlinePreviews: {},
    requests: { rows: [], loaded: false, loading: false, error: '', updatedAt: null,
      filter: 'all', search: '', detailId: '', unread: new Set(), known: new Set(),
      fileName: '', worksheet: '', info: '', updatingId: '', updateError: '' }
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


  // Gunakan createdAt (waktu pencatatan awal), bukan tanggal surat/periode.
  // Data lama tanpa createdAt tetap kosong; jangan menebak dari updatedAt.
  function inputTimestamp(value) {
    if (!value) return null;
    let date;
    if (typeof value.toDate === 'function') date = value.toDate();
    else if (typeof value.seconds === 'number') date = new Date(value.seconds * 1000 + Math.floor(Number(value.nanoseconds || 0) / 1e6));
    else if (typeof value._seconds === 'number') date = new Date(value._seconds * 1000 + Math.floor(Number(value._nanoseconds || 0) / 1e6));
    else date = value instanceof Date ? value : new Date(value);
    return date instanceof Date && Number.isFinite(date.getTime()) ? date : null;
  }

  function inputDateSortValue(value) {
    return inputTimestamp(value)?.toISOString() || '';
  }

  function formatInputDate(value) {
    const date = inputTimestamp(value);
    if (!date) return '—';
    return new Intl.DateTimeFormat('id-ID', {
      day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).format(date);
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
      assignmentType: normalizeAssignmentType(row.assignmentType, row.adjustmentType),
      accountable: row.accountable !== false,
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
          item.letterNumber, item.activity, item.location, assignmentTypeLabel(item.assignmentType),
          ...people.flatMap((person) => [person.name, person.nip, person.unit])
        ].join(' ').toLowerCase();
        return haystack.includes(query);
      })
      .sort((a, b) => b.startDate.localeCompare(a.startDate) || b.letterDate.localeCompare(a.letterDate));
  }

  function peopleList(item) {
    const people = assignmentPeople(item);
    if (!people.length) return '<span class="assignment-sub">Belum ada pegawai</span>';
    return `<div class="assignment-people-cell">${people.map((person) => `<div class="assignment-person">${esc(person.name || 'Pegawai')}</div>`).join('')}</div>`;
  }

  function assignmentUnitList(item) {
    const units = [...new Set(assignmentPeople(item).map((person) => person.unit || 'Belum diisi'))];
    return `<div class="assignment-units-cell">${units.map((unit) => `<div>${esc(unit)}</div>`).join('')}</div>`;
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
      <div class="assignment-head"><div><h1>Surat Tugas</h1><p>Arsip penugasan pegawai yang terhubung dengan proses Tunjangan Kinerja.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <div class="assignment-metrics">${metric('Surat Tugas', yearRows.length, `Tahun ${CURRENT_YEAR}`)}${metric('Sedang Berjalan', active, 'Berdasarkan tanggal hari ini')}${metric('Akan Datang', upcoming, 'Penugasan yang belum dimulai')}${metric('Pegawai Ditugaskan', uniqueEmployeeIds.size, 'Pegawai unik pada tahun berjalan')}</div>
      <section class="card assignment-card"><div class="assignment-card-head"><div><h2>Penugasan Terdekat</h2><p>Surat Tugas berjalan dan akan datang.</p></div><a class="btn btn-secondary btn-sm" href="surat-tugas.html#data">Lihat Semua</a></div>
        ${upcomingRows.length ? `<div class="assignment-upcoming-list">${upcomingRows.map((item) => `<article class="assignment-upcoming"><strong>${esc(item.activity || item.letterNumber || 'Surat Tugas')}</strong><span>${esc(formatDateRange(item))} · ${esc(item.location || 'Lokasi belum diisi')}</span><span>${esc(assignmentPeople(item).map((person) => person.name).filter(Boolean).slice(0, 3).join(', '))}${assignmentPeople(item).length > 3 ? ` +${assignmentPeople(item).length - 3}` : ''}</span><div style="margin-top:8px"><button class="btn btn-secondary btn-sm" type="button" data-assignment-detail="${esc(item.id)}">Detail</button></div></article>`).join('')}</div>` : '<div class="assignment-empty"><strong>Belum ada penugasan</strong>Tambahkan Surat Tugas untuk mulai membangun arsip penugasan.</div>'}
      </section>
    </div>`;
  }

  function dataView() {
    const rows = filteredAssignments();
    return `<div class="assignment-page">
      <div class="assignment-head"><div><h1>Data Surat Tugas</h1><p>Seluruh nama pegawai ditampilkan, dengan unit/bidang pada kolom terpisah. Klik Tanggal Input untuk mengurutkan berdasarkan waktu pencatatan.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <section class="card assignment-filter-card"><div class="assignment-filter-grid">
        <div class="search"><input id="assignment-filter-search" type="search" value="${esc(state.filters.search)}" placeholder="Cari nomor ST, kegiatan, pegawai, NIP, lokasi"></div>
        <div class="field"><label>Tahun</label><select id="assignment-filter-year">${yearOptions(state.filters.year)}</select></div>
        <div class="field"><label>Unit/Bidang</label>${unitOptions(state.filters.unit, true)}</div>
        <div class="field"><label>Status</label><select id="assignment-filter-status"><option value="">Semua Status</option><option value="upcoming" ${state.filters.status === 'upcoming' ? 'selected' : ''}>Akan datang</option><option value="active" ${state.filters.status === 'active' ? 'selected' : ''}>Berjalan</option><option value="done" ${state.filters.status === 'done' ? 'selected' : ''}>Selesai</option></select></div>
        <button class="btn btn-secondary" type="button" data-action="reset-filter">Reset</button>
      </div></section>
      <section class="card assignment-card"><div class="assignment-table-wrap"><table id="assignment-data-table" data-table-search="external" class="data-table assignment-table"><thead><tr><th>No</th><th>Pegawai</th><th>Unit/Bidang</th><th>Nomor Surat</th><th>Tanggal Surat</th><th>Periode Tugas</th><th data-sort-default="desc">Tanggal Input</th><th>Kegiatan</th><th>Lokasi</th><th>Status</th><th>Dokumen</th><th>Aksi</th></tr></thead><tbody>${rows.length ? rows.map((item, index) => {
        const status = statusOf(item);
        const doc = primaryDocument(item);
        return `<tr><td>${index + 1}</td><td>${peopleList(item)}</td><td>${assignmentUnitList(item)}</td><td><strong>${esc(item.letterNumber || '—')}</strong></td><td class="nowrap">${esc(formatDate(item.letterDate))}</td><td class="nowrap">${esc(formatDateRange(item))}</td><td class="nowrap" data-sort-value="${esc(inputDateSortValue(item.createdAt))}" title="${item.createdAt ? 'Waktu data pertama kali disimpan' : 'Waktu input tidak tersedia pada data lama'}">${esc(formatInputDate(item.createdAt))}</td><td><strong>${esc(item.activity || '—')}</strong><div class="assignment-sub">${esc(assignmentTypeLabel(item.assignmentType))}${item.note ? ` · ${esc(item.note)}` : ''}</div></td><td>${esc(item.location || '—')}</td><td><span class="assignment-pill ${status.key}">${esc(status.label)}</span></td><td>${doc?.file?.itemId ? `<span class="assignment-pill">Tersimpan</span><div class="assignment-sub">${esc(doc.file.name || doc.file.remoteName || 'Surat Tugas')}</div>` : '<span class="assignment-sub">Belum ada file</span>'}</td><td><div class="assignment-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-assignment-detail="${esc(item.id)}">Detail</button><button class="btn btn-secondary btn-sm" type="button" data-assignment-edit="${esc(item.id)}">Edit</button></div></td></tr>`;
      }).join('') : '<tr><td colspan="12"><div class="assignment-empty"><strong>Tidak ada data</strong>Tidak ada Surat Tugas yang sesuai filter.</div></td></tr>'}</tbody></table></div></section>
    </div>`;
  }

  function calendarPeople(item) {
    return assignmentPeople(item).filter((person) => !state.calendar.unit || person.unit === state.calendar.unit);
  }

  function calendarRows() {
    const year = Number(state.calendar.year);
    const month = Number(state.calendar.month);
    const start = `${year}-${String(month).padStart(2, '0')}-01`;
    const end = `${year}-${String(month).padStart(2, '0')}-${String(new Date(year, month, 0).getDate()).padStart(2, '0')}`;
    return state.assignments
      .filter((item) => item.startDate && item.startDate <= end && (item.endDate || item.startDate) >= start)
      .filter((item) => !state.calendar.unit || calendarPeople(item).length > 0)
      .filter((item) => !state.calendar.type || normalizeAssignmentType(item.assignmentType) === state.calendar.type)
      .sort((a, b) => a.startDate.localeCompare(b.startDate) || a.letterNumber.localeCompare(b.letterNumber));
  }

  function calendarYearOptions() {
    const current = Number(state.calendar.year);
    const years = new Set([current, CURRENT_YEAR - 2, CURRENT_YEAR - 1, CURRENT_YEAR, CURRENT_YEAR + 1, CURRENT_YEAR + 2]);
    state.assignments.forEach((item) => {
      if (item.startDate) years.add(Number(item.startDate.slice(0, 4)));
      if (item.endDate) years.add(Number(item.endDate.slice(0, 4)));
    });
    return [...years].filter(Number.isFinite).sort((a, b) => b - a)
      .map((year) => `<option value="${year}" ${year === current ? 'selected' : ''}>${year}</option>`).join('');
  }

  function calendarGrid(rows) {
    const year = Number(state.calendar.year);
    const month = Number(state.calendar.month);
    const first = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0).getDate();
    const mondayOffset = (first.getDay() + 6) % 7;
    const cells = [];
    for (let i = 0; i < mondayOffset; i += 1) {
      cells.push('<div class="assignment-calendar-cell muted" aria-hidden="true"></div>');
    }
    for (let day = 1; day <= lastDay; day += 1) {
      const iso = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const date = new Date(year, month - 1, day);
      const events = rows.filter((item) => item.startDate <= iso && item.endDate >= iso)
        .flatMap((item) => {
          const people = calendarPeople(item);
          return (people.length ? people : [{ name: 'Pegawai belum diisi' }])
            .map((person) => ({ item, person }));
        });
      const expanded = state.expandedCalendarDays.has(iso);
      const visible = expanded ? events : events.slice(0, 4);
      const isToday = iso === TODAY;
      const weekend = date.getDay() === 0 || date.getDay() === 6;
      cells.push(`<div class="assignment-calendar-cell ${weekend ? 'weekend' : ''} ${isToday ? 'today' : ''}">
        <div class="assignment-calendar-day">${day}${isToday ? '<span>Hari ini</span>' : ''}</div>
        <div class="assignment-calendar-events">
          ${visible.map(({ item, person }) => `<button type="button" data-assignment-detail="${esc(item.id)}" title="${esc(`${person.name || 'Pegawai'} — ${item.activity || item.letterNumber || 'Surat Tugas'} (${formatDateRange(item)})`)}"><strong>${esc(person.name || 'Pegawai')}</strong><span>${esc(item.activity || item.letterNumber || 'Surat Tugas')}</span></button>`).join('')}
          ${events.length > 4 ? `<button type="button" class="assignment-calendar-more" data-action="calendar-toggle-day" data-calendar-day="${iso}" aria-expanded="${expanded}">${expanded ? 'Tampilkan lebih sedikit' : `+${events.length - 4} pegawai lainnya`}</button>` : ''}
        </div>
      </div>`);
    }
    while (cells.length % 7) cells.push('<div class="assignment-calendar-cell muted" aria-hidden="true"></div>');
    return `<div class="assignment-calendar"><div class="assignment-calendar-weekdays">${['Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab', 'Min'].map((day) => `<span>${day}</span>`).join('')}</div><div class="assignment-calendar-grid">${cells.join('')}</div></div>`;
  }

  function calendarList(rows) {
    return `<div class="assignment-calendar-list">${rows.length ? rows.map((item) => `<button class="assignment-calendar-row" type="button" data-assignment-detail="${esc(item.id)}"><div class="date">${esc(formatDateRange(item))}</div><div class="activity"><strong>${esc(item.activity || item.letterNumber || 'Surat Tugas')}</strong><span>${esc(item.letterNumber || 'Tanpa nomor')}</span></div><div>${esc(calendarPeople(item).map((person) => person.name).filter(Boolean).join(', ') || 'Pegawai belum diisi')}</div><div><span class="assignment-pill ${statusOf(item).key}">${esc(statusOf(item).label)}</span></div></button>`).join('') : '<div class="assignment-empty"><strong>Tidak ada penugasan</strong>Belum ada Surat Tugas pada bulan dan filter yang dipilih.</div>'}</div>`;
  }

  function calendarView() {
    const rows = calendarRows();
    const cal = state.calendar;
    const units = [...new Set([...uniqueUnits(), ...state.assignments.flatMap((item) => assignmentPeople(item).map((person) => person.unit))].filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, 'id'));
    const monthOptions = Array.from({ length: 12 }, (_, index) => {
      const month = index + 1;
      return `<option value="${month}" ${month === cal.month ? 'selected' : ''}>${esc(monthName(month))}</option>`;
    }).join('');
    return `<div class="assignment-page">
      <div class="assignment-head"><div><h1>Kalender Surat Tugas</h1><p>Pantau penugasan pada kalender bulanan. Nama pegawai ditampilkan pada setiap tanggal selama periode Surat Tugas.</p></div><div class="assignment-head-actions"><button class="btn btn-primary" type="button" data-action="add-assignment">+ Tambah Surat Tugas</button></div></div>
      <section class="card assignment-calendar-filter-card"><div class="assignment-calendar-filters">
        <div class="field"><label for="assignment-calendar-month">Bulan</label><select id="assignment-calendar-month">${monthOptions}</select></div>
        <div class="field"><label for="assignment-calendar-year">Tahun</label><select id="assignment-calendar-year">${calendarYearOptions()}</select></div>
        <div class="field"><label for="assignment-calendar-unit">Unit/Bidang</label><select id="assignment-calendar-unit"><option value="">Semua Unit/Bidang</option>${units.map((unit) => `<option value="${esc(unit)}" ${cal.unit === unit ? 'selected' : ''}>${esc(unit)}</option>`).join('')}</select></div>
        <div class="field"><label for="assignment-calendar-type">Jenis Penugasan</label><select id="assignment-calendar-type"><option value="">Semua Jenis</option>${Object.entries(ASSIGNMENT_TYPES).map(([key, label]) => `<option value="${esc(key)}" ${cal.type === key ? 'selected' : ''}>${esc(label)}</option>`).join('')}</select></div>
        <div class="assignment-calendar-view-toggle"><button type="button" class="btn btn-sm ${state.calendarMode === 'list' ? 'btn-primary' : 'btn-secondary'}" data-assignment-calendar-mode="list" aria-pressed="${state.calendarMode === 'list'}">Daftar</button><button type="button" class="btn btn-sm ${state.calendarMode === 'calendar' ? 'btn-primary' : 'btn-secondary'}" data-assignment-calendar-mode="calendar" aria-pressed="${state.calendarMode === 'calendar'}">Kalender</button></div>
      </div></section>
      <section class="card assignment-card"><div class="assignment-card-head assignment-calendar-toolbar"><div><h2>${esc(monthName(cal.month))} ${esc(cal.year)}</h2><p>${rows.length} Surat Tugas pada bulan dan filter aktif. Klik nama pegawai untuk melihat detail.</p></div><div class="assignment-calendar-nav"><button class="btn btn-secondary btn-sm" type="button" data-action="calendar-prev" aria-label="Bulan sebelumnya">←</button><strong>${esc(monthName(cal.month))}</strong><button class="btn btn-secondary btn-sm" type="button" data-action="calendar-next" aria-label="Bulan berikutnya">→</button></div></div>
        ${state.calendarMode === 'calendar' ? calendarGrid(rows) : calendarList(rows)}
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
      <div class="assignment-modal-head"><div><h3>${item.id ? 'Edit Surat Tugas' : 'Tambah Surat Tugas'}</h3><p>Upload surat, pilih satu atau lebih pegawai, kemudian isi rentang tanggal penugasan. Data Surat Tugas akan menjadi data pendukung pada proses Tukin.</p></div><button class="icon-btn" type="button" data-action="close-assignment-modal">×</button></div>
      <div class="assignment-form-grid">
        <div class="field"><label for="assignment-letter-number">Nomor Surat Tugas</label><input id="assignment-letter-number" type="text" value="${esc(item.letterNumber || '')}" placeholder="Contoh: ST/123/Pusdatin/2026"></div>
        <div class="field"><label for="assignment-letter-date">Tanggal Surat</label><input id="assignment-letter-date" type="date" value="${esc(item.letterDate || '')}"></div>
        <div class="field"><label for="assignment-start">Tanggal Mulai Tugas</label><input id="assignment-start" type="date" value="${esc(item.startDate || '')}"></div>
        <div class="field"><label for="assignment-end">Tanggal Selesai Tugas</label><input id="assignment-end" type="date" value="${esc(item.endDate || '')}"></div>
        <div class="field span-2"><label>Pegawai yang Ditugaskan</label>${employeePickerMarkup()}<div class="assignment-field-note">Satu Surat Tugas cukup dicatat satu kali walaupun memuat banyak pegawai.</div></div>
        <div class="field span-2"><label for="assignment-activity">Kegiatan / Keperluan</label><input id="assignment-activity" type="text" value="${esc(item.activity || '')}" placeholder="Contoh: Rapat Koordinasi ..."></div>
        <div class="field"><label for="assignment-location">Lokasi / Tujuan</label><input id="assignment-location" type="text" value="${esc(item.location || '')}" placeholder="Opsional"></div>
        <div class="field"><label for="assignment-type">Jenis Penugasan</label><select id="assignment-type"><option value="official_travel" ${normalizeAssignmentType(item.assignmentType, item.adjustmentType) === 'official_travel' ? 'selected' : ''}>Dinas Luar / Perjalanan Dinas</option><option value="work_arrangement" ${normalizeAssignmentType(item.assignmentType, item.adjustmentType) === 'work_arrangement' ? 'selected' : ''}>Pengaturan Kerja</option><option value="other" ${normalizeAssignmentType(item.assignmentType, item.adjustmentType) === 'other' ? 'selected' : ''}>Lainnya</option></select><div class="assignment-field-note">Dinas Luar / Perjalanan Dinas dapat menyesuaikan potongan menjadi 0% secara otomatis. Pengaturan Kerja dan Lainnya hanya menjadi data pendukung dan tidak mengubah persentase otomatis.</div></div>
        <div class="field span-2"><label class="assignment-accountable"><input type="checkbox" id="assignment-accountable" ${item.accountable!==false?'checked':''}> Surat Tugas dipertanggungjawabkan untuk administrasi kehadiran/Tukin</label><div class="assignment-field-note">Jika tidak dicentang, Surat Tugas tetap terlihat pada pratinjau, tetapi tidak dihitung sebagai dasar penyesuaian Tukin dan tidak dimasukkan dalam ZIP bukti dukung. Pegawai tetap wajib presensi sesuai ketentuan.</div></div><div class="field span-2"><label for="assignment-note">Keterangan</label><textarea id="assignment-note" rows="3" placeholder="Opsional">${esc(item.note || '')}</textarea></div>
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
      <div class="assignment-section"><div class="assignment-section-title">Informasi Surat Tugas</div><div class="assignment-detail-grid"><div><span>Nomor Surat</span><strong>${esc(item.letterNumber || '—')}</strong></div><div><span>Tanggal Surat</span><strong>${esc(formatDate(item.letterDate))}</strong></div><div><span>Status</span><strong>${esc(status.label)}</strong></div><div><span>Jenis Penugasan</span><strong>${esc(assignmentTypeLabel(item.assignmentType))}</strong></div><div><span>Dipertanggungjawabkan</span><strong>${item.accountable===false?'Tidak — hanya pratinjau':'Ya'}</strong></div><div><span>Tanggal Mulai</span><strong>${esc(formatDate(item.startDate))}</strong></div><div><span>Tanggal Selesai</span><strong>${esc(formatDate(item.endDate))}</strong></div><div class="span-2"><span>Lokasi / Tujuan</span><strong>${esc(item.location || '—')}</strong></div><div class="span-all"><span>Kegiatan / Keperluan</span><strong>${esc(item.activity || '—')}</strong></div><div class="span-all"><span>Keterangan</span><strong>${esc(item.note || '—')}</strong></div></div></div>
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

  // Permohonan Perubahan ST: Refresh manual dari Excel; status disimpan ke SharePoint
  // melalui Microsoft Graph saat petugas mengonfirmasi penyelesaian.
  // Penanda BARU tetap lokal per browser, status permohonan bersumber dari Excel.
  // Sumber SharePoint ditetapkan administrator pada file konfigurasi; tidak
  // menerima override dari localStorage atau isian URL pengguna.
  function requestWorkbookUrl() {
    return String(window.ST_CHANGE_REQUESTS_CONFIG?.shareUrl || '').trim();
  }

  function requestStorageKey() {
    // Keep unread markers separate for each workbook; Forms response IDs are
    // only unique within their form, not across different ticket workbooks.
    const source = requestWorkbookUrl();
    let hash = 2166136261;
    for (let i = 0; i < source.length; i += 1) {
      hash ^= source.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return `pkp-st-request-read-v3:${String(user()?.uid || 'anon')}:${(hash >>> 0).toString(36)}`;
  }

  function readRequestMarkers() {
    try {
      const data = JSON.parse(localStorage.getItem(requestStorageKey()) || '{}');
      state.requests.known = new Set(Array.isArray(data.known) ? data.known.map(String) : []);
      state.requests.unread = new Set(Array.isArray(data.unread) ? data.unread.map(String) : []);
    } catch (_) {
      state.requests.known = new Set();
      state.requests.unread = new Set();
    }
    window.STRequestMenuBadgeCount = state.requests.unread.size;
  }

  function saveRequestMarkers() {
    window.STRequestMenuBadgeCount = state.requests.unread.size;
    try {
      localStorage.setItem(requestStorageKey(), JSON.stringify({
        known: [...state.requests.known].slice(-5000),
        unread: [...state.requests.unread].slice(-5000)
      }));
    } catch (error) { console.warn('Penanda permohonan tidak dapat disimpan di browser:', error); }
  }

  function requestDone(item) {
    return /^(selesai|ditolak|dibatalkan|batal|closed|resolved)$/i.test(String(item?.status || '').trim());
  }

  function requestStatusClass(item) {
    if (requestDone(item)) return 'st-req-done';
    if (/proses|verifikasi|tindak lanjut/i.test(String(item.status))) return 'st-req-process';
    return 'st-req-pending';
  }

  function requestStatistics() {
    const rows = state.requests.rows;
    return {
      total: rows.length,
      pending: rows.filter(row => !requestDone(row)).length,
      done: rows.filter(requestDone).length,
      newCount: rows.filter(row => state.requests.unread.has(row.id)).length
    };
  }

  function requestRowsFiltered() {
    const needle = state.requests.search.trim().toLowerCase();
    const filter = state.requests.filter;
    return state.requests.rows.filter(row => {
      if (filter === 'unread' && !state.requests.unread.has(row.id)) return false;
      if (filter === 'pending' && requestDone(row)) return false;
      if (filter === 'done' && !requestDone(row)) return false;
      if (!needle) return true;
      return [row.name,row.id,row.letter,row.description,row.type,row.status].some(value => String(value || '').toLowerCase().includes(needle));
    });
  }

  function requestView() {
    const info = state.requests;
    const summary = requestStatistics();
    const filtered = requestRowsFiltered();
    const sourceUrl = requestWorkbookUrl();
    const sourceLink = /^https:\/\//i.test(sourceUrl) ? `<a class="btn btn-secondary btn-sm" href="${esc(sourceUrl)}" target="_blank" rel="noopener noreferrer">Buka Excel SharePoint ↗</a>` : '';
    const recent = info.updatedAt ? new Intl.DateTimeFormat('id-ID', { day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit' }).format(info.updatedAt) : 'Belum pernah disinkronkan dalam sesi ini';
    const rows = filtered.map(row => {
      const unread = info.unread.has(row.id);
      return `<tr class="${unread ? 'st-req-unread-row' : ''}">
        <td data-label="Diterima"><span class="st-req-small">${esc(row.completed || '—')}</span></td>
        <td data-label="Pemohon"><strong>${esc(row.name)}</strong><small>#${esc(row.id)} · ${esc(row.type)}</small></td>
        <td data-label="Nomor Surat"><span class="st-req-letter">${esc(row.letter)}</span></td>
        <td data-label="Uraian"><span class="st-req-description">${esc(row.description || '—')}</span></td>
        <td data-label="Status"><span class="st-req-pill ${requestStatusClass(row)}">${esc(row.status)}</span>${unread ? '<span class="st-req-new-mark">BARU</span>' : ''}</td>
        <td data-label="Tindakan"><button type="button" class="btn btn-secondary btn-sm" data-action="st-request-detail" data-request-id="${esc(row.id)}">Detail</button></td>
      </tr>`;
    }).join('');
    return `<section class="assignment-page st-request-page">
      <div class="assignment-head st-request-header"><div><h1>Permohonan Perubahan Surat Tugas</h1><p>Data diambil dari sheet Tiket pada workbook respons Microsoft Forms yang sama. Power Automate mengisi tabel TblPermohonanST; klik Refresh untuk memperbarui.</p></div>
        <div class="assignment-head-actions">${sourceLink}<button class="btn btn-primary" type="button" data-action="st-request-refresh" ${info.loading || info.updatingId ? 'disabled aria-busy="true"' : ''}>${info.loading ? 'Mengambil data…' : '↻ Refresh dari SharePoint'}</button></div>
      </div>
      
      <div class="st-req-hint">${info.loading ? 'Sedang membaca Excel SharePoint melalui akun Microsoft Anda…' : `Terakhir diperiksa: <strong>${esc(recent)}</strong>`}${info.fileName ? ` · ${esc(info.fileName)} / ${esc(info.worksheet)}` : ''} <span>Sumber: sheet Tiket / tabel TblPermohonanST (bukan Sheet1). Penyelesaian tiket mengubah kolom Status pada sheet Tiket.</span></div>
      ${info.error ? `<div class="alert alert-warning" role="alert"><div class="alert-title">Gagal mengambil respons Microsoft Forms</div>${esc(info.error)}<div class="st-req-error-help">Pastikan Anda login dengan akun Microsoft yang diberi akses ke file dan file tetap berada di OneDrive/SharePoint organisasi.</div></div>` : ''}
      ${info.info ? `<div class="st-req-message" role="status">${esc(info.info)}</div>` : ''}
      <div class="st-req-metrics">
        ${metric('Total Permohonan',summary.total,'Respons dengan Completion time')}
        ${metric('Belum Selesai',summary.pending,'Status dari Excel')}
        ${metric('Selesai / Ditutup',summary.done,'Status dari Excel')}
        ${metric('Baru / Belum Dibaca',summary.newCount,'Penanda lokal browser')}
      </div>
      <section class="card st-req-list-card">
        <div class="st-req-list-head"><div><h2>Daftar Permohonan</h2><p>${info.loaded ? `${filtered.length} dari ${summary.total} permohonan ditampilkan` : 'Klik Refresh untuk membaca data terbaru dari Excel SharePoint.'}</p></div></div>
        <div class="st-req-controls"><label class="st-req-search-field">Cari permohonan<input id="st-request-search" type="search" placeholder="Nama pegawai, nomor ST, atau uraian..." value="${esc(info.search)}" autocomplete="off"></label>
          <label class="st-req-filter-field">Status<select id="st-request-status-filter"><option value="all" ${info.filter === 'all' ? 'selected' : ''}>Semua</option><option value="unread" ${info.filter === 'unread' ? 'selected' : ''}>Baru / Belum Dibaca</option><option value="pending" ${info.filter === 'pending' ? 'selected' : ''}>Belum Selesai</option><option value="done" ${info.filter === 'done' ? 'selected' : ''}>Selesai / Ditutup</option></select></label></div>
        <div class="st-req-table-scroll"><table class="st-req-table"><thead><tr><th>Diterima</th><th>Pemohon</th><th>Nomor Surat/Perihal</th><th>Deskripsi Permohonan</th><th>Status Excel</th><th>Tindakan</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="6" class="st-req-empty">${info.loaded ? 'Tidak ada permohonan yang sesuai filter.' : 'Belum ada data. Gunakan tombol Refresh dari SharePoint.'}</td></tr>`}</tbody></table></div>
      </section>
    </section>`;
  }

  function requestDetailModal() {
    const id = state.requests.detailId;
    if (!id || state.view !== 'permohonan') return '';
    const item = state.requests.rows.find(row => row.id === id);
    if (!item) return '';
    const isPending = !requestDone(item);
    const isSaving = state.requests.updatingId === item.id;
    const candidates = state.assignments.filter(row => String(row.letterNumber || '').replace(/[^a-z0-9]/gi,'').toLowerCase() === String(item.letter || '').replace(/[^a-z0-9]/gi,'').toLowerCase());
    const assignments = window.UserAccess.canView('surat','data') && candidates.length
      ? `<div class="st-req-detail-field"><span>Surat Tugas terkait di aplikasi</span>${candidates.map(row => `<button type="button" class="btn btn-secondary btn-sm" data-assignment-detail="${esc(row.id)}">${esc(row.letterNumber)} · Lihat Surat Tugas</button>`).join(' ')}</div>` : '';
    return `<div class="assignment-modal-backdrop st-request-detail-backdrop" data-action="st-request-close"><section class="assignment-modal st-request-detail-modal" role="dialog" aria-modal="true" aria-labelledby="st-request-detail-heading">
      <div class="assignment-modal-head"><div><h3 id="st-request-detail-heading">Permohonan #${esc(item.id)}</h3><p>Sheet Tiket pada workbook Forms di SharePoint (status dapat diperbarui)</p></div><button class="icon-btn" type="button" data-action="st-request-close" aria-label="Tutup">×</button></div>
      <div class="st-req-detail-body">
        <div class="st-req-detail-grid"><div class="st-req-detail-field"><span>Nama Pemohon</span><strong>${esc(item.name)}</strong></div><div class="st-req-detail-field"><span>Waktu Pengiriman</span><strong>${esc(item.completed || '—')}</strong></div>
        <div class="st-req-detail-field"><span>Jenis Permohonan</span><strong>${esc(item.type)}</strong></div><div class="st-req-detail-field"><span>Status pada Excel</span><strong><span class="st-req-pill ${requestStatusClass(item)}">${esc(item.status)}</span></strong></div></div>
        <div class="st-req-detail-field"><span>Nomor Surat/Perihal</span><strong>${esc(item.letter)}</strong></div>
        <div class="st-req-detail-field"><span>Rincian Permohonan</span><div class="st-req-detail-description">${esc(item.description || 'Tidak diisi.')}</div></div>
        ${assignments}
        <div class="st-req-detail-footnote">Tombol Tandai Sudah Dibaca sekaligus menyimpan <strong>Selesai</strong> ke kolom Status di Excel SharePoint. Gunakan hanya setelah permohonan ditindaklanjuti.</div>
        ${state.requests.updateError ? `<div class="alert alert-warning" role="alert"><div class="alert-title">Gagal menyimpan status</div>${esc(state.requests.updateError)}</div>` : ''}
      </div>
      <div class="st-req-detail-actions">${isPending ? `<button type="button" class="btn btn-primary" data-action="st-request-read" data-request-id="${esc(item.id)}" ${state.requests.updatingId ? 'disabled aria-busy="true"' : ''}>${isSaving ? 'Menyimpan ke Excel…' : 'Tandai Sudah Dibaca & Selesai'}</button>` : ''}<button type="button" class="btn btn-secondary" data-action="st-request-close" ${isSaving ? 'disabled' : ''}>Tutup</button></div>
    </section></div>`;
  }

  async function refreshRequests() {
    if (state.requests.loading || state.requests.updatingId || !window.UserAccess.canView('surat','permohonan')) return;
    state.requests.loading = true;
    state.requests.error = '';
    state.requests.info = '';
    render();
    try {
      const result = await window.STChangeRequests.fetchRequests();
      const rows = result.rows;
      const existing = state.requests.known;
      const newPending = rows.filter(row => !existing.has(row.id) && !requestDone(row));
      for (const row of rows) existing.add(row.id);
      for (const row of newPending) state.requests.unread.add(row.id);
      // Respons yang sudah ditutup tidak lagi memunculkan penanda tindakan baru.
      for (const row of rows) if (requestDone(row)) state.requests.unread.delete(row.id);
      saveRequestMarkers();
      state.requests.rows = rows;
      state.requests.loaded = true;
      state.requests.updatedAt = new Date();
      state.requests.fileName = result.fileName;
      state.requests.worksheet = result.worksheet;
      state.requests.info = newPending.length
        ? `${newPending.length} permohonan baru yang belum selesai ditemukan.`
        : `Sinkronisasi berhasil. ${rows.length} respons lengkap ditemukan; tidak ada permohonan baru yang belum selesai.`;
    } catch (error) {
      state.requests.error = error?.message || String(error);
    } finally {
      state.requests.loading = false;
      render();
    }
  }

  async function markRequestRead(id) {
    if (state.requests.loading || state.requests.updatingId || !window.UserAccess.canView('surat','permohonan')) return;
    const item = state.requests.rows.find(row => row.id === id);
    if (!item || requestDone(item)) return;
    const confirmation = `Tandai permohonan #${id} dari ${item.name} sebagai Selesai?\n\nTindakan ini akan MENULIS \"Selesai\" ke kolom Status pada Excel SharePoint, bukan hanya menghilangkan notifikasi.`;
    if (!window.confirm(confirmation)) return;
    state.requests.updatingId = id;
    state.requests.updateError = '';
    render();
    try {
      const result = await window.STChangeRequests.completeRequest(id);
      item.status = result.status;
      state.requests.known.add(id);
      state.requests.unread.delete(id);
      saveRequestMarkers();
      state.requests.detailId = '';
      state.requests.info = result.alreadyDone
        ? `Permohonan #${id} sudah berstatus Selesai di Excel SharePoint.`
        : `Status permohonan #${id} berhasil disimpan sebagai Selesai pada Excel SharePoint.`;
    } catch (error) {
      state.requests.updateError = error?.message || String(error);
    } finally {
      state.requests.updatingId = '';
      render();
    }
  }


  function overlays() {
    return `${assignmentDetail()}${assignmentModal()}${previewModal()}${requestDetailModal()}${state.toast ? `<div class="assignment-toast">${esc(state.toast)}</div>` : ''}`;
  }

  function viewMarkup() {
    if (state.view === 'permohonan') return requestView();
    if (state.view === 'data') return dataView();
    if (state.view === 'kalender') return calendarView();
    return dashboardView();
  }

  function render() {
    document.body.classList.toggle('wide-records', state.view === 'data' || state.view === 'permohonan');
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
      if (!['dashboard', 'data', 'kalender'].some(view => window.UserAccess.canView('surat', view))) {
        state.employees = []; state.directory = new Map(); state.assignments = [];
        return;
      }
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
    if (!window.UserAccess.assert('surat',id?'data.edit':'data.create')) return;
    const existing = id ? state.assignments.find((item) => item.id === id) : null;
    const selected = new Set((existing?.employees || []).map((employee) => employee.employeeId));
    state.detailId = '';
    state.previewModal = null;
    state.modal = {
      record: existing ? { ...existing, employees: (existing.employees || []).map((employee) => ({ ...employee })), documents: (existing.documents || []).map((doc) => ({ ...doc, file: doc.file ? { ...doc.file } : null })) } : { letterNumber: '', letterDate: TODAY, startDate: '', endDate: '', assignmentType: 'official_travel', activity: '', location: '', note: '', documents: [] },
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
    if (!window.UserAccess.assert('surat',state.modal?.record?.id?'data.edit':'data.create')) return;
    if (document.getElementById('assignment-file')?.files?.length && !window.UserAccess.assert('surat','documents.upload')) return;
    if (!state.modal) return;
    const selectedIds = selectedEmployeeIdsFromForm();
    const letterNumber = String(document.getElementById('assignment-letter-number')?.value || '').trim();
    const letterDate = String(document.getElementById('assignment-letter-date')?.value || '');
    const startDate = String(document.getElementById('assignment-start')?.value || '');
    const endDate = String(document.getElementById('assignment-end')?.value || '');
    const activity = String(document.getElementById('assignment-activity')?.value || '').trim();
    const assignmentType = normalizeAssignmentType(document.getElementById('assignment-type')?.value || 'official_travel');
    const locationText = String(document.getElementById('assignment-location')?.value || '').trim();
    const note = String(document.getElementById('assignment-note')?.value || '').trim();
    const accountable = Boolean(document.getElementById('assignment-accountable')?.checked);
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
    if (!selectedFile && !existingFile?.itemId && window.UserAccess.can('surat','documents.upload')) { alert('Upload berkas Surat Tugas terlebih dahulu.'); return; }

    const docRef = existing.id ? db().collection('assignmentRecords').doc(existing.id) : db().collection('assignmentRecords').doc();
    const id = docRef.id;
    const assignmentYear = Number(startDate.slice(0, 4)) || CURRENT_YEAR;
    const employees = [...selectedIds].map((employeeId) => {
      const identity = employeeIdentity(employeeId);
      return { employeeId, name: identity.name, nip: identity.nip, unit: identity.unit };
    });
    const draft = { id, letterNumber, letterDate, startDate, endDate, assignmentYear, employees, assignmentType, accountable, activity, location: locationText, note };
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
        employees, assignmentType, accountable, activity, location: locationText, note,
        // Disimpan untuk kompatibilitas data lama. Mesin Tukin memakai assignmentType
        // sebagai business rule utama mulai schemaVersion 2.
        adjustmentType: assignmentType === 'official_travel' ? 'official_duty' : 'manual_review',
        adjustmentPercent: assignmentType === 'official_travel' ? 0 : null,
        documents,
        schemaVersion: 2,
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
      window.AttendanceQuickCheck?.checkSource('surat', id, assignmentYear).then(count => {
        if(count>0) showToast(`${count} potensi irisan dokumen terdeteksi. Periksa Monitoring Konflik Kehadiran.`);
        else if(count<0) console.info('Pemeriksaan register dapat diulang dari menu Monitoring Konflik Kehadiran.');
      });
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
    if (!window.UserAccess.assert('surat','data.delete')) return;
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
    if (!window.UserAccess.can('surat','documents.view')) return;
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
    state.requests.detailId = '';
    state.modal = null;
    state.previewModal = null;
    state.inlinePreviews = {};
    state.detailId = id;
    render();
    loadAssignmentPreview(id);
  }

  async function previewAssignmentDocument(assignmentId, documentId) {
    if (!window.UserAccess.assert('surat','documents.view')) return;
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
    if (!window.UserAccess.assert('surat','documents.download')) return;
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
    state.calendar.year = year;
    state.calendar.month = month;
    state.expandedCalendarDays.clear();
    render();
  }

  app.addEventListener('change', (event) => {
    const target = event.target;
    if (target.matches?.('[data-assignment-employee]')) { updateEmployeeSummary(); return; }
    if (target.id === 'assignment-file') { renderLocalFilePreview(target.files?.[0] || null); return; }
    if (target.id === 'st-request-status-filter') { state.requests.filter = target.value || 'all'; render(); return; }
    if (target.id === 'assignment-filter-year') { state.filters.year = Number(target.value || CURRENT_YEAR); render(); return; }
    if (target.id === 'assignment-filter-unit') { state.filters.unit = String(target.value || ''); render(); return; }
    if (target.id === 'assignment-filter-status') { state.filters.status = String(target.value || ''); render(); return; }
    if (target.id === 'assignment-calendar-month') { state.calendar.month = Number(target.value || CURRENT_MONTH); state.expandedCalendarDays.clear(); render(); return; }
    if (target.id === 'assignment-calendar-year') { state.calendar.year = Number(target.value || CURRENT_YEAR); state.expandedCalendarDays.clear(); render(); return; }
    if (target.id === 'assignment-calendar-unit') { state.calendar.unit = String(target.value || ''); state.expandedCalendarDays.clear(); render(); return; }
    if (target.id === 'assignment-calendar-type') { state.calendar.type = String(target.value || ''); state.expandedCalendarDays.clear(); render(); }
  });

  app.addEventListener('input', (event) => {
    const target = event.target;
    if (target.id === 'st-request-search') {
      state.requests.search = String(target.value || '');
      const cursor = target.selectionStart;
      render();
      requestAnimationFrame(() => {
        const input = document.getElementById('st-request-search');
        if (input) { input.focus(); try { input.setSelectionRange(cursor, cursor); } catch (_) {} }
      });
      return;
    }
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

    if (action === 'st-request-refresh') { refreshRequests(); return; }
    if (action === 'st-request-detail') { state.requests.detailId = String(actionEl.dataset.requestId || ''); render(); return; }
    if (action === 'st-request-close' && state.requests.updatingId) return;
    if (action === 'st-request-close' && actionEl.classList.contains('st-request-detail-backdrop') && event.target !== actionEl) return;
    if (action === 'st-request-close') { state.requests.detailId = ''; render(); return; }
    if (action === 'st-request-read') { markRequestRead(String(actionEl.dataset.requestId || '')); return; }
    if (action === 'add-assignment') { openAssignmentModal(''); return; }
    if (action === 'close-assignment-modal') { state.modal = null; render(); return; }
    if (action === 'close-assignment-detail') { state.detailId = ''; state.inlinePreviews = {}; render(); return; }
    if (action === 'close-preview-modal') { state.previewModal = null; render(); return; }
    if (action === 'save-assignment') { saveAssignment(actionEl); return; }
    if (action === 'reset-filter') { state.filters = { year: CURRENT_YEAR, search: '', unit: '', status: '' }; render(); return; }
    if (action === 'calendar-prev') { changeCalendarMonth(-1); return; }
    if (action === 'calendar-next') { changeCalendarMonth(1); return; }
    if (action === 'calendar-toggle-day') {
      const day = actionEl?.dataset.calendarDay;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(day || '')) return;
      if (state.expandedCalendarDays.has(day)) state.expandedCalendarDays.delete(day);
      else state.expandedCalendarDays.add(day);
      render();
      return;
    }
    const calendarMode = event.target.closest?.('[data-assignment-calendar-mode]')?.dataset.assignmentCalendarMode;
    if (calendarMode) { state.calendarMode = calendarMode === 'list' ? 'list' : 'calendar'; render(); return; }
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
    state.requests.detailId = '';
    render();
  });

  window.addEventListener('beforeunload', () => {
    if (window.__assignmentLocalPreviewUrl) URL.revokeObjectURL(window.__assignmentLocalPreviewUrl);
  });

  async function init() {
    if (!(await window.UserAccess.requireModule('surat'))) return;
    state.view = viewFromHash();
    readRequestMarkers();
    if (!location.hash || !window.UserAccess.canView('surat',state.view)) {
      state.view = window.UserAccess.firstView('surat') || 'dashboard';
      history.replaceState(null, '', '#'+state.view);
    }
    if (!location.hash) history.replaceState(null, '', '#dashboard');
    await loadData();
    // Tautan dari Monitoring Konflik membuka dokumen sumber yang tepat.
    try {
      const recordId = new URLSearchParams(location.search || '').get('record');
      if (recordId && window.UserAccess.canView('surat', 'data')) {
        const found = state.assignments.find(row => row.id === recordId);
        if (found) {
          state.filters.year = Number(found.assignmentYear || CURRENT_YEAR);
          state.view = 'data';
          openAssignmentDetail(found.id);
        }
      }
    } catch (error) { console.warn('Tautan Surat Tugas tidak dapat dibuka:', error); }
  }

  init();
})();
