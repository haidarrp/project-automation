(function () {
  'use strict';

  const app = document.getElementById('app');
  const rules = window.CutiRules;
  const sharePoint = window.SharePointStorage;
  const now = new Date();
  const CURRENT_YEAR = now.getFullYear();
  const CURRENT_MONTH = now.getMonth() + 1;
  const TODAY = rules.toIsoDate(now);
  const LEAVE_TYPES = rules.leaveTypes;
  const DOCUMENT_TYPES = Object.freeze(['Surat Cuti', 'Surat Keterangan', 'Dokumen Pendukung', 'Dokumen Lainnya']);
  const VIEW_LABELS = Object.freeze({
    dashboard: 'Dashboard',
    data: 'Data Cuti',
    kalender: 'Kalender',
    pegawai: 'Data Pegawai',
    saldo: 'Saldo Cuti',
    laporan: 'Laporan'
  });

  const state = {
    loading: true,
    error: '',
    toast: '',
    view: 'dashboard',
    employees: [],
    directory: new Map(),
    leaves: [],
    balances: [],
    loadedYears: new Set(),
    loadingYear: null,
    filters: {
      year: CURRENT_YEAR,
      search: '',
      month: '',
      unit: '',
      type: ''
    },
    calendarFilters: {
      year: CURRENT_YEAR,
      month: CURRENT_MONTH,
      unit: '',
      type: ''
    },
    calendarMode: 'calendar',
    reportFilters: {
      year: CURRENT_YEAR,
      startDate: `${CURRENT_YEAR}-01-01`,
      endDate: `${CURRENT_YEAR}-12-31`,
      employeeId: '',
      unit: '',
      type: ''
    },
    modal: null,
    balanceModal: null,
    documentModal: null,
    drawer: null
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

  function normalizeView(value) {
    const key = String(value || '').toLowerCase();
    return Object.prototype.hasOwnProperty.call(VIEW_LABELS, key) ? key : 'dashboard';
  }

  function viewFromHash() {
    return normalizeView(String(location.hash || '').replace(/^#/, ''));
  }

  function employeeIdentity(id) {
    const row = state.directory.get(String(id || '')) || {};
    return {
      name: String(row.name || '').trim(),
      nip: String(row.nip || '').replace(/\D/g, ''),
      unit: String(row.unit || '').trim()
    };
  }

  function employeeOrder(id) {
    const employee = state.employees.find((item) => item.id === id);
    return Number(employee?.order || 9999);
  }

  function activeEmployees() {
    return state.employees
      .filter((item) => item.active !== false)
      .map((item) => ({ ...item, ...employeeIdentity(item.id) }))
      .filter((item) => item.name)
      .sort((a, b) => Number(a.order || 9999) - Number(b.order || 9999) || a.name.localeCompare(b.name, 'id'));
  }

  function unitLabel(employeeId) {
    return employeeIdentity(employeeId).unit || 'Belum diisi';
  }

  function uniqueUnits() {
    return [...new Set(activeEmployees().map((item) => item.unit).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'id'));
  }

  function dateObj(value) {
    return rules.parseDate(String(value || ''));
  }

  function formatDate(value) {
    const date = dateObj(value);
    if (!date) return '—';
    return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
  }

  function formatDateRange(row) {
    const start = String(row?.startDate || '');
    const end = String(row?.endDate || '');
    if (!start) return '—';
    if (!end || end === start) return formatDate(start);
    return `${formatDate(start)} – ${formatDate(end)}`;
  }

  function monthName(month) {
    const value = Number(month || 1);
    return new Intl.DateTimeFormat('id-ID', { month: 'long' }).format(new Date(2026, Math.max(0, value - 1), 1));
  }

  function monthRange(year, month) {
    const y = Number(year);
    const m = Number(month);
    if (!y || !m) return { start: '', end: '' };
    const last = new Date(y, m, 0).getDate();
    return {
      start: `${y}-${String(m).padStart(2, '0')}-01`,
      end: `${y}-${String(m).padStart(2, '0')}-${String(last).padStart(2, '0')}`
    };
  }

  function addDaysIso(value, days) {
    const date = dateObj(value);
    if (!date) return '';
    date.setDate(date.getDate() + Number(days || 0));
    return rules.toIsoDate(date);
  }

  function yearFromDate(value) {
    return Number(String(value || '').slice(0, 4)) || CURRENT_YEAR;
  }

  function normalizeDocument(raw, index) {
    const row = raw && typeof raw === 'object' ? raw : {};
    const storedFile = row.file && typeof row.file === 'object' ? { ...row.file } : null;
    if (storedFile) delete storedFile.file;
    return {
      id: String(row.id || `doc-${index + 1}`),
      type: String(row.type || 'Dokumen Cuti'),
      name: String(row.name || ''),
      number: String(row.number || ''),
      note: String(row.note || ''),
      file: storedFile
    };
  }

  function formatFileSize(value) {
    const bytes = Number(value || 0);
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(bytes < 10 * 1024 ? 1 : 0)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0)} MB`;
  }

  function cutiSharePointPath(item) {
    const identity = employeeIdentity(item?.employeeId);
    const nip = identity.nip || 'TANPA-NIP';
    const employeeFolder = sharePoint?.safeName?.(`${nip}_${identity.name || 'Pegawai'}`, nip) || `${nip}_${identity.name || 'Pegawai'}`;
    const recordFolder = sharePoint?.safeName?.(`${item?.startDate || 'tanggal'}_${item?.id || 'cuti'}`, item?.id || 'cuti') || String(item?.id || 'cuti');
    return [String(item?.leaveYear || yearFromDate(item?.startDate)), employeeFolder, recordFolder];
  }

  function cutiRemoteFileName(documentId, file) {
    const original = String(file?.name || 'Bukti_Cuti');
    const dot = original.lastIndexOf('.');
    const ext = dot > 0 && original.length - dot <= 10 ? original.slice(dot) : '';
    const base = ext ? original.slice(0, dot) : original;
    const prefix = String(documentId || 'dokumen').replace(/[^A-Za-z0-9_-]/g, '').slice(-28) || 'dokumen';
    const safeBase = sharePoint?.safeName?.(base, 'Bukti_Cuti') || base;
    const maxBase = Math.max(20, 116 - prefix.length - ext.length);
    return `${prefix}_${safeBase.slice(0, maxBase)}${ext}`;
  }

  function normalizeLeave(data, id) {
    const row = data || {};
    const startDate = String(row.startDate || '');
    const endDate = String(row.endDate || '');
    const calculatedDays = Number.isFinite(Number(row.calculatedDays))
      ? Number(row.calculatedDays)
      : rules.countWorkingDays(startDate, endDate);
    const days = Number.isFinite(Number(row.days)) ? Number(row.days) : calculatedDays;
    return {
      ...row,
      id,
      employeeId: String(row.employeeId || ''),
      leaveType: LEAVE_TYPES.includes(String(row.leaveType || '')) ? String(row.leaveType) : String(row.leaveType || 'Lainnya'),
      startDate,
      endDate,
      calculatedDays,
      days,
      manualOverride: Boolean(row.manualOverride) || days !== calculatedDays,
      leaveYear: Number(row.leaveYear || yearFromDate(startDate)),
      documentNumber: String(row.documentNumber || ''),
      documentDate: String(row.documentDate || ''),
      note: String(row.note || ''),
      documents: (Array.isArray(row.documents) ? row.documents : []).map(normalizeDocument)
    };
  }

  function normalizeBalance(data, id) {
    const row = data || {};
    return {
      ...row,
      id,
      employeeId: String(row.employeeId || ''),
      year: Number(row.year || CURRENT_YEAR),
      entitlement: Number(row.entitlement || 0),
      carryOver: Number(row.carryOver || 0),
      usage: Number(row.usage || 0),
      remaining: Number(row.remaining || 0)
    };
  }

  function leavesByYear(year) {
    return state.leaves.filter((item) => Number(item.leaveYear) === Number(year));
  }

  function annualUsage(employeeId, year, sourceLeaves) {
    return (sourceLeaves || state.leaves)
      .filter((item) => item.employeeId === employeeId && Number(item.leaveYear) === Number(year) && item.leaveType === 'Cuti Tahunan')
      .reduce((sum, item) => sum + Number(item.days || 0), 0);
  }

  function storedBalance(employeeId, year) {
    return state.balances.find((item) => item.employeeId === employeeId && Number(item.year) === Number(year)) || null;
  }

  function balanceSummary(employeeId, year) {
    const stored = storedBalance(employeeId, year) || {};
    const entitlement = Number(stored.entitlement || 0);
    const carryOver = Number(stored.carryOver || 0);
    const usage = annualUsage(employeeId, year);
    return {
      id: stored.id || balanceDocId(employeeId, year),
      employeeId,
      year: Number(year),
      entitlement,
      carryOver,
      usage,
      remaining: entitlement + carryOver - usage
    };
  }

  function balanceDocId(employeeId, year) {
    return `${String(employeeId || '')}__${Number(year || CURRENT_YEAR)}`;
  }

  function yearOptions(selected) {
    const values = new Set([...Array.from({ length: 11 }, (_, index) => CURRENT_YEAR - index), CURRENT_YEAR + 1, Number(selected)]);
    state.leaves.forEach((item) => values.add(Number(item.leaveYear)));
    state.balances.forEach((item) => values.add(Number(item.year)));
    return [...values].filter((value) => Number.isFinite(value) && value > 2000 && value < 2200).sort((a, b) => b - a);
  }

  function yearSelect(id, selected, extraAttrs) {
    return `<select id="${esc(id)}" ${extraAttrs || ''}>${yearOptions(selected).map((year) => `<option value="${year}" ${Number(selected) === year ? 'selected' : ''}>${year}</option>`).join('')}</select>`;
  }

  function monthSelect(id, selected, allowAll) {
    return `<select id="${esc(id)}">${allowAll ? `<option value="">Semua Bulan</option>` : ''}${Array.from({ length: 12 }, (_, i) => i + 1).map((month) => `<option value="${String(month).padStart(2, '0')}" ${Number(selected) === month ? 'selected' : ''}>${monthName(month)}</option>`).join('')}</select>`;
  }

  function unitSelect(id, selected, allowAll) {
    const units = uniqueUnits();
    return `<select id="${esc(id)}">${allowAll ? '<option value="">Semua Unit/Bidang</option>' : ''}${units.map((unit) => `<option value="${esc(unit)}" ${selected === unit ? 'selected' : ''}>${esc(unit)}</option>`).join('')}</select>`;
  }

  function typeSelect(id, selected, allowAll) {
    return `<select id="${esc(id)}">${allowAll ? '<option value="">Semua Jenis Cuti</option>' : ''}${LEAVE_TYPES.map((type) => `<option value="${esc(type)}" ${selected === type ? 'selected' : ''}>${esc(type)}</option>`).join('')}</select>`;
  }

  function employeeSelect(id, selected, allowAll) {
    const employees = activeEmployees();
    return `<select id="${esc(id)}">${allowAll ? '<option value="">Semua Pegawai</option>' : '<option value="">Pilih pegawai</option>'}${employees.map((employee) => `<option value="${esc(employee.id)}" ${selected === employee.id ? 'selected' : ''}>${esc(employee.name)}${employee.nip ? ` · ${esc(employee.nip)}` : ''}</option>`).join('')}</select>`;
  }

  function metric(label, value, note) {
    return `<div class="card leave-metric"><div class="leave-metric-label">${esc(label)}</div><div><div class="leave-metric-value">${esc(value)}</div><div class="leave-metric-note">${esc(note || '')}</div></div></div>`;
  }

  function usageByMonth(year) {
    const totals = Array(12).fill(0);
    leavesByYear(year).forEach((item) => {
      const working = rules.workingDates(item.startDate, item.endDate).filter((date) => yearFromDate(date) === Number(year));
      working.forEach((date) => { totals[Number(date.slice(5, 7)) - 1] += 1; });
      const delta = Number(item.days || 0) - working.length;
      if (delta !== 0) {
        const index = Math.max(0, Math.min(11, Number(item.startDate.slice(5, 7) || 1) - 1));
        totals[index] += delta;
      }
    });
    return totals;
  }

  function chartBars(values, labels, formatter) {
    const max = Math.max(1, ...values.map((value) => Math.max(0, Number(value || 0))));
    return `<div class="leave-bar-chart">${values.map((value, index) => `<div class="leave-bar-item" title="${esc(labels[index])}: ${esc(formatter ? formatter(value) : value)}"><div class="leave-bar-track"><div class="leave-bar-fill" style="height:${Math.max(3, Math.round((Math.max(0, value) / max) * 100))}%"></div></div><span>${esc(labels[index])}</span><strong>${esc(formatter ? formatter(value) : value)}</strong></div>`).join('')}</div>`;
  }

  function dashboardView() {
    const year = Number(state.filters.year || CURRENT_YEAR);
    const yearLeaves = leavesByYear(year);
    const todayLeaves = yearLeaves.filter((item) => item.startDate <= TODAY && item.endDate >= TODAY);
    const upcomingLimit = addDaysIso(TODAY, 7);
    const upcoming = yearLeaves
      .filter((item) => item.startDate > TODAY && item.startDate <= upcomingLimit)
      .sort((a, b) => a.startDate.localeCompare(b.startDate));
    const totalDays = yearLeaves.reduce((sum, item) => sum + Number(item.days || 0), 0);
    const monthly = usageByMonth(year);
    const typeTotals = LEAVE_TYPES.map((type) => yearLeaves.filter((item) => item.leaveType === type).reduce((sum, item) => sum + Number(item.days || 0), 0));
    const typeMax = Math.max(1, ...typeTotals);
    const summaries = activeEmployees().map((employee) => ({ employee, balance: balanceSummary(employee.id, year) }));
    const annualUsers = summaries.filter((item) => item.balance.usage > 0).length;
    const averageUsage = summaries.length ? summaries.reduce((sum, item) => sum + item.balance.usage, 0) / summaries.length : 0;
    const averageRemaining = summaries.length ? summaries.reduce((sum, item) => sum + item.balance.remaining, 0) / summaries.length : 0;
    const exhausted = summaries.filter((item) => (item.balance.entitlement + item.balance.carryOver) > 0 && item.balance.remaining <= 0).length;

    return `<div class="leave-page">
      <div class="leave-head"><div><h1>Dashboard Cuti</h1><p>Monitoring internal Bagian Umum dan Tata Usaha terhadap pencatatan cuti pegawai Pusdatin.</p></div><div class="leave-head-actions"><div class="field leave-year-field"><label>Tahun</label>${yearSelect('leave-dashboard-year', year)}</div><button class="btn btn-primary" type="button" data-action="add-leave">+ Tambah Data Cuti</button></div></div>
      <div class="leave-metrics">
        ${metric('Pegawai Pusdatin', activeEmployees().length, 'Pegawai aktif pada master')}
        ${metric('Sedang Cuti Hari Ini', todayLeaves.length, TODAY === `${year}-${TODAY.slice(5)}` ? formatDate(TODAY) : `Filter tahun ${year}`)}
        ${metric('Akan Cuti 7 Hari', upcoming.length, 'Mulai dalam 7 hari ke depan')}
        ${metric('Total Hari Cuti', totalDays, `Seluruh jenis cuti tahun ${year}`)}
      </div>
      <div class="leave-dashboard-grid">
        <section class="card leave-card"><div class="leave-card-head"><div><h2>Pegawai Cuti Hari Ini</h2><p>Pegawai dengan periode cuti yang mencakup tanggal hari ini.</p></div><span class="card-subtitle">${todayLeaves.length} pegawai</span></div>${leaveMiniTable(todayLeaves, 'today')}</section>
        <section class="card leave-card"><div class="leave-card-head"><div><h2>Cuti Mendatang</h2><p>Data cuti yang mulai dalam tujuh hari ke depan.</p></div><span class="card-subtitle">${upcoming.length} data</span></div>${leaveMiniTable(upcoming, 'upcoming')}</section>
      </div>
      <div class="leave-dashboard-grid">
        <section class="card leave-card leave-chart-card"><div class="leave-card-head"><div><h2>Penggunaan Cuti per Bulan</h2><p>Jumlah hari cuti Januari–Desember pada tahun ${year}.</p></div></div>${chartBars(monthly, ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'])}</section>
        <section class="card leave-card leave-chart-card"><div class="leave-card-head"><div><h2>Komposisi Jenis Cuti</h2><p>Distribusi jumlah hari berdasarkan jenis cuti.</p></div></div><div class="leave-composition">${LEAVE_TYPES.map((type, index) => `<div class="leave-composition-row"><div><span>${esc(type)}</span><strong>${typeTotals[index]} hari</strong></div><div class="leave-composition-track"><div style="width:${Math.round((typeTotals[index] / typeMax) * 100)}%"></div></div></div>`).join('')}</div></section>
      </div>
      <section class="card leave-card"><div class="leave-card-head"><div><h2>Monitoring Saldo Cuti Tahunan</h2><p>Informasi hak, penggunaan, dan sisa berdasarkan saldo yang dicatat TU. Jenis cuti selain Cuti Tahunan tidak mengurangi saldo ini.</p></div><div class="leave-summary-chips"><span>Rata-rata digunakan <strong>${averageUsage.toFixed(1)}</strong></span><span>Rata-rata sisa <strong>${averageRemaining.toFixed(1)}</strong></span><span>Belum menggunakan <strong>${Math.max(0, summaries.length - annualUsers)}</strong></span><span>Saldo habis <strong>${exhausted}</strong></span></div></div>${balanceTableMarkup(summaries.slice(0, 12), false)}</section>
    </div>`;
  }

  function leaveMiniTable(rows, kind) {
    if (!rows.length) return '<div class="leave-empty"><strong>Tidak ada data</strong>Belum terdapat pegawai pada kategori ini.</div>';
    if (kind === 'today') {
      return `<div class="leave-table-wrap"><table class="data-table leave-table"><thead><tr><th>Pegawai</th><th>Unit/Bidang</th><th>Jenis Cuti</th><th>Periode</th></tr></thead><tbody>${rows.slice(0, 8).map((item) => `<tr><td><button class="leave-link" type="button" data-leave-detail="${esc(item.id)}">${esc(employeeIdentity(item.employeeId).name || 'Pegawai')}</button></td><td>${esc(unitLabel(item.employeeId))}</td><td>${esc(item.leaveType)}</td><td class="nowrap">${esc(formatDateRange(item))}</td></tr>`).join('')}</tbody></table></div>`;
    }
    return `<div class="leave-table-wrap"><table class="data-table leave-table"><thead><tr><th>Pegawai</th><th>Unit/Bidang</th><th>Mulai</th><th>Jenis Cuti</th><th>Jumlah Hari</th></tr></thead><tbody>${rows.slice(0, 8).map((item) => `<tr><td><button class="leave-link" type="button" data-leave-detail="${esc(item.id)}">${esc(employeeIdentity(item.employeeId).name || 'Pegawai')}</button></td><td>${esc(unitLabel(item.employeeId))}</td><td class="nowrap">${esc(formatDate(item.startDate))}</td><td>${esc(item.leaveType)}</td><td>${esc(item.days)}</td></tr>`).join('')}</tbody></table></div>`;
  }

  function recordMatchesMonth(item, year, month) {
    if (!month) return true;
    const range = monthRange(year, month);
    return item.startDate <= range.end && item.endDate >= range.start;
  }

  function filteredDataLeaves() {
    const q = state.filters.search.trim().toLowerCase();
    return state.leaves
      .filter((item) => Number(item.leaveYear) === Number(state.filters.year))
      .filter((item) => recordMatchesMonth(item, state.filters.year, state.filters.month))
      .filter((item) => !state.filters.unit || employeeIdentity(item.employeeId).unit === state.filters.unit)
      .filter((item) => !state.filters.type || item.leaveType === state.filters.type)
      .filter((item) => {
        if (!q) return true;
        const identity = employeeIdentity(item.employeeId);
        return [identity.name, identity.nip, identity.unit, item.leaveType, item.documentNumber, item.note].some((value) => String(value || '').toLowerCase().includes(q));
      })
      .sort((a, b) => b.startDate.localeCompare(a.startDate) || employeeIdentity(a.employeeId).name.localeCompare(employeeIdentity(b.employeeId).name, 'id'));
  }

  function dataFilterCard() {
    return `<section class="card leave-filter-card"><div class="leave-filter-grid">
      <div class="search"><input id="leave-search" type="search" value="${esc(state.filters.search)}" placeholder="Cari nama/NIP..." autocomplete="off"></div>
      <div class="field"><label>Tahun</label>${yearSelect('leave-data-year', state.filters.year)}</div>
      <div class="field"><label>Bulan/Periode</label>${monthSelect('leave-data-month', state.filters.month, true)}</div>
      <div class="field"><label>Unit/Bidang</label>${unitSelect('leave-data-unit', state.filters.unit, true)}</div>
      <div class="field"><label>Jenis Cuti</label>${typeSelect('leave-data-type', state.filters.type, true)}</div>
      <div class="leave-filter-actions"><button class="btn btn-secondary btn-sm" type="button" data-action="reset-data-filter">Reset</button></div>
    </div></section>`;
  }

  function dataView() {
    const rows = filteredDataLeaves();
    return `<div class="leave-page">
      <div class="leave-head"><div><h1>Data Cuti</h1><p>Pencatatan cuti yang telah diajukan pegawai. Modul ini tidak menjalankan workflow persetujuan.</p></div><div class="leave-head-actions"><button class="btn btn-primary" type="button" data-action="add-leave">+ Tambah Data Cuti</button></div></div>
      ${dataFilterCard()}
      <section class="card leave-card"><div class="leave-card-head"><div><h2>Daftar Cuti</h2><p>Filter dapat dikombinasikan berdasarkan tahun, bulan, unit, jenis cuti, serta nama/NIP.</p></div><span class="card-subtitle">${rows.length} data</span></div>
      <div class="leave-table-wrap"><table class="data-table leave-table leave-table-wide"><thead><tr><th>Pegawai</th><th>NIP</th><th>Unit/Bidang</th><th>Jenis Cuti</th><th>Periode</th><th>Hari</th><th>Dokumen</th><th>Aksi</th></tr></thead><tbody>${rows.length ? rows.map((item) => {
        const identity = employeeIdentity(item.employeeId);
        return `<tr><td><button class="leave-link" type="button" data-leave-detail="${esc(item.id)}">${esc(identity.name || 'Pegawai')}</button></td><td class="nowrap">${esc(identity.nip || '—')}</td><td>${esc(identity.unit || 'Belum diisi')}</td><td>${esc(item.leaveType)}</td><td class="nowrap">${esc(formatDateRange(item))}</td><td>${esc(item.days)}${item.manualOverride ? '<span class="leave-sub">koreksi manual</span>' : ''}</td><td>${item.documentNumber ? `${esc(item.documentNumber)}<span class="leave-sub">${esc(formatDate(item.documentDate))}</span>` : '—'}</td><td><div class="leave-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-leave-detail="${esc(item.id)}">Detail</button><button class="btn btn-secondary btn-sm" type="button" data-leave-edit="${esc(item.id)}">Edit</button><button class="btn btn-danger btn-sm" type="button" data-leave-delete="${esc(item.id)}">Hapus</button></div></td></tr>`;
      }).join('') : '<tr><td colspan="8"><div class="leave-empty"><strong>Data tidak ditemukan</strong>Ubah filter atau tambahkan data cuti.</div></td></tr>'}</tbody></table></div></section>
    </div>`;
  }

  function calendarRecords() {
    const filter = state.calendarFilters;
    const range = monthRange(filter.year, filter.month);
    return state.leaves
      .filter((item) => item.startDate <= range.end && item.endDate >= range.start)
      .filter((item) => !filter.unit || employeeIdentity(item.employeeId).unit === filter.unit)
      .filter((item) => !filter.type || item.leaveType === filter.type)
      .sort((a, b) => a.startDate.localeCompare(b.startDate) || employeeOrder(a.employeeId) - employeeOrder(b.employeeId));
  }

  function calendarGrid(rows) {
    const year = Number(state.calendarFilters.year);
    const month = Number(state.calendarFilters.month);
    const first = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0).getDate();
    const mondayOffset = (first.getDay() + 6) % 7;
    const cells = [];
    for (let i = 0; i < mondayOffset; i += 1) cells.push('<div class="leave-calendar-cell muted"></div>');
    for (let day = 1; day <= lastDay; day += 1) {
      const iso = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const date = new Date(year, month - 1, day);
      const events = rows.filter((item) => item.startDate <= iso && item.endDate >= iso);
      const isToday = iso === TODAY;
      const weekend = date.getDay() === 0 || date.getDay() === 6;
      cells.push(`<div class="leave-calendar-cell ${weekend ? 'weekend' : ''} ${isToday ? 'today' : ''}"><div class="leave-calendar-day">${day}${isToday ? '<span>Hari ini</span>' : ''}</div><div class="leave-calendar-events">${events.slice(0, 4).map((item) => `<button type="button" data-leave-detail="${esc(item.id)}" title="${esc(employeeIdentity(item.employeeId).name)} — ${esc(item.leaveType)}"><strong>${esc(employeeIdentity(item.employeeId).name || 'Pegawai')}</strong><span>${esc(item.leaveType)}</span></button>`).join('')}${events.length > 4 ? `<div class="leave-calendar-more">+${events.length - 4} lainnya</div>` : ''}</div></div>`);
    }
    while (cells.length % 7) cells.push('<div class="leave-calendar-cell muted"></div>');
    return `<div class="leave-calendar"><div class="leave-calendar-weekdays">${['Sen','Sel','Rab','Kam','Jum','Sab','Min'].map((day) => `<span>${day}</span>`).join('')}</div><div class="leave-calendar-grid">${cells.join('')}</div></div>`;
  }

  function calendarList(rows) {
    return `<div class="leave-table-wrap"><table class="data-table leave-table"><thead><tr><th>Pegawai</th><th>Unit/Bidang</th><th>Jenis Cuti</th><th>Periode</th><th>Hari</th></tr></thead><tbody>${rows.length ? rows.map((item) => `<tr><td><button class="leave-link" type="button" data-leave-detail="${esc(item.id)}">${esc(employeeIdentity(item.employeeId).name || 'Pegawai')}</button></td><td>${esc(unitLabel(item.employeeId))}</td><td>${esc(item.leaveType)}</td><td class="nowrap">${esc(formatDateRange(item))}</td><td>${esc(item.days)}</td></tr>`).join('') : '<tr><td colspan="5"><div class="leave-empty"><strong>Tidak ada data</strong>Tidak ada cuti pada bulan dan filter yang dipilih.</div></td></tr>'}</tbody></table></div>`;
  }

  function calendarView() {
    const rows = calendarRecords();
    const f = state.calendarFilters;
    return `<div class="leave-page">
      <div class="leave-head"><div><h1>Kalender Cuti</h1><p>Monitoring visual periode cuti pegawai tanpa deteksi benturan atau pembatasan jumlah pegawai.</p></div><div class="leave-head-actions"><button class="btn btn-secondary" type="button" data-action="calendar-prev">← Bulan Sebelumnya</button><button class="btn btn-secondary" type="button" data-action="calendar-next">Bulan Berikutnya →</button></div></div>
      <section class="card leave-filter-card"><div class="leave-filter-grid calendar-filters"><div class="field"><label>Bulan</label>${monthSelect('leave-calendar-month', f.month, false)}</div><div class="field"><label>Tahun</label>${yearSelect('leave-calendar-year', f.year)}</div><div class="field"><label>Unit/Bidang</label>${unitSelect('leave-calendar-unit', f.unit, true)}</div><div class="field"><label>Jenis Cuti</label>${typeSelect('leave-calendar-type', f.type, true)}</div><div class="leave-view-toggle"><button type="button" class="btn btn-sm ${state.calendarMode === 'list' ? 'btn-primary' : 'btn-secondary'}" data-calendar-mode="list">Daftar</button><button type="button" class="btn btn-sm ${state.calendarMode === 'calendar' ? 'btn-primary' : 'btn-secondary'}" data-calendar-mode="calendar">Kalender</button></div></div></section>
      <section class="card leave-card"><div class="leave-card-head"><div><h2>${esc(monthName(f.month))} ${esc(f.year)}</h2><p>${rows.length} record cuti pada filter aktif.</p></div></div>${state.calendarMode === 'calendar' ? calendarGrid(rows) : calendarList(rows)}</section>
    </div>`;
  }

  function employeeRows() {
    const q = state.filters.search.trim().toLowerCase();
    return activeEmployees().filter((employee) => !q || [employee.name, employee.nip, employee.unit].some((value) => String(value || '').toLowerCase().includes(q)));
  }

  function latestLeave(employeeId, year) {
    return state.leaves
      .filter((item) => item.employeeId === employeeId && Number(item.leaveYear) === Number(year))
      .sort((a, b) => b.startDate.localeCompare(a.startDate))[0] || null;
  }

  function employeeView() {
    const year = Number(state.filters.year);
    const rows = employeeRows();
    return `<div class="leave-page">
      <div class="leave-head"><div><h1>Data Pegawai</h1><p>Ringkasan penggunaan cuti dari master pegawai existing. Klik pegawai untuk melihat riwayat dan saldo tahun yang dipilih.</p></div><div class="leave-head-actions"><div class="field leave-year-field"><label>Tahun</label>${yearSelect('leave-employee-year', year)}</div></div></div>
      <section class="card leave-filter-card"><div class="leave-filter-grid employee-filter"><div class="search"><input id="leave-employee-search" type="search" value="${esc(state.filters.search)}" placeholder="Cari nama/NIP..." autocomplete="off"></div></div></section>
      <section class="card leave-card"><div class="leave-card-head"><div><h2>Daftar Pegawai</h2><p>Penggunaan dan sisa pada tabel mengacu pada saldo Cuti Tahunan.</p></div><span class="card-subtitle">${rows.length} pegawai</span></div><div class="leave-table-wrap"><table class="data-table leave-table"><thead><tr><th>Pegawai</th><th>Unit/Bidang</th><th>Hak + Saldo</th><th>Digunakan</th><th>Sisa</th><th>Cuti Terakhir</th><th>Aksi</th></tr></thead><tbody>${rows.length ? rows.map((employee) => {
        const balance = balanceSummary(employee.id, year);
        const latest = latestLeave(employee.id, year);
        return `<tr><td><button class="leave-link" type="button" data-employee-leave-detail="${esc(employee.id)}">${esc(employee.name)}</button><span class="leave-sub">${employee.nip ? `NIP ${esc(employee.nip)}` : 'NIP belum tersedia'}</span></td><td>${esc(employee.unit || 'Belum diisi')}</td><td>${balance.entitlement + balance.carryOver}</td><td>${balance.usage}</td><td><strong>${balance.remaining}</strong></td><td>${latest ? `${esc(formatDateRange(latest))}<span class="leave-sub">${esc(latest.leaveType)}</span>` : '—'}</td><td><button class="btn btn-secondary btn-sm" type="button" data-employee-leave-detail="${esc(employee.id)}">Detail</button></td></tr>`;
      }).join('') : '<tr><td colspan="7"><div class="leave-empty"><strong>Pegawai tidak ditemukan</strong>Ubah kata kunci pencarian.</div></td></tr>'}</tbody></table></div></section>
    </div>`;
  }

  function balanceTableMarkup(rows, withActions) {
    return `<div class="leave-table-wrap"><table class="data-table leave-table"><thead><tr><th>Pegawai</th><th>Hak</th><th>Saldo Tahun Lalu</th><th>Digunakan</th><th>Sisa</th>${withActions ? '<th>Aksi</th>' : ''}</tr></thead><tbody>${rows.length ? rows.map(({ employee, balance }) => `<tr><td>${esc(employee.name)}<span class="leave-sub">${esc(employee.unit || 'Belum diisi')}</span></td><td>${balance.entitlement}</td><td>${balance.carryOver}</td><td>${balance.usage}</td><td><strong>${balance.remaining}</strong></td>${withActions ? `<td><button class="btn btn-secondary btn-sm" type="button" data-balance-edit="${esc(employee.id)}">Edit Saldo</button></td>` : ''}</tr>`).join('') : `<tr><td colspan="${withActions ? 6 : 5}"><div class="leave-empty"><strong>Belum ada pegawai</strong>Master pegawai belum tersedia.</div></td></tr>`}</tbody></table></div>`;
  }

  function balanceView() {
    const year = Number(state.filters.year);
    const rows = activeEmployees().map((employee) => ({ employee, balance: balanceSummary(employee.id, year) }));
    return `<div class="leave-page">
      <div class="leave-head"><div><h1>Saldo Cuti</h1><p>Saldo Cuti Tahunan dikelola per pegawai per tahun dan terpisah dari riwayat cuti.</p></div><div class="leave-head-actions"><div class="field leave-year-field"><label>Tahun</label>${yearSelect('leave-balance-year', year)}</div></div></div>
      <div class="alert alert-warning"><div class="alert-title">Perhitungan saldo</div>Sisa = Hak Cuti + Saldo yang Dibawa dari Tahun Sebelumnya − Penggunaan Cuti Tahunan. Nilai hak dan saldo awal dapat disesuaikan TU sesuai dokumen/ketentuan yang berlaku.</div>
      <section class="card leave-card"><div class="leave-card-head"><div><h2>Saldo Cuti per Pegawai</h2><p>Jenis cuti selain Cuti Tahunan tidak otomatis mengurangi saldo Cuti Tahunan.</p></div><span class="card-subtitle">${rows.length} pegawai</span></div>${balanceTableMarkup(rows, true)}</section>
    </div>`;
  }

  function reportRows() {
    const f = state.reportFilters;
    return state.leaves
      .filter((item) => Number(item.leaveYear) === Number(f.year))
      .filter((item) => !f.startDate || item.endDate >= f.startDate)
      .filter((item) => !f.endDate || item.startDate <= f.endDate)
      .filter((item) => !f.employeeId || item.employeeId === f.employeeId)
      .filter((item) => !f.unit || employeeIdentity(item.employeeId).unit === f.unit)
      .filter((item) => !f.type || item.leaveType === f.type)
      .sort((a, b) => a.startDate.localeCompare(b.startDate) || employeeIdentity(a.employeeId).name.localeCompare(employeeIdentity(b.employeeId).name, 'id'));
  }

  function reportView() {
    const f = state.reportFilters;
    const rows = reportRows();
    const totalDays = rows.reduce((sum, item) => sum + Number(item.days || 0), 0);
    return `<div class="leave-page leave-report-page">
      <div class="leave-head"><div><h1>Laporan Cuti</h1><p>Rekap di layar berdasarkan kombinasi filter tahun, tanggal, pegawai, unit, dan jenis cuti.</p></div><div class="leave-head-actions"><button class="btn btn-secondary" type="button" data-action="print-report">Cetak Rekap</button></div></div>
      <section class="card leave-filter-card"><div class="leave-report-filters">
        <div class="field"><label>Tahun</label>${yearSelect('leave-report-year', f.year)}</div>
        <div class="field"><label>Dari Tanggal</label><input id="leave-report-start" type="date" value="${esc(f.startDate)}"></div>
        <div class="field"><label>Sampai Tanggal</label><input id="leave-report-end" type="date" value="${esc(f.endDate)}"></div>
        <div class="field"><label>Pegawai</label>${employeeSelect('leave-report-employee', f.employeeId, true)}</div>
        <div class="field"><label>Unit/Bidang</label>${unitSelect('leave-report-unit', f.unit, true)}</div>
        <div class="field"><label>Jenis Cuti</label>${typeSelect('leave-report-type', f.type, true)}</div>
      </div></section>
      <div class="leave-metrics leave-report-metrics">${metric('Jumlah Record', rows.length, 'Sesuai filter')}${metric('Total Hari', totalDays, 'Akumulasi jumlah hari')}</div>
      <section class="card leave-card leave-report-card"><div class="leave-card-head"><div><h2>Rekap Cuti</h2><p>Periode laporan ${esc(formatDate(f.startDate))} s.d. ${esc(formatDate(f.endDate))}.</p></div></div><div class="leave-table-wrap"><table class="data-table leave-table leave-table-wide"><thead><tr><th>No</th><th>Pegawai</th><th>NIP</th><th>Unit</th><th>Jenis Cuti</th><th>Periode</th><th>Jumlah Hari</th></tr></thead><tbody>${rows.length ? rows.map((item, index) => {
        const identity = employeeIdentity(item.employeeId);
        return `<tr><td>${index + 1}</td><td>${esc(identity.name || 'Pegawai')}</td><td>${esc(identity.nip || '—')}</td><td>${esc(identity.unit || 'Belum diisi')}</td><td>${esc(item.leaveType)}</td><td class="nowrap">${esc(formatDateRange(item))}</td><td>${esc(item.days)}</td></tr>`;
      }).join('') : '<tr><td colspan="7"><div class="leave-empty"><strong>Tidak ada data</strong>Tidak ada record yang sesuai filter laporan.</div></td></tr>'}</tbody></table></div></section>
    </div>`;
  }

  function leaveDrawer(item) {
    const identity = employeeIdentity(item.employeeId);
    const docs = item.documents || [];
    return `<div class="leave-drawer-backdrop" data-action="close-drawer"></div><aside class="leave-drawer" role="dialog" aria-modal="true" aria-label="Detail cuti">
      <div class="leave-drawer-head"><div><h3>${esc(identity.name || 'Pegawai')}</h3><p>Detail pencatatan cuti internal Pusdatin.</p></div><button class="icon-btn" type="button" data-action="close-drawer">×</button></div>
      <div class="leave-detail-section"><div class="leave-section-title">Informasi Pegawai</div><div class="leave-detail-grid"><div><span>Nama</span><strong>${esc(identity.name || '—')}</strong></div><div><span>NIP</span><strong>${esc(identity.nip || '—')}</strong></div><div class="span-2"><span>Unit/Bidang</span><strong>${esc(identity.unit || 'Belum diisi')}</strong></div></div></div>
      <div class="leave-detail-section"><div class="leave-section-title">Informasi Cuti</div><div class="leave-detail-grid"><div><span>Jenis Cuti</span><strong>${esc(item.leaveType)}</strong></div><div><span>Tahun</span><strong>${esc(item.leaveYear)}</strong></div><div><span>Tanggal Mulai</span><strong>${esc(formatDate(item.startDate))}</strong></div><div><span>Tanggal Selesai</span><strong>${esc(formatDate(item.endDate))}</strong></div><div><span>Jumlah Hari</span><strong>${esc(item.days)}${item.manualOverride ? ' (koreksi manual)' : ''}</strong></div><div><span>Hasil Kalkulasi</span><strong>${esc(item.calculatedDays)} hari kerja</strong></div><div><span>Nomor Surat/Dokumen</span><strong>${esc(item.documentNumber || '—')}</strong></div><div><span>Tanggal Surat</span><strong>${esc(formatDate(item.documentDate))}</strong></div><div class="span-2"><span>Catatan</span><strong>${item.note ? esc(item.note) : '—'}</strong></div></div></div>
      <div class="leave-document-head"><div><div class="leave-section-title">Bukti Dukung Cuti</div><p>File disimpan pada SharePoint menggunakan lokasi yang sama dengan TUKIN, di dalam folder khusus CUTI.</p></div><button class="btn btn-secondary btn-sm" type="button" data-add-leave-document="${esc(item.id)}">+ Tambah Dokumen</button></div>
      <div class="leave-document-list">${docs.length ? docs.map((doc) => {
        const file = doc.file || null;
        const fileDetail = file?.itemId ? `<br><span class="leave-file-meta">SharePoint · ${esc(file.name || file.remoteName || 'File')}${file.size ? ` · ${esc(formatFileSize(file.size))}` : ''}</span>` : '<br><span class="leave-file-meta leave-file-missing">File belum tersimpan di SharePoint</span>';
        const fileActions = file?.itemId ? `<button class="btn btn-secondary btn-sm" type="button" data-download-leave-document="${esc(doc.id)}" data-leave-id="${esc(item.id)}">Unduh</button>${file.webUrl ? `<a class="btn btn-secondary btn-sm" href="${esc(file.webUrl)}" target="_blank" rel="noopener noreferrer">SharePoint ↗</a>` : ''}` : '';
        return `<div class="leave-document-row"><div><strong>${esc(doc.name || doc.type)}</strong><span>${esc(doc.type)}${doc.number ? ` · No. ${esc(doc.number)}` : ''}${doc.note ? `<br>${esc(doc.note)}` : ''}${fileDetail}</span></div><div class="leave-row-actions">${fileActions}<button class="btn btn-secondary btn-sm" type="button" data-edit-leave-document="${esc(doc.id)}" data-leave-id="${esc(item.id)}">Edit</button><button class="btn btn-danger btn-sm" type="button" data-delete-leave-document="${esc(doc.id)}" data-leave-id="${esc(item.id)}">Hapus</button></div></div>`;
      }).join('') : '<div class="leave-empty"><strong>Belum ada bukti dukung</strong>Tambahkan dokumen cuti untuk mengunggah file ke SharePoint.</div>'}</div>
      <div class="leave-modal-actions"><button class="btn btn-secondary" type="button" data-leave-edit="${esc(item.id)}">Edit Data</button><button class="btn btn-danger" type="button" data-leave-delete="${esc(item.id)}">Hapus</button></div>
    </aside>`;
  }

  function employeeDrawer(employeeId) {
    const identity = employeeIdentity(employeeId);
    const year = Number(state.filters.year);
    const history = state.leaves.filter((item) => item.employeeId === employeeId && Number(item.leaveYear) === year).sort((a, b) => b.startDate.localeCompare(a.startDate));
    const balance = balanceSummary(employeeId, year);
    const latest = history[0] || null;
    return `<div class="leave-drawer-backdrop" data-action="close-drawer"></div><aside class="leave-drawer" role="dialog" aria-modal="true" aria-label="Detail cuti pegawai">
      <div class="leave-drawer-head"><div><h3>${esc(identity.name || 'Pegawai')}</h3><p>${identity.nip ? `NIP ${esc(identity.nip)} · ` : ''}${esc(identity.unit || 'Unit belum diisi')}</p></div><button class="icon-btn" type="button" data-action="close-drawer">×</button></div>
      <div class="leave-section-title">Ringkasan Cuti ${year}</div>
      <div class="leave-detail-grid"><div><span>Hak Cuti Tahunan</span><strong>${balance.entitlement}</strong></div><div><span>Saldo Tahun Sebelumnya</span><strong>${balance.carryOver}</strong></div><div><span>Digunakan</span><strong>${balance.usage}</strong></div><div><span>Sisa</span><strong>${balance.remaining}</strong></div><div><span>Jumlah Kejadian Cuti</span><strong>${history.length}</strong></div><div><span>Cuti Terakhir</span><strong>${latest ? esc(formatDateRange(latest)) : '—'}</strong></div></div>
      <div class="leave-section-title">Riwayat Cuti</div>
      <div class="leave-document-list">${history.length ? history.map((item) => `<button class="leave-history-row" type="button" data-leave-detail="${esc(item.id)}"><span><strong>${esc(item.leaveType)}</strong><small>${esc(formatDateRange(item))} · ${esc(item.days)} hari</small></span><span>${esc(item.leaveYear)}</span></button>`).join('') : '<div class="leave-empty"><strong>Belum ada riwayat</strong>Tidak ada data cuti pada tahun yang dipilih.</div>'}</div>
    </aside>`;
  }

  function leaveModal() {
    if (!state.modal) return '';
    const item = state.modal.record || {};
    const calculated = Number.isFinite(Number(item.calculatedDays)) ? Number(item.calculatedDays) : rules.countWorkingDays(item.startDate, item.endDate);
    const days = Number.isFinite(Number(item.days)) ? Number(item.days) : calculated;
    const identity = employeeIdentity(item.employeeId);
    return `<div class="leave-modal-backdrop" data-action="close-leave-modal"><div class="leave-modal" role="dialog" aria-modal="true" data-leave-modal-panel>
      <div class="leave-modal-head"><div><h3>${item.id ? 'Edit Data Cuti' : 'Tambah Data Cuti'}</h3><p>Catat data cuti yang sudah diajukan. Tidak ada proses approval atau status pengajuan pada modul ini.</p></div><button class="icon-btn" type="button" data-action="close-leave-modal">×</button></div>
      <div class="leave-modal-grid">
        <div class="field span-2"><label>Pegawai</label>${employeeSelect('leave-employee', item.employeeId || '', false)}<div class="leave-field-note" id="leave-employee-meta">${identity.name ? `${esc(identity.nip ? `NIP ${identity.nip} · ` : '')}${esc(identity.unit || 'Unit/Bidang belum diisi')}` : 'NIP dan Unit/Bidang akan diambil dari master pegawai.'}</div></div>
        <div class="field span-2"><label>Jenis Cuti</label>${typeSelect('leave-type', item.leaveType || 'Cuti Tahunan', false)}</div>
        <div class="field"><label for="leave-start">Tanggal Mulai</label><input id="leave-start" type="date" value="${esc(item.startDate || '')}"></div>
        <div class="field"><label for="leave-end">Tanggal Selesai</label><input id="leave-end" type="date" value="${esc(item.endDate || '')}"></div>
        <div class="field"><label for="leave-days">Jumlah Hari</label><input id="leave-days" type="number" min="0" step="1" value="${esc(days)}"><div class="leave-field-note" id="leave-calculation-note">Hasil kalkulasi: ${calculated} hari kerja (Sabtu/Minggu tidak dihitung).</div></div>
        <div class="field leave-calculation-action"><label>&nbsp;</label><button class="btn btn-secondary btn-sm" type="button" data-action="use-calculated-days">Gunakan Hasil Kalkulasi</button></div>
        <div class="field"><label for="leave-document-number">Nomor Surat/Dokumen <span class="leave-optional">(opsional)</span></label><input id="leave-document-number" type="text" value="${esc(item.documentNumber || '')}"></div>
        <div class="field"><label for="leave-document-date">Tanggal Surat <span class="leave-optional">(opsional)</span></label><input id="leave-document-date" type="date" value="${esc(item.documentDate || '')}"></div>
        <div class="field span-2"><label for="leave-note">Catatan</label><textarea id="leave-note" rows="3" placeholder="Informasi tambahan (opsional)">${esc(item.note || '')}</textarea></div>
      </div>
      <div class="leave-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-leave-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-leave">Simpan</button></div>
    </div></div>`;
  }

  function balanceModal() {
    if (!state.balanceModal) return '';
    const employeeId = state.balanceModal.employeeId;
    const year = Number(state.balanceModal.year);
    const identity = employeeIdentity(employeeId);
    const balance = balanceSummary(employeeId, year);
    return `<div class="leave-modal-backdrop" data-action="close-balance-modal"><div class="leave-modal leave-balance-modal" role="dialog" aria-modal="true" data-balance-modal-panel>
      <div class="leave-modal-head"><div><h3>Edit Saldo Cuti Tahunan</h3><p>${esc(identity.name || 'Pegawai')} · Tahun ${year}</p></div><button class="icon-btn" type="button" data-action="close-balance-modal">×</button></div>
      <div class="leave-modal-grid">
        <div class="field"><label>Hak Cuti</label><input id="balance-entitlement" type="number" min="0" step="1" value="${esc(balance.entitlement)}"></div>
        <div class="field"><label>Saldo dari Tahun Sebelumnya</label><input id="balance-carry" type="number" min="0" step="1" value="${esc(balance.carryOver)}"></div>
        <div class="leave-balance-preview span-2"><span>Penggunaan Cuti Tahunan</span><strong>${balance.usage} hari</strong><span>Sisa setelah disimpan</span><strong id="balance-preview-remaining">${balance.remaining} hari</strong></div>
      </div>
      <div class="leave-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-balance-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-balance">Simpan Saldo</button></div>
    </div></div>`;
  }

  function documentModal() {
    if (!state.documentModal) return '';
    const doc = state.documentModal.document || {};
    const storedFile = doc.file || null;
    const storedFileLabel = storedFile?.itemId
      ? `<div class="leave-current-file"><span>File saat ini</span><strong>${esc(storedFile.name || storedFile.remoteName || 'File SharePoint')}</strong>${storedFile.size ? `<small>${esc(formatFileSize(storedFile.size))}</small>` : ''}</div>`
      : '';
    return `<div class="leave-modal-backdrop" data-action="close-document-modal"><div class="leave-modal leave-document-modal" role="dialog" aria-modal="true" data-document-modal-panel>
      <div class="leave-modal-head"><div><h3>${doc.id ? 'Edit Dokumen Cuti' : 'Tambah Dokumen Cuti'}</h3><p>File bukti dukung disimpan di SharePoint pada folder CUTI. Metadata dokumen tetap dicatat pada aplikasi.</p></div><button class="icon-btn" type="button" data-action="close-document-modal">×</button></div>
      <div class="leave-modal-grid">
        <div class="field"><label>Jenis Dokumen</label><select id="leave-doc-type">${DOCUMENT_TYPES.map((type) => `<option value="${esc(type)}" ${(doc.type || DOCUMENT_TYPES[0]) === type ? 'selected' : ''}>${esc(type)}</option>`).join('')}</select></div>
        <div class="field"><label>Nomor Dokumen</label><input id="leave-doc-number" type="text" value="${esc(doc.number || '')}" placeholder="Opsional"></div>
        <div class="field span-2"><label>Nama Dokumen</label><input id="leave-doc-name" type="text" value="${esc(doc.name || '')}" placeholder="Contoh: Surat Cuti Tahunan"></div>
        <div class="field span-2"><label>File Bukti Dukung${storedFile?.itemId ? ' <span class="leave-optional">(pilih file baru untuk mengganti)</span>' : ''}</label><input id="leave-doc-file" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png">${storedFileLabel}<div class="leave-field-note">Saat upload pertama, aplikasi dapat meminta Anda menghubungkan akun Microsoft 365 Kementerian PKP.</div></div>
        <div class="field span-2"><label>Keterangan</label><textarea id="leave-doc-note" rows="3" placeholder="Opsional">${esc(doc.note || '')}</textarea></div>
      </div>
      <div class="leave-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-document-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-leave-document">Simpan Dokumen</button></div>
    </div></div>`;
  }

  function overlays() {
    let drawer = '';
    if (state.drawer?.type === 'leave') {
      const item = state.leaves.find((row) => row.id === state.drawer.id);
      if (item) drawer = leaveDrawer(item);
    } else if (state.drawer?.type === 'employee') {
      drawer = employeeDrawer(state.drawer.employeeId);
    }
    return `${drawer}${leaveModal()}${balanceModal()}${documentModal()}${state.toast ? `<div class="leave-toast">${esc(state.toast)}</div>` : ''}`;
  }

  function viewMarkup() {
    if (state.view === 'data') return dataView();
    if (state.view === 'kalender') return calendarView();
    if (state.view === 'pegawai') return employeeView();
    if (state.view === 'saldo') return balanceView();
    if (state.view === 'laporan') return reportView();
    return dashboardView();
  }

  function render() {
    if (state.loading) {
      app.innerHTML = window.AppShell.render({
        module: 'leave',
        view: state.view,
        viewLabel: VIEW_LABELS[state.view],
        content: '<div class="card leave-loading"><strong>Memuat modul Cuti</strong>Membaca master pegawai, riwayat cuti, dan saldo cuti dari Cloud Firestore...</div>'
      });
      return;
    }
    const yearLoading = state.loadingYear ? `<div class="alert alert-info"><div class="alert-title">Memuat data tahun ${esc(state.loadingYear)}</div>Riwayat dan saldo cuti sedang diambil dari cache/Cloud Firestore.</div>` : '';
    const content = `${state.error ? `<div class="alert alert-warning"><div class="alert-title">Data belum dapat dimuat sempurna</div>${esc(state.error)}</div>` : ''}${yearLoading}${viewMarkup()}`;
    app.innerHTML = window.AppShell.render({ module: 'leave', view: state.view, viewLabel: VIEW_LABELS[state.view], content, overlays: overlays() });
  }

  async function audit(action, target, detail) {
    try {
      await db().collection('adminAudit').doc().set({
        action: String(action || ''),
        target: String(target || ''),
        detail: String(detail || ''),
        actorUid: user()?.uid || '',
        actorEmail: user()?.email || '',
        createdAt: serverTimestamp()
      });
    } catch (error) {
      console.warn('Audit cuti gagal ditulis:', error);
    }
  }

  async function fetchYearData(year) {
    const numericYear = Number(year || CURRENT_YEAR);
    const [leaveSnap, balanceSnap] = await Promise.all([
      db().collection('leaveRecords').where('leaveYear', '==', numericYear).get(),
      db().collection('leaveBalances').where('year', '==', numericYear).get()
    ]);
    return {
      year: numericYear,
      leaves: leaveSnap.docs.map((doc) => normalizeLeave(doc.data(), doc.id)),
      balances: balanceSnap.docs.map((doc) => normalizeBalance(doc.data(), doc.id))
    };
  }

  function mergeYearData(payload) {
    const year = Number(payload.year);
    state.leaves = state.leaves.filter((item) => Number(item.leaveYear) !== year).concat(payload.leaves || []);
    state.balances = state.balances.filter((item) => Number(item.year) !== year).concat(payload.balances || []);
    state.loadedYears.add(year);
  }

  async function ensureYearData(year, options) {
    const numericYear = Number(year || CURRENT_YEAR);
    const force = Boolean(options?.force);
    if (!force && state.loadedYears.has(numericYear)) return;
    if (options?.showLoading) {
      state.loadingYear = numericYear;
      render();
    }
    try {
      mergeYearData(await fetchYearData(numericYear));
    } catch (error) {
      state.error = error.message || String(error);
    } finally {
      if (state.loadingYear === numericYear) state.loadingYear = null;
      if (options?.showLoading) render();
    }
  }

  async function loadData() {
    state.loading = true;
    state.error = '';
    render();
    try {
      const [masterBundle, currentYearData] = await Promise.all([
        window.MasterDataService.getBundle(false),
        fetchYearData(CURRENT_YEAR)
      ]);
      state.employees = [...masterBundle.employees];
      state.directory = new Map(masterBundle.directory);
      state.leaves = [];
      state.balances = [];
      state.loadedYears = new Set();
      mergeYearData(currentYearData);
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
    clearTimeout(window.__leaveToastTimer);
    window.__leaveToastTimer = setTimeout(() => {
      state.toast = '';
      render();
    }, 3200);
  }

  function openLeaveModal(id) {
    const existing = id ? state.leaves.find((item) => item.id === id) : null;
    state.drawer = null;
    state.modal = {
      record: existing ? { ...existing, documents: (existing.documents || []).map((doc) => ({ ...doc })) } : { leaveType: 'Cuti Tahunan', startDate: '', endDate: '', calculatedDays: 0, days: 0 },
      keepManual: Boolean(existing?.manualOverride)
    };
    render();
  }

  function updateEmployeeMeta() {
    const employeeId = String(document.getElementById('leave-employee')?.value || '');
    const identity = employeeIdentity(employeeId);
    const target = document.getElementById('leave-employee-meta');
    if (target) target.textContent = employeeId ? `${identity.nip ? `NIP ${identity.nip} · ` : ''}${identity.unit || 'Unit/Bidang belum diisi'}` : 'NIP dan Unit/Bidang akan diambil dari master pegawai.';
  }

  function updateLeaveCalculation(forceDays) {
    const startDate = String(document.getElementById('leave-start')?.value || '');
    const endDate = String(document.getElementById('leave-end')?.value || '');
    const calculated = rules.countWorkingDays(startDate, endDate);
    const note = document.getElementById('leave-calculation-note');
    if (note) note.textContent = `Hasil kalkulasi: ${calculated} hari kerja (Sabtu/Minggu tidak dihitung).`;
    if (state.modal) state.modal.record.calculatedDays = calculated;
    const days = document.getElementById('leave-days');
    if (days && (forceDays || !state.modal?.keepManual)) days.value = String(calculated);
    if (forceDays && state.modal) state.modal.keepManual = false;
  }

  async function syncBalanceWithLeaves(employeeId, year, sourceLeaves) {
    if (!employeeId || !year) return;
    const current = storedBalance(employeeId, year) || {};
    const entitlement = Number(current.entitlement || 0);
    const carryOver = Number(current.carryOver || 0);
    const usage = annualUsage(employeeId, year, sourceLeaves);
    const remaining = entitlement + carryOver - usage;
    await db().collection('leaveBalances').doc(balanceDocId(employeeId, year)).set({
      employeeId,
      year: Number(year),
      entitlement,
      carryOver,
      usage,
      remaining,
      schemaVersion: 1,
      updatedAt: serverTimestamp(),
      updatedBy: user()?.email || ''
    }, { merge: true });
  }

  async function saveLeave(button) {
    if (!state.modal) return;
    const employeeId = String(document.getElementById('leave-employee')?.value || '');
    const leaveType = String(document.getElementById('leave-type')?.value || '');
    const startDate = String(document.getElementById('leave-start')?.value || '');
    const endDate = String(document.getElementById('leave-end')?.value || '');
    const calculatedDays = rules.countWorkingDays(startDate, endDate);
    const daysRaw = String(document.getElementById('leave-days')?.value || '');
    const days = Number(daysRaw);
    const documentNumber = String(document.getElementById('leave-document-number')?.value || '').trim();
    const documentDate = String(document.getElementById('leave-document-date')?.value || '');
    const note = String(document.getElementById('leave-note')?.value || '').trim();

    if (!employeeId) { alert('Pegawai wajib dipilih.'); return; }
    if (!LEAVE_TYPES.includes(leaveType)) { alert('Jenis cuti wajib dipilih.'); return; }
    if (!startDate) { alert('Tanggal mulai wajib diisi.'); return; }
    if (!endDate) { alert('Tanggal selesai wajib diisi.'); return; }
    if (endDate < startDate) { alert('Tanggal selesai tidak boleh lebih awal dari tanggal mulai.'); return; }
    if (!Number.isFinite(days) || days < 0) { alert('Jumlah hari harus berupa angka dan tidak boleh negatif.'); return; }

    const existingId = String(state.modal.record?.id || '');
    const oldRecord = existingId ? state.leaves.find((item) => item.id === existingId) : null;
    const ref = existingId ? db().collection('leaveRecords').doc(existingId) : db().collection('leaveRecords').doc();
    const leaveYear = yearFromDate(startDate);
    const payload = {
      employeeId,
      leaveType,
      startDate,
      endDate,
      calculatedDays,
      days,
      manualOverride: days !== calculatedDays,
      leaveYear,
      documentNumber,
      documentDate,
      note,
      documents: oldRecord?.documents || state.modal.record?.documents || [],
      schemaVersion: 1,
      updatedAt: serverTimestamp(),
      updatedBy: user()?.email || ''
    };
    if (!existingId) {
      payload.createdAt = serverTimestamp();
      payload.createdBy = user()?.email || '';
    }

    button.disabled = true;
    button.textContent = 'Menyimpan...';
    try {
      await ref.set(payload, { merge: true });
      const nextRecord = normalizeLeave({ ...payload, createdAt: oldRecord?.createdAt }, ref.id);
      const nextLeaves = state.leaves.filter((item) => item.id !== ref.id).concat(nextRecord);
      const keys = new Set();
      if (oldRecord?.leaveType === 'Cuti Tahunan') keys.add(`${oldRecord.employeeId}|${oldRecord.leaveYear}`);
      if (leaveType === 'Cuti Tahunan') keys.add(`${employeeId}|${leaveYear}`);
      for (const key of keys) {
        const [balanceEmployeeId, balanceYear] = key.split('|');
        await syncBalanceWithLeaves(balanceEmployeeId, Number(balanceYear), nextLeaves);
      }
      await audit(existingId ? 'UPDATE_LEAVE' : 'CREATE_LEAVE', `leaveRecords/${ref.id}`, `${employeeIdentity(employeeId).name || employeeId} · ${leaveType} · ${startDate} s.d. ${endDate}`);
      state.modal = null;
      const yearsToReload = [...new Set([Number(oldRecord?.leaveYear || 0), Number(leaveYear)].filter(Boolean))];
      await Promise.all(yearsToReload.map((year) => ensureYearData(year, { force: true })));
      render();
      showToast(existingId ? 'Data cuti berhasil diperbarui.' : 'Data cuti berhasil ditambahkan.');
    } catch (error) {
      button.disabled = false;
      button.textContent = 'Simpan';
      alert(`Data cuti gagal disimpan: ${error.message || error}`);
    }
  }

  async function deleteLeave(id) {
    const item = state.leaves.find((row) => row.id === id);
    if (!item) return;
    if (!window.confirm(`Hapus data ${item.leaveType} atas nama ${employeeIdentity(item.employeeId).name || 'pegawai'}?`)) return;
    try {
      await db().collection('leaveRecords').doc(id).delete();
      if (item.leaveType === 'Cuti Tahunan') {
        const nextLeaves = state.leaves.filter((row) => row.id !== id);
        await syncBalanceWithLeaves(item.employeeId, item.leaveYear, nextLeaves);
      }

      const remoteFiles = (item.documents || []).map((doc) => doc.file).filter((file) => file?.driveId && file?.itemId);
      let cleanupFailures = [];
      if (remoteFiles.length) {
        try { cleanupFailures = await sharePoint.deleteAttachments(remoteFiles, true); }
        catch (sharePointError) { cleanupFailures = remoteFiles.map((file) => ({ item: file, error: sharePointError })); }
      }

      await audit('DELETE_LEAVE', `leaveRecords/${id}`, `${employeeIdentity(item.employeeId).name || item.employeeId} · ${item.leaveType}`);
      state.drawer = null;
      await ensureYearData(item.leaveYear, { force: true });
      render();
      showToast(cleanupFailures.length ? 'Data cuti dihapus, tetapi sebagian file SharePoint belum berhasil dibersihkan.' : 'Data cuti dan bukti dukung berhasil dihapus.');
    } catch (error) {
      alert(`Data cuti gagal dihapus: ${error.message || error}`);
    }
  }

  function openBalanceModal(employeeId) {
    state.balanceModal = { employeeId, year: Number(state.filters.year) };
    render();
  }

  function updateBalancePreview() {
    if (!state.balanceModal) return;
    const entitlement = Number(document.getElementById('balance-entitlement')?.value || 0);
    const carry = Number(document.getElementById('balance-carry')?.value || 0);
    const usage = annualUsage(state.balanceModal.employeeId, state.balanceModal.year);
    const target = document.getElementById('balance-preview-remaining');
    if (target) target.textContent = `${entitlement + carry - usage} hari`;
  }

  async function saveBalance(button) {
    const ctx = state.balanceModal;
    if (!ctx) return;
    const entitlement = Number(String(document.getElementById('balance-entitlement')?.value || ''));
    const carryOver = Number(String(document.getElementById('balance-carry')?.value || ''));
    if (!Number.isFinite(entitlement) || entitlement < 0 || !Number.isFinite(carryOver) || carryOver < 0) {
      alert('Hak dan saldo cuti harus berupa angka yang valid dan tidak boleh negatif.');
      return;
    }
    const usage = annualUsage(ctx.employeeId, ctx.year);
    const remaining = entitlement + carryOver - usage;
    button.disabled = true;
    try {
      await db().collection('leaveBalances').doc(balanceDocId(ctx.employeeId, ctx.year)).set({
        employeeId: ctx.employeeId,
        year: Number(ctx.year),
        entitlement,
        carryOver,
        usage,
        remaining,
        schemaVersion: 1,
        updatedAt: serverTimestamp(),
        updatedBy: user()?.email || ''
      }, { merge: true });
      await audit('UPDATE_LEAVE_BALANCE', `leaveBalances/${balanceDocId(ctx.employeeId, ctx.year)}`, `${employeeIdentity(ctx.employeeId).name || ctx.employeeId} · ${ctx.year}`);
      state.balanceModal = null;
      await ensureYearData(ctx.year, { force: true });
      render();
      showToast('Saldo cuti berhasil diperbarui.');
    } catch (error) {
      button.disabled = false;
      alert(`Saldo gagal disimpan: ${error.message || error}`);
    }
  }

  function openDocumentModal(leaveId, documentId) {
    const item = state.leaves.find((row) => row.id === leaveId);
    if (!item) return;
    const doc = documentId ? (item.documents || []).find((row) => row.id === documentId) : null;
    state.documentModal = { leaveId, document: doc ? { ...doc } : {} };
    render();
  }

  async function saveLeaveDocument(button) {
    const ctx = state.documentModal;
    if (!ctx) return;
    const item = state.leaves.find((row) => row.id === ctx.leaveId);
    if (!item) return;
    const type = String(document.getElementById('leave-doc-type')?.value || 'Dokumen Lainnya');
    const selectedFile = document.getElementById('leave-doc-file')?.files?.[0] || null;
    const existingFile = ctx.document?.file || null;
    const nameInput = String(document.getElementById('leave-doc-name')?.value || '').trim();
    const name = nameInput || selectedFile?.name || ctx.document?.name || '';
    const number = String(document.getElementById('leave-doc-number')?.value || '').trim();
    const note = String(document.getElementById('leave-doc-note')?.value || '').trim();
    if (!name) { alert('Nama dokumen wajib diisi.'); return; }
    if (!selectedFile && !existingFile?.itemId) { alert('Pilih file bukti dukung yang akan disimpan di SharePoint.'); return; }

    const documents = (item.documents || []).map((doc) => ({ ...doc, file: doc.file ? { ...doc.file } : null }));
    const existingId = String(ctx.document?.id || '');
    const id = existingId || `doc-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    let uploadedFile = null;
    let storedFile = existingFile ? { ...existingFile } : null;
    let firestoreSaved = false;

    button.disabled = true;
    button.textContent = selectedFile ? 'Mengunggah...' : 'Menyimpan...';
    try {
      if (selectedFile) {
        const localAttachment = sharePoint.normalizeAttachment(selectedFile, 'cuti');
        localAttachment.id = `cuti-${id}`;
        uploadedFile = await sharePoint.uploadAttachment(localAttachment, {
          kind: 'cuti',
          path: cutiSharePointPath(item),
          remoteName: cutiRemoteFileName(id, selectedFile)
        }, true);
        storedFile = sharePoint.serializableAttachment(uploadedFile);
      }

      const payload = { id, type, name, number, note, file: storedFile };
      const index = documents.findIndex((doc) => doc.id === id);
      if (index >= 0) documents[index] = payload;
      else documents.push(payload);

      button.textContent = 'Menyimpan...';
      await db().collection('leaveRecords').doc(item.id).set({ documents, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });
      firestoreSaved = true;

      if (selectedFile && existingFile?.driveId && existingFile?.itemId && existingFile.itemId !== storedFile?.itemId) {
        try { await sharePoint.deleteAttachment(existingFile, true); }
        catch (cleanupError) { console.warn('File bukti dukung lama tidak berhasil dihapus dari SharePoint:', cleanupError); }
      }

      await audit(existingId ? 'UPDATE_LEAVE_DOCUMENT' : 'CREATE_LEAVE_DOCUMENT', `leaveRecords/${item.id}`, `${employeeIdentity(item.employeeId).name || item.employeeId} · ${name}`);
      state.documentModal = null;
      await ensureYearData(item.leaveYear, { force: true });
      state.drawer = { type: 'leave', id: item.id };
      render();
      showToast(existingId ? 'Dokumen cuti berhasil diperbarui.' : 'Bukti dukung cuti berhasil diunggah ke SharePoint.');
    } catch (error) {
      if (uploadedFile?.driveId && uploadedFile?.itemId && !firestoreSaved) {
        try { await sharePoint.deleteAttachment(uploadedFile, true); } catch (cleanupError) { console.warn(cleanupError); }
      }
      button.disabled = false;
      button.textContent = 'Simpan Dokumen';
      alert(`Dokumen cuti gagal disimpan: ${error.message || error}`);
    }
  }

  async function downloadLeaveDocument(leaveId, documentId, button) {
    const item = state.leaves.find((row) => row.id === leaveId);
    const doc = (item?.documents || []).find((row) => row.id === documentId);
    if (!doc?.file?.driveId || !doc?.file?.itemId) {
      alert('File SharePoint untuk dokumen ini tidak tersedia.');
      return;
    }
    const originalLabel = button?.textContent || 'Unduh';
    if (button) { button.disabled = true; button.textContent = 'Mengunduh...'; }
    try {
      const file = await sharePoint.downloadAttachment({ ...doc.file, file: null }, true);
      if (!file) throw new Error('File tidak berhasil diunduh dari SharePoint.');
      const url = URL.createObjectURL(file);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = doc.file.name || doc.file.remoteName || doc.name || 'Dokumen_Cuti';
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      alert(`File bukti dukung gagal diunduh: ${error.message || error}`);
    } finally {
      if (button) { button.disabled = false; button.textContent = originalLabel; }
    }
  }

  async function deleteLeaveDocument(leaveId, documentId) {
    const item = state.leaves.find((row) => row.id === leaveId);
    if (!item) return;
    const doc = (item.documents || []).find((row) => row.id === documentId);
    if (!doc) return;
    if (!window.confirm(`Hapus dokumen "${doc.name || doc.type}" beserta file bukti dukungnya?`)) return;
    try {
      const documents = (item.documents || []).filter((row) => row.id !== documentId);
      await db().collection('leaveRecords').doc(item.id).set({ documents, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });

      let cleanupFailed = false;
      if (doc.file?.driveId && doc.file?.itemId) {
        try { await sharePoint.deleteAttachment(doc.file, true); }
        catch (sharePointError) { cleanupFailed = true; console.warn('File SharePoint tidak berhasil dihapus:', sharePointError); }
      }

      await audit('DELETE_LEAVE_DOCUMENT', `leaveRecords/${item.id}`, `${employeeIdentity(item.employeeId).name || item.employeeId} · ${doc.name || doc.type}`);
      await ensureYearData(item.leaveYear, { force: true });
      state.drawer = { type: 'leave', id: item.id };
      render();
      showToast(cleanupFailed ? 'Metadata dokumen dihapus, tetapi file SharePoint belum berhasil dibersihkan.' : 'Dokumen dan file bukti dukung berhasil dihapus.');
    } catch (error) {
      alert(`Dokumen cuti gagal dihapus: ${error.message || error}`);
    }
  }

  function resetDataFilters() {
    state.filters.search = '';
    state.filters.month = '';
    state.filters.unit = '';
    state.filters.type = '';
    state.filters.year = CURRENT_YEAR;
    render();
  }

  function changeCalendarMonth(delta) {
    const date = new Date(Number(state.calendarFilters.year), Number(state.calendarFilters.month) - 1 + delta, 1);
    state.calendarFilters.year = date.getFullYear();
    state.calendarFilters.month = date.getMonth() + 1;
    render();
  }

  function printReport() {
    window.print();
  }

  app.addEventListener('input', (event) => {
    if (event.target?.id === 'leave-search' || event.target?.id === 'leave-employee-search') {
      state.filters.search = String(event.target.value || '');
      const id = event.target.id;
      clearTimeout(window.__leaveSearchTimer);
      window.__leaveSearchTimer = setTimeout(() => {
        render();
        const input = document.getElementById(id);
        if (input) { input.focus(); input.setSelectionRange(input.value.length, input.value.length); }
      }, 240);
      return;
    }
    if (event.target?.id === 'leave-days' && state.modal) {
      state.modal.keepManual = true;
      return;
    }
    if (event.target?.id === 'balance-entitlement' || event.target?.id === 'balance-carry') updateBalancePreview();
  });

  app.addEventListener('change', (event) => {
    const id = event.target?.id || '';
    if (id === 'leave-dashboard-year' || id === 'leave-data-year' || id === 'leave-employee-year' || id === 'leave-balance-year') {
      state.filters.year = Number(event.target.value || CURRENT_YEAR);
      render();
      ensureYearData(state.filters.year, { showLoading: true });
      return;
    }
    if (id === 'leave-data-month') { state.filters.month = String(event.target.value || ''); render(); return; }
    if (id === 'leave-data-unit') { state.filters.unit = String(event.target.value || ''); render(); return; }
    if (id === 'leave-data-type') { state.filters.type = String(event.target.value || ''); render(); return; }
    if (id === 'leave-calendar-month') { state.calendarFilters.month = Number(event.target.value || CURRENT_MONTH); render(); return; }
    if (id === 'leave-calendar-year') { state.calendarFilters.year = Number(event.target.value || CURRENT_YEAR); render(); ensureYearData(state.calendarFilters.year, { showLoading: true }); return; }
    if (id === 'leave-calendar-unit') { state.calendarFilters.unit = String(event.target.value || ''); render(); return; }
    if (id === 'leave-calendar-type') { state.calendarFilters.type = String(event.target.value || ''); render(); return; }
    if (id === 'leave-report-year') {
      const year = Number(event.target.value || CURRENT_YEAR);
      state.reportFilters.year = year;
      state.reportFilters.startDate = `${year}-01-01`;
      state.reportFilters.endDate = `${year}-12-31`;
      render();
      ensureYearData(year, { showLoading: true });
      return;
    }
    if (id === 'leave-report-start') { state.reportFilters.startDate = String(event.target.value || ''); render(); return; }
    if (id === 'leave-report-end') { state.reportFilters.endDate = String(event.target.value || ''); render(); return; }
    if (id === 'leave-report-employee') { state.reportFilters.employeeId = String(event.target.value || ''); render(); return; }
    if (id === 'leave-report-unit') { state.reportFilters.unit = String(event.target.value || ''); render(); return; }
    if (id === 'leave-report-type') { state.reportFilters.type = String(event.target.value || ''); render(); return; }
    if (id === 'leave-employee') { updateEmployeeMeta(); return; }
    if (id === 'leave-start' || id === 'leave-end') { updateLeaveCalculation(false); }
  });

  app.addEventListener('click', (event) => {
    const actionEl = event.target.closest?.('[data-action]');
    const action = actionEl?.dataset.action;
    if (action === 'close-leave-modal' && actionEl.classList.contains('leave-modal-backdrop') && event.target !== actionEl) return;
    if (action === 'close-balance-modal' && actionEl.classList.contains('leave-modal-backdrop') && event.target !== actionEl) return;
    if (action === 'close-document-modal' && actionEl.classList.contains('leave-modal-backdrop') && event.target !== actionEl) return;

    if (action === 'add-leave') { openLeaveModal(''); return; }
    if (action === 'close-leave-modal') { state.modal = null; render(); return; }
    if (action === 'close-balance-modal') { state.balanceModal = null; render(); return; }
    if (action === 'close-document-modal') { state.documentModal = null; render(); return; }
    if (action === 'close-drawer') { state.drawer = null; render(); return; }
    if (action === 'save-leave') { saveLeave(actionEl); return; }
    if (action === 'save-balance') { saveBalance(actionEl); return; }
    if (action === 'save-leave-document') { saveLeaveDocument(actionEl); return; }
    if (action === 'use-calculated-days') { updateLeaveCalculation(true); return; }
    if (action === 'reset-data-filter') { resetDataFilters(); return; }
    if (action === 'calendar-prev') { changeCalendarMonth(-1); return; }
    if (action === 'calendar-next') { changeCalendarMonth(1); return; }
    if (action === 'print-report') { printReport(); return; }

    const mode = event.target.closest?.('[data-calendar-mode]')?.dataset.calendarMode;
    if (mode) { state.calendarMode = mode === 'list' ? 'list' : 'calendar'; render(); return; }

    const detailId = event.target.closest?.('[data-leave-detail]')?.dataset.leaveDetail;
    if (detailId) { state.drawer = { type: 'leave', id: detailId }; render(); return; }

    const employeeDetail = event.target.closest?.('[data-employee-leave-detail]')?.dataset.employeeLeaveDetail;
    if (employeeDetail) { state.drawer = { type: 'employee', employeeId: employeeDetail }; render(); return; }

    const editId = event.target.closest?.('[data-leave-edit]')?.dataset.leaveEdit;
    if (editId) { openLeaveModal(editId); return; }

    const deleteId = event.target.closest?.('[data-leave-delete]')?.dataset.leaveDelete;
    if (deleteId) { deleteLeave(deleteId); return; }

    const balanceEmployee = event.target.closest?.('[data-balance-edit]')?.dataset.balanceEdit;
    if (balanceEmployee) { openBalanceModal(balanceEmployee); return; }

    const addDocId = event.target.closest?.('[data-add-leave-document]')?.dataset.addLeaveDocument;
    if (addDocId) { openDocumentModal(addDocId, ''); return; }

    const editDoc = event.target.closest?.('[data-edit-leave-document]');
    if (editDoc) { openDocumentModal(editDoc.dataset.leaveId, editDoc.dataset.editLeaveDocument); return; }

    const downloadDoc = event.target.closest?.('[data-download-leave-document]');
    if (downloadDoc) { downloadLeaveDocument(downloadDoc.dataset.leaveId, downloadDoc.dataset.downloadLeaveDocument, downloadDoc); return; }

    const deleteDoc = event.target.closest?.('[data-delete-leave-document]');
    if (deleteDoc) { deleteLeaveDocument(deleteDoc.dataset.leaveId, deleteDoc.dataset.deleteLeaveDocument); return; }
  });

  window.addEventListener('hashchange', () => {
    state.view = viewFromHash();
    state.modal = null;
    state.balanceModal = null;
    state.documentModal = null;
    state.drawer = null;
    render();
  });

  async function init() {
    await window.FirebaseClient.requireAdmin();
    state.view = viewFromHash();
    if (!location.hash) history.replaceState(null, '', '#dashboard');
    await loadData();
  }

  init();
})();
