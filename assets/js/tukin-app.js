(function () {
  'use strict';

  const cfg = window.TUKIN_CONFIG;
  const rules = window.TukinRules;
  const parser = window.TukinParser;
  const generator = window.TukinGenerator;
  const storage = window.TukinStorage;
  const resources = window.ResourceLoader;
  const sharePoint = window.SharePointStorage;
  const microsoftAuth = window.MicrosoftAuth;
  const supportIntegration = window.TukinSupportingDataIntegration;
  const app = document.getElementById('app');
  const now = new Date();

  let historyUnsubscribe = null;

  const state = {
    view: 'dashboard',
    step: 'period',
    period: { month: now.getMonth() + 1, year: now.getFullYear() },
    settings: { holidays: [], ramadanEnabled: false, ramadanStart: '', ramadanEnd: '' },
    files: [],
    validationResults: [],
    employees: [],
    busy: false,
    busyAction: '',
    drawerEmployeeKey: null,
    editRecordKey: null,
    editDraft: null,
    editEmployeeKey: null,
    generated: null,
    search: '',
    history: [],
    historyLoaded: false,
    historyBusy: false,
    historyPreview: null,
    editingHistoryId: null,
    editingProcessedAt: null,
    lastSavedRunId: null,
    resultMode: 'new',
    currentRunId: null,
    sharePointRun: null,
    pendingSharePointDeletes: [],
    sharePointBusy: false,
    supportIntegration: { status: 'idle', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: '' }
  };

  function esc(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function money(value) {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency', currency: 'IDR', maximumFractionDigits: 0
    }).format(Number(value || 0));
  }

  function pct(value) {
    return `${Number(value || 0).toFixed(2).replace('.', ',')}%`;
  }

  function bytes(value) {
    const v = Number(value || 0);
    if (v < 1024) return `${v} B`;
    if (v < 1048576) return `${Math.round(v / 1024)} KB`;
    return `${(v / 1048576).toFixed(1)} MB`;
  }

  function periodLabel(period = state.period) {
    return `${cfg.MONTHS[period.month - 1]} ${period.year}`;
  }

  function attendanceRange(period = state.period) {
    return rules.attendancePeriod(period);
  }

  function range() {
    return attendanceRange(state.period);
  }

  function dateInputBounds() {
    const rg = range();
    return { min: rules.dateKey(rg.start), max: rules.dateKey(rg.end) };
  }

  function formatDateTime(iso) {
    if (!iso) return '-';
    try {
      return new Intl.DateTimeFormat('id-ID', {
        timeZone: cfg.TIME_ZONE,
        day: '2-digit', month: 'short', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      }).format(new Date(iso));
    } catch (_) {
      return new Date(iso).toLocaleString('id-ID');
    }
  }

  function employeeByKey(key) {
    return state.employees.find((employee) => rules.employeeKey(employee) === key) || null;
  }

  function recordByKey(employee, key) {
    return employee?.records?.[key] || null;
  }

  function calcAll() {
    state.employees.forEach((employee) => parser.recalculateEmployee(employee, state.settings));
    state.generated = null;
  }

  function icon(name) {
    const icons = {
      process: '<path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M5 18v3h14v-3"/>',
      history: '<path d="M4 6h16"/><path d="M4 12h16"/><path d="M4 18h10"/>',
      tukin: '<path d="M5 4h14v16H5z"/><path d="M8 8h8M8 12h8M8 16h5"/>',
      overtime: '<path d="M12 7v5l3 2"/><circle cx="12" cy="12" r="8"/>',
      chevron: '<path d="m9 18 6-6-6-6"/>',
      home: '<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M9 21v-6h6v6"/>'
    };
    return `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[name] || icons.tukin}</svg>`;
  }

  function attachmentOpenLink(item, label) {
    if (!item?.webUrl) return '';
    return `<a class="text-link-btn" href="${esc(item.webUrl)}" target="_blank" rel="noopener noreferrer">${esc(label || 'Buka SharePoint')} ↗</a>`;
  }

  function sharePointPanel() {
    const status = microsoftAuth?.getStatus?.() || { configured: false, ready: false, libraryLoaded: false, connected: false };
    const d = status.diagnostics || {};
    if (!status.configured) {
      const missing = [];
      if (!d.configLoaded) missing.push('microsoft-config.js tidak termuat');
      if (!d.tenantIdPresent) missing.push('tenantId');
      if (!d.clientIdPresent) missing.push('clientId');
      if (!d.shareUrlPresent) missing.push('shareUrl');
      return `<div class="alert alert-warning"><div class="alert-title">SharePoint belum dikonfigurasi</div>Periksa <code>assets/js/microsoft-config.js</code>. Komponen yang belum tersedia: <strong>${esc(missing.join(', ') || 'konfigurasi Microsoft')}</strong>.</div>`;
    }
    if (!status.libraryLoaded) {
      return `<div class="alert alert-warning"><div class="alert-title">MSAL belum termuat</div><code>tenantId</code> dan <code>clientId</code> sudah terbaca, tetapi <code>assets/vendor/msal-browser.bundle.js</code> belum tersedia/termuat. Jalankan <code>npm install</code> lalu <code>npm run build</code>, kemudian deploy ulang.</div>`;
    }
    if (status.connected) {
      return `<div class="alert alert-success"><div class="alert-title">SharePoint terhubung</div>${esc(status.name || status.email || 'Akun Microsoft 365')} ${status.email ? `· ${esc(status.email)}` : ''}${state.sharePointRun?.runFolderWebUrl ? ` · <a href="${esc(state.sharePointRun.runFolderWebUrl)}" target="_blank" rel="noopener noreferrer">Buka folder run ↗</a>` : ''}</div>`;
    }
    return `<div class="alert alert-warning"><div class="alert-title">Hubungkan SharePoint</div>Hubungkan akun Microsoft 365 Kementerian PKP agar file Excel asli dan bukti dukung tersimpan dan dapat dibuka lintas perangkat.<div style="margin-top:10px"><button class="btn btn-secondary btn-sm" data-action="connect-sharepoint" type="button" ${state.sharePointBusy ? 'disabled' : ''}>${state.sharePointBusy ? 'Menghubungkan...' : 'Hubungkan SharePoint'}</button></div></div>`;
  }

  function shell(content) {
    const view = ['dashboard', 'history', 'process'].includes(state.view) ? state.view : 'dashboard';
    const leaf = view === 'dashboard' ? 'Dashboard' : view === 'history' ? 'Riwayat' : 'Proses Tukin';
    return window.AppShell.render({
      module: 'tukin',
      view,
      viewLabel: leaf,
      content,
      overlays: `${renderDrawer()}${renderEditModal()}${renderEmployeeModal()}${renderHistoryPreview()}`
    });
  }

  function stepper() {
    const steps = [['period', 'Periode'], ['upload', 'Upload'], ['validation', 'Validasi / Simpan']];
    const idx = steps.findIndex((item) => item[0] === state.step);
    return `<div class="card tukin-stepper">${steps.map((step, i) => `${i ? '<div class="tukin-step-line"></div>' : ''}<div class="tukin-step ${i < idx ? 'done' : i === idx ? 'active' : ''}"><span class="tukin-step-dot">${i < idx ? '✓' : i + 1}</span><span>${step[1]}</span></div>`).join('')}</div>`;
  }

  function render() {
    let body = '';
    if (state.view === 'dashboard') {
      body = renderDashboard();
    } else if (state.view === 'history') {
      body = renderHistory();
    } else {
      if (state.step === 'period') body = renderPeriod();
      else if (state.step === 'upload') body = renderUpload();
      else if (state.step === 'validation') body = renderValidation();
      else body = renderResult();
      body = `${stepper()}${body}`;
    }
    app.innerHTML = shell(`${sharePointPanel()}${body}`);
    bind();
  }

  function dashboardMetric(value, label, suffix = '') {
    return `<div class="card metric"><div class="metric-top"><div class="metric-label">${esc(label)}</div><div class="metric-icon">${icon('tukin')}</div></div><div class="metric-value">${value}${suffix ? `<span class="metric-suffix">${esc(suffix)}</span>` : ''}</div></div>`;
  }

  function renderDashboard() {
    const history = state.history || [];
    const latest = history[0] || null;
    const summary = latest?.summary || {};
    const recent = history.slice(0, 5);
    const recentRows = recent.length ? recent.map((item) => {
      const itemSummary = item.summary || {};
      return `<div class="dashboard-history-row">
        <div><div class="history-period-name">${esc(periodLabel(item.period))}</div><div class="history-period-sub">Tunjangan Kinerja</div></div>
        <div>${Number(itemSummary.employees || 0)}</div>
        <div class="tukin-money">${money(itemSummary.totalCutAmount || 0)}</div>
        <div>${esc(formatDateTime(item.updatedAt || item.processedAt))}</div>
        <div><span class="status-pill">Selesai</span></div>
        <button class="chev-btn" type="button" data-view-tukin-history="${esc(item.id)}" aria-label="Lihat ${esc(periodLabel(item.period))}">${icon('chevron')}</button>
      </div>`;
    }).join('') : '<div class="empty-state"><strong>Belum ada riwayat Tukin</strong>Proses Tunjangan Kinerja pertama akan tampil di sini.</div>';

    return `<div class="page-head"><div><h1>Dashboard Tunjangan Kinerja</h1><div class="small">Ringkasan proses dan hasil Tunjangan Kinerja.</div></div><a class="btn btn-primary" href="tukin.html#process">+ Proses Tukin</a></div>
      <div class="dashboard-metrics">
        ${dashboardMetric(Number(summary.employees || 0), 'Pegawai Periode Terakhir')}
        ${dashboardMetric(money(summary.totalCutAmount || 0), 'Total Potongan')}
        ${dashboardMetric(Number(summary.adjustedRecords || 0), 'Koreksi')}
        ${dashboardMetric(Number(summary.flaggedRecords || 0), 'Perlu Verifikasi')}
      </div>
      <div class="dashboard-grid"><div class="card period-card"><div class="dashboard-card-head"><div class="card-title">Periode Terakhir</div><div class="dashboard-card-note">${latest ? 'Selesai' : '—'}</div></div><div class="period-box"><div class="period-label">Periode Tukin</div><div class="period-main">${latest ? esc(periodLabel(latest.period)) : 'Belum ada'}</div><div class="period-meta"><div><div class="period-meta-value">${Number(summary.employees || 0)}</div><div class="period-meta-label">Pegawai</div></div><div><div class="period-meta-value">${Number(summary.adjustedRecords || 0)}</div><div class="period-meta-label">Koreksi</div></div></div></div><div class="period-actions">${latest ? `<button class="text-link-btn" data-view-tukin-history="${esc(latest.id)}" type="button">Lihat hasil →</button>` : ''}</div></div></div>
      <div class="card dashboard-history"><div class="dashboard-history-head"><div class="card-title">Riwayat Tukin Terbaru</div>${history.length ? '<a class="text-link-btn" href="tukin.html#history">Lihat semua</a>' : ''}</div>${history.length ? '<div class="dashboard-history-cols"><div>Periode</div><div>Pegawai</div><div>Total Potongan</div><div>Diproses</div><div>Status</div><div></div></div>' : ''}${recentRows}</div>`;
  }

  function renderPeriod() {
    const months = cfg.MONTHS.map((month, i) => `<option value="${i + 1}" ${state.period.month === i + 1 ? 'selected' : ''}>${month}</option>`).join('');
    const years = Array.from({ length: 8 }, (_, i) => now.getFullYear() - 3 + i)
      .map((year) => `<option value="${year}" ${state.period.year === year ? 'selected' : ''}>${year}</option>`).join('');
    const rg = range();
    const bounds = dateInputBounds();
    const chips = state.settings.holidays.length
      ? state.settings.holidays.map((key) => `<span class="tukin-chip">${esc(rules.formatDate(rules.dateFromKey(key), true))}<button data-remove-holiday="${esc(key)}" type="button">×</button></span>`).join('')
      : '<span class="card-subtitle">Belum ada tanggal merah tambahan.</span>';

    return `<div class="page-title"><div><h2>Proses Tunjangan Kinerja</h2><p>Pilih periode pembayaran. Dasar absensi ditetapkan otomatis tanggal 11 dua bulan sebelumnya sampai tanggal 10 bulan sebelumnya.</p></div></div>
      ${state.editingHistoryId ? '<div class="alert alert-warning"><div class="alert-title">Mode edit riwayat</div>Perubahan akan memperbarui riwayat yang sama setelah tombol Simpan digunakan. Download hanya mengunduh paket ZIP ke perangkat.</div>' : ''}
      <div class="card card-pad">
        <div class="card-title">Periode Pembayaran</div>
        <div class="form-grid section-gap"><div class="field"><label>Bulan Tukin</label><select id="tukin-month">${months}</select></div><div class="field"><label>Tahun</label><select id="tukin-year">${years}</select></div></div>
        <div class="tukin-period-box"><div class="tukin-readonly"><span class="card-subtitle">Periode Tukin</span><strong>${esc(periodLabel())}</strong></div><div class="tukin-readonly"><span class="card-subtitle">Dasar Absensi</span><strong>${esc(rules.formatPeriodRange(rg))}</strong></div></div>
        <div class="tukin-setting"><div class="card-title">Tanggal Merah Tambahan</div><div class="card-subtitle">Tambahkan manual bila hari libur tidak tercatat pada file presensi.</div><div class="tukin-date-row"><div class="field"><label>Tanggal</label><input id="holiday-date" type="date" min="${bounds.min}" max="${bounds.max}"></div><button class="btn btn-secondary" data-action="add-holiday" type="button">+ Tambah Tanggal Merah</button></div><div class="tukin-chip-list">${chips}</div></div>
        <div class="tukin-setting"><div class="tukin-setting-head"><div><div class="card-title">Jam Kerja Ramadan</div><div class="card-subtitle">Senin–Kamis 08.00–15.00, Jumat 08.00–15.30. Fleksibilitas 60 menit mengikuti jam masuk Ramadan.</div></div><label class="tukin-toggle"><input id="ramadan-enabled" type="checkbox" ${state.settings.ramadanEnabled ? 'checked' : ''}> Aktifkan</label></div>
          <div class="tukin-date-row ${state.settings.ramadanEnabled ? '' : 'hidden'}" id="ramadan-range"><div class="field"><label>Mulai Ramadan</label><input id="ramadan-start" type="date" min="${bounds.min}" max="${bounds.max}" value="${esc(state.settings.ramadanStart)}"></div><div class="field"><label>Selesai Ramadan</label><input id="ramadan-end" type="date" min="${bounds.min}" max="${bounds.max}" value="${esc(state.settings.ramadanEnd)}"></div></div>
        </div>
        <div class="actions"><span></span><button class="btn btn-primary" data-action="period-next" type="button">${state.employees.length ? 'Kembali ke Validasi →' : 'Lanjut Upload →'}</button></div>
      </div>`;
  }

  function renderUpload() {
    return `<div class="page-title"><div><h2>Upload Absensi Pegawai</h2><p>Tukin ${esc(periodLabel())} · dasar absensi ${esc(rules.formatPeriodRange(range()))}.</p></div></div>
      <div class="card card-pad"><div id="tukin-dropzone" class="dropzone"><div class="drop-icon">⇧</div><h3>Tarik file Excel absensi per pegawai ke sini</h3><p>Dapat memilih banyak file .xlsx/.xls sekaligus.</p><div style="margin-top:14px"><button class="btn btn-secondary btn-sm" data-action="choose-files" type="button">Pilih File</button></div><input id="tukin-files" class="file-input" type="file" multiple accept=".xlsx,.xls"></div>
      ${state.files.length ? `<div class="section-gap table-wrap"><table class="file-table"><thead><tr><th>No.</th><th>Nama File</th><th>Ukuran</th><th>Status</th><th></th></tr></thead><tbody>${state.files.map((file, i) => `<tr><td>${i + 1}</td><td>${esc(file.name)}</td><td>${bytes(file.size)}</td><td><span class="status ok">● Siap</span></td><td class="text-right"><button class="btn btn-secondary btn-sm" data-remove-file="${i}" type="button">Hapus</button></td></tr>`).join('')}</tbody></table></div>` : ''}
      <div class="actions"><button class="btn btn-secondary" data-action="back-period" type="button">← Kembali</button><div class="actions-right">${state.files.length ? '<button class="btn btn-secondary" data-action="choose-files" type="button">+ Tambah File</button><button class="btn btn-primary" data-action="calculate" type="button">Hitung & Validasi →</button>' : ''}</div></div></div>`;
  }

  function supportIntegrationAlert() {
    const info = state.supportIntegration || {};
    if (info.status === 'loading') {
      return `<div class="alert alert-info"><div class="alert-title">Mencocokkan Cuti & Surat Tugas</div>Data pendukung pegawai sedang dicocokkan dengan tanggal yang memerlukan verifikasi.</div>`;
    }
    if (info.status === 'error') {
      return `<div class="alert alert-warning"><div class="alert-title">Integrasi data pendukung belum diterapkan</div>${esc(info.message || 'Data Cuti/Surat Tugas tidak dapat dibaca oleh akun ini.')} Koreksi manual tetap dapat dilakukan.</div>`;
    }
    if (info.status !== 'applied') return '';

    const adjusted = Number(info.adjustedRecords || 0);
    const cutiAdjusted = Number(info.cutiAdjustedRecords || 0);
    const assignmentAdjusted = Number(info.assignmentAdjustedRecords || 0);
    const assignmentInformational = Number(info.assignmentInformationalRecords || 0);
    const conflicts = Number(info.conflictRecords || 0);
    const manual = Number(info.preservedManualRecords || 0);
    const evidence = Number(info.evidenceFiles || 0);
    if (!adjusted && !assignmentInformational && !conflicts && !manual) return '';

    const success = adjusted
      ? `<div class="alert alert-success"><div class="alert-title">Penyesuaian otomatis dari data pendukung</div>${adjusted} tanggal otomatis disesuaikan menjadi <strong>0%</strong>${cutiAdjusted ? ` · ${cutiAdjusted} dari Cuti` : ''}${assignmentAdjusted ? ` · ${assignmentAdjusted} dari Surat Tugas Dinas Luar / Perjalanan Dinas` : ''}. Nilai hasil presensi tetap ditampilkan pada kolom <strong>Otomatis</strong>.${evidence ? ` ${evidence} file bukti dukung unik siap disalin satu kali saat Simpan.` : ''}${manual ? ` ${manual} koreksi manual dipertahankan.` : ''}</div>`
      : '';
    const information = assignmentInformational
      ? `<div class="alert alert-info"><div class="alert-title">${assignmentInformational} tanggal memiliki Surat Tugas sebagai data pendukung</div>Jenis <strong>Pengaturan Kerja</strong> atau <strong>Lainnya</strong> tidak mengubah persentase secara otomatis. Jika diperlukan, penyesuaian dapat dilakukan melalui tombol <strong>Edit</strong> pada Proses Tukin.</div>`
      : '';
    const conflict = conflicts
      ? `<div class="alert alert-warning"><div class="alert-title">${conflicts} konflik Cuti dan Surat Tugas</div>Pada tanggal yang sama ditemukan Cuti dan Surat Tugas. Sistem tidak menerapkan 0% secara otomatis; tanggal tersebut tetap memerlukan verifikasi manual.</div>`
      : '';
    return `${success}${information}${conflict}`;
  }

  async function syncSupportingAdjustments(options) {
    if (!supportIntegration?.apply || !state.employees.length) return null;
    const opts = options || {};
    state.supportIntegration = { ...state.supportIntegration, status: 'loading', message: '' };
    if (opts.render !== false) render();
    try {
      const result = await supportIntegration.apply(state.employees, state.period);
      const removedEvidence = Array.isArray(result?.removedEvidence) ? result.removedEvidence : [];
      if (removedEvidence.length) {
        const queuedIds = new Set(state.pendingSharePointDeletes.map((item) => item?.itemId).filter(Boolean));
        for (const item of removedEvidence) {
          if (!item?.driveId || !item?.itemId || queuedIds.has(item.itemId)) continue;
          state.pendingSharePointDeletes.push({ ...item, file: null });
          queuedIds.add(item.itemId);
        }
      }
      const { removedEvidence: _removedEvidence, ...status } = result || {};
      state.supportIntegration = { ...status, message: '' };
      state.generated = null;
      return status;
    } catch (error) {
      console.warn('Integrasi data pendukung ke Tukin tidak dapat diterapkan:', error);
      const code = String(error?.code || '');
      const permissionMessage = code.includes('permission-denied')
        ? 'Akun yang digunakan tidak memiliki hak baca terhadap data Cuti/Surat Tugas.'
        : (error?.message || String(error));
      state.supportIntegration = { status: 'error', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: permissionMessage };
      return null;
    } finally {
      if (opts.render !== false) render();
    }
  }

  function renderValidation() {
    const errors = state.validationResults.filter((item) => !item.ok);
    const summaries = state.employees.map((employee) => ({ employee, summary: rules.summarizeEmployee(employee) }));
    const totalCut = summaries.reduce((sum, item) => sum + item.summary.cutAmount, 0);
    const totalNeed = summaries.reduce((sum, item) => sum + item.summary.flaggedRecords, 0);
    const adjusted = summaries.reduce((sum, item) => sum + item.summary.adjustedRecords, 0);
    const query = state.search.trim().toLowerCase();
    const rows = summaries.filter((item) => !query || item.employee.name.toLowerCase().includes(query) || String(item.employee.nip).includes(query));

    return `<div class="page-title"><div><h2>Validasi & Verifikasi Perhitungan</h2><p>Periksa hasil perhitungan otomatis. Buka Detail untuk menelusuri tanggal, alasan potongan, koreksi dan bukti dukung.</p></div></div>
      ${state.editingHistoryId ? `<div class="alert alert-warning"><div class="alert-title">Mengedit riwayat ${esc(periodLabel())}</div>Setelah koreksi selesai, klik Simpan untuk memperbarui riwayat. Gunakan Download bila ingin mengunduh paket ZIP ke perangkat.</div>` : ''}
      ${supportIntegrationAlert()}
      ${errors.length ? `<div class="alert alert-danger"><div class="alert-title">${errors.length} file bermasalah</div>File bermasalah harus diperbaiki/dihapus sebelum data disimpan atau didownload.</div>` : ''}
      <div class="tukin-kpi section-gap"><div class="card"><span>Pegawai</span><strong>${state.employees.length}</strong></div><div class="card"><span>Perlu Diverifikasi</span><strong>${totalNeed}</strong></div><div class="card"><span>Sudah Dikoreksi</span><strong>${adjusted}</strong></div><div class="card"><span>Total Potongan Rp</span><strong>${money(totalCut)}</strong></div></div>
      <div class="toolbar"><div class="search"><input id="tukin-search" placeholder="Cari nama atau NIP" value="${esc(state.search)}"></div><div class="card-subtitle">% SKP sementara 0% untuk seluruh pegawai</div></div>
      <div class="table-wrap"><table class="data-table"><thead><tr><th>No</th><th>Nama / NIP</th><th>Anak Satker</th><th>Hari Kerja</th><th>% Pot. Absensi</th><th>% Pot. SKP</th><th>% Pot. Final</th><th>Besaran Tukin</th><th>Potongan Tukin</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${rows.map((item, i) => {
        const employee = item.employee;
        const summary = item.summary;
        return `<tr class="${employee.masterMatched ? '' : 'tukin-master-miss'}"><td>${i + 1}</td><td><strong>${esc(employee.name)}</strong><div class="card-subtitle">${esc(employee.nip || '-')}</div></td><td>${esc(employee.anakSatker || '-')}</td><td>${summary.workDays}</td><td class="tukin-pct ${summary.attendancePercent ? 'tukin-danger' : ''}">${pct(summary.attendancePercent)}</td><td>${pct(summary.skpPercent)}</td><td class="tukin-pct">${pct(summary.finalPercent)}</td><td class="tukin-money">${money(summary.tukin)}</td><td class="tukin-money">${money(summary.cutAmount)}</td><td>${summary.flaggedRecords ? `<span class="tukin-badge warn">${summary.flaggedRecords} perlu cek</span>` : '<span class="tukin-badge">Tidak ada isu</span>'}${summary.adjustedRecords ? ` <span class="tukin-badge edit">${summary.adjustedRecords} koreksi</span>` : ''}${employee.masterMatched ? '' : ' <span class="tukin-badge warn">Master belum cocok</span>'}</td><td><button class="btn btn-secondary btn-sm" data-detail-employee="${esc(rules.employeeKey(employee))}" type="button">Detail</button></td></tr>`;
      }).join('') || '<tr><td colspan="11" class="text-center">Tidak ada data.</td></tr>'}</tbody></table></div>
      ${state.validationResults.length ? `<div class="section-gap card"><div class="table-wrap" style="border:0"><table class="file-table"><thead><tr><th>File</th><th>Status</th><th>Keterangan</th></tr></thead><tbody>${state.validationResults.map((result) => `<tr><td>${esc(result.fileName)}</td><td><span class="status ${result.ok ? 'ok' : 'error'}">${result.ok ? '● Berhasil' : '● Gagal'}</span></td><td>${esc(result.ok ? (result.warnings?.join(' · ') || `Pegawai: ${result.employee}`) : result.error)}</td></tr>`).join('')}</tbody></table></div></div>` : ''}
      <div class="actions"><button class="btn btn-secondary" data-action="back-upload" type="button">← Kembali ke Upload</button><div class="actions-right"><button class="btn btn-secondary" data-action="sync-cuti" type="button" ${state.busy || state.supportIntegration?.status === 'loading' ? 'disabled' : ''}>Sinkronkan Cuti & Surat Tugas</button>${state.employees.some((employee) => (employee.sourceFiles || []).some((item) => item?.driveId && item?.itemId)) ? `<button class="btn btn-secondary" data-action="reload-sharepoint-attendance" type="button" ${state.busy ? 'disabled' : ''}>Muat Ulang Absensi dari SharePoint</button>` : ''}<button class="btn btn-secondary" data-action="settings" type="button">Atur Periode/Tanggal Merah</button><button class="btn btn-primary" data-action="save-run" type="button" ${errors.length || !state.employees.length || state.busy || state.supportIntegration?.status === 'loading' ? 'disabled' : ''}>${state.busy && state.busyAction === 'save' ? 'Menyimpan...' : 'Simpan'}</button><button class="btn btn-secondary" data-action="download-run" type="button" ${errors.length || !state.employees.length || state.busy || state.supportIntegration?.status === 'loading' ? 'disabled' : ''}>${state.busy && state.busyAction === 'download' ? 'Menyiapkan...' : 'Download'}</button></div></div>`;
  }

  function renderResult() {
    const total = state.employees.reduce((sum, employee) => sum + rules.summarizeEmployee(employee).cutAmount, 0);
    const updated = state.resultMode === 'updated';
    return `<div class="page-title"><div><h2>${updated ? 'Riwayat Tunjangan Kinerja Diperbarui' : 'Dokumen Tunjangan Kinerja Siap'}</h2><p>${updated ? 'Koreksi terakhir telah disimpan dan paket ZIP baru telah dibuat.' : 'Paket ZIP sudah dibuat dari hasil verifikasi terakhir. Data perhitungan tersinkron ke Firestore; file Excel asli, bukti dukung, dan rekap tersimpan di SharePoint.'}</p></div></div><div class="card result-hero"><div class="success-mark">✓</div><h3>${updated ? 'Perubahan tersimpan' : 'Generate selesai'}</h3><p>${esc(state.generated?.name || generator.zipFileName(state.period))}</p><div class="tukin-result-grid"><div><span>Periode Tukin</span><strong>${esc(periodLabel())}</strong></div><div><span>Dasar Absensi</span><strong>${esc(rules.formatPeriodRange(range()))}</strong></div><div><span>Jumlah Pegawai</span><strong>${state.employees.length}</strong></div><div><span>Total Potongan Tukin</span><strong>${money(total)}</strong></div></div><div class="download-list"><div class="download-row"><div class="file-icon">ZIP</div><div><div class="download-name">${esc(state.generated?.name || 'Paket Tunjangan Kinerja.zip')}</div><div class="download-meta">Berisi rekap Excel, folder pegawai, file absensi asli dan bukti dukung koreksi.</div></div>${state.generated ? '<button class="btn btn-primary btn-sm" data-action="download-again" type="button">Unduh Lagi</button>' : ''}</div></div><div class="actions"><button class="btn btn-secondary" data-action="edit-again" type="button">← Edit / Verifikasi Kembali</button><div class="actions-right"><button class="btn btn-secondary" data-action="go-tukin-history" type="button">Lihat Riwayat</button><button class="btn btn-primary" data-action="new-process" type="button">Proses Periode Baru</button></div></div></div>`;
  }

  function renderHistory() {
    if (state.historyBusy && !state.historyLoaded) {
      return '<div class="page-title"><div><h2>Riwayat Tunjangan Kinerja</h2><p>Memuat riwayat proses bersama dari Cloud Firestore.</p></div></div><div class="card card-pad text-center">Memuat riwayat...</div>';
    }
    const rows = state.history.map((item) => {
      const rg = attendanceRange(item.period);
      const summary = item.summary || {};
      const changedAt = item.updatedAt || item.processedAt;
      const changedLabel = item.updatedAt ? 'Diubah' : 'Diproses';
      const manageable = Boolean(storage.canManage?.(item));
      const actorName = item.ownerName || item.ownerEmail || 'Akun lama';
      const actorSub = item.ownerName && item.ownerEmail ? `<div class="card-subtitle">${esc(item.ownerEmail)}</div>` : '';
      const manageActions = manageable
        ? `<button class="btn btn-secondary btn-sm" data-edit-tukin-history="${esc(item.id)}" type="button">Verifikasi/Edit</button><button class="btn btn-danger btn-sm" data-delete-tukin-history="${esc(item.id)}" type="button">Hapus</button>`
        : '';
      return `<tr><td><strong>${esc(periodLabel(item.period))}</strong><div class="card-subtitle">${esc(rules.formatPeriodRange(rg))}</div></td><td>${Number(summary.employees || 0)}</td><td class="tukin-money">${money(summary.totalCutAmount || 0)}</td><td>${Number(summary.adjustedRecords || 0)}</td><td><strong>${esc(actorName)}</strong>${actorSub}</td><td><span class="history-time-label">${changedLabel}</span><br>${esc(formatDateTime(changedAt))}</td><td><span class="status-pill">Selesai</span></td><td><div class="history-actions tukin-history-actions"><button class="btn btn-secondary btn-sm" data-view-tukin-history="${esc(item.id)}" type="button">Lihat</button>${manageActions}${item.sharePoint?.runFolderWebUrl ? `<a class="btn btn-secondary btn-sm" href="${esc(item.sharePoint.runFolderWebUrl)}" target="_blank" rel="noopener noreferrer">SharePoint ↗</a>` : ''}<button class="btn btn-secondary btn-sm" data-download-tukin-history="${esc(item.id)}" type="button">Download</button></div></td></tr>`;
    }).join('');

    return `<div class="page-title"><div><h2>Riwayat Tunjangan Kinerja</h2><p>Riwayat bersama seluruh akun aplikasi. Data perhitungan tersimpan di Firestore dan file biner tersimpan di SharePoint sehingga dapat diakses lintas akun/perangkat sesuai hak akses Microsoft 365.</p></div><button class="btn btn-primary" data-action="new-process-from-history" type="button">+ Proses Tukin Baru</button></div>
      <div class="card table-wrap"><table class="data-table history-table"><thead><tr><th>Periode Tukin / Dasar Absensi</th><th>Pegawai</th><th>Total Potongan</th><th>Koreksi</th><th>Diproses Oleh</th><th>Terakhir Diubah</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${rows || '<tr><td colspan="8" class="text-center" style="padding:36px;color:#6b7c93">Belum ada riwayat Tunjangan Kinerja.</td></tr>'}</tbody></table></div>
      <div class="footer-note">Data perhitungan Tukin tersinkron melalui Cloud Firestore. File Excel asli dan bukti dukung menggunakan SharePoint sebagai penyimpanan utama; IndexedDB tetap dipakai sebagai cache lokal.</div>`;
  }

  function assignmentReferenceMarkup(record) {
    const items = record?.assignmentAdjustment?.assignments || [];
    if (!items.length) return '';
    return items.map((item) => `<div class="tukin-cuti-ref"><span class="tukin-badge edit">Surat Tugas</span><span>${esc(item.letterNumber || 'Tanpa nomor')} · ${esc(item.assignmentTypeLabel || 'Dinas Luar / Perjalanan Dinas')} · ${esc(item.startDate || '')}${item.endDate && item.endDate !== item.startDate ? ` s.d. ${esc(item.endDate)}` : ''}${item.activity ? ` · ${esc(item.activity)}` : ''}</span></div>`).join('');
  }

  function adjustmentLabel(record) {
    if (record?.adjustmentSource === 'cuti') return 'disesuaikan otomatis · cuti';
    if (record?.adjustmentSource === 'surat_tugas') return 'disesuaikan otomatis · surat tugas';
    if (record?.adjustmentSource === 'conflict') return 'konflik · perlu verifikasi';
    return record?.adjustedPercent !== null && record?.adjustedPercent !== undefined ? 'disesuaikan manual' : '';
  }

  function renderDrawer() {
    if (!state.drawerEmployeeKey) return '';
    const employee = employeeByKey(state.drawerEmployeeKey);
    if (!employee) return '';
    const summary = rules.summarizeEmployee(employee);
    const records = Object.values(employee.records || {})
      .filter((record) => record.needsVerification || record.adjustedPercent !== null || (record.evidence || []).length || record.adjustmentNote)
      .sort((a, b) => a.key.localeCompare(b.key));

    return `<div class="drawer-backdrop" data-action="close-drawer"></div><aside class="drawer tukin-drawer"><div class="tukin-drawer-header"><div class="drawer-header"><div><h3>Detail Potongan Pegawai</h3><div class="card-subtitle">Log tanggal yang menyebabkan potongan atau memerlukan verifikasi.</div></div><button class="icon-btn" data-action="close-drawer" type="button">×</button></div><div class="profile"><div class="profile-avatar">${esc(employee.name.charAt(0).toUpperCase())}</div><div><div class="profile-name">${esc(employee.name)}</div><div class="profile-meta">NIP ${esc(employee.nip || '-')} · Anak Satker ${esc(employee.anakSatker || '-')}</div></div><button class="btn btn-secondary btn-sm" style="margin-left:auto" data-edit-employee="${esc(rules.employeeKey(employee))}" type="button">Data Pegawai</button></div></div>
      <div class="tukin-detail-summary"><div><span>Pot. Absensi</span><strong>${pct(summary.attendancePercent)}</strong></div><div><span>Pot. SKP</span><strong>${pct(summary.skpPercent)}</strong></div><div><span>Pot. Final</span><strong>${pct(summary.finalPercent)}</strong></div><div><span>Potongan Rupiah</span><strong>${money(summary.cutAmount)}</strong></div></div>
      ${(employee.sourceFiles || []).length ? `<div class="card card-pad" style="margin-bottom:14px"><div class="card-title">File Absensi Asli</div><div class="tukin-files">${employee.sourceFiles.map((file) => `<div>${esc(file.name || file.file?.name || 'File Absensi')} ${attachmentOpenLink(file, 'Buka')}</div>`).join('')}</div></div>` : ''}
      <div class="table-wrap"><table class="data-table"><thead><tr><th>Tanggal</th><th>Masuk</th><th>Pulang</th><th>Status</th><th>TL</th><th>PSW</th><th>Otomatis</th><th>Final</th><th>Alasan / Bukti</th><th>Aksi</th></tr></thead><tbody>${records.map((record) => `<tr><td class="nowrap">${esc(rules.formatDate(record.date, true))}</td><td>${esc(rules.formatMinutes(record.inMinutes))}</td><td>${esc(rules.formatMinutes(record.outMinutes))}</td><td>${esc(record.status || '-')}${record.flags?.length ? `<div class="card-subtitle">${esc(record.flags.join(', '))}</div>` : ''}</td><td>${esc(record.tlCategory)}<div class="card-subtitle">${pct(record.tlPercent)}</div></td><td>${esc(record.pswCategory)}<div class="card-subtitle">${pct(record.pswPercent)}</div></td><td class="tukin-pct">${pct(record.autoTotalPercent)}</td><td class="tukin-pct ${record.adjustedPercent !== null ? 'tukin-ok' : ''}">${pct(rules.recordFinalPercent(record))}${adjustmentLabel(record) ? `<div class="card-subtitle">${esc(adjustmentLabel(record))}</div>` : ''}</td><td><div class="tukin-reason">${esc(record.adjustmentNote || record.reason || '-')}</div>${record.cutiAdjustment?.automatic ? `<div class="tukin-cuti-ref"><span class="tukin-badge edit">Data Cuti</span><span>${esc(record.cutiAdjustment.leaveType || 'Cuti')} · ${esc(record.cutiAdjustment.startDate || '')}${record.cutiAdjustment.endDate && record.cutiAdjustment.endDate !== record.cutiAdjustment.startDate ? ` s.d. ${esc(record.cutiAdjustment.endDate)}` : ''}</span></div>` : ''}${assignmentReferenceMarkup(record)}${record.adjustmentConflict?.automatic ? '<div class="tukin-cuti-ref"><span class="tukin-badge warn">Konflik Data</span><span>Cuti dan Surat Tugas ditemukan pada tanggal yang sama.</span></div>' : ''}${(record.evidence || []).length ? `<div class="tukin-files">${record.evidence.map((evidence) => `${esc(evidence.name)} ${attachmentOpenLink(evidence, 'Buka')}`).join('<br>')}</div>` : ''}</td><td><button class="btn btn-secondary btn-sm" data-edit-record="${esc(record.key)}" data-employee="${esc(rules.employeeKey(employee))}" type="button">Edit</button></td></tr>`).join('') || '<tr><td colspan="10" class="text-center">Tidak ada tanggal yang menghasilkan potongan atau memerlukan verifikasi.</td></tr>'}</tbody></table></div></aside>`;
  }

  function renderEditModal() {
    if (!state.editRecordKey || !state.editDraft) return '';
    const employee = employeeByKey(state.editDraft.employeeKey);
    const record = recordByKey(employee, state.editRecordKey);
    if (!employee || !record) return '';
    const automaticContext = record.adjustmentConflict?.automatic
      ? `<div class="alert alert-warning"><div class="alert-title">Konflik Cuti dan Surat Tugas</div>${esc(record.adjustmentNote || '')}</div>`
      : record.cutiAdjustment?.automatic
        ? `<div class="alert alert-success"><div class="alert-title">Penyesuaian otomatis dari Data Cuti</div>${esc(record.adjustmentNote || '')}</div>`
        : record.assignmentAdjustment?.adjustmentApplied || record.assignmentAdjustment?.automatic
          ? `<div class="alert alert-success"><div class="alert-title">Penyesuaian otomatis dari Surat Tugas</div>${esc(record.adjustmentNote || '')}</div>`
          : record.assignmentAdjustment?.informational
            ? `<div class="alert alert-info"><div class="alert-title">Data pendukung Surat Tugas</div>${esc(record.adjustmentNote || '')}</div>`
            : '';
    return `<div class="tukin-modal-backdrop"><div class="tukin-modal"><div class="tukin-modal-head"><div><h3>Koreksi Perhitungan</h3><p>${esc(employee.name)} · ${esc(rules.formatDate(record.date, true))}</p></div><button class="icon-btn" data-action="close-edit" type="button">×</button></div><div class="alert alert-warning"><div class="alert-title">Hasil otomatis ${pct(record.autoTotalPercent)}</div>${esc(record.reason || '')}</div>${automaticContext}
      <div class="tukin-form-grid section-gap"><div class="field"><label>% Potongan Hasil Penyesuaian</label><input id="edit-percent" type="number" min="0" max="2.5" step="0.01" value="${esc(state.editDraft.percent)}"></div><div class="field"><label>Bukti Dukung</label><input id="edit-evidence" type="file" multiple accept="application/pdf,image/png,image/jpeg,image/webp"></div><div class="full field"><label>Keterangan</label><textarea id="edit-note" placeholder="Contoh: Dinas berdasarkan Surat Tugas ...">${esc(state.editDraft.note)}</textarea></div></div>
      <div class="tukin-evidence-list">${state.editDraft.evidence.length ? state.editDraft.evidence.map((evidence, i) => `<div class="tukin-evidence"><span>${esc(evidence.name)} · ${bytes(evidence.size || 0)} ${attachmentOpenLink(evidence, 'Buka')}</span><button data-remove-evidence="${i}" type="button">Hapus</button></div>`).join('') : '<div class="card-subtitle">Belum ada bukti dukung.</div>'}</div>
      <div class="actions"><button class="btn btn-secondary" data-action="reset-auto" type="button">Gunakan Hasil Otomatis</button><div class="actions-right"><button class="btn btn-secondary" data-action="close-edit" type="button">Batal</button><button class="btn btn-primary" data-action="save-edit" type="button">Simpan Koreksi</button></div></div></div></div>`;
  }

  function renderEmployeeModal() {
    if (!state.editEmployeeKey) return '';
    const employee = employeeByKey(state.editEmployeeKey);
    if (!employee) return '';
    return `<div class="tukin-modal-backdrop"><div class="tukin-modal"><div class="tukin-modal-head"><div><h3>Data Pegawai</h3><p>${esc(employee.name)} · ${esc(employee.nip || '-')}</p></div><button class="icon-btn" data-action="close-employee-edit" type="button">×</button></div><div class="tukin-form-grid"><div class="field"><label>Anak Satker</label><input id="employee-anak" value="${esc(employee.anakSatker || '')}" maxlength="20"></div><div class="field"><label>Besaran Tukin</label><input id="employee-tukin" type="number" min="0" step="1000" value="${Number(employee.tukin || cfg.DEFAULT_TUKIN)}"></div></div><div class="actions"><span></span><div class="actions-right"><button class="btn btn-secondary" data-action="close-employee-edit" type="button">Batal</button><button class="btn btn-primary" data-action="save-employee" type="button">Simpan</button></div></div></div></div>`;
  }

  function renderHistoryPreview() {
    const run = state.historyPreview;
    if (!run) return '';
    const summaries = (run.employees || []).map((employee) => ({ employee, summary: rules.summarizeEmployee(employee) }));
    const totalCut = summaries.reduce((sum, item) => sum + item.summary.cutAmount, 0);
    const adjusted = summaries.reduce((sum, item) => sum + item.summary.adjustedRecords, 0);
    const rg = attendanceRange(run.period);
    const manageable = Boolean(storage.canManage?.(run));
    return `<div class="drawer-backdrop" data-action="close-history-preview"></div><aside class="drawer tukin-drawer"><div class="tukin-drawer-header"><div class="drawer-header"><div><h3>Riwayat Tukin ${esc(periodLabel(run.period))}</h3><div class="card-subtitle">Dasar absensi ${esc(rules.formatPeriodRange(rg))}</div></div><button class="icon-btn" data-action="close-history-preview" type="button">×</button></div></div>
      <div class="tukin-detail-summary"><div><span>Pegawai</span><strong>${summaries.length}</strong></div><div><span>Total Potongan</span><strong>${money(totalCut)}</strong></div><div><span>Koreksi</span><strong>${adjusted}</strong></div><div><span>Terakhir Diubah</span><strong>${esc(formatDateTime(run.updatedAt || run.processedAt))}</strong></div></div>
      <div class="table-wrap"><table class="data-table"><thead><tr><th>No</th><th>Nama / NIP</th><th>Pot. Absensi</th><th>Pot. Final</th><th>Potongan Rp</th></tr></thead><tbody>${summaries.map((item, i) => `<tr><td>${i + 1}</td><td><strong>${esc(item.employee.name)}</strong><div class="card-subtitle">${esc(item.employee.nip || '-')}</div></td><td>${pct(item.summary.attendancePercent)}</td><td>${pct(item.summary.finalPercent)}</td><td>${money(item.summary.cutAmount)}</td></tr>`).join('')}</tbody></table></div>
      <div class="actions">${manageable ? `<button class="btn btn-secondary" data-edit-tukin-history="${esc(run.id)}" type="button">Verifikasi / Edit</button>` : '<span class="card-subtitle">Hanya pembuat atau administrator yang dapat mengedit.</span>'}<div class="actions-right">${run.sharePoint?.runFolderWebUrl ? `<a class="btn btn-secondary" href="${esc(run.sharePoint.runFolderWebUrl)}" target="_blank" rel="noopener noreferrer">Buka SharePoint ↗</a>` : ''}<button class="btn btn-primary" data-download-tukin-history="${esc(run.id)}" type="button">Download</button></div></div></aside>`;
  }

  function validateRamadan() {
    if (!state.settings.ramadanEnabled) return true;
    const bounds = dateInputBounds();
    const start = state.settings.ramadanStart;
    const end = state.settings.ramadanEnd;
    if (!start || !end) {
      alert('Isi tanggal mulai dan selesai Ramadan.');
      return false;
    }
    if (start > end || start < bounds.min || end > bounds.max) {
      alert('Rentang Ramadan harus berada di dalam periode absensi.');
      return false;
    }
    return true;
  }

  function resetDownstream() {
    state.files = [];
    state.validationResults = [];
    state.employees = [];
    state.generated = null;
    state.drawerEmployeeKey = null;
    state.editingHistoryId = null;
    state.editingProcessedAt = null;
    state.resultMode = 'new';
    state.currentRunId = null;
    state.sharePointRun = null;
    state.pendingSharePointDeletes = [];
    state.supportIntegration = { status: 'idle', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: '' };
  }

  function resetProcess() {
    state.view = 'process';
    state.step = 'period';
    state.period = { month: now.getMonth() + 1, year: now.getFullYear() };
    state.settings = { holidays: [], ramadanEnabled: false, ramadanStart: '', ramadanEnd: '' };
    state.files = [];
    state.validationResults = [];
    state.employees = [];
    state.generated = null;
    state.drawerEmployeeKey = null;
    state.editRecordKey = null;
    state.editDraft = null;
    state.editEmployeeKey = null;
    state.search = '';
    state.editingHistoryId = null;
    state.editingProcessedAt = null;
    state.lastSavedRunId = null;
    state.resultMode = 'new';
    state.currentRunId = null;
    state.sharePointRun = null;
    state.pendingSharePointDeletes = [];
    state.supportIntegration = { status: 'idle', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: '' };
  }

  function addFiles(fileList) {
    for (const file of Array.from(fileList || [])) {
      if (!/\.xlsx?$/i.test(file.name) && !/\.xls$/i.test(file.name)) continue;
      if (!state.files.some((existing) => existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified)) {
        state.files.push(file);
      }
    }
    state.validationResults = [];
    state.employees = [];
    state.generated = null;
    state.supportIntegration = { status: 'idle', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: '' };
    if (state.files.length) {
      resources?.warm?.('xlsx');
      window.MasterDataService?.getEmployees?.(false).catch(() => {});
    }
    render();
  }

  async function calculate() {
    if (!state.files.length || state.busy) return;
    if (!validateRamadan()) return;
    state.busy = true;
    render();
    try {
      await resources?.ensure?.('xlsx');
      const result = await parser.parseFiles(state.files, state.period, state.settings);
      state.validationResults = result.results;
      state.employees = result.employees;
      state.employees.forEach((employee) => sharePoint?.normalizeEmployeeAttachments?.(employee));
      await syncSupportingAdjustments({ render: false });
      if (!state.currentRunId) state.currentRunId = newRunId();
      state.step = 'validation';
      resources?.warm?.(['exceljs', 'jszip', 'filesaver']);
    } catch (error) {
      alert(`Gagal memproses data: ${error.message || error}`);
    } finally {
      state.busy = false;
      render();
    }
  }

  function runSummary(employees) {
    const summaries = employees.map((employee) => rules.summarizeEmployee(employee));
    return {
      employees: employees.length,
      totalCutAmount: summaries.reduce((sum, summary) => sum + summary.cutAmount, 0),
      adjustedRecords: summaries.reduce((sum, summary) => sum + summary.adjustedRecords, 0),
      flaggedRecords: summaries.reduce((sum, summary) => sum + summary.flaggedRecords, 0),
      attendancePercentTotal: Number(summaries.reduce((sum, summary) => sum + summary.attendancePercent, 0).toFixed(2))
    };
  }

  function newRunId() {
    if (window.crypto?.randomUUID) return `tukin-${crypto.randomUUID()}`;
    return `tukin-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  async function saveCurrentRun(generatedName) {
    const isEdit = Boolean(state.editingHistoryId);
    const timestamp = new Date().toISOString();
    const run = {
      id: state.editingHistoryId || state.currentRunId || newRunId(),
      period: { ...state.period },
      settings: {
        holidays: [...state.settings.holidays],
        ramadanEnabled: Boolean(state.settings.ramadanEnabled),
        ramadanStart: state.settings.ramadanStart || '',
        ramadanEnd: state.settings.ramadanEnd || ''
      },
      validationResults: state.validationResults,
      employees: state.employees,
      summary: runSummary(state.employees),
      generatedName: generatedName || generator.zipFileName(state.period),
      processedAt: state.editingProcessedAt || timestamp,
      updatedAt: isEdit ? timestamp : null,
      sharePoint: state.sharePointRun || null
    };
    const savedRun = await storage.saveRun(run);
    state.editingHistoryId = savedRun.id;
    state.currentRunId = savedRun.id;
    state.sharePointRun = savedRun.sharePoint || state.sharePointRun || null;
    state.editingProcessedAt = savedRun.processedAt;
    state.lastSavedRunId = savedRun.id;
    state.resultMode = isEdit ? 'updated' : 'new';
    return savedRun;
  }

  async function saveRun() {
    if (state.busy || !state.employees.length) return;
    state.busy = true;
    state.busyAction = 'save';
    render();
    try {
      if (!state.currentRunId) state.currentRunId = state.editingHistoryId || newRunId();

      // Simpan = sinkronkan data Cuti & Surat Tugas terbaru, salin file ke SharePoint, buat rekap,
      // lalu simpan state perhitungan ke Firestore. Tidak ada download ke perangkat.
      await syncSupportingAdjustments({ render: false });
      await sharePoint?.ensureReady?.(true);

      const draftRun = {
        id: state.currentRunId,
        period: { ...state.period },
        employees: state.employees,
        sharePoint: state.sharePointRun || null
      };
      const syncResult = await sharePoint.syncRunFiles(draftRun, true);
      state.sharePointRun = syncResult.sharePoint;

      await resources?.ensure?.('exceljs');
      const recap = await generator.generateRecap(state.employees, state.period);
      const recapRun = {
        id: state.currentRunId,
        period: { ...state.period },
        employees: state.employees,
        sharePoint: state.sharePointRun
      };
      await sharePoint.uploadRecap(recapRun, recap, true);
      state.sharePointRun = recapRun.sharePoint;

      await saveCurrentRun(generator.zipFileName(state.period));

      if (state.pendingSharePointDeletes.length) {
        const activeRemoteKeys = new Set();
        for (const employee of state.employees || []) {
          for (const source of employee.sourceFiles || []) {
            if (source?.driveId && source?.itemId) activeRemoteKeys.add(`${source.driveId}:${source.itemId}`);
          }
          for (const record of Object.values(employee.records || {})) {
            for (const evidence of record.evidence || []) {
              if (evidence?.driveId && evidence?.itemId) activeRemoteKeys.add(`${evidence.driveId}:${evidence.itemId}`);
            }
          }
        }

        const uniqueDeletes = [];
        const seen = new Set();
        for (const item of state.pendingSharePointDeletes) {
          const key = `${item?.driveId || ''}:${item?.itemId || ''}`;
          if (!item?.driveId || !item?.itemId || seen.has(key) || activeRemoteKeys.has(key)) continue;
          seen.add(key);
          uniqueDeletes.push(item);
        }
        const failures = uniqueDeletes.length ? await sharePoint.deleteAttachments(uniqueDeletes, true) : [];
        if (failures.length) console.warn('Sebagian file SharePoint lama tidak berhasil dibersihkan:', failures);
        state.pendingSharePointDeletes = [];
      }

      state.generated = null;
      alert(`Data Tunjangan Kinerja ${periodLabel()} berhasil disimpan. File absensi, rekap, dan bukti dukung tersimpan di SharePoint; tidak ada file yang diunduh ke perangkat.`);
    } catch (error) {
      console.error(error);
      alert(`Gagal menyimpan data Tukin: ${error.message || error}`);
    } finally {
      state.busy = false;
      state.busyAction = '';
      render();
    }
  }

  async function downloadRun() {
    if (state.busy || !state.employees.length) return;
    state.busy = true;
    state.busyAction = 'download';
    render();
    try {
      // Download = gunakan kondisi data yang sedang tampil dan ambil paket ZIP ke
      // perangkat. Aksi ini tidak menyimpan ulang Firestore maupun menyalin file ke
      // folder run SharePoint.
      await syncSupportingAdjustments({ render: false });
      await resources?.ensure?.(['exceljs', 'jszip', 'filesaver']);
      state.generated = await generator.generateZip(state.employees, state.period);
      generator.download(state.generated);
    } catch (error) {
      console.error(error);
      alert(`Gagal mendownload paket Tukin: ${error.message || error}`);
    } finally {
      state.busy = false;
      state.busyAction = '';
      render();
    }
  }

  function collectSourceFiles(employees) {
    const map = new Map();
    employees.forEach((employee) => {
      (employee.sourceFiles || []).forEach((item) => {
        const file = item?.file || ((typeof File !== 'undefined' && item instanceof File) ? item : null);
        if (!file) return;
        const key = `${file.name}|${file.size}|${file.lastModified || 0}`;
        if (!map.has(key)) map.set(key, file);
      });
    });
    return [...map.values()];
  }

  async function reloadAttendanceFromSharePoint() {
    if (state.busy) return;
    const remoteSources = [];
    for (const employee of state.employees || []) {
      for (const attachment of employee.sourceFiles || []) {
        if (attachment?.driveId && attachment?.itemId) remoteSources.push(attachment);
      }
    }
    if (!remoteSources.length) {
      alert('Tidak ada file absensi SharePoint pada proses ini.');
      return;
    }

    state.busy = true;
    render();
    try {
      await sharePoint.ensureReady(true);
      await resources?.ensure?.('xlsx');

      const remoteByFile = new Map();
      const downloadedFiles = [];
      for (const attachment of remoteSources) {
        const file = await sharePoint.downloadAttachment({ ...attachment, file: null }, true);
        if (!file) continue;
        remoteByFile.set(file, attachment);
        downloadedFiles.push(file);
      }
      if (!downloadedFiles.length) throw new Error('File absensi tidak berhasil diunduh dari SharePoint.');

      const previous = new Map((state.employees || []).map((employee) => [rules.employeeKey(employee), employee]));
      const parsed = await parser.parseFiles(downloadedFiles, state.period, state.settings);

      for (const employee of parsed.employees) {
        const oldEmployee = previous.get(rules.employeeKey(employee));
        employee.sourceFiles = (employee.sourceFiles || []).map((file) => {
          const remote = remoteByFile.get(file);
          if (!remote) return file;
          return {
            ...remote,
            file,
            name: remote.name || file.name,
            size: Number(file.size || remote.size || 0),
            type: file.type || remote.type || '',
            lastModified: Number(file.lastModified || remote.lastModified || 0),
            sourceSize: Number(file.size || 0),
            sourceLastModified: Number(file.lastModified || Date.now()),
            localOnly: false,
            provider: 'sharepoint'
          };
        });

        if (oldEmployee) {
          employee.anakSatker = oldEmployee.anakSatker || employee.anakSatker;
          employee.tukin = Number(oldEmployee.tukin || employee.tukin || cfg.DEFAULT_TUKIN);
          employee.skpDeductionPercent = Number(oldEmployee.skpDeductionPercent || 0);
          employee.masterMatched = oldEmployee.masterMatched ?? employee.masterMatched;
          Object.entries(employee.records || {}).forEach(([key, record]) => {
            const oldRecord = oldEmployee.records?.[key];
            if (!oldRecord) return;
            record.adjustedPercent = oldRecord.adjustedPercent ?? null;
            record.adjustmentNote = oldRecord.adjustmentNote || '';
            record.adjustmentSource = oldRecord.adjustmentSource || '';
            record.cutiAdjustment = oldRecord.cutiAdjustment ? { ...oldRecord.cutiAdjustment } : null;
            record.assignmentAdjustment = oldRecord.assignmentAdjustment ? { ...oldRecord.assignmentAdjustment, assignments: (oldRecord.assignmentAdjustment.assignments || []).map((item) => ({ ...item })) } : null;
            record.adjustmentConflict = oldRecord.adjustmentConflict ? { ...oldRecord.adjustmentConflict, assignmentIds: [...(oldRecord.adjustmentConflict.assignmentIds || [])] } : null;
            record.evidence = [...(oldRecord.evidence || [])];
          });
        }
        sharePoint.normalizeEmployeeAttachments(employee);
      }

      state.employees = parsed.employees;
      state.validationResults = parsed.results;
      state.files = downloadedFiles;
      state.generated = null;
      await syncSupportingAdjustments({ render: false });
      alert('File absensi terbaru dari SharePoint telah dimuat ulang. Perhitungan otomatis dan integrasi Cuti/Surat Tugas diperbarui; koreksi manual/bukti dukung yang sudah ada dipertahankan berdasarkan tanggal.');
    } catch (error) {
      console.error(error);
      alert(`Gagal memuat ulang absensi dari SharePoint: ${error.message || error}`);
    } finally {
      state.busy = false;
      render();
    }
  }

  async function refreshHistory(showLoading = true) {
    if (!storage) return;
    if (showLoading) {
      state.historyBusy = true;
      render();
    }
    try {
      state.history = await storage.listRuns();
      state.historyLoaded = true;
    } catch (error) {
      console.error(error);
      state.history = [];
      state.historyLoaded = true;
      if (showLoading) alert('Riwayat Tukin tidak dapat dimuat dari Firebase maupun cache lokal.');
    } finally {
      state.historyBusy = false;
      if (showLoading) render();
    }
  }

  async function viewHistory(id) {
    state.historyBusy = true;
    render();
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan.');
      state.historyPreview = run;
    } catch (error) {
      alert(error.message || String(error));
    } finally {
      state.historyBusy = false;
      render();
    }
  }

  async function editHistory(id) {
    state.historyBusy = true;
    render();
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan.');
      if (!storage.canManage?.(run)) throw new Error('Riwayat ini dibuat oleh akun lain. Anda dapat melihat atau mendownload hasilnya, tetapi hanya pembuat atau administrator yang dapat mengedit.');
      if (run.cloudOnly && !run.binaryFilesAvailable) {
        alert('Riwayat lama ini tidak memiliki pointer file SharePoint. Data perhitungan tersedia, tetapi file biner mungkin tidak tersedia pada perangkat ini.');
      }
      state.period = { ...run.period };
      state.settings = {
        holidays: [...(run.settings?.holidays || [])],
        ramadanEnabled: Boolean(run.settings?.ramadanEnabled),
        ramadanStart: run.settings?.ramadanStart || '',
        ramadanEnd: run.settings?.ramadanEnd || ''
      };
      state.validationResults = Array.isArray(run.validationResults) ? run.validationResults : [];
      state.employees = Array.isArray(run.employees) ? run.employees : [];
      state.files = collectSourceFiles(state.employees);
      state.generated = null;
      state.search = '';
      state.drawerEmployeeKey = null;
      state.editRecordKey = null;
      state.editDraft = null;
      state.editEmployeeKey = null;
      state.editingHistoryId = run.id;
      state.currentRunId = run.id;
      state.sharePointRun = run.sharePoint || null;
      state.pendingSharePointDeletes = [];
      state.editingProcessedAt = run.processedAt || null;
      state.lastSavedRunId = run.id;
      state.resultMode = 'history';
      state.historyPreview = null;
      state.view = 'process';
      state.step = 'validation';
      history.replaceState(null, '', 'tukin.html#process');
    } catch (error) {
      alert(error.message || String(error));
    } finally {
      state.historyBusy = false;
      render();
    }
  }

  async function attendanceSourceChangedOnSharePoint(run) {
    for (const employee of run?.employees || []) {
      for (const attachment of employee.sourceFiles || []) {
        if (!attachment?.driveId || !attachment?.itemId || !attachment?.eTag) continue;
        const latest = await sharePoint.refreshAttachmentMetadata({ ...attachment, file: null }, true);
        if (latest?.eTag && latest.eTag !== attachment.eTag) return true;
      }
    }
    return false;
  }

  async function downloadHistory(id) {
    if (state.historyBusy || state.busy) return;
    state.historyBusy = true;
    render();
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan.');
      if (run.cloudOnly && !run.binaryFilesAvailable) {
        const approved = window.confirm('Riwayat lama ini tidak memiliki file SharePoint yang lengkap. Download dapat menghasilkan paket tanpa sebagian file biner. Lanjutkan?');
        if (!approved) return;
      }
      if (run.binaryFilesAvailable) {
        await sharePoint?.ensureReady?.(true);
        const attendanceChanged = await attendanceSourceChangedOnSharePoint(run);
        if (attendanceChanged) {
          const message = storage.canManage?.(run)
            ? 'File absensi di SharePoint telah berubah sejak perhitungan terakhir. Buka Verifikasi/Edit, klik "Muat Ulang Absensi dari SharePoint", lalu klik Simpan agar rekap konsisten dengan file sumber terbaru.'
            : 'File absensi di SharePoint telah berubah sejak perhitungan terakhir. Pemilik riwayat atau administrator perlu memuat ulang absensi dan menyimpan perhitungan sebelum paket ZIP dapat didownload.';
          alert(message);
          return;
        }
      }
      await resources?.ensure?.(['exceljs', 'jszip', 'filesaver']);
      const generated = await generator.generateZip(run.employees || [], run.period);
      generator.download(generated);
    } catch (error) {
      alert(`Gagal download: ${error.message || error}`);
    } finally {
      state.historyBusy = false;
      render();
    }
  }

  async function deleteHistory(id) {
    const item = state.history.find((run) => run.id === id);
    if (item && !storage.canManage?.(item)) {
      alert('Riwayat ini dibuat oleh akun lain dan hanya dapat dihapus oleh pembuat atau administrator.');
      return;
    }
    const label = item ? periodLabel(item.period) : 'ini';
    if (!window.confirm(`Hapus riwayat Tunjangan Kinerja ${label}?\n\nData riwayat Firestore akan dihapus. Folder run dan file terkait di SharePoint juga akan dicoba dihapus otomatis.`)) return;
    state.historyBusy = true;
    render();
    try {
      let fullRun = null;
      try { fullRun = await storage.getRun(id); } catch (_) {}
      await storage.deleteRun(id);
      if (fullRun?.sharePoint?.runFolderItemId) {
        try {
          await sharePoint?.ensureReady?.(true);
          await sharePoint?.deleteRunFolder?.(fullRun.sharePoint, true);
        } catch (sharePointError) {
          console.warn('Folder run SharePoint tidak berhasil dibersihkan:', sharePointError);
          alert('Riwayat Firestore telah dihapus, tetapi folder SharePoint tidak berhasil dihapus otomatis. Folder dapat dibersihkan manual.');
        }
      }
      if (state.historyPreview?.id === id) state.historyPreview = null;
      if (state.editingHistoryId === id) {
        state.editingHistoryId = null;
        state.editingProcessedAt = null;
      }
      state.history = state.history.filter((run) => run.id !== id);
    } catch (error) {
      alert(`Riwayat gagal dihapus: ${error.message || error}`);
    } finally {
      state.historyBusy = false;
      render();
    }
  }

  function openRecordEdit(employeeKey, recordKey) {
    const employee = employeeByKey(employeeKey);
    const record = recordByKey(employee, recordKey);
    if (!record) return;
    state.editRecordKey = recordKey;
    state.editDraft = {
      employeeKey,
      percent: rules.recordFinalPercent(record),
      note: record.adjustmentNote || '',
      evidence: [...(record.evidence || [])]
    };
    render();
  }

  function syncDraftFromForm() {
    if (!state.editDraft) return;
    const percent = document.getElementById('edit-percent');
    const note = document.getElementById('edit-note');
    if (percent) state.editDraft.percent = percent.value;
    if (note) state.editDraft.note = note.value;
  }

  function saveRecordEdit() {
    const employee = employeeByKey(state.editDraft.employeeKey);
    const record = recordByKey(employee, state.editRecordKey);
    if (!record) return;
    const value = Number(document.getElementById('edit-percent')?.value);
    if (!Number.isFinite(value) || value < 0 || value > 2.5) {
      alert('Persentase penyesuaian harus antara 0% sampai 2,5%.');
      return;
    }
    const keptRemoteIds = new Set((state.editDraft.evidence || []).map((item) => item?.itemId).filter(Boolean));
    for (const item of record.evidence || []) {
      if (item?.driveId && item?.itemId && !keptRemoteIds.has(item.itemId)) {
        state.pendingSharePointDeletes.push({ ...item, file: null });
      }
    }
    const assignmentContext = record.assignmentAdjustment?.linkedAutomatically || record.assignmentAdjustment?.assignments?.length
      ? {
          ...record.assignmentAdjustment,
          automatic: false,
          adjustmentApplied: false,
          informational: true,
          manualDecision: true,
          assignments: (record.assignmentAdjustment.assignments || []).map((item) => ({ ...item }))
        }
      : null;
    record.adjustedPercent = Number(value.toFixed(2));
    record.adjustmentNote = String(document.getElementById('edit-note')?.value || '').trim();
    record.adjustmentSource = 'manual';
    record.cutiAdjustment = null;
    record.assignmentAdjustment = assignmentContext;
    record.adjustmentConflict = null;
    record.evidence = [...state.editDraft.evidence];
    state.generated = null;
    state.editRecordKey = null;
    state.editDraft = null;
    render();
  }

  function resetRecordAuto() {
    const employee = employeeByKey(state.editDraft.employeeKey);
    const record = recordByKey(employee, state.editRecordKey);
    if (!record) return;
    record.adjustedPercent = null;
    record.adjustmentNote = '';
    record.adjustmentSource = '';
    record.cutiAdjustment = null;
    record.assignmentAdjustment = null;
    record.adjustmentConflict = null;
    for (const item of record.evidence || []) {
      if (item?.driveId && item?.itemId) state.pendingSharePointDeletes.push({ ...item, file: null });
    }
    record.evidence = [];
    state.generated = null;
    state.editRecordKey = null;
    state.editDraft = null;
    render();
  }

  function setViewFromHash() {
    const route = String(location.hash || '').replace(/^#/, '').toLowerCase();
    if (route === 'history') state.view = 'history';
    else if (route === 'process') state.view = 'process';
    else state.view = 'dashboard';
  }

  function goToHistory() {
    state.view = 'history';
    state.historyPreview = null;
    history.replaceState(null, '', 'tukin.html#history');
    render();
  }

  function bind() {
    document.querySelector('[data-action="connect-sharepoint"]')?.addEventListener('click', async () => {
      if (state.sharePointBusy) return;
      state.sharePointBusy = true;
      render();
      try {
        await microsoftAuth?.connect?.();
        await sharePoint?.resolveAppRoot?.(true);
      } catch (error) {
        console.error(error);
        alert(`Koneksi SharePoint gagal: ${error.message || error}`);
      } finally {
        state.sharePointBusy = false;
        render();
      }
    });
    document.querySelectorAll('a[href="tukin.html#dashboard"]').forEach((link) => link.addEventListener('click', (event) => {
      event.preventDefault();
      state.view = 'dashboard';
      state.historyPreview = null;
      history.replaceState(null, '', 'tukin.html#dashboard');
      render();
    }));
    document.querySelectorAll('a[href="tukin.html#process"]').forEach((link) => link.addEventListener('click', (event) => {
      event.preventDefault();
      state.view = 'process';
      state.historyPreview = null;
      history.replaceState(null, '', 'tukin.html#process');
      render();
    }));
    document.querySelectorAll('a[href="tukin.html#history"]').forEach((link) => link.addEventListener('click', (event) => {
      event.preventDefault();
      goToHistory();
    }));

    document.getElementById('tukin-month')?.addEventListener('change', (event) => {
      const old = state.period.month;
      state.period.month = Number(event.target.value);
      if (old !== state.period.month) {
        state.settings.holidays = [];
        state.settings.ramadanStart = '';
        state.settings.ramadanEnd = '';
        resetDownstream();
      }
      render();
    });
    document.getElementById('tukin-year')?.addEventListener('change', (event) => {
      const old = state.period.year;
      state.period.year = Number(event.target.value);
      if (old !== state.period.year) {
        state.settings.holidays = [];
        state.settings.ramadanStart = '';
        state.settings.ramadanEnd = '';
        resetDownstream();
      }
      render();
    });
    document.querySelector('[data-action="add-holiday"]')?.addEventListener('click', () => {
      const key = document.getElementById('holiday-date')?.value || '';
      const bounds = dateInputBounds();
      if (!key || key < bounds.min || key > bounds.max) {
        alert('Pilih tanggal yang berada pada periode absensi.');
        return;
      }
      state.settings.holidays = rules.normalizeHolidays([...state.settings.holidays, key]);
      calcAll();
      render();
    });
    document.querySelectorAll('[data-remove-holiday]').forEach((button) => button.addEventListener('click', () => {
      state.settings.holidays = state.settings.holidays.filter((key) => key !== button.dataset.removeHoliday);
      calcAll();
      render();
    }));
    document.getElementById('ramadan-enabled')?.addEventListener('change', (event) => {
      state.settings.ramadanEnabled = event.target.checked;
      calcAll();
      render();
    });
    document.getElementById('ramadan-start')?.addEventListener('change', (event) => {
      state.settings.ramadanStart = event.target.value;
      calcAll();
    });
    document.getElementById('ramadan-end')?.addEventListener('change', (event) => {
      state.settings.ramadanEnd = event.target.value;
      calcAll();
    });
    document.querySelector('[data-action="period-next"]')?.addEventListener('click', () => {
      if (!validateRamadan()) return;
      state.step = state.employees.length ? 'validation' : 'upload';
      render();
    });
    document.querySelector('[data-action="back-period"]')?.addEventListener('click', () => { state.step = 'period'; render(); });
    document.querySelectorAll('[data-action="choose-files"]').forEach((button) => button.addEventListener('click', () => document.getElementById('tukin-files')?.click()));
    document.getElementById('tukin-files')?.addEventListener('change', (event) => addFiles(event.target.files));
    const dropzone = document.getElementById('tukin-dropzone');
    if (dropzone) {
      ['dragenter', 'dragover'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add('dragover'); }));
      ['dragleave', 'drop'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove('dragover'); }));
      dropzone.addEventListener('drop', (event) => addFiles(event.dataTransfer.files));
    }
    document.querySelectorAll('[data-remove-file]').forEach((button) => button.addEventListener('click', () => {
      state.files.splice(Number(button.dataset.removeFile), 1);
      state.validationResults = [];
      state.employees = [];
      state.supportIntegration = { status: 'idle', leaveRecords: 0, assignmentRecords: 0, adjustedRecords: 0, cutiAdjustedRecords: 0, assignmentAdjustedRecords: 0, conflictRecords: 0, clearedRecords: 0, matchedEmployees: 0, unresolvedEmployees: 0, preservedManualRecords: 0, evidenceFiles: 0, message: '' };
      render();
    }));
    document.querySelector('[data-action="calculate"]')?.addEventListener('click', calculate);
    document.querySelector('[data-action="back-upload"]')?.addEventListener('click', () => { state.step = 'upload'; render(); });
    document.querySelector('[data-action="settings"]')?.addEventListener('click', () => { state.step = 'period'; render(); });
    document.getElementById('tukin-search')?.addEventListener('input', (event) => {
      state.search = event.target.value;
      clearTimeout(window.__tukinSearch);
      window.__tukinSearch = setTimeout(render, 160);
    });
    document.querySelectorAll('[data-detail-employee]').forEach((button) => button.addEventListener('click', () => { state.drawerEmployeeKey = button.dataset.detailEmployee; render(); }));
    document.querySelectorAll('[data-action="close-drawer"]').forEach((button) => button.addEventListener('click', () => { state.drawerEmployeeKey = null; render(); }));
    document.querySelectorAll('[data-edit-record]').forEach((button) => button.addEventListener('click', () => openRecordEdit(button.dataset.employee, button.dataset.editRecord)));
    document.querySelector('[data-action="close-edit"]')?.addEventListener('click', () => { state.editRecordKey = null; state.editDraft = null; render(); });
    document.getElementById('edit-evidence')?.addEventListener('change', (event) => {
      syncDraftFromForm();
      for (const file of Array.from(event.target.files || [])) {
        if (!/^application\/pdf$|^image\//i.test(file.type)) {
          alert(`File ${file.name} bukan PDF/image.`);
          continue;
        }
        state.editDraft.evidence.push({ id: `${Date.now()}-${Math.random()}`, name: file.name, size: file.size, type: file.type, file });
      }
      render();
    });
    document.querySelectorAll('[data-remove-evidence]').forEach((button) => button.addEventListener('click', () => {
      syncDraftFromForm();
      const index = Number(button.dataset.removeEvidence);
      state.editDraft.evidence.splice(index, 1);
      render();
    }));
    document.querySelector('[data-action="save-edit"]')?.addEventListener('click', saveRecordEdit);
    document.querySelector('[data-action="reset-auto"]')?.addEventListener('click', resetRecordAuto);
    document.querySelectorAll('[data-edit-employee]').forEach((button) => button.addEventListener('click', () => { state.editEmployeeKey = button.dataset.editEmployee; render(); }));
    document.querySelector('[data-action="close-employee-edit"]')?.addEventListener('click', () => { state.editEmployeeKey = null; render(); });
    document.querySelector('[data-action="save-employee"]')?.addEventListener('click', () => {
      const employee = employeeByKey(state.editEmployeeKey);
      if (!employee) return;
      const rawAnakSatker = String(document.getElementById('employee-anak')?.value || '').trim();
      const anakSatker = /^\d$/.test(rawAnakSatker) ? rawAnakSatker.padStart(2, '0') : rawAnakSatker;
      const tukin = Number(document.getElementById('employee-tukin')?.value);
      if (!anakSatker || !Number.isFinite(tukin) || tukin < 0) {
        alert('Anak Satker dan besaran Tukin harus valid.');
        return;
      }
      employee.anakSatker = anakSatker;
      employee.tukin = tukin;
      employee.masterMatched = true;
      state.editEmployeeKey = null;
      state.generated = null;
      render();
    });
    document.querySelector('[data-action="sync-cuti"]')?.addEventListener('click', async () => {
      const result = await syncSupportingAdjustments();
      if (result) {
        const count = Number(result.adjustedRecords || 0);
        const evidenceCount = Number(result.evidenceFiles || 0);
        const informationalCount = Number(result.assignmentInformationalRecords || 0);
        const conflictCount = Number(result.conflictRecords || 0);
        const evidenceText = evidenceCount ? ` ${evidenceCount} file bukti unik siap disalin satu kali saat Simpan.` : '';
        const informationText = informationalCount ? ` ${informationalCount} tanggal Surat Tugas Pengaturan Kerja/Lainnya ditampilkan sebagai data pendukung tanpa mengubah persentase otomatis.` : '';
        const conflictText = conflictCount ? ` ${conflictCount} tanggal memiliki konflik Cuti dan Surat Tugas dan tetap memerlukan verifikasi.` : '';
        alert(count ? `${count} tanggal berhasil disesuaikan otomatis berdasarkan Cuti/Surat Tugas.${informationText}${evidenceText}${conflictText}` : `Sinkronisasi Cuti & Surat Tugas selesai.${informationText}${evidenceText}${conflictText}`);
      }
    });
    document.querySelector('[data-action="reload-sharepoint-attendance"]')?.addEventListener('click', reloadAttendanceFromSharePoint);
    document.querySelector('[data-action="save-run"]')?.addEventListener('click', saveRun);
    document.querySelector('[data-action="download-run"]')?.addEventListener('click', downloadRun);
    document.querySelector('[data-action="download-again"]')?.addEventListener('click', () => state.generated && generator.download(state.generated));
    document.querySelector('[data-action="edit-again"]')?.addEventListener('click', () => { state.step = 'validation'; render(); });
    document.querySelector('[data-action="new-process"]')?.addEventListener('click', () => {
      resetProcess();
      history.replaceState(null, '', 'tukin.html#process');
      render();
    });
    document.querySelector('[data-action="new-process-from-history"]')?.addEventListener('click', () => {
      resetProcess();
      history.replaceState(null, '', 'tukin.html#process');
      render();
    });
    document.querySelector('[data-action="go-tukin-history"]')?.addEventListener('click', goToHistory);

    document.querySelectorAll('[data-view-tukin-history]').forEach((button) => button.addEventListener('click', () => viewHistory(button.dataset.viewTukinHistory)));
    document.querySelectorAll('[data-edit-tukin-history]').forEach((button) => button.addEventListener('click', () => editHistory(button.dataset.editTukinHistory)));
    document.querySelectorAll('[data-download-tukin-history]').forEach((button) => button.addEventListener('click', () => downloadHistory(button.dataset.downloadTukinHistory)));
    document.querySelectorAll('[data-delete-tukin-history]').forEach((button) => button.addEventListener('click', () => deleteHistory(button.dataset.deleteTukinHistory)));
    document.querySelectorAll('[data-action="close-history-preview"]').forEach((button) => button.addEventListener('click', () => { state.historyPreview = null; render(); }));
  }

  window.addEventListener('hashchange', () => {
    const previous = state.view;
    setViewFromHash();
    state.historyPreview = null;
    if (state.view === 'history' || state.view === 'dashboard') {
      render();
    } else if (previous !== 'process') {
      render();
    }
  });

  async function init() {
    await window.FirebaseClient.requireAuth();
    await window.AppSettingsService?.loadAndApply?.();
    try { await microsoftAuth?.initialize?.(); } catch (error) { console.warn('Inisialisasi Microsoft gagal:', error); }
    setViewFromHash();
    if (!location.hash) history.replaceState(null, '', 'tukin.html#dashboard');
    state.historyBusy = state.view === 'history' || state.view === 'dashboard';
    render();
    await storage.migrateLegacyOnce?.();
    historyUnsubscribe = storage.subscribeRuns?.((history) => {
      state.history = history;
      state.historyLoaded = true;
      state.historyBusy = false;
      if (state.view === 'history' || state.view === 'dashboard') render();
    }, async (error) => {
      console.warn('Sinkronisasi realtime Tukin gagal:', error);
      await refreshHistory(false);
      if (state.view === 'history' || state.view === 'dashboard') render();
    }) || null;
    render();
  }

  window.addEventListener('beforeunload', () => historyUnsubscribe?.());
  init();
})();
