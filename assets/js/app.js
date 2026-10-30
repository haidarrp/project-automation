(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;
  const parser = window.ExcelParser;
  const generator = window.DocumentGenerator;
  const storage = window.AppStorage;
  const resources = window.ResourceLoader;
  const app = document.getElementById('app');

  let historyUnsubscribe = null;

  const state = {
    started: false,
    view: 'dashboard',
    processStep: 'period',
    period: { month: new Date().getMonth() + 1, year: new Date().getFullYear() },
    files: [],
    validation: null,
    employees: [],
    filter: '',
    drawerKey: null,
    detailMode: 'calendar',
    detailFilter: 'all',
    detailDateKey: null,
    reviewMode: 'calendar',
    reviewGroup: 'employee',
    reviewEmployeeKey: null,
    reviewDateKey: null,
    reviewFilter: 'all',
    reviewTableFilter: 'overtime',
    reviewEditKey: null,
    pendingHoliday: null,
    dirtyKeys: new Set(),
    processing: { percent: 0, active: 0 },
    generated: null,
    currentRun: null,
    resultFromHistory: false,
    holidays: [],
    editingHistoryId: null,
    resultMode: 'new',
    resultTab: 'summary',
    resultSearch: '',
    resultCategory: 'all',
    resultDate: 'all',
    busy: false,
    history: [],
    historyBusy: false
  };

  function esc(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function periodLabel(period) {
    return `${cfg.INDONESIAN_MONTHS[period.month - 1]} ${period.year}`;
  }

  function formatDateTime(iso) {
    try {
      return new Intl.DateTimeFormat('id-ID', {
        timeZone: cfg.TIME_ZONE, day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit'
      }).format(new Date(iso));
    } catch (_) {
      return new Date(iso).toLocaleString('id-ID');
    }
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes)) return '-';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function periodDateBounds(period) {
    const days = new Date(period.year, period.month, 0).getDate();
    return {
      min: `${period.year}-${rules.pad2(period.month)}-01`,
      max: `${period.year}-${rules.pad2(period.month)}-${rules.pad2(days)}`
    };
  }

  function holidayLabel(key) {
    try {
      return rules.formatIndonesianDate(rules.dateFromKey(key), true);
    } catch (_) {
      return key;
    }
  }

  function renderHolidayManager() {
    const bounds = periodDateBounds(state.period);
    const holidays = rules.normalizeHolidays(state.holidays);
    const chips = holidays.length
      ? holidays.map((key) => `<span class="holiday-chip"><span>${esc(holidayLabel(key))}</span><button type="button" data-remove-holiday="${esc(key)}" aria-label="Hapus tanggal merah">x</button></span>`).join('')
      : '<span class="holiday-empty">Belum ada tanggal merah tambahan.</span>';
    return `<div class="holiday-box"><div class="holiday-head"><div><div class="card-title">Tanggal Merah</div><div class="card-subtitle">Tanggal yang ditambahkan diperlakukan seperti weekend untuk perhitungan lembur dan SPKL.</div></div></div><div class="holiday-form"><input id="holiday-date" type="date" min="${bounds.min}" max="${bounds.max}"><button class="btn btn-secondary btn-sm" type="button" data-action="add-holiday">+ Tambah Tanggal Merah</button></div><div class="holiday-list">${chips}</div></div>`;
  }

  function dateBelongsToPeriod(key, period) {
    return String(key || '').startsWith(`${period.year}-${rules.pad2(period.month)}-`);
  }

  // Perubahan tanggal merah selalu ditinjau sebelum dihitung ulang.
  // Koreksi manual sebelumnya dipertahankan, bukan ditimpa otomatis.
  function addHoliday() {
    if (!canCorrectOvertime(true)) return;
    const input = document.getElementById('holiday-date');
    const key = String(input?.value || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key) || !dateBelongsToPeriod(key, state.period)) {
      alert(`Pilih tanggal pada periode ${periodLabel(state.period)}.`);
      return;
    }
    if (state.holidays.includes(key)) return;
    state.pendingHoliday = { key, operation: 'add' };
    renderStay();
  }

  function removeHoliday(key) {
    if (!canCorrectOvertime(true) || !state.holidays.includes(key)) return;
    state.pendingHoliday = { key, operation: 'remove' };
    renderStay();
  }

  function canCorrectOvertime(holiday = false) {
    const action = state.editingHistoryId ? 'history.edit' : holiday ? 'process.holiday' : 'process.correct';
    return window.UserAccess.assert('lembur', action);
  }

  function reviewHolidayImpact() {
    const pending = state.pendingHoliday;
    if (!pending) return null;
    const holidays = rules.normalizeHolidays(pending.operation === 'add'
      ? [...state.holidays, pending.key] : state.holidays.filter(key => key !== pending.key));
    const affected = [];
    state.employees.forEach(employee => {
      const record = employee.records?.[pending.key];
      if (!record) return;
      const next = rules.calculateOvertime(record.date, record.inMinutes, record.outMinutes, record.status, holidays);
      const before = autoHours(record);
      const manual = isManualOvertime(record);
      if (before !== next.hours || manual) affected.push({ employee, record, before, next: next.hours, manual });
    });
    return { holidays, affected, changed: affected.filter(item => item.before !== item.next).length,
      manualCount: affected.filter(item => item.manual).length };
  }

  function confirmHolidayChange() {
    if (!state.pendingHoliday || !canCorrectOvertime(true)) return;
    const pending = state.pendingHoliday;
    const impact = reviewHolidayImpact();
    state.holidays = impact.holidays;
    state.employees.forEach(employee => {
      const record = employee.records?.[pending.key];
      if (!record) return;
      const keepManual = isManualOvertime(record);
      const calculation = rules.calculateOvertime(record.date, record.inMinutes, record.outMinutes, record.status, state.holidays);
      record.originalOvertimeHours = calculation.hours;
      if (record.documentBlocksOvertime) record.overtimeHours = 0;
      else if (!keepManual) record.overtimeHours = calculation.hours;
      else record.manualOverride = true;
      record.normalEndMinutes = calculation.normalEndMinutes;
      record.adjustedWorkEndMinutes = calculation.adjustedWorkEndMinutes;
      record.overtimeStartMinutes = calculation.overtimeStartMinutes;
      record.isHoliday = rules.isHoliday(record.date, state.holidays);
    });
    state.dirtyKeys.add(`holiday:${pending.key}`);
    state.generated = null;
    state.pendingHoliday = null;
    renderStay();
  }

  function icon(name) {
    const icons = {
      home: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M9 21v-6h6v6"/></svg>',
      process: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M5 18v3h14v-3"/></svg>',
      history: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16"/><path d="M4 12h16"/><path d="M4 18h10"/></svg>',
      plus: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>',
      arrowRight: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12h14"/><path d="m14 7 5 5-5 5"/></svg>',
      chevronRight: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9 18 6-6-6-6"/></svg>',
      calendar: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="15" rx="2"/><path d="M8 3v4M16 3v4M4 9h16"/></svg>',
      upload: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4"/><path d="m8 8 4-4 4 4"/><path d="M4 18v2h16v-2"/></svg>',
      review: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h10M4 18h8"/><path d="m16 17 2 2 3-4"/></svg>',
      fileCheck: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h10l4 4v12H5z"/><path d="M15 4v4h4"/><path d="m8 14 2 2 5-5"/></svg>',
      users: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3"/><path d="M3.8 19c.7-3 2.4-4.5 5.2-4.5S13.5 16 14.2 19"/><circle cx="17" cy="9" r="2.2"/><path d="M15.7 14.8c2.6-.1 4.1 1.3 4.5 4.2"/></svg>',
      clock: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="M12 7v5l3 2"/></svg>',
      timer: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h12"/><path d="M8 3v5l4 4 4-4V3"/><path d="M8 21v-5l4-4 4 4v5"/><path d="M6 21h12"/></svg>',
      meal: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3v7"/><path d="M3 3v5a2 2 0 0 0 4 0V3"/><path d="M5 10v11"/><path d="M15 3c3 2 4 5 4 8v10"/><path d="M15 3v10h4"/></svg>'
    };
    return icons[name] || icons.fileCheck;
  }

  function shell(content) {
    const view = state.view === 'process' && (state.resultFromHistory || state.editingHistoryId) && window.UserAccess.canView('lembur','history') ? 'history' : ['dashboard', 'history', 'process'].includes(state.view) ? state.view : 'dashboard';
    const viewLabel = view === 'dashboard' ? 'Dashboard' : view === 'history' ? 'Riwayat' : 'Proses Lembur';
    return window.AppShell.render({
      module: 'lembur',
      view,
      viewLabel,
      content,
      overlays: renderDrawer()
    });
  }

  function renderWelcome() {
    app.innerHTML = `
      <section class="welcome-shell">
        <div class="welcome-minimal">
          <img class="welcome-logo" src="assets/img/logo-pkp.png?v=2" alt="Kementerian Perumahan dan Kawasan Permukiman">
          <h1 class="welcome-title">Generator Dokumen Lembur</h1>
          <button class="welcome-start" type="button" data-action="enter-app">Mulai ${icon('arrowRight')}</button>
        </div>
      </section>`;
    bindEvents();
  }

  async function refreshHistory(showAlert) {
    state.historyBusy = true;
    try {
      state.history = await storage.listHistory();
    } catch (error) {
      console.error(error);
      state.history = [];
      if (showAlert) alert(`Riwayat Lembur tidak dapat dimuat: ${error.message || error}`);
    } finally {
      state.historyBusy = false;
    }
  }

  function renderDashboard() {
    const history = state.history;
    const latest = history[0];
    const summary = latest?.summary || { employees: 0, overtimeEmployees: 0, totalHours: 0, mealDays: 0 };
    const recent = history.slice(0, 3);
    const recentRows = recent.length
      ? recent.map((item) => `<div class="dashboard-history-row">
          <div><div class="history-period-name">${esc(periodLabel(item.period))}</div><div class="history-period-sub">Dokumen lembur</div></div>
          <div>${Number(item.summary?.employees || 0)}</div>
          <div>${Number(item.summary?.totalHours || 0)} jam</div>
          <div>${esc(formatDateTime(item.updatedAt || item.processedAt))}</div>
          <div><span class="status-pill">Selesai</span></div>
          <button class="chev-btn" type="button" data-history-id="${esc(item.id)}" aria-label="Lihat ${esc(periodLabel(item.period))}">${icon('chevronRight')}</button>
        </div>`).join('')
      : '<div class="empty-state"><strong>Belum ada riwayat</strong>Proses pertama akan tampil di sini.</div>';

    const content = `
      <div class="page-head">
        <div><h1>Dashboard</h1><div class="small">Ringkasan periode terakhir</div></div>
        <button class="btn btn-primary" data-action="start-process" type="button">${icon('plus')} Proses Baru</button>
      </div>

      <div class="dashboard-metrics">
        ${metric(summary.employees, 'Pegawai')}
        ${metric(summary.overtimeEmployees, 'Pegawai Lembur')}
        ${metric(summary.totalHours, 'Total Jam Lembur')}
        ${metric(summary.mealDays, 'Hari Uang Makan')}
      </div>

      <div class="dashboard-grid">
        <div class="card period-card">
          <div class="dashboard-card-head"><div class="card-title">Periode Terakhir</div><div class="dashboard-card-note">${latest ? 'Selesai' : '—'}</div></div>
          <div class="period-box">
            <div class="period-label">Periode</div>
            <div class="period-main">${latest ? esc(periodLabel(latest.period)) : 'Belum ada'}</div>
            <div class="period-meta">
              <div><div class="period-meta-value">${Number(summary.employees || 0)}</div><div class="period-meta-label">Pegawai</div></div>
              <div><div class="period-meta-value">${Number(summary.totalHours || 0)} jam</div><div class="period-meta-label">Lembur</div></div>
            </div>
          </div>
          <div class="period-actions">${latest ? `<button class="text-link-btn" data-history-id="${esc(latest.id)}" type="button">Lihat hasil ${icon('chevronRight')}</button>` : ''}</div>
        </div>
      </div>

      <div class="card dashboard-history">
        <div class="dashboard-history-head"><div class="card-title">Riwayat Terbaru</div>${history.length ? '<button class="text-link-btn" data-action="go-history" type="button">Lihat semua</button>' : ''}</div>
        ${history.length ? '<div class="dashboard-history-cols"><div>Periode</div><div>Pegawai</div><div>Lembur</div><div>Diproses</div><div>Status</div><div></div></div>' : ''}
        ${recentRows}
      </div>`;
    app.innerHTML = shell(content);
    bindEvents();
  }

  function metric(value, label) {
    const meta = {
      'Pegawai': { icon: 'users', suffix: '' },
      'Pegawai Lembur': { icon: 'clock', suffix: '' },
      'Total Jam Lembur': { icon: 'timer', suffix: 'jam' },
      'Hari Uang Makan': { icon: 'meal', suffix: 'hari' }
    }[label] || { icon: 'fileCheck', suffix: '' };
    return `<div class="card metric"><div class="metric-top"><div class="metric-label">${esc(label)}</div><div class="metric-icon">${icon(meta.icon)}</div></div><div class="metric-value">${Number(value || 0)}${meta.suffix ? `<span class="metric-suffix">${meta.suffix}</span>` : ''}</div></div>`;
  }

  function stepper(active) {
    const steps = [['period','Periode'],['upload','Upload'],['review','Review'],['result','Hasil']];
    const order = { period:0, upload:1, validation:1, processing:1, review:2, confirm:2, result:3 };
    const idx = order[active] ?? 0;
    return `<div class="card stepper">${steps.map((s,i) => `${i ? '<div class="step-line"></div>' : ''}<div class="step ${i < idx ? 'done' : i === idx ? 'active' : ''}"><span class="step-dot">${i < idx ? '✓' : i+1}</span><span class="step-label">${s[1]}</span></div>`).join('')}</div>`;
  }

  function renderProcess() {
    let body = '';
    switch (state.processStep) {
      case 'period': body = renderPeriod(); break;
      case 'upload': body = renderUpload(); break;
      case 'validation': body = renderValidation(); break;
      case 'processing': body = renderProcessing(); break;
      case 'review': body = renderReview(); break;
      case 'confirm': body = renderConfirm(); break;
      case 'result': body = renderResult(); break;
      default: body = renderPeriod();
    }
    app.innerHTML = shell(`${stepper(state.processStep)}${body}`); bindEvents();
  }

  function renderPeriod() {
    const months = cfg.INDONESIAN_MONTHS.map((m,i) => `<option value="${i+1}" ${state.period.month===i+1?'selected':''}>${m}</option>`).join('');
    const y = new Date().getFullYear();
    const years = Array.from({length:7},(_,i)=>y-3+i).map(v=>`<option ${state.period.year===v?'selected':''}>${v}</option>`).join('');
    return `<div class="page-title"><div><h2>Proses Lembur Baru</h2><p>Tentukan periode data presensi yang akan diproses.</p></div></div>
      <div class="card card-pad"><div class="card-title">Periode Lembur</div><div class="card-subtitle">Bulan dan tahun digunakan untuk memfilter data pada setiap file presensi.</div>
      <div class="form-grid section-gap"><div class="field"><label>Bulan</label><select id="period-month">${months}</select></div><div class="field"><label>Tahun</label><select id="period-year">${years}</select></div></div>
      ${renderHolidayManager()}
      <div class="actions"><span></span><button class="btn btn-primary" data-action="period-next">Selanjutnya -></button></div></div>`;
  }

  function renderUpload() {
    return `<div class="page-title"><div><h2>Upload File Presensi</h2><p>${esc(periodLabel(state.period))} · Mendukung banyak file .xlsx/.xls sekaligus.</p></div></div>
      <div class="card card-pad"><div id="dropzone" class="dropzone"><div class="drop-icon">⇧</div><h3>Tarik dan letakkan file presensi di sini</h3><p>atau pilih file dari komputer.</p><div style="margin-top:14px"><button class="btn btn-secondary btn-sm" data-action="choose-files">Pilih File</button></div><input class="file-input" id="file-input" type="file" multiple accept=".xlsx,.xls"></div>
      ${state.files.length ? `<div class="section-gap table-wrap"><table class="file-table"><thead><tr><th>No.</th><th>Nama File</th><th>Ukuran</th><th>Status</th><th></th></tr></thead><tbody>${state.files.map((file,i)=>`<tr><td>${i+1}</td><td>${esc(file.name)}</td><td>${formatBytes(file.size)}</td><td><span class="status ok">● Siap</span></td><td class="text-right"><button class="btn btn-secondary btn-sm" data-remove-file="${i}">Hapus</button></td></tr>`).join('')}</tbody></table></div>` : ''}
      <div class="actions"><button class="btn btn-secondary" data-action="back-period">← Kembali</button><div class="actions-right">${state.files.length ? '<button class="btn btn-secondary" data-action="choose-files">＋ Tambah File</button><button class="btn btn-primary" data-action="validate">Validasi Data →</button>' : ''}</div></div></div>`;
  }

  function renderValidation() {
    const v = state.validation || {results:[],errors:[],employees:[]};
    const ok = v.results.filter(r=>r.ok).length;
    const bad = v.results.length-ok;
    return `<div class="page-title"><div><h2>Validasi Data Presensi</h2><p>Hasil pemeriksaan struktur file dan periode data.</p></div></div>
      <div class="summary-cards"><div class="card summary-card"><div class="summary-icon">▤</div><div><div class="summary-value">${v.results.length}</div><div class="summary-label">File Diunggah</div></div></div><div class="card summary-card"><div class="summary-icon">✓</div><div><div class="summary-value">${ok}</div><div class="summary-label">Berhasil Dibaca</div></div></div><div class="card summary-card"><div class="summary-icon">!</div><div><div class="summary-value">${bad}</div><div class="summary-label">File Bermasalah</div></div></div></div>
      <div class="section-gap ${bad ? 'alert alert-danger' : 'alert alert-success'}"><div class="alert-title">${bad ? 'Validasi belum dapat dilanjutkan' : 'Data siap diproses'}</div>${bad ? 'Perbaiki atau hapus file bermasalah, kemudian lakukan validasi ulang.' : `Periode data: ${esc(periodLabel(state.period))} · Pegawai ditemukan: ${v.employees.length} · File valid: ${ok}`}</div>
      ${cfg.RULES.LATEST_OVERTIME_ARRIVAL_CONFIRMATION_REQUIRED ? '<div class="section-gap alert alert-warning"><div class="alert-title">Item aturan yang perlu dikonfirmasi</div>Source existing tidak konsisten untuk batas jam datang terakhir: komentar Main.gs menyebut 09:31, sedangkan konfigurasi Database.gs menetapkan 09:00. MVP ini memakai nilai konfigurasi 09:00 dan menandainya untuk konfirmasi.</div>' : ''}
      <div class="section-gap card"><div class="table-wrap" style="border:0"><table class="file-table"><thead><tr><th>Nama File</th><th>Status</th><th>Keterangan</th></tr></thead><tbody>${v.results.map(r=>`<tr><td>${esc(r.fileName)}</td><td><span class="status ${r.ok?'ok':'error'}">${r.ok?'● Berhasil':'● Gagal'}</span></td><td>${esc(r.ok ? `Pegawai: ${r.employee}` : r.error)}</td></tr>`).join('')}</tbody></table></div></div>
      <div class="actions"><button class="btn btn-secondary" data-action="back-upload">← Kembali</button>${bad ? '' : '<button class="btn btn-primary" data-action="process-data">Proses Data →</button>'}</div>`;
  }

  function renderProcessing() {
    const labels = ['Membaca file','Mengidentifikasi pegawai','Memvalidasi periode','Menghitung jam lembur','Menyiapkan hasil review'];
    return `<div class="page-title"><div><h2>Memproses Data Presensi</h2><p>Sistem menyiapkan data lembur untuk direview.</p></div></div><div class="card progress-wrap"><div class="progress-title"><h3>Memproses Data Presensi</h3><p>Jangan menutup halaman selama proses berlangsung.</p></div><div class="progress-bar"><div class="progress-fill" style="width:${state.processing.percent}%"></div></div><div class="progress-meta"><span>${state.processing.percent}%</span><span>${state.files.length} file</span></div><ul class="process-list">${labels.map((l,i)=>`<li class="${i<state.processing.active?'done':i===state.processing.active?'active':''}"><span class="process-bullet">${i<state.processing.active?'✓':i+1}</span>${l}</li>`).join('')}</ul></div>`;
  }

  // Kalender menggunakan tanggal lokal pukul 12.00 agar aman dari pergeseran zona waktu.
  function monthKeys(period) {
    const days = new Date(period.year, period.month, 0).getDate();
    return Array.from({ length: days }, (_, i) => `${period.year}-${rules.pad2(period.month)}-${rules.pad2(i + 1)}`);
  }

  function shortWeekday(key) {
    const weekdays = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
    return weekdays[rules.dateFromKey(key).getDay()];
  }

  function fullDayLabel(key) { return rules.formatIndonesianDate(rules.dateFromKey(key), true); }
  function autoHours(record) { return Number(record?.originalOvertimeHours ?? record?.overtimeHours ?? 0); }
  function finalHours(record) { return Number(record?.overtimeHours || 0); }
  function pendingDocumentReviews() {
    return window.LemburDocumentIntegration?.remaining?.(state.employees) || 0;
  }
  function documentReviewMarkup(record) {
    const sources=record?.documentSources||[];
    if(!sources.length)return '';
    const items=sources.map(item=>`<li>${esc(item.label)}</li>`).join('');
    return `<div class="alert ${record.documentBlocksOvertime?'alert-warning':'alert-info'}" style="margin-top:12px">
      <div class="alert-title">Verifikasi Cuti / Surat Tugas ${record.documentReviewPending?'— belum selesai':'— telah selesai'}</div>
      <ul style="margin:8px 0 8px 20px;padding:0">${items}</ul>
      ${record.documentBlocksOvertime?esc(record.documentBlockNote):'Dokumen tetap perlu diverifikasi. Surat Tugas tidak dipertanggungjawabkan tidak membatalkan klaim lembur secara otomatis.'}
    </div>`;
  }
  function renderDocumentPeriodReview() {
    const groups=window.LemburDocumentIntegration?.reviewGroups?.(state.employees)||[];
    if(!groups.length)return '';
    const pending=groups.filter(group=>!group.verified).length;
    return `<section class="lm-period-review" id="lm-document-period-review" aria-label="Verifikasi Cuti dan Surat Tugas per periode">
      <div class="lm-period-review-title"><div><h3>Verifikasi Cuti / Surat Tugas per Periode</h3>
        <p><b>Cukup satu kali verifikasi untuk setiap dokumen.</b> Keputusan berlaku untuk semua tanggal dan pegawai terkait yang terdapat dalam proses lembur ini. Koreksi jam presensi tetap dilakukan secara terpisah.</p></div>
        <span class="lm-review-counter ${pending?'pending':'complete'}">${pending?`${pending} dokumen perlu verifikasi`:'Semua dokumen diverifikasi'}</span></div>
      <div class="lm-period-review-items">${groups.map(group=>{
        const noClaim=group.blockByDefault;
        const status=group.verified?'Sudah diverifikasi':'Perlu verifikasi';
        const isLeave=group.kind==='cuti';
        const value=group.verified?group.choice:(noClaim?'deny':'allow');
        return `<article class="lm-period-document ${group.verified?'verified':'pending'}">
          <div class="lm-period-document-head"><div><span class="lm-document-type ${isLeave?'leave':'assignment'}">${isLeave?'Cuti':'Surat Tugas'}</span><strong>${esc(group.label)}</strong></div><span class="lm-document-status ${group.verified?'verified':'pending'}">${status}</span></div>
          <div class="lm-document-meta"><span><b>Periode:</b> ${esc(group.startDate)} s.d. ${esc(group.endDate)}</span><span><b>Cakupan:</b> ${group.dates.length} tanggal, ${group.employees.length} pegawai (${group.records.length} tanggal-pegawai)</span>
            <span><b>Pegawai:</b> ${esc(group.employees.slice(0,4).join(', '))}${group.employees.length>4?` dan ${group.employees.length-4} lainnya`:''}</span></div>
          <div class="lm-document-controls"><label>Hasil Setelah Verifikasi
            <select class="lm-form-input" data-lm-document-choice="${esc(group.key)}" aria-label="Keputusan klaim lembur untuk ${esc(group.label)}">
              <option value="allow" ${value==='allow'?'selected':''} ${noClaim?'disabled':''}>Dapat klaim lembur (sesuai perhitungan presensi)</option>
              <option value="deny" ${value==='deny'?'selected':''}>Tidak dapat klaim lembur (0 jam)</option>
            </select></label>
            <button type="button" class="btn btn-primary btn-sm" data-lm-document-save="${esc(group.key)}">${group.verified?'Perbarui Verifikasi Periode':'Simpan Verifikasi Periode'}</button></div>
          <p class="lm-document-hint">Rekomendasi awal: <b>${noClaim?'Tidak dapat klaim lembur':'Dapat klaim lembur jika memenuhi aturan presensi'}</b>. ${noClaim?'Dokumen ini masih berlaku untuk pertanggungjawaban. Jika pegawai seharusnya dapat mengklaim, ubah status Surat Tugas atau validitas Cuti pada modul sumber terlebih dahulu.':'Operator dapat memilih Tidak dapat klaim apabila dokumen atau realisasi tidak mendukung klaim.'} ${group.verified?'Keputusan telah disimpan pada seluruh tanggal dan pegawai terkait.':'Keputusan belum disimpan.'}</p>
        </article>`;
      }).join('')}</div></section>`;
  }
  function saveDocumentPeriodReview(key) {
    if(!canCorrectOvertime())return;
    const groups=window.LemburDocumentIntegration?.reviewGroups?.(state.employees)||[];
    const group=groups.find(item=>item.key===key);
    if(!group)return alert('Dokumen tidak ditemukan. Sinkronkan ulang data proses.');
    const control=[...document.querySelectorAll('[data-lm-document-choice]')].find(el=>el.dataset.lmDocumentChoice===key);
    if(!control)return;
    try {
      const actor=window.FirebaseClient?.getCurrentUser?.()||{};
      window.LemburDocumentIntegration.verifySource(state.employees,key,group.signature,control.value,actor);
      for(const {employee,date} of group.records)state.dirtyKeys.add(`${rules.employeeKey(employee)}|${date}`);
      state.generated=null;
      renderStay();
    }catch(error){alert(error.message||String(error));}
  }
  function scrollToDocumentPeriodReview() {
    document.getElementById('lm-document-period-review')?.scrollIntoView({behavior:'smooth',block:'start'});
  }
  function isManualOvertime(record) {
    return Boolean(record && !record.documentBlocksOvertime && (record.manualOverride === true || finalHours(record) !== autoHours(record)));
  }
  function hasOvertime(record) { return record && finalHours(record) > 0; }
  function mealEligible(record) { return finalHours(record) >= Number(cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS || 2); }
  function activeReviewEmployee() {
    const list = state.employees || [];
    let current = list.find(e => rules.employeeKey(e) === state.reviewEmployeeKey);
    if (!current) {
      current = list.find(e => Object.values(e.records || {}).some(r=>r.documentReviewPending))
        || list.find(e => Object.values(e.records || {}).some(hasOvertime)) || list[0];
      state.reviewEmployeeKey = current ? rules.employeeKey(current) : null;
    }
    return current || null;
  }
  function activeReviewDate() {
    if (!dateBelongsToPeriod(state.reviewDateKey, state.period)) {
      const employee = activeReviewEmployee();
      const first = Object.keys(employee?.records || {}).sort().find(key => employee.records[key]?.documentReviewPending)
        || Object.keys(employee?.records || {}).sort().find(key => hasOvertime(employee.records[key]));
      state.reviewDateKey = first || monthKeys(state.period)[0];
    }
    return state.reviewDateKey;
  }
  function activeResultDate(employee, period) {
    if (!dateBelongsToPeriod(state.detailDateKey, period)) {
      state.detailDateKey = Object.keys(employee.records || {}).sort().find(k => hasOvertime(employee.records[k])) || monthKeys(period)[0];
    }
    return state.detailDateKey;
  }
  function monthCalendarMarkup(period, renderCell) {
    const keys = monthKeys(period);
    const before = (rules.dateFromKey(keys[0]).getDay() + 6) % 7;
    const empty = Array.from({ length: before }, () => '<span class="lm-cal-blank" aria-hidden="true"></span>').join('');
    return `<div class="lm-cal-weekdays">${['Sen','Sel','Rab','Kam','Jum','Sab','Min'].map((day,i) => `<span class="${i > 4 ? 'lm-wknd' : ''}">${day}</span>`).join('')}</div>
      <div class="lm-calendar">${empty}${keys.map(renderCell).join('')}</div>`;
  }
  function employeeStats(employee) {
    const records = Object.values(employee?.records || {}).filter(hasOvertime);
    return { hours: records.reduce((sum,r) => sum + finalHours(r), 0), days: records.length,
      meals: records.filter(mealEligible).length };
  }
  function reviewRows() {
    const rows = [];
    state.employees.forEach(employee => {
      Object.keys(employee.records || {}).sort().forEach(key => {
        const record = employee.records[key];
        if (!record) return;
        if (state.reviewTableFilter === 'overtime' && !(autoHours(record) > 0 || finalHours(record) > 0 || isManualOvertime(record) || record.documentSources?.length)) return;
        if (state.reviewTableFilter === 'adjusted' && !isManualOvertime(record)) return;
        if (state.reviewTableFilter === 'documents' && !record.documentReviewPending) return;
        rows.push({ employee, record, key });
      });
    });
    const q = state.filter.trim().toLowerCase();
    return rows.filter(({ employee, record, key }) => !q || employee.name.toLowerCase().includes(q)
      || String(employee.nip || '').includes(q) || fullDayLabel(key).toLowerCase().includes(q));
  }
  function reviewStatus(record, key) {
    if (record?.documentReviewPending) return 'needs-review';
    if (isManualOvertime(record)) return 'adjusted';
    if (rules.isWeekendLike(rules.dateFromKey(key), state.holidays)) return 'holiday';
    if (hasOvertime(record)) return 'overtime';
    return 'empty';
  }
  function renderReviewEmployeeList() {
    const selected = activeReviewEmployee();
    const q = state.filter.trim().toLowerCase();
    const list = state.employees.filter(e => !q || String(e.name || '').toLowerCase().includes(q) || String(e.nip || '').includes(q));
    return `<aside class="lm-person-list"><div class="lm-person-heading">Daftar Pegawai <span>${list.length}</span></div>
      <div class="lm-person-scroller">${list.length ? list.map(e => {
        const stats = employeeStats(e), key = rules.employeeKey(e);
        const changed = Object.values(e.records || {}).filter(isManualOvertime).length;
        const pending = Object.values(e.records || {}).filter(r=>r.documentReviewPending).length;
        return `<button type="button" class="lm-person ${selected && rules.employeeKey(selected) === key ? 'selected' : ''}" data-review-person="${esc(key)}" aria-pressed="${selected && rules.employeeKey(selected) === key}">
          <span class="lm-person-name">${esc(e.name)}</span><span class="lm-person-meta">${esc(e.nip || 'NIP belum tersedia')}</span>
          <span class="lm-person-stats">${stats.days} hari · ${stats.hours} jam ${changed ? `· ${changed} koreksi` : ''}${pending ? ` · ${pending} perlu verifikasi` : ''}</span></button>`;
      }).join('') : '<div class="lm-empty">Pegawai tidak ditemukan.</div>'}</div></aside>`;
  }
  function renderReviewEmployeeCalendar() {
    const employee = activeReviewEmployee();
    const selected = activeReviewDate();
    if (!employee) return '<div class="lm-empty">Tidak ada data pegawai yang dapat diperiksa.</div>';
    const calendar = monthCalendarMarkup(state.period, key => {
      const record = employee.records?.[key];
      const status = reviewStatus(record, key);
      const weekend = rules.isWeekendLike(rules.dateFromKey(key), state.holidays);
      const faded = state.reviewFilter === 'overtime' && !hasOvertime(record) && !record?.documentSources?.length
        || state.reviewFilter === 'adjusted' && !isManualOvertime(record)
        || state.reviewFilter === 'holiday' && !weekend
        || state.reviewFilter === 'documents' && !record?.documentReviewPending;
      return `<button type="button" class="lm-cal-day lm-${status} ${weekend ? 'lm-rest' : ''} ${selected === key ? 'selected' : ''} ${faded ? 'lm-dim' : ''}" data-review-day="${key}" aria-pressed="${selected === key}" aria-label="${esc(fullDayLabel(key))}, ${record ? `${finalHours(record)} jam lembur` : 'tidak ada catatan presensi'}">
        <span class="lm-date">${Number(key.slice(-2))}</span><span class="lm-dayname">${shortWeekday(key)}</span><strong>${record && finalHours(record) > 0 ? `${finalHours(record)} jam` : '–'}</strong>
        <small>${record?.documentReviewPending?'Perlu verifikasi':record?.documentBlocksOvertime?'Dibatasi dokumen':isManualOvertime(record) ? 'Koreksi' : weekend ? rules.isHoliday(rules.dateFromKey(key), state.holidays) ? 'Tanggal merah' : 'Libur' : hasOvertime(record) ? 'Lembur' : record ? '0 jam' : 'Tanpa data'}</small>
      </button>`;
    });
    return `<div class="lm-cal-layout">${renderReviewEmployeeList()}<section class="lm-cal-panel">
      <div class="lm-cal-panel-head"><div><strong>${esc(employee.name)}</strong><span>${esc(periodLabel(state.period))} · ${esc(employee.nip || '-')}</span></div><div class="lm-calendar-help">Klik tanggal untuk detail dan koreksi.</div></div>
      ${calendar}${renderReviewDayPanel(employee, selected)}</section></div>`;
  }
  function renderReviewDayPanel(employee, key) {
    const record = employee?.records?.[key];
    const category = resultCategory(record || { date: rules.dateFromKey(key) }, state.holidays);
    return `<section class="lm-day-panel" id="lm-review-day-panel" aria-live="polite">
      <div class="lm-day-panel-head"><div><strong>${esc(fullDayLabel(key))}</strong><span class="category-pill ${category.rest ? 'rest-day' : ''}">${esc(category.label)}</span></div>
      ${record ? record.documentSources?.length ? `<button type="button" class="btn btn-primary btn-sm" data-action="open-period-document-review">Buka Verifikasi Periode</button>` : `<button type="button" class="btn btn-primary btn-sm" data-review-edit="${esc(rules.employeeKey(employee))}|${esc(key)}">Koreksi Jam Lembur</button>` : ''}</div>
      ${record ? `<div class="lm-day-grid">
        <div><span>Jam Masuk</span><strong>${esc(rules.formatMinutes(record.inMinutes))}</strong></div>
        <div><span>Mulai Lembur</span><strong>${esc(rules.formatMinutes(record.overtimeStartMinutes))}</strong></div>
        <div><span>Jam Pulang</span><strong>${esc(rules.formatMinutes(record.outMinutes))}</strong></div>
        <div><span>Hasil Otomatis</span><strong>${autoHours(record)} jam</strong></div>
        <div class="lm-final-value"><span>Hasil Final</span><strong>${finalHours(record)} jam</strong></div>
        <div><span>Uang Makan Lembur</span><strong>${mealEligible(record) ? 'Memenuhi' : 'Tidak'}</strong></div>
      </div><div class="lm-day-subrow"><span>Status presensi: ${esc(record.status || '-')}</span>${isManualOvertime(record) ? `<span>Alasan koreksi: ${esc(record.adjustmentReason || 'Koreksi lama, alasan belum dicatat')}</span>` : ''}${record.documentSources?.length?`<span>${record.documentReviewPending?'Perlu verifikasi dokumen':'Dokumen diverifikasi'}${record.documentBlocksOvertime?' · klaim lembur ditolak':''}</span>`:''}</div>${documentReviewMarkup(record)}`
      : '<div class="lm-empty lm-inline">Tidak ada data presensi untuk tanggal ini. Data lembur baru tidak dapat dibuat melalui editor.</div>'}</section>`;
  }
  function dailyReviewStats(key) {
    let count = 0, hours = 0, corrected = 0, records = 0, pending = 0;
    state.employees.forEach(e => { const r = e.records?.[key]; if(!r) return; records += 1;
      if(hasOvertime(r)){count += 1; hours += finalHours(r);} if(isManualOvertime(r)) corrected += 1; if(r.documentReviewPending)pending += 1; });
    return { count, hours, corrected, records, pending };
  }
  function renderReviewDateCalendar() {
    const selected = activeReviewDate();
    const calendar = monthCalendarMarkup(state.period, key => {
      const info = dailyReviewStats(key);
      const date = rules.dateFromKey(key);
      const weekend = rules.isWeekendLike(date, state.holidays);
      const faded = state.reviewFilter === 'overtime' && !info.count || state.reviewFilter === 'adjusted' && !info.corrected
        || state.reviewFilter === 'holiday' && !weekend || state.reviewFilter === 'documents' && !info.pending;
      return `<button type="button" class="lm-cal-day ${info.pending ? 'lm-needs-review' : info.corrected ? 'lm-adjusted' : info.count ? 'lm-overtime' : 'lm-empty-day'} ${weekend ? 'lm-rest' : ''} ${selected === key ? 'selected' : ''} ${faded ? 'lm-dim' : ''}" data-review-day="${key}" aria-pressed="${selected === key}">
        <span class="lm-date">${Number(key.slice(-2))}</span><span class="lm-dayname">${shortWeekday(key)}</span><strong>${info.hours ? `${info.hours} jam` : '–'}</strong>
        <small>${info.pending ? `${info.pending} perlu verifikasi` : info.count ? `${info.count} pegawai` : weekend ? 'Libur' : 'Tanpa lembur'}</small></button>`;
    });
    const rows = state.employees.map(e => ({ employee:e, record:e.records?.[selected] })).filter(x => x.record);
    return `<section class="lm-date-view"><div class="lm-cal-panel-head"><div><strong>Kalender per Tanggal</strong><span>${esc(periodLabel(state.period))} · Total pegawai dan jam lembur per hari</span></div></div>
      ${calendar}<div class="lm-day-panel" id="lm-review-day-panel"><div class="lm-day-panel-head"><strong>${esc(fullDayLabel(selected))}</strong><span class="lm-count">${rows.length} data presensi</span></div>
      ${rows.length ? `<div class="table-wrap"><table class="data-table lm-rows"><thead><tr><th>Pegawai</th><th>Masuk</th><th>Mulai</th><th>Pulang</th><th>Otomatis</th><th>Final</th><th>Aksi</th></tr></thead><tbody>${rows.map(({employee,record}) => `<tr><td><b>${esc(employee.name)}</b></td><td>${esc(rules.formatMinutes(record.inMinutes))}</td><td>${esc(rules.formatMinutes(record.overtimeStartMinutes))}</td><td>${esc(rules.formatMinutes(record.outMinutes))}</td><td>${autoHours(record)} jam</td><td>${finalHours(record)} jam ${isManualOvertime(record) ? '<span class="lm-mod-label">Koreksi</span>' : ''}${record.documentSources?.length?`<span class="lm-mod-label">${record.documentReviewPending?'Perlu verifikasi dokumen':record.documentBlocksOvertime?'Tidak dapat diklaim':'Dokumen diperiksa'}</span>`:''}</td><td>${record.documentSources?.length?'<button class="btn btn-secondary btn-sm" type="button" data-action="open-period-document-review">Verifikasi Periode</button>':`<button class="btn btn-secondary btn-sm" type="button" data-review-edit="${esc(rules.employeeKey(employee))}|${esc(selected)}">Edit</button>`}</td></tr>`).join('')}</tbody></table></div>` : '<div class="lm-empty lm-inline">Tidak ada data presensi pada tanggal ini.</div>'}</div></section>`;
  }
  function renderReviewTable() {
    const rows = reviewRows();
    return `<div class="lm-table-toolbar"><div class="search"><input id="review-search" value="${esc(state.filter)}" placeholder="Cari nama, NIP, atau tanggal..."></div>
      <select id="review-table-filter" class="lm-select" aria-label="Filter tabel"><option value="overtime" ${state.reviewTableFilter === 'overtime' ? 'selected' : ''}>Ada lembur / koreksi</option><option value="all" ${state.reviewTableFilter === 'all' ? 'selected' : ''}>Semua data presensi</option><option value="adjusted" ${state.reviewTableFilter === 'adjusted' ? 'selected' : ''}>Hanya koreksi</option><option value="documents" ${state.reviewTableFilter === 'documents' ? 'selected' : ''}>Perlu verifikasi dokumen</option></select>
      <span class="lm-count">${rows.length} baris</span></div>
      <div class="card table-wrap"><table class="data-table lm-rows"><thead><tr><th>No.</th><th>Pegawai</th><th>Tanggal</th><th>Kategori</th><th>Masuk</th><th>Pulang</th><th>Status</th><th>Otomatis</th><th>Final</th><th>Aksi</th></tr></thead><tbody>
        ${rows.length ? rows.map(({employee,record,key},i) => { const category = resultCategory(record, state.holidays); return `<tr><td>${i+1}</td><td><button type="button" class="employee-link" data-review-person="${esc(rules.employeeKey(employee))}">${esc(employee.name)}</button></td><td class="nowrap">${esc(fullDayLabel(key))}</td><td><span class="category-pill ${category.rest ? 'rest-day' : ''}">${category.label}</span></td><td>${esc(rules.formatMinutes(record.inMinutes))}</td><td>${esc(rules.formatMinutes(record.outMinutes))}</td><td>${esc(record.status || '-')}</td><td>${autoHours(record)} jam</td><td><b>${finalHours(record)} jam</b> ${isManualOvertime(record) ? '<span class="lm-mod-label">Koreksi</span>' : ''}${record.documentSources?.length?`<span class="lm-mod-label">${record.documentReviewPending?'Perlu verifikasi dokumen':record.documentBlocksOvertime?'Tidak dapat diklaim':'Dokumen diperiksa'}</span>`:''}</td><td>${record.documentSources?.length?'<button class="btn btn-secondary btn-sm" type="button" data-action="open-period-document-review">Verifikasi Periode</button>':`<button class="btn btn-secondary btn-sm" type="button" data-review-edit="${esc(rules.employeeKey(employee))}|${esc(key)}">Edit</button>`}</td></tr>`; }).join('') : '<tr><td colspan="10" class="lm-empty">Tidak ada data pada filter ini.</td></tr>'}
      </tbody></table></div>`;
  }
  function renderReview() {
    const summary = rules.summarize(state.employees);
    const isEdit = Boolean(state.editingHistoryId);
    const actions = isEdit ? `<button class="btn btn-secondary" data-action="cancel-history-edit">Kembali ke Riwayat</button><button class="btn btn-primary" data-action="save-history-edit" ${state.busy ? 'disabled' : ''}>${state.busy ? 'Menyimpan...' : 'Simpan & Generate Ulang'}</button>`
      : `<button class="btn btn-secondary" data-action="back-validation">Kembali</button><button class="btn btn-primary" data-action="to-confirm" ${pendingDocumentReviews()?'disabled':''}>Konfirmasi & Lanjut</button>`;
    return `<div class="lm-editor" data-history-edit="${isEdit}">
      <div class="page-title"><div><h2>${isEdit ? 'Edit Data Lembur' : 'Review Data Lembur'}</h2><p>${esc(periodLabel(state.period))} · Tinjau hasil otomatis dan koreksi jam lembur final.</p></div></div>
      ${isEdit ? '<div class="alert alert-warning"><div class="alert-title">Mode edit riwayat</div>Koreksi akan mengganti hasil tersimpan setelah Anda menekan Simpan & Generate Ulang.</div>' : ''}
      ${pendingDocumentReviews()?`<div class="alert alert-warning"><div class="alert-title">${pendingDocumentReviews()} tanggal-pegawai memerlukan verifikasi dokumen</div>Gunakan panel <b>Verifikasi Cuti / Surat Tugas per Periode</b> di bawah. Cukup satu kali verifikasi per dokumen, dan keputusan berlaku pada semua tanggal serta pegawai terkait.</div>`:''}
      <div class="grid-4 section-gap">${metric(summary.employees,'Pegawai')}${metric(summary.overtimeEmployees,'Pegawai Lembur')}${metric(summary.totalHours,'Total Jam Lembur')}${metric(summary.mealDays,'Hari Uang Makan')}</div>
      ${renderHolidayManager()}
      ${renderDocumentPeriodReview()}
      <div class="card lm-review-card"><div class="lm-review-head"><div><div class="card-title">Pemeriksaan dan Koreksi</div><div class="card-subtitle">Gunakan kalender untuk meninjau per pegawai atau per tanggal; tabel tetap tersedia.</div></div>
      <div class="lm-segment" role="group" aria-label="Tampilan editor">${['calendar','table'].map((mode)=>`<button type="button" data-review-mode="${mode}" aria-pressed="${state.reviewMode === mode}" class="${state.reviewMode === mode ? 'active' : ''}">${mode === 'calendar' ? 'Kalender' : 'Tabel'}</button>`).join('')}</div></div>
      ${state.reviewMode === 'calendar' ? `<div class="lm-review-controls"><div class="lm-segment" role="group" aria-label="Perspektif kalender"><button type="button" class="${state.reviewGroup === 'employee' ? 'active' : ''}" data-review-group="employee">Per Pegawai</button><button type="button" class="${state.reviewGroup === 'date' ? 'active' : ''}" data-review-group="date">Per Tanggal</button></div>
        ${state.reviewGroup === 'employee' ? `<input id="lm-person-search" class="lm-input" type="search" placeholder="Cari pegawai / NIP" value="${esc(state.filter)}">` : ''}
        <select id="lm-review-filter" class="lm-select" aria-label="Filter kalender"><option value="all" ${state.reviewFilter === 'all' ? 'selected' : ''}>Semua Tanggal</option><option value="overtime" ${state.reviewFilter === 'overtime' ? 'selected' : ''}>Ada Lembur</option><option value="adjusted" ${state.reviewFilter === 'adjusted' ? 'selected' : ''}>Sudah Dikoreksi</option><option value="holiday" ${state.reviewFilter === 'holiday' ? 'selected' : ''}>Hari Libur</option><option value="documents" ${state.reviewFilter === 'documents' ? 'selected' : ''}>Perlu Verifikasi</option></select></div>
      <div class="lm-legend"><span><i class="lm-dot lm-dot-orange"></i>Perlu verifikasi</span><span><i class="lm-dot lm-dot-green"></i>Ada lembur</span><span><i class="lm-dot lm-dot-blue"></i>Sudah dikoreksi</span><span><i class="lm-dot lm-dot-gold"></i>Libur/tanggal merah</span><span><i class="lm-dot lm-dot-gray"></i>Tanpa lembur</span></div>
      ${state.reviewGroup === 'employee' ? renderReviewEmployeeCalendar() : renderReviewDateCalendar()}` : renderReviewTable()}
      </div><div class="lm-editor-bottom"><div><strong>${state.dirtyKeys.size ? `${state.dirtyKeys.size} perubahan belum disimpan` : 'Tidak ada perubahan yang menunggu disimpan'}</strong>
      <span>Jam otomatis tetap tercatat; hasil final dan alasan koreksi disimpan saat proses diselesaikan.</span></div><div class="lm-editor-buttons">${actions}</div></div>
    </div>`;
  }

  function renderReviewEditModal() {
    if (!state.reviewEditKey) return '';
    const [employeeKey, dateKey] = state.reviewEditKey.split('|');
    const employee = state.employees.find(e => rules.employeeKey(e) === employeeKey);
    const record = employee?.records?.[dateKey];
    if (!record) return '';
    const automatic = autoHours(record);
    return `<div class="lm-overlay" data-action="close-review-edit-backdrop"><section class="lm-small-modal" role="dialog" aria-modal="true" aria-labelledby="lm-edit-title">
      <header class="lm-small-head"><div><h3 id="lm-edit-title">Koreksi Jam Lembur</h3><p>${esc(employee.name)} · ${esc(fullDayLabel(dateKey))}</p></div><button type="button" class="icon-btn" data-action="close-review-edit" aria-label="Tutup">×</button></header>
      <div class="lm-small-body"><div class="lm-correction-comparison"><div><span>Hasil otomatis</span><strong>${automatic} jam</strong></div><span class="lm-arrow">→</span><div><span>Hasil final</span><strong>${finalHours(record)} jam</strong></div></div>
        <div class="lm-day-grid lm-edit-presence"><div><span>Jam masuk</span><strong>${esc(rules.formatMinutes(record.inMinutes))}</strong></div><div><span>Mulai lembur</span><strong>${esc(rules.formatMinutes(record.overtimeStartMinutes))}</strong></div><div><span>Jam pulang</span><strong>${esc(rules.formatMinutes(record.outMinutes))}</strong></div></div>
        ${documentReviewMarkup(record)}
        <label class="lm-form-label" for="lm-edit-hours">Jam lembur final</label><select class="lm-form-input" id="lm-edit-hours">${(record.documentBlocksOvertime?[0]:[0,1,2,3,4]).map(n => `<option value="${n}" ${n === finalHours(record) ? 'selected' : ''}>${n} jam${n === automatic ? ' (otomatis)' : ''}</option>`).join('')}</select>
        <label class="lm-form-label" for="lm-edit-reason">Alasan koreksi <span>(wajib apabila berbeda dari otomatis)</span></label>
        <textarea class="lm-form-input lm-textarea" id="lm-edit-reason" maxlength="600" rows="3" placeholder="Jelaskan koreksi dan dasar penyesuaian">${esc(record.adjustmentReason || '')}</textarea>
        ${record.adjustmentAt ? `<p class="lm-audit">Koreksi terakhir: ${esc(record.adjustmentBy || '-')}, ${esc(formatDateTime(record.adjustmentAt))}</p>` : ''}
        ${Array.isArray(record.adjustmentHistory) && record.adjustmentHistory.length ? `<details class="lm-audit-history"><summary>Riwayat Koreksi (${record.adjustmentHistory.length})</summary><div>${record.adjustmentHistory.slice().reverse().map(item => `<p><strong>${Number(item.before) || 0} → ${Number(item.after) || 0} jam</strong><span>${esc(item.reason || '-')}<br>${esc(item.operator || '-')} · ${esc(item.timestamp ? formatDateTime(item.timestamp) : '-')}</span></p>`).join('')}</div></details>` : ''}
        <div class="lm-modal-note">Perubahan hanya pada hasil final. Data presensi asal tidak diubah.</div></div>
      <footer class="lm-small-foot"><button class="btn btn-secondary" type="button" data-action="close-review-edit">Batal</button><button class="btn btn-primary" type="button" data-action="save-review-edit">Simpan Koreksi</button></footer>
    </section></div>`;
  }
  function renderHolidayPreviewModal() {
    if (!state.pendingHoliday) return '';
    const pending = state.pendingHoliday;
    const impact = reviewHolidayImpact();
    return `<div class="lm-overlay" data-action="close-holiday-preview-backdrop"><section class="lm-small-modal" role="dialog" aria-modal="true" aria-labelledby="lm-holiday-title">
      <header class="lm-small-head"><div><h3 id="lm-holiday-title">Konfirmasi Tanggal Merah</h3><p>${pending.operation === 'add' ? 'Tambahkan' : 'Hapus'} ${esc(fullDayLabel(pending.key))}</p></div><button type="button" class="icon-btn" data-action="close-holiday-preview">×</button></header>
      <div class="lm-small-body"><div class="lm-impact-stats"><div><strong>${impact.changed}</strong><span>perhitungan otomatis berubah</span></div><div><strong>${impact.manualCount}</strong><span>koreksi manual dipertahankan</span></div></div>
      <p class="lm-quiet">Sistem akan menghitung ulang hasil otomatis pada tanggal ini. Nilai final dari koreksi manual tidak akan dihapus.</p>
      ${impact.affected.length ? `<div class="lm-impact-list">${impact.affected.slice(0,8).map(item => `<div><span>${esc(item.employee.name)}${item.manual ? ' · Koreksi manual' : ''}</span><b>${item.before} → ${item.next} jam</b></div>`).join('')}${impact.affected.length > 8 ? `<p>Dan ${impact.affected.length - 8} pegawai lainnya.</p>` : ''}</div>` : '<div class="lm-empty lm-inline">Tidak ada perubahan jam otomatis pada data yang tersedia.</div>'}
      </div><footer class="lm-small-foot"><button class="btn btn-secondary" type="button" data-action="close-holiday-preview">Batal</button><button class="btn btn-primary" type="button" data-action="confirm-holiday-preview">Terapkan Perubahan</button></footer></section></div>`;
  }

  function openReviewEdit(key) {
    if (!canCorrectOvertime()) return;
    const [employeeKey, dateKey] = key.split('|');
    const record=state.employees.find(e=>rules.employeeKey(e)===employeeKey)?.records?.[dateKey];
    if(!record)return;
    if(record.documentSources?.length){scrollToDocumentPeriodReview();return;}
    state.reviewEditKey = key;
    renderStay();
  }
  function saveReviewEdit() {
    if (!state.reviewEditKey || !canCorrectOvertime()) return;
    const [employeeKey, dateKey] = state.reviewEditKey.split('|');
    const employee = state.employees.find(e => rules.employeeKey(e) === employeeKey);
    const record = employee?.records?.[dateKey];
    if (!record) return;
    const value = Number(document.getElementById('lm-edit-hours')?.value);
    const reason = String(document.getElementById('lm-edit-reason')?.value || '').trim() || (record.documentBlocksOvertime ? record.documentBlockNote : '');
    const automatic = autoHours(record);
    if (!Number.isInteger(value) || value < 0 || value > 4) return alert('Jam lembur final harus antara 0 sampai 4 jam.');
    if (record.documentBlocksOvertime && value !== 0) return alert('Cuti atau Surat Tugas yang dipertanggungjawabkan tidak dapat diklaim sebagai lembur.');
    if (record.documentReviewPending) return alert('Selesaikan Verifikasi Cuti / Surat Tugas per Periode sebelum melakukan koreksi jam.');
    if (value !== automatic && reason.length < 3) return alert('Isi alasan koreksi minimal 3 karakter.');
    if (value !== finalHours(record) || (value !== automatic && reason !== (record.adjustmentReason || '')) || (value === automatic && isManualOvertime(record))) {
      const actor = window.FirebaseClient?.getCurrentUser?.();
      const at = new Date().toISOString();
      const historyItem = { before: finalHours(record), after: value, automatic, reason: value === automatic ? 'Dikembalikan ke nilai otomatis' : reason,
        operator: actor?.displayName || actor?.email || 'Operator', operatorUid: actor?.uid || '', timestamp: at };
      record.adjustmentHistory = [...(Array.isArray(record.adjustmentHistory) ? record.adjustmentHistory : []), historyItem].slice(-12);
      record.adjustmentAt = at;
      record.adjustmentBy = historyItem.operator;
      record.adjustmentByUid = historyItem.operatorUid;
      record.adjustmentReason = value === automatic ? '' : reason;
      record.manualOverride = !record.documentBlocksOvertime && value !== automatic;
      record.overtimeHours = value;
      state.dirtyKeys.add(`${employeeKey}|${dateKey}`);
      state.generated = null;
    }
    state.reviewEditKey = null;
    state.reviewEmployeeKey = employeeKey;
    state.reviewDateKey = dateKey;
    renderStay();
  }
  function renderConfirm() {
    const s = rules.summarize(state.employees);
    return `<div class="page-title"><div><h2>Siap Membuat Dokumen</h2><p>Periksa ringkasan sebelum file Excel dibuat.</p></div></div><div class="card card-pad">
      <div class="two-col"><div><div class="card-title">Ringkasan Proses</div><div class="detail-grid section-gap"><div><div class="detail-label">Periode</div><div class="detail-value">${esc(periodLabel(state.period))}</div></div><div><div class="detail-label">Jumlah Pegawai</div><div class="detail-value">${s.employees}</div></div><div><div class="detail-label">Pegawai Lembur</div><div class="detail-value">${s.overtimeEmployees}</div></div><div><div class="detail-label">Total Lembur</div><div class="detail-value">${s.totalHours} jam</div></div><div><div class="detail-label">Hari Uang Makan</div><div class="detail-value">${s.mealDays} hari</div></div></div></div>
      <div><div class="card-title">Dokumen yang akan dibuat</div><ul class="confirm-list section-gap"><li><span class="check">☑</span>Rekapitulasi Lembur</li><li><span class="check">☑</span>Daftar Hadir Kerja Lembur</li><li><span class="check">☑</span>SPKL Hari Kerja</li><li><span class="check">☑</span>SPKL Weekend</li></ul><div class="card-subtitle">SPKL Hari Kerja dan WEEKEND berada dalam satu workbook Excel.</div></div></div>
      <div class="actions"><button class="btn btn-secondary" data-action="back-review">← Kembali ke Review</button><button class="btn btn-primary" data-action="generate" ${state.busy||pendingDocumentReviews()?'disabled':''}>${state.busy?'Membuat Dokumen...':'Generate Dokumen'}</button></div></div>`;
  }

  function resultCategory(record, holidays) {
    const holiday = rules.isHoliday(record.date, holidays);
    const weekend = rules.isWeekend(record.date);
    if (holiday) return { key: 'holiday', label: 'Tanggal Merah', rest: true };
    if (weekend) return { key: 'weekend', label: 'Weekend', rest: true };
    return { key: 'workday', label: 'Hari Kerja', rest: false };
  }

  function resultEmployeeRecap(employees) {
    return (employees || []).map((employee) => {
      const records = Object.values(employee.records || {}).filter((record) => Number(record.overtimeHours || 0) > 0);
      return {
        employee,
        totalHours: records.reduce((sum, record) => sum + Number(record.overtimeHours || 0), 0),
        overtimeDays: records.length,
        mealDays: records.filter((record) => Number(record.overtimeHours || 0) >= Number(cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS || 2)).length
      };
    }).filter((item) => item.totalHours > 0);
  }

  function resultDetailRows(employees, holidays) {
    return rules.collectOvertimeRows(employees || []).map((item) => ({
      ...item,
      category: resultCategory(item.record, holidays),
      meal: Number(item.record.overtimeHours || 0) >= Number(cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS || 2)
    }));
  }

  function renderResultTabs() {
    const tabs = [
      ['summary', 'Ringkasan'],
      ['employees', 'Rekap Pegawai'],
      ['details', 'Rincian Lembur']
    ];
    return `<div class="result-tabs" role="tablist">${tabs.map(([key, label]) => `<button class="result-tab ${state.resultTab === key ? 'active' : ''}" type="button" data-result-tab="${key}" role="tab" aria-selected="${state.resultTab === key ? 'true' : 'false'}">${label}</button>`).join('')}</div>`;
  }

  function renderResultSummary(run, period, summary, holidays) {
    return `<div class="result-panel">
      <div class="grid-4">${metric(summary.employees,'Pegawai')}${metric(summary.overtimeEmployees,'Pegawai Lembur')}${metric(summary.totalHours,'Total Jam Lembur')}${metric(summary.mealDays,'Hari Uang Makan')}</div>
      ${holidays.length ? `<div class="result-holidays"><strong>Tanggal merah:</strong> ${holidays.map((key)=>esc(holidayLabel(key))).join(', ')}</div>` : ''}
      <div class="result-section-head"><div><div class="card-title">Dokumen Hasil Generate</div><div class="card-subtitle">Data rekap dapat diperiksa langsung di tab Rekap Pegawai dan Rincian Lembur sebelum dokumen diunduh.</div></div></div>
      <div class="download-list"><div class="download-row"><div class="file-icon">▤</div><div><div class="download-name">Rekapitulasi Lembur</div><div class="download-meta">${esc(periodLabel(period))}</div></div><button class="btn btn-secondary btn-sm" data-download="recap">Download</button></div><div class="download-row"><div class="file-icon">▦</div><div><div class="download-name">Daftar Hadir Kerja Lembur</div><div class="download-meta">Workbook dengan selector tanggal</div></div><button class="btn btn-secondary btn-sm" data-download="daily">Download</button></div><div class="download-row"><div class="file-icon">▧</div><div><div class="download-name">SPKL ${esc(periodLabel(period))}.xlsx</div><div class="download-meta">Sheet Hari Kerja + WEEKEND</div></div><button class="btn btn-secondary btn-sm" data-download="spkl">Download</button></div></div>
    </div>`;
  }

  function renderResultEmployees(employees) {
    const q = state.resultSearch.trim().toLowerCase();
    const recap = resultEmployeeRecap(employees).filter(({ employee }) => {
      if (!q) return true;
      return String(employee.name || '').toLowerCase().includes(q)
        || String(employee.nip || '').toLowerCase().includes(q)
        || String(employee.satkerCode || '').toLowerCase().includes(q);
    });
    return `<div class="result-panel">
      <div class="result-toolbar"><div class="search"><input id="result-search" placeholder="Cari nama, NIP, atau kode satker..." value="${esc(state.resultSearch)}"></div><span class="result-count">${recap.length} pegawai lembur</span></div>
      <div class="card table-wrap result-table-wrap"><table class="data-table result-table"><thead><tr><th>No.</th><th>Pegawai</th><th>NIP</th><th>Kode Satker</th><th>Total Jam</th><th>Hari Lembur</th><th>Uang Makan</th><th>Aksi</th></tr></thead><tbody>${recap.length ? recap.map((item, index) => `<tr><td>${index + 1}</td><td><strong>${esc(item.employee.name)}</strong></td><td class="nowrap">${esc(item.employee.nip || '-')}</td><td>${esc(item.employee.satkerCode || '-')}</td><td><strong>${item.totalHours} jam</strong></td><td>${item.overtimeDays} hari</td><td>${item.mealDays} hari</td><td><button class="btn btn-secondary btn-sm" type="button" data-result-employee="${esc(rules.employeeKey(item.employee))}">Detail</button></td></tr>`).join('') : '<tr><td colspan="8" class="text-center result-empty">Tidak ada pegawai yang sesuai pencarian.</td></tr>'}</tbody></table></div>
      <div class="footer-note">Rekap ini bersifat read-only. Gunakan tombol Edit Data apabila diperlukan koreksi terhadap hasil perhitungan.</div>
    </div>`;
  }

  function renderResultDetails(employees, holidays) {
    const allRows = resultDetailRows(employees, holidays);
    const dateOptions = [...new Map(allRows.map((item) => [item.key, item.record.date])).entries()].sort((a, b) => a[0].localeCompare(b[0]));
    const q = state.resultSearch.trim().toLowerCase();
    const rows = allRows.filter((item) => {
      if (q && !String(item.employee.name || '').toLowerCase().includes(q) && !String(item.employee.nip || '').toLowerCase().includes(q) && !String(item.employee.satkerCode || '').toLowerCase().includes(q)) return false;
      if (state.resultCategory !== 'all' && item.category.key !== state.resultCategory) return false;
      if (state.resultDate !== 'all' && item.key !== state.resultDate) return false;
      return true;
    });
    return `<div class="result-panel">
      <div class="result-toolbar result-toolbar-multi">
        <div class="search"><input id="result-search" placeholder="Cari pegawai, NIP, atau kode satker..." value="${esc(state.resultSearch)}"></div>
        <select id="result-category" class="result-filter"><option value="all" ${state.resultCategory === 'all' ? 'selected' : ''}>Semua Kategori</option><option value="workday" ${state.resultCategory === 'workday' ? 'selected' : ''}>Hari Kerja</option><option value="weekend" ${state.resultCategory === 'weekend' ? 'selected' : ''}>Weekend</option><option value="holiday" ${state.resultCategory === 'holiday' ? 'selected' : ''}>Tanggal Merah</option></select>
        <select id="result-date" class="result-filter"><option value="all">Semua Tanggal</option>${dateOptions.map(([key, date]) => `<option value="${esc(key)}" ${state.resultDate === key ? 'selected' : ''}>${esc(rules.formatIndonesianDate(date, false))}</option>`).join('')}</select>
        <span class="result-count">${rows.length} kejadian</span>
      </div>
      <div class="card table-wrap result-table-wrap"><table class="data-table result-table result-detail-table"><thead><tr><th>No.</th><th>Pegawai</th><th>Tanggal</th><th>Kategori</th><th>Jam Masuk</th><th>Mulai Lembur</th><th>Jam Pulang</th><th>Jam Lembur</th><th>Uang Makan</th></tr></thead><tbody>${rows.length ? rows.map((item, index) => `<tr><td>${index + 1}</td><td><button class="employee-link" type="button" data-result-employee="${esc(rules.employeeKey(item.employee))}">${esc(item.employee.name)}</button><div class="result-row-meta">${esc(item.employee.nip || '-')} · ${esc(item.employee.satkerCode || '-')}</div></td><td class="nowrap">${esc(rules.formatIndonesianDate(item.record.date, false))}</td><td><span class="category-pill ${item.category.rest ? 'rest-day' : ''}">${esc(item.category.label)}</span></td><td>${rules.formatMinutes(item.record.inMinutes)}</td><td>${rules.formatMinutes(item.record.overtimeStartMinutes)}</td><td>${rules.formatMinutes(item.record.outMinutes)}</td><td><strong>${Number(item.record.overtimeHours || 0)} jam</strong></td><td><span class="status ${item.meal ? 'ok' : ''}">${item.meal ? 'Ya' : 'Tidak'}</span></td></tr>`).join('') : '<tr><td colspan="9" class="text-center result-empty">Tidak ada rincian lembur yang sesuai filter.</td></tr>'}</tbody></table></div>
      <div class="footer-note">Rincian ini merupakan hasil final dari data yang sudah direview. Perubahan hanya dilakukan melalui Edit Data lalu Generate Ulang.</div>
    </div>`;
  }

  function renderResult() {
    const run = state.currentRun;
    const period = run?.period || state.period;
    const employees = run?.employees || state.employees || [];
    const summary = run?.summary || rules.summarize(employees);
    const holidays = rules.normalizeHolidays(run?.holidays || state.holidays);
    const title = state.resultMode === 'updated' ? 'Perubahan Berhasil Disimpan' : state.resultMode === 'history' ? 'Hasil Proses Tersimpan' : 'Dokumen Berhasil Dibuat';
    const timeValue = run?.updatedAt || run?.processedAt;
    const timeLabel = run?.updatedAt ? 'Diperbarui' : 'Diproses';
    const canManageRun = !run?.id || storage.canManage?.(run);
    const panel = state.resultTab === 'employees'
      ? renderResultEmployees(employees)
      : state.resultTab === 'details'
        ? renderResultDetails(employees, holidays)
        : renderResultSummary(run, period, summary, holidays);
    return `<div class="card result-head"><div class="result-head-main"><div class="success-mark">✓</div><div><h3>${title}</h3><p>${esc(periodLabel(period))}${timeValue ? ` - ${timeLabel} ${esc(formatDateTime(timeValue))}` : ''}</p></div></div><div class="result-head-note">Hasil generate dapat diperiksa langsung pada aplikasi sebelum dokumen diunduh.</div></div>
      ${renderResultTabs()}
      ${panel}
      <div class="result-actions actions"><button class="btn btn-secondary" data-action="go-history">Lihat Riwayat</button><div class="actions-right">${run?.id && canManageRun ? '<button class="btn btn-secondary" data-action="edit-current-run">Edit Data</button>' : ''}<button class="btn btn-primary" data-action="new-period">Proses Periode Baru</button></div></div>`;
  }

  function renderHistory() {
    const history = state.history;
    const rows = history.map((item) => {
      const changedAt = item.updatedAt || item.processedAt;
      const changedLabel = item.updatedAt ? 'Diubah' : 'Diproses';
      const manageable = Boolean(storage.canManage?.(item));
      const actorName = item.ownerName || item.ownerEmail || 'Akun lama';
      const actorSub = item.ownerName && item.ownerEmail ? `<div class="card-subtitle">${esc(item.ownerEmail)}</div>` : '';
      const manageActions = manageable
        ? `<button class="btn btn-secondary btn-sm" data-edit-history-id="${esc(item.id)}">Verifikasi/Edit</button><button class="btn btn-danger btn-sm" data-delete-history-id="${esc(item.id)}">Hapus</button>`
        : '';
      return `<tr><td>${esc(periodLabel(item.period))}</td><td>${item.summary.employees}</td><td>${item.summary.totalHours} jam</td><td>${(item.holidays || []).length}</td><td><strong>${esc(actorName)}</strong>${actorSub}</td><td><span class="history-time-label">${changedLabel}</span><br>${esc(formatDateTime(changedAt))}</td><td><span class="status-pill">Selesai</span></td><td><div class="history-actions"><button class="btn btn-secondary btn-sm" data-history-id="${esc(item.id)}">Lihat</button>${manageActions}</div></td></tr>`;
    }).join('');
    const content = `<div class="page-title"><div><h2>Riwayat Lembur</h2><p>Riwayat bersama seluruh akun aplikasi. Proses yang dibuat oleh akun lain dapat dilihat, sedangkan perubahan dan penghapusan dibatasi untuk pembuat riwayat atau administrator.</p></div></div><div class="card table-wrap"><table class="data-table history-table"><thead><tr><th>Periode</th><th>Pegawai</th><th>Total Lembur</th><th>Tanggal Merah</th><th>Diproses Oleh</th><th>Terakhir Diubah</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${rows || '<tr><td colspan="8" class="text-center" style="padding:36px;color:#6b7c93">Belum ada riwayat.</td></tr>'}</tbody></table></div><div class="footer-note">Riwayat Lembur tersinkron melalui koleksi bersama Cloud Firestore. Semua akun terverifikasi dapat melihat riwayat yang sama.</div>`;
    app.innerHTML = shell(content); bindEvents();
  }

  function renderDrawer() {
    if (state.reviewEditKey) return renderReviewEditModal();
    if (state.pendingHoliday) return renderHolidayPreviewModal();
    if (!state.drawerKey) return '';
    const employee = state.employees.find(e => rules.employeeKey(e) === state.drawerKey);
    if (!employee) return '';
    const period = state.currentRun?.period || state.period;
    const holidays = rules.normalizeHolidays(state.currentRun?.holidays || state.holidays);
    const stats = employeeStats(employee);
    const keys = monthKeys(period);
    const selected = activeResultDate(employee, period);
    const selectedRecord = employee.records?.[selected];
    const classFor = (key, record) => {
      if (rules.isHoliday(rules.dateFromKey(key), holidays)) return hasOvertime(record) ? 'lm-result-holiday' : 'lm-result-rest';
      if (rules.isWeekend(rules.dateFromKey(key))) return hasOvertime(record) ? 'lm-result-weekend' : 'lm-result-rest';
      return hasOvertime(record) ? 'lm-overtime' : 'lm-empty-day';
    };
    const calendar = monthCalendarMarkup(period, (key) => {
      const record = employee.records?.[key];
      const weekendLike = rules.isWeekendLike(rules.dateFromKey(key), holidays);
      const status = resultCategory({ date: rules.dateFromKey(key) }, holidays);
      const show = state.detailFilter === 'all' || state.detailFilter === 'overtime' && hasOvertime(record)
        || state.detailFilter === 'workday' && hasOvertime(record) && !weekendLike
        || state.detailFilter === 'rest' && hasOvertime(record) && weekendLike;
      return `<button type="button" class="lm-cal-day ${classFor(key,record)} ${selected === key ? 'selected' : ''} ${!show ? 'lm-dim' : ''}" data-result-day="${key}" aria-pressed="${selected === key}" aria-label="${esc(fullDayLabel(key))}, ${record ? `${finalHours(record)} jam lembur` : 'tidak ada catatan lembur'}">
        <span class="lm-date">${Number(key.slice(-2))}</span><span class="lm-dayname">${shortWeekday(key)}</span><strong>${hasOvertime(record) ? `${finalHours(record)} jam` : '–'}</strong>
        <small>${hasOvertime(record) ? status.label : weekendLike ? 'Libur' : 'Tanpa lembur'}</small></button>`;
    });
    const cat = resultCategory({ date: rules.dateFromKey(selected) }, holidays);
    const rows = keys.map(key => ({ key, record: employee.records?.[key] })).filter(x => hasOvertime(x.record)).filter(x => {
      const rest = rules.isWeekendLike(rules.dateFromKey(x.key), holidays);
      return state.detailFilter === 'all' || state.detailFilter === 'overtime' || state.detailFilter === 'workday' && !rest || state.detailFilter === 'rest' && rest;
    });
    return `<div class="lm-overlay lm-result-overlay" data-action="close-drawer-backdrop">
      <section class="lm-modal" role="dialog" aria-modal="true" aria-labelledby="lm-detail-heading">
        <header class="lm-modal-header"><div class="lm-modal-headline"><div><h2 id="lm-detail-heading">Detail Rekap Lembur Pegawai</h2><p>Hasil perhitungan lembur dan uang makan lembur · ${esc(periodLabel(period))}</p></div>
          <button type="button" class="icon-btn" aria-label="Tutup detail pegawai" data-action="close-drawer">×</button></div>
          <div class="lm-modal-person"><span class="lm-avatar">${esc((employee.name || 'P').trim().charAt(0).toUpperCase())}</span><div><strong>${esc(employee.name)}</strong><span>NIP ${esc(employee.nip || '-')} · Anak Satker ${esc(employee.satkerCode || '-')}</span></div><span class="lm-readonly-tag">Baca-saja</span></div>
        </header><div class="lm-modal-body">
          <div class="lm-result-metrics"><div><span>Total Jam Lembur</span><strong>${stats.hours} <small>jam</small></strong></div><div><span>Hari Lembur</span><strong>${stats.days} <small>hari</small></strong></div><div class="lm-highlight"><span>Uang Makan Lembur</span><strong>${stats.meals} <small>hari</small></strong></div><div><span>Hari Tanpa Lembur</span><strong>${keys.length - stats.days} <small>hari</small></strong></div></div>
          <div class="lm-review-head lm-result-mode-head"><div><div class="card-title">Rincian Kehadiran Lembur</div><div class="card-subtitle">Pilih tanggal untuk melihat jam masuk, mulai lembur, pulang dan perhitungan.</div></div>
            <div class="lm-segment" role="group" aria-label="Tampilan detail lembur"><button type="button" data-detail-mode="calendar" class="${state.detailMode === 'calendar' ? 'active' : ''}">Kalender</button><button type="button" data-detail-mode="table" class="${state.detailMode === 'table' ? 'active' : ''}">Tabel</button></div></div>
          <div class="lm-detail-controls"><select class="lm-select" id="lm-detail-filter" aria-label="Filter detail lembur"><option value="all" ${state.detailFilter === 'all' ? 'selected' : ''}>Semua Tanggal</option><option value="overtime" ${state.detailFilter === 'overtime' ? 'selected' : ''}>Ada Lembur</option><option value="workday" ${state.detailFilter === 'workday' ? 'selected' : ''}>Lembur Hari Kerja</option><option value="rest" ${state.detailFilter === 'rest' ? 'selected' : ''}>Lembur Hari Libur</option></select></div>
          ${state.detailMode === 'calendar' ? `<div class="lm-result-calendar">${calendar}
            <div class="lm-legend"><span><i class="lm-dot lm-dot-green"></i>Lembur hari kerja</span><span><i class="lm-dot lm-dot-gold"></i>Lembur akhir pekan / tanggal merah</span><span><i class="lm-dot lm-dot-gray"></i>Tanpa lembur</span></div>
            <section class="lm-day-panel lm-readonly-panel" id="lm-detail-day-panel" aria-live="polite"><div class="lm-day-panel-head"><strong>${esc(fullDayLabel(selected))}</strong><span class="category-pill ${cat.rest ? 'rest-day' : ''}">${esc(cat.label)}</span></div>
              ${selectedRecord && finalHours(selectedRecord) > 0 ? `<div class="lm-day-grid"><div><span>Jam Masuk</span><strong>${esc(rules.formatMinutes(selectedRecord.inMinutes))}</strong></div><div><span>Mulai Lembur</span><strong>${esc(rules.formatMinutes(selectedRecord.overtimeStartMinutes))}</strong></div><div><span>Jam Pulang</span><strong>${esc(rules.formatMinutes(selectedRecord.outMinutes))}</strong></div><div><span>Total Lembur</span><strong>${finalHours(selectedRecord)} jam</strong></div><div><span>Uang Makan Lembur</span><strong>${mealEligible(selectedRecord) ? 'Ya' : 'Tidak'}</strong></div><div><span>Status Presensi</span><strong>${esc(selectedRecord.status || '-')}</strong></div></div>`
                : '<div class="lm-empty lm-inline">Tidak ada lembur tercatat pada tanggal ini.</div>'}</section></div>`
            : `<div class="lm-table-count">${rows.length} hari lembur sesuai filter.</div><div class="table-wrap lm-result-table-scroll"><table class="data-table lm-rows"><thead><tr><th>Tanggal</th><th>Kategori</th><th>Masuk</th><th>Mulai Lembur</th><th>Pulang</th><th>Durasi</th><th>Uang Makan</th><th>Aksi</th></tr></thead><tbody>${rows.length ? rows.map(({key,record}) => { const cat = resultCategory(record,holidays); return `<tr><td class="nowrap">${esc(fullDayLabel(key))}</td><td><span class="category-pill ${cat.rest ? 'rest-day' : ''}">${esc(cat.label)}</span></td><td>${esc(rules.formatMinutes(record.inMinutes))}</td><td>${esc(rules.formatMinutes(record.overtimeStartMinutes))}</td><td>${esc(rules.formatMinutes(record.outMinutes))}</td><td>${finalHours(record)} jam</td><td>${mealEligible(record) ? 'Ya' : 'Tidak'}</td><td><button type="button" class="btn btn-secondary btn-sm" data-result-day="${key}">Lihat</button></td></tr>`; }).join('') : '<tr><td colspan="8" class="lm-empty">Tidak ada lembur sesuai filter.</td></tr>'}</tbody></table></div>`}
        </div><footer class="lm-modal-footer"><span>Data rekap bersifat baca-saja. Koreksi dilakukan melalui Edit Data Lembur.</span><button type="button" class="btn btn-secondary" data-action="close-drawer">Selesai</button></footer>
      </section></div>`;
  }

  function renderStay() {
    const x = window.scrollX || 0, y = window.scrollY || 0;
    const overlay = document.querySelector('.lm-modal-body');
    const overlayY = overlay?.scrollTop ?? null;
    render();
    window.scrollTo(x, y);
    if (overlayY !== null) document.querySelector('.lm-modal-body')?.scrollTo(0, overlayY);
  }

  function render() {
    if (!state.started) return renderWelcome();
    if (state.view === 'dashboard') return renderDashboard();
    if (state.view === 'history') return renderHistory();
    return renderProcess();
  }

  function resetProcess() {
    state.view = 'process';
    state.processStep = 'period';
    state.files = [];
    state.validation = null;
    state.employees = [];
    state.filter = '';
    state.reviewMode = 'calendar';
    state.reviewGroup = 'employee';
    state.reviewEmployeeKey = null;
    state.reviewDateKey = null;
    state.reviewFilter = 'all';
    state.reviewTableFilter = 'overtime';
    state.reviewEditKey = null;
    state.pendingHoliday = null;
    state.dirtyKeys.clear();
    state.detailMode = 'calendar';
    state.detailFilter = 'all';
    state.detailDateKey = null;
    state.generated = null;
    state.currentRun = null;
    state.resultFromHistory = false;
    state.holidays = [];
    state.editingHistoryId = null;
    state.resultMode = 'new';
    state.resultTab = 'summary';
    state.resultSearch = '';
    state.resultCategory = 'all';
    state.resultDate = 'all';
    state.drawerKey = null;
  }

  function addFiles(fileList) {
    const incoming = Array.from(fileList || []).filter(f=>/\.(xlsx|xls)$/i.test(f.name));
    const existing = new Map(state.files.map(f=>[`${f.name}|${f.size}|${f.lastModified}`,f]));
    incoming.forEach(f=>existing.set(`${f.name}|${f.size}|${f.lastModified}`,f));
    state.files = [...existing.values()]; state.validation = null;
    if (state.files.length) {
      resources?.warm?.('xlsx');
      window.MasterDataService?.getEmployees?.(false).catch(() => {});
    }
    render();
  }

  async function validateFiles() {
    if (!state.files.length || state.busy) return;
    state.busy = true;
    try {
      await resources?.ensure?.('xlsx');
      const result = await parser.parseFiles(state.files, state.period, state.holidays);
      state.validation = result; state.processStep = 'validation';
    } catch (error) {
      alert(error.message || String(error));
    } finally { state.busy = false; render(); }
  }

  async function processData() {
    if (!state.validation || state.validation.errors.length) return;
    state.processStep = 'processing'; state.processing = {percent:8,active:0}; render();
    const stages = [20,38,58,78,92];
    for (let i=0;i<stages.length;i+=1) {
      state.processing = {percent:stages[i],active:i}; render(); await new Promise(r=>setTimeout(r,180));
      if (i===1) await rules.assignSatkerCodes(state.validation.employees);
    }
    try {
      // Pemeriksaan wajib sebelum masuk Review; jangan menghasilkan klaim jika sumber tidak terbaca.
      await window.LemburDocumentIntegration.apply(state.validation.employees,state.period);
    }catch(error){
      state.processStep='validation';
      alert(`Gagal memverifikasi Cuti/Surat Tugas: ${error.message||error}`);
      render();return;
    }
    state.employees = state.validation.employees;
    state.reviewEmployeeKey = null;
    state.reviewDateKey = null;
    state.dirtyKeys.clear();
    resources?.warm?.(['exceljs', 'filesaver']);
    state.processing = {percent:100,active:5}; render(); await new Promise(r=>setTimeout(r,220));
    state.processStep = 'review'; render();
  }

  // Mode tabel juga menggunakan modal yang sama agar koreksi selalu tercatat.
  function updateOvertime(key) { openReviewEdit(key); }

  async function generateDocs() {
    if (!window.UserAccess.assert('lembur',state.editingHistoryId?'history.edit':'process.generate')) return;
    if (state.busy) return;
    state.busy = true; render();
    try {
      await window.LemburDocumentIntegration.apply(state.employees,state.period);
      if(pendingDocumentReviews())throw new Error(`${pendingDocumentReviews()} tanggal Cuti/Surat Tugas belum diverifikasi. Kembali ke Review dan selesaikan dahulu.`);
      await resources?.ensure?.('exceljs');
      const holidays = rules.normalizeHolidays(state.holidays);
      state.generated = await generator.generateAll(state.employees, state.period, holidays);
      const now = new Date().toISOString();
      const existing = state.editingHistoryId ? (state.currentRun || await storage.getRun(state.editingHistoryId)) : null;
      const run = state.editingHistoryId
        ? {
            id: state.editingHistoryId,
            period: {...state.period},
            processedAt: existing?.processedAt || now,
            updatedAt: now,
            summary: rules.summarize(state.employees),
            employees: state.employees,
            holidays
          }
        : {
            id: `${state.period.year}-${rules.pad2(state.period.month)}-${Date.now()}`,
            period: {...state.period},
            processedAt: now,
            updatedAt: null,
            summary: rules.summarize(state.employees),
            employees: state.employees,
            holidays
          };
      const savedMeta = await storage.saveRun(run);
      state.currentRun = { ...run, ...savedMeta };
      state.resultFromHistory = Boolean(state.editingHistoryId);
      state.resultMode = state.editingHistoryId ? 'updated' : 'new';
      state.resultTab = 'summary';
      state.resultSearch = '';
      state.resultCategory = 'all';
      state.resultDate = 'all';
      state.editingHistoryId = null;
      state.reviewEditKey = null;
      state.pendingHoliday = null;
      state.dirtyKeys.clear();
      state.processStep = 'result';
    } catch (error) {
      alert(`Gagal membuat dokumen: ${error.message || error}`);
    } finally { state.busy=false; render(); }
  }

  async function download(kind) {
    if (!window.UserAccess.assert('lembur','history.download')) return;
    const run = state.currentRun || {period:state.period,employees:state.employees,holidays:state.holidays};
    if (!run.employees?.length) return;
    try {
      await resources?.ensure?.(['exceljs', 'filesaver']);
      let file = state.generated && state.generated[kind];
      const holidays = rules.normalizeHolidays(run.holidays || state.holidays);
      if (!file) {
        if (kind==='recap') file = await generator.generateRecap(run.employees,run.period,holidays);
        if (kind==='daily') file = await generator.generateDaily(run.employees,run.period,holidays);
        if (kind==='spkl') file = await generator.generateSpkl(run.employees,run.period,holidays);
      }
      generator.downloadFile(file);
    } catch (error) { alert(`Gagal menyiapkan file: ${error.message || error}`); }
  }

  async function openHistory(id) {
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan atau sudah dihapus.');
      state.view='process';
      state.processStep='result';
      state.period={...run.period};
      state.employees=run.employees;
      state.holidays=rules.normalizeHolidays(run.holidays || []);
      state.currentRun=run;
      state.generated=null;
      state.resultFromHistory=true;
      state.editingHistoryId=null;
      state.resultMode='history';
      state.resultTab='summary';
      state.resultSearch='';
      state.resultCategory='all';
      state.resultDate='all';
      state.filter='';
      state.drawerKey=null;
      state.detailDateKey=null;
      state.detailMode='calendar';
      state.detailFilter='all';
      render();
    } catch (error) {
      alert(error.message || String(error));
    }
  }

  async function editHistory(id) {
    if (!window.UserAccess.assert('lembur','history.edit')) return;
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan atau sudah dihapus.');
      if (!storage.canManage?.(run)) throw new Error('Riwayat ini dibuat oleh akun lain. Anda dapat melihat hasilnya, tetapi hanya pembuat atau administrator yang dapat mengedit.');
      state.view='process';
      state.processStep='review';
      state.period={...run.period};
      state.employees=run.employees;
      state.holidays=rules.normalizeHolidays(run.holidays || []);
      state.currentRun=run;
      state.generated=null;
      state.resultFromHistory=true;
      state.editingHistoryId=run.id;
      state.resultMode='history';
      state.resultTab='summary';
      state.resultSearch='';
      state.resultCategory='all';
      state.resultDate='all';
      state.filter='';
      state.drawerKey=null;
      state.reviewMode='calendar';
      state.reviewGroup='employee';
      state.reviewEmployeeKey=null;
      state.reviewDateKey=null;
      state.reviewFilter='all';
      state.reviewTableFilter='overtime';
      state.reviewEditKey=null;
      state.pendingHoliday=null;
      state.dirtyKeys.clear();
      await window.LemburDocumentIntegration.apply(state.employees,state.period);
      render();
    } catch (error) {
      alert(error.message || String(error));
    }
  }

  function cancelHistoryEdit() {
    state.editingHistoryId = null;
    state.currentRun = null;
    state.generated = null;
    state.employees = [];
    state.holidays = [];
    state.filter = '';
    state.drawerKey = null;
    state.reviewEditKey = null;
    state.pendingHoliday = null;
    state.dirtyKeys.clear();
    state.view = 'history';
    render();
  }

  async function deleteHistory(id) {
    if (!window.UserAccess.assert('lembur','history.delete')) return;
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan atau sudah dihapus.');
      if (!storage.canManage?.(run)) throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat dihapus oleh pembuat atau administrator.');
      const label = periodLabel(run.period);
      const approved = window.confirm(`Hapus riwayat proses ${label}?\n\nData hasil proses periode ini akan dihapus dari Firestore dan tindakan ini tidak dapat dibatalkan.`);
      if (!approved) return;
      await storage.deleteRun(id);
      state.history = state.history.filter((item) => item.id !== id);
      if (state.currentRun?.id === id) {
        state.currentRun = null;
        state.generated = null;
        state.resultFromHistory = false;
        state.editingHistoryId = null;
        state.holidays = [];
      }
      render();
    } catch (error) {
      alert(`Riwayat gagal dihapus: ${error.message || error}`);
    }
  }

  function bindEvents() {
    document.querySelectorAll('[data-nav]').forEach(btn=>btn.addEventListener('click',(event)=>{ event.preventDefault(); const target=btn.dataset.nav; if(target==='process') resetProcess(); else { state.view=target; state.editingHistoryId=null; } history.replaceState(null,'',`lembur.html#${target}`); render(); }));
    document.querySelector('[data-action="enter-app"]')?.addEventListener('click',()=>{ state.started=true; resetProcess(); history.replaceState(null,'','lembur.html#process'); render(); });
    document.querySelector('[data-action="start-process"]')?.addEventListener('click',()=>{ resetProcess(); history.replaceState(null,'','lembur.html#process'); render(); });
    document.querySelector('[data-action="go-history"]')?.addEventListener('click',()=>{ state.view='history'; state.editingHistoryId=null; history.replaceState(null,'','lembur.html#history'); render(); });

    document.getElementById('period-month')?.addEventListener('change',(e)=>{ state.period.month=Number(e.target.value); state.holidays=state.holidays.filter((key)=>dateBelongsToPeriod(key,state.period)); state.validation=null; render(); });
    document.getElementById('period-year')?.addEventListener('change',(e)=>{ state.period.year=Number(e.target.value); state.holidays=state.holidays.filter((key)=>dateBelongsToPeriod(key,state.period)); state.validation=null; render(); });
    document.querySelector('[data-action="period-next"]')?.addEventListener('click',()=>{ state.period.month=Number(document.getElementById('period-month').value); state.period.year=Number(document.getElementById('period-year').value); state.holidays=state.holidays.filter((key)=>dateBelongsToPeriod(key,state.period)); state.validation=null; state.processStep='upload'; render(); });
    document.querySelector('[data-action="add-holiday"]')?.addEventListener('click',addHoliday);
    document.querySelectorAll('[data-remove-holiday]').forEach(btn=>btn.addEventListener('click',()=>removeHoliday(btn.dataset.removeHoliday)));

    document.querySelector('[data-action="back-period"]')?.addEventListener('click',()=>{ state.processStep='period'; render(); });
    document.querySelectorAll('[data-action="choose-files"]').forEach(btn=>btn.addEventListener('click',()=>document.getElementById('file-input')?.click()));
    document.getElementById('file-input')?.addEventListener('change',e=>addFiles(e.target.files));
    const dz=document.getElementById('dropzone'); if(dz){ ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('dragover')})); ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('dragover')})); dz.addEventListener('drop',e=>addFiles(e.dataTransfer.files)); }
    document.querySelectorAll('[data-remove-file]').forEach(btn=>btn.addEventListener('click',()=>{ state.files.splice(Number(btn.dataset.removeFile),1); state.validation=null; render(); }));
    document.querySelector('[data-action="validate"]')?.addEventListener('click',validateFiles);
    document.querySelector('[data-action="back-upload"]')?.addEventListener('click',()=>{state.processStep='upload';render();});
    document.querySelector('[data-action="process-data"]')?.addEventListener('click',processData);
    document.querySelector('[data-action="back-validation"]')?.addEventListener('click',()=>{state.processStep='validation';render();});
    document.querySelector('[data-action="to-confirm"]')?.addEventListener('click',()=>{if(pendingDocumentReviews())return alert('Selesaikan verifikasi seluruh Cuti/Surat Tugas per periode terlebih dahulu.');state.processStep='confirm';render();});
    document.querySelector('[data-action="back-review"]')?.addEventListener('click',()=>{state.processStep='review';render();});
    document.querySelector('[data-action="generate"]')?.addEventListener('click',generateDocs);
    document.querySelector('[data-action="save-history-edit"]')?.addEventListener('click',generateDocs);
    document.querySelector('[data-action="cancel-history-edit"]')?.addEventListener('click',cancelHistoryEdit);
    document.querySelector('[data-action="go-dashboard"]')?.addEventListener('click',()=>{ state.started=true; state.view='dashboard'; state.editingHistoryId=null; history.replaceState(null,'','lembur.html#dashboard'); render(); });
    document.querySelector('[data-action="new-period"]')?.addEventListener('click',()=>{resetProcess(); history.replaceState(null,'','lembur.html#process'); render();});
    document.querySelector('[data-action="edit-current-run"]')?.addEventListener('click',()=>{ if(state.currentRun?.id) editHistory(state.currentRun.id); });
    document.querySelectorAll('[data-history-id]').forEach(btn=>btn.addEventListener('click',()=>openHistory(btn.dataset.historyId)));
    document.querySelectorAll('[data-edit-history-id]').forEach(btn=>btn.addEventListener('click',()=>editHistory(btn.dataset.editHistoryId)));
    document.querySelectorAll('[data-delete-history-id]').forEach(btn=>btn.addEventListener('click',()=>deleteHistory(btn.dataset.deleteHistoryId)));
    document.querySelectorAll('[data-download]').forEach(btn=>btn.addEventListener('click',()=>download(btn.dataset.download)));
    document.querySelectorAll('[data-result-tab]').forEach((btn) => btn.addEventListener('click', () => { state.resultTab = btn.dataset.resultTab || 'summary'; state.resultSearch = ''; state.resultCategory = 'all'; state.resultDate = 'all'; state.drawerKey = null; state.detailDateKey = null; render(); }));
    document.querySelectorAll('[data-result-employee]').forEach(btn => btn.addEventListener('click', () => {
      state.drawerKey = btn.dataset.resultEmployee;
      state.detailMode = 'calendar';
      state.detailFilter = 'all';
      state.detailDateKey = null;
      render();
    }));
    document.getElementById('result-search')?.addEventListener('input', (event) => { state.resultSearch = event.target.value; clearTimeout(window.__resultSearchTimer); window.__resultSearchTimer = setTimeout(render, 180); });
    document.getElementById('result-category')?.addEventListener('change', (event) => { state.resultCategory = event.target.value || 'all'; render(); });
    document.getElementById('result-date')?.addEventListener('change', (event) => { state.resultDate = event.target.value || 'all'; render(); });
    // Kalender edit dan tabel menggunakan sumber data yang sama. Hanya satu record diubah tiap modal.
    document.querySelectorAll('[data-review-mode]').forEach(btn => btn.addEventListener('click', () => {
      state.reviewMode = btn.dataset.reviewMode;
      renderStay();
    }));
    document.querySelectorAll('[data-review-group]').forEach(btn => btn.addEventListener('click', () => {
      state.reviewGroup = btn.dataset.reviewGroup;
      state.reviewDateKey = null;
      renderStay();
    }));
    document.querySelectorAll('[data-review-person]').forEach(btn => btn.addEventListener('click', () => {
      state.reviewEmployeeKey = btn.dataset.reviewPerson;
      state.reviewMode = 'calendar';
      state.reviewGroup = 'employee';
      state.reviewDateKey = null;
      renderStay();
    }));
    document.querySelectorAll('[data-review-day]').forEach(btn => btn.addEventListener('click', () => {
      state.reviewDateKey = btn.dataset.reviewDay;
      renderStay();
    }));
    document.getElementById('lm-review-filter')?.addEventListener('change', e => {
      state.reviewFilter = e.target.value;
      renderStay();
    });
    document.getElementById('review-table-filter')?.addEventListener('change', e => {
      state.reviewTableFilter = e.target.value;
      renderStay();
    });
    const bindReviewSearch = (id) => {
      document.getElementById(id)?.addEventListener('input', e => {
        state.filter = e.target.value;
        const caret = e.target.selectionStart;
        clearTimeout(window.__reviewTimer);
        window.__reviewTimer = setTimeout(() => {
          renderStay();
          const next = document.getElementById(id);
          if (next) { next.focus(); if (Number.isInteger(caret)) next.setSelectionRange(caret,caret); }
        }, 220);
      });
    };
    bindReviewSearch('review-search');
    bindReviewSearch('lm-person-search');
    document.querySelectorAll('[data-lm-document-save]').forEach(btn=>btn.addEventListener('click',()=>saveDocumentPeriodReview(btn.dataset.lmDocumentSave)));
    document.querySelectorAll('[data-action="open-period-document-review"]').forEach(btn=>btn.addEventListener('click',scrollToDocumentPeriodReview));
    document.querySelectorAll('[data-review-edit]').forEach(btn => btn.addEventListener('click', () => openReviewEdit(btn.dataset.reviewEdit)));
    document.querySelectorAll('[data-action="save-review-edit"]').forEach(btn => btn.addEventListener('click', saveReviewEdit));
    document.querySelectorAll('[data-action="close-review-edit"]').forEach(btn => btn.addEventListener('click', () => { state.reviewEditKey = null; renderStay(); }));
    document.querySelectorAll('[data-action="close-review-edit-backdrop"]').forEach(backdrop => backdrop.addEventListener('click', e => { if (e.target === backdrop) {state.reviewEditKey = null; renderStay();} }));
    document.querySelectorAll('[data-action="close-holiday-preview"]').forEach(btn => btn.addEventListener('click', () => {state.pendingHoliday=null;renderStay();}));
    document.querySelectorAll('[data-action="close-holiday-preview-backdrop"]').forEach(backdrop => backdrop.addEventListener('click', e => {if(e.target===backdrop){state.pendingHoliday=null;renderStay();}}));
    document.querySelectorAll('[data-action="confirm-holiday-preview"]').forEach(btn => btn.addEventListener('click', confirmHolidayChange));
    document.querySelectorAll('[data-detail-mode]').forEach(btn => btn.addEventListener('click', () => {state.detailMode=btn.dataset.detailMode;renderStay();}));
    document.querySelectorAll('[data-result-day]').forEach(btn => btn.addEventListener('click', () => {
      state.detailDateKey = btn.dataset.resultDay;
      state.detailMode = 'calendar';
      renderStay();
      if (btn.closest('.lm-result-table-scroll')) document.getElementById('lm-detail-day-panel')?.scrollIntoView({block:'nearest'});
    }));
    document.getElementById('lm-detail-filter')?.addEventListener('change', e => {state.detailFilter=e.target.value;renderStay();});
    document.querySelectorAll('[data-action="close-drawer"]').forEach(btn => btn.addEventListener('click', () => {state.drawerKey=null;render();}));
    document.querySelectorAll('[data-action="close-drawer-backdrop"]').forEach(backdrop => backdrop.addEventListener('click', e => {if(e.target===backdrop){state.drawerKey=null;render();}}));
  }

  function applyHashRoute() {
    const route = String(location.hash || '').replace(/^#/, '').toLowerCase();
    state.started = true;
    if (route === 'history') { state.view = 'history'; state.editingHistoryId = null; }
    else if (route === 'process') { resetProcess(); }
    else { state.view = 'dashboard'; state.editingHistoryId = null; if (route !== 'dashboard') history.replaceState(null, '', 'lembur.html#dashboard'); }
  }

  window.addEventListener('hashchange', () => {
    const route = String(location.hash || '').replace(/^#/, '').toLowerCase();
    state.started = true;
    if (route === 'history') { state.view = 'history'; state.editingHistoryId = null; render(); }
    else if (route === 'process') { resetProcess(); render(); }
    else { state.view = 'dashboard'; state.editingHistoryId = null; render(); }
  });

  async function init() {
    if (!(await window.UserAccess.requireModule('lembur'))) return;
    await window.AppSettingsService?.loadAndApply?.();
    applyHashRoute();
    if (!window.UserAccess.canView('lembur',state.view)) { state.view=window.UserAccess.firstView('lembur'); history.replaceState(null,'',`lembur.html#${state.view}`); }
    if (window.UserAccess.can('lembur','process.generate')) await storage.migrateLegacyOnce?.();
    state.historyBusy = state.view === 'history' || state.view === 'dashboard';
    historyUnsubscribe = storage.subscribeHistory?.((history) => {
      state.history = history;
      state.historyBusy = false;
      if (state.view === 'history' || state.view === 'dashboard') render();
    }, async (error) => {
      console.warn('Sinkronisasi realtime Lembur gagal:', error);
      await refreshHistory(false);
      if (state.view === 'history' || state.view === 'dashboard') render();
    }) || null;
    render();
  }

  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    if (state.reviewEditKey) {state.reviewEditKey=null;renderStay();}
    else if (state.pendingHoliday) {state.pendingHoliday=null;renderStay();}
    else if (state.drawerKey) {state.drawerKey=null;render();}
    else return;
    event.preventDefault();
  });
  window.addEventListener('beforeunload', () => historyUnsubscribe?.());
  init();
})();
