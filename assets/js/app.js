(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;
  const parser = window.ExcelParser;
  const generator = window.DocumentGenerator;
  const storage = window.AppStorage;
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

  function recalculateDate(dateKey) {
    state.employees.forEach((employee) => {
      const record = employee.records && employee.records[dateKey];
      if (!record) return;
      const overtime = rules.calculateOvertime(record.date, record.inMinutes, record.outMinutes, record.status, state.holidays);
      record.overtimeHours = overtime.hours;
      record.originalOvertimeHours = overtime.hours;
      record.normalEndMinutes = overtime.normalEndMinutes;
      record.adjustedWorkEndMinutes = overtime.adjustedWorkEndMinutes;
      record.overtimeStartMinutes = overtime.overtimeStartMinutes;
      record.isHoliday = rules.isHoliday(record.date, state.holidays);
    });
    state.generated = null;
  }

  function addHoliday() {
    const input = document.getElementById('holiday-date');
    const key = String(input?.value || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) {
      alert('Pilih tanggal merah terlebih dahulu.');
      return;
    }
    if (!dateBelongsToPeriod(key, state.period)) {
      alert(`Tanggal merah harus berada pada periode ${periodLabel(state.period)}.`);
      return;
    }
    const before = state.holidays.includes(key);
    state.holidays = rules.normalizeHolidays([...state.holidays, key]);
    if (!before) recalculateDate(key);
    state.validation = state.employees.length ? state.validation : null;
    render();
  }

  function removeHoliday(key) {
    if (!state.holidays.includes(key)) return;
    state.holidays = state.holidays.filter((item) => item !== key);
    recalculateDate(key);
    state.validation = state.employees.length ? state.validation : null;
    render();
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
    const view = ['dashboard', 'history', 'process'].includes(state.view) ? state.view : 'dashboard';
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
          <img class="welcome-logo" src="assets/img/logo-pkp.png" alt="Kementerian Perumahan dan Kawasan Permukiman">
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

  function reviewRows() {
    const rows = [];
    state.employees.forEach((employee) => {
      Object.keys(employee.records || {}).sort().forEach((key) => {
        const record = employee.records[key];
        if (!record || (Number(record.originalOvertimeHours || 0) <= 0 && Number(record.overtimeHours || 0) <= 0)) return;
        rows.push({ employee, record, key });
      });
    });
    return rows.filter(({employee,record}) => {
      const q = state.filter.trim().toLowerCase();
      if (!q) return true;
      return employee.name.toLowerCase().includes(q) || rules.formatIndonesianDate(record.date,false).toLowerCase().includes(q);
    });
  }

  function renderReview() {
    const summary = rules.summarize(state.employees);
    const rows = reviewRows();
    const isEdit = Boolean(state.editingHistoryId);
    const actions = isEdit
      ? `<div class="actions"><button class="btn btn-secondary" data-action="cancel-history-edit"><- Kembali ke Riwayat</button><button class="btn btn-primary" data-action="save-history-edit" ${state.busy?'disabled':''}>${state.busy?'Menyimpan...':'Simpan & Generate Ulang'}</button></div>`
      : `<div class="actions"><button class="btn btn-secondary" data-action="back-validation"><- Kembali</button><button class="btn btn-primary" data-action="to-confirm">Konfirmasi & Lanjut -></button></div>`;
    return `<div class="page-title"><div><h2>${isEdit ? 'Edit Data Lembur' : 'Review Data Lembur'}</h2><p>${esc(periodLabel(state.period))} - Koreksi jam lembur dan tanggal merah sebelum dokumen dibuat.</p></div></div>
      ${isEdit ? '<div class="alert alert-warning"><div class="alert-title">Mode edit riwayat</div>Perubahan akan mengganti data periode tersimpan yang sama dan dokumen akan dibuat ulang berdasarkan hasil edit terbaru.</div>' : ''}
      <div class="grid-4 section-gap">${metric(summary.employees,'Pegawai')}${metric(summary.overtimeEmployees,'Pegawai Lembur')}${metric(summary.totalHours,'Total Jam Lembur')}${metric(summary.mealDays,'Hari Uang Makan')}</div>
      ${renderHolidayManager()}
      <div class="toolbar"><div class="search"><input id="review-search" placeholder="Cari pegawai atau tanggal..." value="${esc(state.filter)}"></div><span style="font-size:11px;color:#6b7c93">${rows.length} baris lembur</span></div>
      <div class="card table-wrap"><table class="data-table"><thead><tr><th>No.</th><th>Pegawai</th><th>Tanggal</th><th>Kategori</th><th>Jam Masuk</th><th>Jam Pulang</th><th>Status</th><th>Jam Lembur</th></tr></thead><tbody>${rows.length ? rows.map((item,i)=>{ const holiday=rules.isHoliday(item.record.date,state.holidays); const weekend=rules.isWeekend(item.record.date); const category=holiday ? 'Tanggal merah' : weekend ? 'Weekend' : 'Hari kerja'; return `<tr><td>${i+1}</td><td><button class="employee-link" data-employee-key="${esc(rules.employeeKey(item.employee))}">${esc(item.employee.name)}</button></td><td class="nowrap">${esc(rules.formatIndonesianDate(item.record.date,false))}</td><td><span class="category-pill ${holiday || weekend ? 'rest-day' : ''}">${category}</span></td><td>${rules.formatMinutes(item.record.inMinutes)}</td><td>${rules.formatMinutes(item.record.outMinutes)}</td><td>${esc(item.record.status || '-')}</td><td><select class="select-mini" data-overtime-key="${esc(rules.employeeKey(item.employee))}|${esc(item.key)}">${[0,1,2,3,4].map(v=>`<option value="${v}" ${Number(item.record.overtimeHours)===v?'selected':''}>${v} jam</option>`).join('')}</select></td></tr>`; }).join('') : '<tr><td colspan="8" class="text-center" style="padding:28px;color:#6b7c93">Tidak ada data lembur pada filter ini.</td></tr>'}</tbody></table></div>
      <div class="footer-note">Perubahan pada Jam Lembur menjadi nilai final. Tanggal merah menggunakan aturan perhitungan weekend dan ditempatkan pada SPKL WEEKEND.</div>
      ${actions}`;
  }

  function renderConfirm() {
    const s = rules.summarize(state.employees);
    return `<div class="page-title"><div><h2>Siap Membuat Dokumen</h2><p>Periksa ringkasan sebelum file Excel dibuat.</p></div></div><div class="card card-pad">
      <div class="two-col"><div><div class="card-title">Ringkasan Proses</div><div class="detail-grid section-gap"><div><div class="detail-label">Periode</div><div class="detail-value">${esc(periodLabel(state.period))}</div></div><div><div class="detail-label">Jumlah Pegawai</div><div class="detail-value">${s.employees}</div></div><div><div class="detail-label">Pegawai Lembur</div><div class="detail-value">${s.overtimeEmployees}</div></div><div><div class="detail-label">Total Lembur</div><div class="detail-value">${s.totalHours} jam</div></div><div><div class="detail-label">Hari Uang Makan</div><div class="detail-value">${s.mealDays} hari</div></div></div></div>
      <div><div class="card-title">Dokumen yang akan dibuat</div><ul class="confirm-list section-gap"><li><span class="check">☑</span>Rekapitulasi Lembur</li><li><span class="check">☑</span>Daftar Hadir Kerja Lembur</li><li><span class="check">☑</span>SPKL Hari Kerja</li><li><span class="check">☑</span>SPKL Weekend</li></ul><div class="card-subtitle">SPKL Hari Kerja dan WEEKEND berada dalam satu workbook Excel.</div></div></div>
      <div class="actions"><button class="btn btn-secondary" data-action="back-review">← Kembali ke Review</button><button class="btn btn-primary" data-action="generate" ${state.busy?'disabled':''}>${state.busy?'Membuat Dokumen...':'Generate Dokumen'}</button></div></div>`;
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
    if (!state.drawerKey) return '';
    const employee = state.employees.find(e=>rules.employeeKey(e)===state.drawerKey);
    if (!employee) return '';
    const holidays = rules.normalizeHolidays(state.currentRun?.holidays || state.holidays);
    const records = Object.keys(employee.records || {}).sort().map((key) => ({ key, record: employee.records[key] })).filter((item) => Number(item.record.overtimeHours || 0) > 0);
    const total = records.reduce((sum, item) => sum + Number(item.record.overtimeHours || 0), 0);
    const mealDays = records.filter((item) => Number(item.record.overtimeHours || 0) >= Number(cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS || 2)).length;
    return `<div class="drawer-backdrop" data-action="close-drawer"></div><aside class="drawer"><div class="drawer-header"><h3>Detail Pegawai</h3><button class="icon-btn" data-action="close-drawer">×</button></div><div class="profile"><div class="profile-avatar">◉</div><div><div class="profile-name">${esc(employee.name)}</div><div class="profile-meta">NIP ${esc(employee.nip || '-')}</div></div></div><div class="detail-grid"><div><div class="detail-label">Kode Satker</div><div class="detail-value">${esc(employee.satkerCode || '-')}</div></div><div><div class="detail-label">Total Lembur</div><div class="detail-value">${total} jam</div></div><div><div class="detail-label">Hari Lembur</div><div class="detail-value">${records.length} hari</div></div><div><div class="detail-label">Uang Makan</div><div class="detail-value">${mealDays} hari</div></div></div><div class="card-title">Data Lembur</div><div class="table-wrap section-gap"><table class="data-table drawer-overtime-table"><thead><tr><th>Tanggal</th><th>Kategori</th><th>Masuk</th><th>Mulai</th><th>Pulang</th><th>Lembur</th><th>Uang Makan</th></tr></thead><tbody>${records.map((item) => { const category = resultCategory(item.record, holidays); const meal = Number(item.record.overtimeHours || 0) >= Number(cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS || 2); return `<tr><td>${esc(rules.formatIndonesianDate(item.record.date,false))}</td><td><span class="category-pill ${category.rest ? 'rest-day' : ''}">${esc(category.label)}</span></td><td>${rules.formatMinutes(item.record.inMinutes)}</td><td>${rules.formatMinutes(item.record.overtimeStartMinutes)}</td><td>${rules.formatMinutes(item.record.outMinutes)}</td><td>${item.record.overtimeHours} jam</td><td>${meal ? 'Ya' : 'Tidak'}</td></tr>`; }).join('')}</tbody></table></div></aside>`;
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
    state.files = [...existing.values()]; state.validation = null; render();
  }

  async function validateFiles() {
    if (!state.files.length || state.busy) return;
    state.busy = true;
    try {
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
    state.employees = state.validation.employees;
    state.processing = {percent:100,active:5}; render(); await new Promise(r=>setTimeout(r,220));
    state.processStep = 'review'; render();
  }

  function updateOvertime(key, value) {
    const [employeeKey, dateKey] = key.split('|');
    const employee = state.employees.find(e=>rules.employeeKey(e)===employeeKey);
    if (!employee || !employee.records[dateKey]) return;
    employee.records[dateKey].overtimeHours = Math.max(0,Math.min(4,Number(value)||0));
    state.generated = null;
    render();
  }

  async function generateDocs() {
    if (state.busy) return;
    state.busy = true; render();
    try {
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
      await refreshHistory(false);
      state.currentRun = { ...run, ...savedMeta };
      state.resultFromHistory = Boolean(state.editingHistoryId);
      state.resultMode = state.editingHistoryId ? 'updated' : 'new';
      state.resultTab = 'summary';
      state.resultSearch = '';
      state.resultCategory = 'all';
      state.resultDate = 'all';
      state.editingHistoryId = null;
      state.processStep = 'result';
    } catch (error) {
      alert(`Gagal membuat dokumen: ${error.message || error}`);
    } finally { state.busy=false; render(); }
  }

  async function download(kind) {
    const run = state.currentRun || {period:state.period,employees:state.employees,holidays:state.holidays};
    if (!run.employees?.length) return;
    try {
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
      render();
    } catch (error) {
      alert(error.message || String(error));
    }
  }

  async function editHistory(id) {
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
    state.view = 'history';
    render();
  }

  async function deleteHistory(id) {
    try {
      const run = await storage.getRun(id);
      if (!run) throw new Error('Riwayat tidak ditemukan atau sudah dihapus.');
      if (!storage.canManage?.(run)) throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat dihapus oleh pembuat atau administrator.');
      const label = periodLabel(run.period);
      const approved = window.confirm(`Hapus riwayat proses ${label}?\n\nData hasil proses periode ini akan dihapus dari Firestore dan tindakan ini tidak dapat dibatalkan.`);
      if (!approved) return;
      await storage.deleteRun(id);
      await refreshHistory(false);
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
    document.querySelector('[data-action="to-confirm"]')?.addEventListener('click',()=>{state.processStep='confirm';render();});
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
    document.querySelectorAll('[data-result-tab]').forEach((btn) => btn.addEventListener('click', () => { state.resultTab = btn.dataset.resultTab || 'summary'; state.resultSearch = ''; state.resultCategory = 'all'; state.resultDate = 'all'; state.drawerKey = null; render(); }));
    document.querySelectorAll('[data-result-employee]').forEach((btn) => btn.addEventListener('click', () => { state.drawerKey = btn.dataset.resultEmployee; render(); }));
    document.getElementById('result-search')?.addEventListener('input', (event) => { state.resultSearch = event.target.value; clearTimeout(window.__resultSearchTimer); window.__resultSearchTimer = setTimeout(render, 180); });
    document.getElementById('result-category')?.addEventListener('change', (event) => { state.resultCategory = event.target.value || 'all'; render(); });
    document.getElementById('result-date')?.addEventListener('change', (event) => { state.resultDate = event.target.value || 'all'; render(); });
    document.querySelectorAll('[data-employee-key]').forEach(btn=>btn.addEventListener('click',()=>{state.drawerKey=btn.dataset.employeeKey;render();}));
    document.querySelectorAll('[data-action="close-drawer"]').forEach(btn=>btn.addEventListener('click',()=>{state.drawerKey=null;render();}));
    document.querySelectorAll('[data-overtime-key]').forEach(sel=>sel.addEventListener('change',()=>updateOvertime(sel.dataset.overtimeKey,sel.value)));
    document.getElementById('review-search')?.addEventListener('input',e=>{state.filter=e.target.value; clearTimeout(window.__reviewTimer); window.__reviewTimer=setTimeout(render,180);});
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
    await window.FirebaseClient.requireAuth();
    await window.AppSettingsService?.loadAndApply?.();
    await window.MasterDataService?.getEmployees?.(false);
    applyHashRoute();
    await refreshHistory(false);
    historyUnsubscribe = storage.subscribeHistory?.((history) => {
      state.history = history;
      if (state.view === 'history' || state.view === 'dashboard') render();
    }, (error) => console.warn('Sinkronisasi realtime Lembur gagal:', error)) || null;
    render();
  }

  window.addEventListener('beforeunload', () => historyUnsubscribe?.());
  init();
})();
