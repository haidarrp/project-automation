(function () {
  'use strict';

  const app = document.getElementById('app');
  const cfg = window.APP_CONFIG;
  const lemburStorage = window.AppStorage;
  const tukinStorage = window.TukinStorage;

  const state = { loading: true, lembur: [], tukin: [], error: '' };

  function esc(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function icon(name) {
    const icons = {
      home: '<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M9 21v-6h6v6"/>',
      tukin: '<path d="M5 4h14v16H5z"/><path d="M8 8h8M8 12h8M8 16h5"/>',
      overtime: '<circle cx="12" cy="12" r="8"/><path d="M12 7v5l3 2"/>',
      users: '<circle cx="9" cy="8" r="3"/><path d="M3.8 19c.7-3 2.4-4.5 5.2-4.5S13.5 16 14.2 19"/><circle cx="17" cy="9" r="2.2"/><path d="M15.7 14.8c2.6-.1 4.1 1.3 4.5 4.2"/>',
      money: '<path d="M4 6h16v12H4z"/><path d="M8 10h.01M16 14h.01"/><circle cx="12" cy="12" r="2.2"/>',
      timer: '<path d="M9 3h6"/><path d="M12 7v5l3 2"/><circle cx="12" cy="13" r="7"/>',
      history: '<path d="M4 6h16M4 12h16M4 18h10"/>',
      chevron: '<path d="m9 18 6-6-6-6"/>',
      plus: '<path d="M12 5v14M5 12h14"/>'
    };
    return `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[name] || icons.home}</svg>`;
  }

  function periodLabel(period) {
    if (!period || !period.month || !period.year) return 'Belum ada';
    return `${cfg.INDONESIAN_MONTHS[Number(period.month) - 1]} ${period.year}`;
  }

  function formatDateTime(iso) {
    if (!iso) return '—';
    try {
      return new Intl.DateTimeFormat('id-ID', {
        timeZone: cfg.TIME_ZONE,
        day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit'
      }).format(new Date(iso));
    } catch (_) {
      return new Date(iso).toLocaleString('id-ID');
    }
  }

  function money(value) {
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(Number(value || 0));
  }

  function latestTimestamp(run) {
    return new Date(run?.updatedAt || run?.processedAt || 0).getTime() || 0;
  }

  function shell(content) {
    return window.AppShell.render({ module: 'dashboard', view: '', viewLabel: 'Dashboard', content });
  }

  function metric(label, value, note, iconName) {
    return `<div class="card main-metric"><div class="main-metric-head"><div class="main-metric-label">${esc(label)}</div><div class="main-metric-icon">${icon(iconName)}</div></div><div><div class="main-metric-value">${value}</div><div class="main-metric-note">${esc(note)}</div></div></div>`;
  }

  function moduleCard(type, run) {
    const isTukin = type === 'tukin';
    const summary = run?.summary || {};
    const processHref = isTukin ? 'tukin.html#process' : 'lembur.html#process';
    const historyHref = isTukin ? 'tukin.html#history' : 'lembur.html#history';
    const title = isTukin ? 'Tunjangan Kinerja' : 'Lembur';
    const subtitle = isTukin ? 'Perhitungan, verifikasi, dan generate dokumen Tukin.' : 'Perhitungan lembur dan generate dokumen pendukung.';
    const stats = isTukin
      ? [
          [Number(summary.employees || 0), 'Pegawai'],
          [money(summary.totalCutAmount || 0), 'Total potongan'],
          [Number(summary.adjustedRecords || 0), 'Koreksi']
        ]
      : [
          [Number(summary.employees || 0), 'Pegawai'],
          [Number(summary.overtimeEmployees || 0), 'Pegawai lembur'],
          [`${Number(summary.totalHours || 0)} jam`, 'Total lembur']
        ];
    return `<div class="card main-module-card">
      <div class="main-module-head">
        <div class="main-module-identity"><div class="main-module-icon">${icon(isTukin ? 'tukin' : 'overtime')}</div><div><div class="main-module-title">${title}</div><div class="main-module-subtitle">${subtitle}</div></div></div>
        <div class="main-module-status">${run ? 'Data tersedia' : 'Belum ada data'}</div>
      </div>
      <div class="main-module-period"><div class="main-module-period-label">Periode terakhir</div><div class="main-module-period-value">${esc(periodLabel(run?.period))}</div><div class="main-module-stats">${stats.map(([value, label]) => `<div class="main-module-stat"><strong>${typeof value === 'number' ? value : esc(value)}</strong><span>${esc(label)}</span></div>`).join('')}</div></div>
      <div class="main-module-foot"><div class="main-module-links"><a class="btn btn-primary btn-sm" href="${processHref}">${icon('plus')} Proses Baru</a><a class="btn btn-secondary btn-sm" href="${historyHref}">${icon('history')} Riwayat</a></div><div class="main-updated">${run ? `Diperbarui ${esc(formatDateTime(run.updatedAt || run.processedAt))}` : 'Belum pernah diproses'}</div></div>
    </div>`;
  }

  function recentRows() {
    const rows = [];
    state.tukin.forEach((run) => rows.push({ type: 'tukin', run, time: latestTimestamp(run) }));
    state.lembur.forEach((run) => rows.push({ type: 'lembur', run, time: latestTimestamp(run) }));
    rows.sort((a, b) => b.time - a.time);
    const recent = rows.slice(0, 6);
    if (!recent.length) return '<div class="empty-state"><strong>Belum ada riwayat proses</strong>Riwayat Tunjangan Kinerja dan Lembur akan tampil di sini setelah proses pertama.</div>';
    return recent.map(({ type, run }) => {
      const isTukin = type === 'tukin';
      const summary = run.summary || {};
      const result = isTukin
        ? `${Number(summary.employees || 0)} pegawai · ${money(summary.totalCutAmount || 0)}`
        : `${Number(summary.overtimeEmployees || 0)} pegawai lembur · ${Number(summary.totalHours || 0)} jam`;
      const href = isTukin ? 'tukin.html#history' : 'lembur.html#history';
      return `<div class="main-recent-row">
        <div class="main-recent-module"><span class="main-recent-module-mark">${icon(isTukin ? 'tukin' : 'overtime')}</span><span>${isTukin ? 'Tunjangan Kinerja' : 'Lembur'}</span></div>
        <div><div class="main-recent-period">${esc(periodLabel(run.period))}</div><div class="main-recent-meta">${isTukin ? 'Periode Tukin' : 'Periode lembur'}</div></div>
        <div>${esc(result)}</div>
        <div>${esc(formatDateTime(run.updatedAt || run.processedAt))}</div>
        <a class="chev-btn" href="${href}" aria-label="Buka riwayat ${isTukin ? 'Tunjangan Kinerja' : 'Lembur'}">${icon('chevron')}</a>
      </div>`;
    }).join('');
  }

  function render() {
    if (state.loading) {
      app.innerHTML = shell('<div class="card main-dashboard-loading"><strong>Memuat dashboard</strong>Membaca riwayat Tunjangan Kinerja dan Lembur dari Cloud Firestore...</div>');
      return;
    }

    const latestTukin = state.tukin[0] || null;
    const latestLembur = state.lembur[0] || null;
    const tukinSummary = latestTukin?.summary || {};
    const lemburSummary = latestLembur?.summary || {};
    const latestActivity = [...state.tukin, ...state.lembur].sort((a, b) => latestTimestamp(b) - latestTimestamp(a))[0];

    const content = `<div class="main-dashboard">
      <div class="main-dashboard-hero"><div><h1>Dashboard Utama</h1><p>Ringkasan proses Tunjangan Kinerja dan Lembur Pusat Data dan Informasi.</p></div><div class="main-dashboard-actions"><a class="btn btn-secondary" href="tukin.html#process">${icon('tukin')} Proses Tukin</a><a class="btn btn-primary" href="lembur.html#process">${icon('overtime')} Proses Lembur</a></div></div>

      ${state.error ? `<div class="alert alert-warning"><div class="alert-title">Sebagian riwayat tidak dapat dibaca</div>${esc(state.error)}</div>` : ''}

      <div class="main-dashboard-metrics">
        ${metric('Pegawai Tukin Terakhir', Number(tukinSummary.employees || 0), latestTukin ? periodLabel(latestTukin.period) : 'Belum ada proses Tukin', 'users')}
        ${metric('Total Potongan Tukin', money(tukinSummary.totalCutAmount || 0), latestTukin ? 'Hasil periode terakhir' : 'Belum ada data', 'money')}
        ${metric('Pegawai Lembur Terakhir', Number(lemburSummary.overtimeEmployees || 0), latestLembur ? periodLabel(latestLembur.period) : 'Belum ada proses lembur', 'overtime')}
        ${metric('Total Jam Lembur', `${Number(lemburSummary.totalHours || 0)} jam`, latestLembur ? 'Hasil periode terakhir' : 'Belum ada data', 'timer')}
      </div>

      <div class="main-module-grid">${moduleCard('tukin', latestTukin)}${moduleCard('lembur', latestLembur)}</div>

      <div class="card main-recent"><div class="main-recent-head"><div><div class="card-title">Aktivitas Terbaru</div><div class="card-subtitle">Riwayat proses kedua modul pada akun Firebase ini.</div></div><div class="main-updated">${latestActivity ? `Terakhir ${esc(formatDateTime(latestActivity.updatedAt || latestActivity.processedAt))}` : 'Belum ada aktivitas'}</div></div>
        ${(state.tukin.length || state.lembur.length) ? '<div class="main-recent-table-head"><div>Modul</div><div>Periode</div><div>Ringkasan</div><div>Diproses</div><div></div></div>' : ''}
        ${recentRows()}
      </div>
    </div>`;
    app.innerHTML = shell(content);
  }

  async function init() {
    await window.FirebaseClient.requireAuth();
    await window.AppSettingsService?.loadAndApply?.();
    const legacyRoute = String(location.hash || '').replace(/^#/, '').toLowerCase();
    if (legacyRoute === 'process' || legacyRoute === 'history') {
      location.replace(`lembur.html#${legacyRoute}`);
      return;
    }
    render();
    try {
      state.lembur = lemburStorage?.listHistory ? await lemburStorage.listHistory() : [];
    } catch (error) {
      state.lembur = [];
      state.error = `Riwayat Lembur: ${error.message || error}`;
    }
    try {
      state.tukin = await tukinStorage.listRuns();
    } catch (error) {
      state.tukin = [];
      state.error = [state.error, `Riwayat Tukin: ${error.message || error}`].filter(Boolean).join(' · ');
    }
    state.loading = false;
    render();
  }

  init();
})();
