(function () {
  'use strict';

  const app = document.getElementById('app');
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const TRAINING_STATUSES = Object.freeze([
    'Diusulkan',
    'Mendaftar/Terdaftar',
    'Lolos Seleksi',
    'Tidak Lolos Seleksi',
    'Mengikuti',
    'Lulus',
    'Tidak Lulus'
  ]);
  const DOCUMENT_TYPES = Object.freeze([
    'Sertifikat',
    'Surat Tugas',
    'Bukti Keikutsertaan',
    'Bukti Kelulusan',
    'Dokumen Lainnya'
  ]);

  const state = {
    loading: true,
    error: '',
    toast: '',
    view: 'dashboard',
    employees: [],
    directory: new Map(),
    trainings: [],
    filters: {
      search: '',
      startMonth: `${now.getFullYear()}-01`,
      endMonth: currentMonth
    },
    modal: null,
    drawer: null,
    documentModal: null
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
    if (value === 'pelatihan' || value === 'pegawai' || value === 'dashboard') return value;
    return 'dashboard';
  }

  function viewLabel(value = state.view) {
    if (value === 'pelatihan') return 'Data Pelatihan';
    if (value === 'pegawai') return 'Data Pegawai';
    return 'Dashboard';
  }

  function viewFromHash() {
    return normalizeView(String(location.hash || '').replace(/^#/, '').toLowerCase());
  }

  function employeeIdentity(id) {
    const row = state.directory.get(String(id || '')) || {};
    return {
      name: String(row.name || '').trim(),
      nip: String(row.nip || '').replace(/\D/g, ''),
      unit: String(row.unit || '').trim()
    };
  }

  function employeeName(id) {
    return employeeIdentity(id).name;
  }

  function unresolvedEmployeeCount() {
    return state.employees.filter((item) => item.active !== false && !employeeName(item.id)).length;
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

  function dateObj(value) {
    if (!value) return null;
    const date = new Date(`${value}T00:00:00`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function formatDate(value) {
    const date = dateObj(value);
    if (!date) return '—';
    return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
  }

  function formatDateRange(training) {
    const start = String(training?.startDate || '');
    const end = String(training?.endDate || '');
    if (!start) return '—';
    if (!end || end === start) return formatDate(start);
    return `${formatDate(start)} – ${formatDate(end)}`;
  }

  function monthStart(value) {
    return /^\d{4}-\d{2}$/.test(String(value || '')) ? `${value}-01` : '';
  }

  function monthEnd(value) {
    if (!/^\d{4}-\d{2}$/.test(String(value || ''))) return '';
    const [year, month] = value.split('-').map(Number);
    const last = new Date(year, month, 0).getDate();
    return `${value}-${String(last).padStart(2, '0')}`;
  }

  function monthLabel(value) {
    if (!/^\d{4}-\d{2}$/.test(String(value || ''))) return 'Semua periode';
    const [year, month] = value.split('-').map(Number);
    return new Intl.DateTimeFormat('id-ID', { month: 'long', year: 'numeric' }).format(new Date(year, month - 1, 1));
  }

  function periodLabel() {
    const start = state.filters.startMonth;
    const end = state.filters.endMonth;
    if (start && end) return start === end ? monthLabel(start) : `${monthLabel(start)} – ${monthLabel(end)}`;
    if (start) return `Mulai ${monthLabel(start)}`;
    if (end) return `Sampai ${monthLabel(end)}`;
    return 'Semua periode';
  }

  function trainingMatchesPeriod(training) {
    const start = String(training?.startDate || '');
    const end = String(training?.endDate || training?.startDate || '');
    if (!start) return false;
    const min = monthStart(state.filters.startMonth);
    const max = monthEnd(state.filters.endMonth);
    if (min && end < min) return false;
    if (max && start > max) return false;
    return true;
  }

  function trainingsInPeriod() {
    return state.trainings
      .filter(trainingMatchesPeriod)
      .sort((a, b) => String(b.startDate || '').localeCompare(String(a.startDate || '')) || String(a.name || '').localeCompare(String(b.name || ''), 'id'));
  }

  function normalizeDocument(raw, index) {
    const row = raw && typeof raw === 'object' ? raw : {};
    return {
      id: String(row.id || `doc-${index + 1}`),
      type: DOCUMENT_TYPES.includes(String(row.type || '')) ? String(row.type) : String(row.type || 'Dokumen Lainnya'),
      name: String(row.name || ''),
      number: String(row.number || ''),
      note: String(row.note || '')
    };
  }

  function normalizeParticipant(raw, fallbackId) {
    const row = raw && typeof raw === 'object' ? raw : {};
    const employeeId = String(row.employeeId || row.id || fallbackId || '');
    const status = TRAINING_STATUSES.includes(String(row.status || '')) ? String(row.status) : 'Diusulkan';
    return {
      employeeId,
      status,
      note: String(row.note || ''),
      documents: (Array.isArray(row.documents) ? row.documents : []).map(normalizeDocument)
    };
  }

  function normalizeTraining(data, id) {
    const source = data || {};
    const participantIds = [...new Set((Array.isArray(source.participantIds) ? source.participantIds : []).map(String).filter(Boolean))];
    const participantMap = new Map();
    (Array.isArray(source.participants) ? source.participants : []).forEach((item) => {
      const participant = normalizeParticipant(item);
      if (participant.employeeId) participantMap.set(participant.employeeId, participant);
    });
    participantIds.forEach((employeeId) => {
      if (!participantMap.has(employeeId)) participantMap.set(employeeId, normalizeParticipant({}, employeeId));
    });
    const allIds = [...new Set([...participantIds, ...participantMap.keys()])];
    return {
      ...source,
      id,
      name: String(source.name || ''),
      organizer: String(source.organizer || ''),
      startDate: String(source.startDate || ''),
      endDate: String(source.endDate || ''),
      participantIds: allIds,
      participants: allIds.map((employeeId) => participantMap.get(employeeId) || normalizeParticipant({}, employeeId))
    };
  }

  function participantFor(training, employeeId) {
    const id = String(employeeId || '');
    return (training?.participants || []).find((item) => item.employeeId === id) || normalizeParticipant({}, id);
  }

  function filteredTrainings() {
    const q = state.filters.search.trim().toLowerCase();
    const rows = trainingsInPeriod();
    if (!q) return rows;
    return rows.filter((training) => {
      const participantNames = (training.participantIds || []).map(employeeName).join(' ');
      const statuses = (training.participants || []).map((item) => item.status).join(' ');
      return [training.name, training.organizer, participantNames, statuses]
        .some((value) => String(value || '').toLowerCase().includes(q));
    });
  }

  function employeeHistory(employeeId) {
    return trainingsInPeriod().filter((training) => (training.participantIds || []).includes(employeeId));
  }

  function filteredEmployees() {
    const q = state.filters.search.trim().toLowerCase();
    const rows = activeEmployees();
    if (!q) return rows;
    return rows.filter((employee) => [employee.name, employee.nip, employee.unit].some((value) => String(value || '').toLowerCase().includes(q)));
  }

  function statusClass(status) {
    const map = {
      'Diusulkan': 'neutral',
      'Mendaftar/Terdaftar': 'info',
      'Lolos Seleksi': 'success',
      'Tidak Lolos Seleksi': 'warning',
      'Mengikuti': 'active',
      'Lulus': 'success',
      'Tidak Lulus': 'danger'
    };
    return map[status] || 'neutral';
  }

  function statusBadge(status) {
    const value = TRAINING_STATUSES.includes(String(status || '')) ? String(status) : 'Diusulkan';
    return `<span class="training-status-badge ${statusClass(value)}">${esc(value)}</span>`;
  }

  function identityNotice() {
    const count = unresolvedEmployeeCount();
    if (!count) return '';
    return `<div class="alert alert-warning training-identity-alert"><div class="alert-title">${esc(count)} identitas pegawai belum tersedia</div>Lengkapi Nama/NIP pada menu Administrasi agar identitas pegawai dapat digunakan pada arsip pelatihan.</div>`;
  }

  function metric(label, value, note) {
    return `<div class="card training-metric"><div class="training-metric-label">${esc(label)}</div><div><div class="training-metric-value">${esc(value)}</div><div class="training-metric-note">${esc(note)}</div></div></div>`;
  }

  function filterCard() {
    const placeholder = state.view === 'pegawai' ? 'Cari nama/NIP pegawai...' : 'Cari pelatihan, penyelenggara, peserta, atau status...';
    return `<section class="card training-filter-card">
      <div class="training-filter-row">
        <div class="search"><input id="training-search" type="search" placeholder="${placeholder}" value="${esc(state.filters.search)}" autocomplete="off"></div>
        <div class="field"><label for="training-start-month">Dari Bulan</label><input id="training-start-month" type="month" value="${esc(state.filters.startMonth)}"></div>
        <div class="field"><label for="training-end-month">Sampai Bulan</label><input id="training-end-month" type="month" value="${esc(state.filters.endMonth)}"></div>
        <div class="training-filter-actions"><button class="btn btn-secondary btn-sm" type="button" data-action="reset-filter">Reset</button></div>
      </div>
    </section>`;
  }

  function employeeMetrics() {
    const employees = activeEmployees();
    const histories = new Map(employees.map((item) => [item.id, employeeHistory(item.id)]));
    const trained = employees.filter((item) => (histories.get(item.id) || []).length > 0).length;
    const participation = [...histories.values()].reduce((sum, items) => sum + items.length, 0);
    return `<div class="training-metrics">
      ${metric('Total Pegawai', employees.length, 'Pegawai aktif pada master')}
      ${metric('Memiliki Riwayat', trained, periodLabel())}
      ${metric('Belum Ada Riwayat', Math.max(0, employees.length - trained), periodLabel())}
      ${metric('Keikutsertaan', participation, 'Total relasi peserta-pelatihan')}
    </div>`;
  }

  function trainingMetrics() {
    const rows = trainingsInPeriod();
    const participantIds = new Set(rows.flatMap((item) => item.participantIds || []));
    const participation = rows.reduce((sum, item) => sum + (item.participantIds || []).length, 0);
    const completed = rows.reduce((sum, item) => sum + (item.participants || []).filter((row) => row.status === 'Lulus').length, 0);
    return `<div class="training-metrics">
      ${metric('Pelatihan Tercatat', rows.length, periodLabel())}
      ${metric('Pegawai Terlibat', participantIds.size, 'Peserta unik pada periode')}
      ${metric('Keikutsertaan', participation, 'Total peserta seluruh pelatihan')}
      ${metric('Status Lulus', completed, 'Relasi peserta berstatus Lulus')}
    </div>`;
  }

  function employeeTable() {
    const rows = filteredEmployees();
    return `<section class="card training-card">
      <div class="training-card-head"><div><h2>Data Pegawai</h2><p>Klik nama pegawai untuk melihat riwayat pelatihan pada rentang bulan yang dipilih.</p></div><span class="card-subtitle">${rows.length} pegawai</span></div>
      <div class="training-table-wrap"><table class="data-table training-table"><thead><tr><th>Nama Pegawai</th><th>Jumlah Pelatihan</th><th>Pelatihan Terakhir</th><th>Status Terakhir</th><th>Tanggal Terakhir</th><th>Aksi</th></tr></thead><tbody>
        ${rows.length ? rows.map((employee) => {
          const history = employeeHistory(employee.id);
          const latest = history[0] || null;
          const participant = latest ? participantFor(latest, employee.id) : null;
          return `<tr>
            <td><button class="training-name-button" type="button" data-employee-detail="${esc(employee.id)}">${esc(employee.name)}</button><span class="training-sub">${employee.nip ? `NIP ${esc(employee.nip)}` : 'NIP belum tersedia'}</span></td>
            <td><span class="training-count ${history.length ? '' : 'zero'}">${history.length}</span></td>
            <td>${latest ? `${esc(latest.name)}<span class="training-sub">${esc(latest.organizer || '—')}</span>` : '<span class="training-sub">Belum ada pelatihan pada periode ini</span>'}</td>
            <td>${participant ? statusBadge(participant.status) : '—'}</td>
            <td class="nowrap">${latest ? esc(formatDateRange(latest)) : '—'}</td>
            <td><button class="btn btn-secondary btn-sm" type="button" data-employee-detail="${esc(employee.id)}">Detail</button></td>
          </tr>`;
        }).join('') : '<tr><td colspan="6"><div class="empty-state"><strong>Pegawai tidak ditemukan</strong>Ubah kata kunci atau rentang bulan yang digunakan.</div></td></tr>'}
      </tbody></table></div>
      <div class="training-period-note">Periode aktif: ${esc(periodLabel())}. Pelatihan yang melintasi batas bulan tetap dihitung apabila tanggalnya beririsan dengan periode.</div>
    </section>`;
  }

  function trainingTable() {
    const rows = filteredTrainings();
    return `<section class="card training-card">
      <div class="training-card-head"><div><h2>Data Pelatihan</h2><p>Satu pelatihan dicatat satu kali dan memiliki status/catatan yang melekat pada masing-masing peserta.</p></div><span class="card-subtitle">${rows.length} pelatihan</span></div>
      <div class="training-table-wrap"><table class="data-table training-table"><thead><tr><th>Nama Pelatihan</th><th>Tanggal</th><th>Penyelenggara</th><th>Peserta</th><th>Aksi</th></tr></thead><tbody>
        ${rows.length ? rows.map((training) => `<tr>
          <td><button class="training-name-button" type="button" data-training-detail="${esc(training.id)}">${esc(training.name || 'Tanpa nama')}</button></td>
          <td class="nowrap">${esc(formatDateRange(training))}</td>
          <td>${esc(training.organizer || '—')}</td>
          <td><span class="training-count ${(training.participantIds || []).length ? '' : 'zero'}">${(training.participantIds || []).length}</span></td>
          <td><div class="training-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-training-detail="${esc(training.id)}">Detail</button><button class="btn btn-secondary btn-sm" type="button" data-training-edit="${esc(training.id)}">Edit</button><button class="btn btn-danger btn-sm" type="button" data-training-delete="${esc(training.id)}">Hapus</button></div></td>
        </tr>`).join('') : '<tr><td colspan="5"><div class="empty-state"><strong>Belum ada data pelatihan</strong>Tambahkan data baru atau ubah filter periode.</div></td></tr>'}
      </tbody></table></div>
      <div class="training-period-note">Periode aktif: ${esc(periodLabel())}. Filter menggunakan rentang tanggal pelatihan.</div>
    </section>`;
  }

  function dashboardStatusSummary(rows) {
    const counts = new Map(TRAINING_STATUSES.map((status) => [status, 0]));
    rows.forEach((training) => (training.participants || []).forEach((participant) => {
      const status = TRAINING_STATUSES.includes(participant.status) ? participant.status : 'Diusulkan';
      counts.set(status, (counts.get(status) || 0) + 1);
    }));
    const total = [...counts.values()].reduce((sum, value) => sum + value, 0);
    return `<section class="card training-card training-dashboard-status"><div class="training-card-head"><div><h2>Komposisi Status Peserta</h2><p>Status peserta pada periode yang dipilih.</p></div><span class="card-subtitle">${total} keikutsertaan</span></div><div class="training-status-summary">${TRAINING_STATUSES.map((status) => `<div class="training-status-summary-row"><span>${statusBadge(status)}</span><strong>${counts.get(status) || 0}</strong></div>`).join('')}</div></section>`;
  }

  function dashboardRecentTraining(rows) {
    const recent = rows.slice().sort((a, b) => String(b.startDate || '').localeCompare(String(a.startDate || ''))).slice(0, 5);
    return `<section class="card training-card"><div class="training-card-head"><div><h2>Pelatihan Terbaru</h2><p>Data pelatihan terbaru pada periode aktif.</p></div><a class="btn btn-secondary btn-sm" href="pelatihan.html#pelatihan">Lihat Semua</a></div><div class="training-table-wrap"><table class="data-table training-table"><thead><tr><th>Pelatihan</th><th>Tanggal</th><th>Penyelenggara</th><th>Peserta</th><th>Aksi</th></tr></thead><tbody>${recent.length ? recent.map((training) => `<tr><td><button class="training-name-button" type="button" data-training-detail="${esc(training.id)}">${esc(training.name || 'Tanpa nama')}</button></td><td class="nowrap">${esc(formatDateRange(training))}</td><td>${esc(training.organizer || '—')}</td><td><span class="training-count">${(training.participantIds || []).length}</span></td><td><button class="btn btn-secondary btn-sm" type="button" data-training-detail="${esc(training.id)}">Detail</button></td></tr>`).join('') : '<tr><td colspan="5"><div class="empty-state"><strong>Belum ada data pelatihan</strong>Tambahkan pelatihan untuk menampilkan ringkasan dashboard.</div></td></tr>'}</tbody></table></div></section>`;
  }

  function dashboardView() {
    const rows = trainingsInPeriod();
    const employees = activeEmployees();
    const participantIds = new Set(rows.flatMap((item) => item.participantIds || []));
    const participation = rows.reduce((sum, item) => sum + (item.participantIds || []).length, 0);
    const completed = rows.reduce((sum, item) => sum + (item.participants || []).filter((participant) => participant.status === 'Lulus').length, 0);
    return `<div class="training-page"><div class="training-head"><div><h1>Dashboard Pelatihan</h1><p>Ringkasan pencatatan pelatihan pegawai Pusat Data dan Informasi.</p></div><div class="training-head-actions"><a class="btn btn-secondary" href="pelatihan.html#pegawai">Data Pegawai</a><button class="btn btn-primary" type="button" data-action="add-training">+ Tambah Pelatihan</button></div></div>${filterCard()}${identityNotice()}<div class="training-metrics">${metric('Total Pegawai', employees.length, 'Pegawai aktif pada master')}${metric('Pelatihan Tercatat', rows.length, periodLabel())}${metric('Pegawai Terlibat', participantIds.size, 'Peserta unik pada periode')}${metric('Status Lulus', completed, `${participation} total keikutsertaan`)}</div><div class="training-dashboard-grid">${dashboardStatusSummary(rows)}${dashboardRecentTraining(rows)}</div></div>`;
  }

  function employeeView() {
    return `<div class="training-page">
      <div class="training-head"><div><h1>Data Pegawai</h1><p>Arsip ringkas pelatihan pegawai Pusat Data dan Informasi berdasarkan periode.</p></div><div class="training-head-actions"><a class="btn btn-secondary" href="pelatihan.html#pelatihan">Data Pelatihan</a></div></div>
      ${filterCard()}
      ${identityNotice()}
      ${employeeMetrics()}
      ${employeeTable()}
    </div>`;
  }

  function trainingView() {
    return `<div class="training-page">
      <div class="training-head"><div><h1>Data Pelatihan</h1><p>Pencatatan dan arsip pelatihan pegawai dengan status, catatan, dan metadata dokumen per peserta.</p></div><div class="training-head-actions"><a class="btn btn-secondary" href="pelatihan.html#pegawai">Data Pegawai</a><button class="btn btn-primary" type="button" data-action="add-training">+ Tambah Pelatihan</button></div></div>
      ${filterCard()}
      ${identityNotice()}
      ${trainingMetrics()}
      ${trainingTable()}
    </div>`;
  }

  function trainingDrawer(training) {
    const participants = (training.participantIds || [])
      .filter((id) => Boolean(employeeName(id)))
      .slice()
      .sort((a, b) => employeeOrder(a) - employeeOrder(b) || employeeName(a).localeCompare(employeeName(b), 'id'));
    return `<div class="training-drawer-backdrop" data-action="close-drawer"></div><aside class="training-drawer" role="dialog" aria-modal="true" aria-label="Detail pelatihan">
      <div class="training-drawer-head"><div><h3>${esc(training.name || 'Pelatihan')}</h3><p>Detail arsip pelatihan dan status masing-masing peserta.</p></div><button class="icon-btn" type="button" data-action="close-drawer" aria-label="Tutup">×</button></div>
      <div class="training-detail-meta">
        <div class="training-detail-box"><div class="training-detail-label">Tanggal</div><div class="training-detail-value">${esc(formatDateRange(training))}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Penyelenggara</div><div class="training-detail-value">${esc(training.organizer || '—')}</div></div>
      </div>
      <div class="training-list-title">Peserta (${participants.length})</div>
      <div class="training-history-list">${participants.length ? participants.map((id) => {
        const participant = participantFor(training, id);
        return `<button class="training-person-row training-person-button" type="button" data-participant-detail="${esc(id)}" data-participant-training="${esc(training.id)}"><span><strong>${esc(employeeName(id))}</strong>${participant.note ? `<small>${esc(participant.note)}</small>` : ''}</span>${statusBadge(participant.status)}</button>`;
      }).join('') : '<div class="training-participant-empty">Belum ada peserta tercatat.</div>'}</div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-training-edit="${esc(training.id)}">Edit</button><button class="btn btn-danger" type="button" data-training-delete="${esc(training.id)}">Hapus</button></div>
    </aside>`;
  }

  function employeeDrawer(employeeId) {
    const identity = employeeIdentity(employeeId);
    const name = identity.name || 'Nama pegawai belum tersedia';
    const history = employeeHistory(employeeId);
    return `<div class="training-drawer-backdrop" data-action="close-drawer"></div><aside class="training-drawer" role="dialog" aria-modal="true" aria-label="Riwayat pelatihan pegawai">
      <div class="training-drawer-head"><div><h3>${esc(name)}</h3><p>${identity.nip ? `NIP ${esc(identity.nip)} · ` : ''}Riwayat pelatihan pada ${esc(periodLabel())}.</p></div><button class="icon-btn" type="button" data-action="close-drawer" aria-label="Tutup">×</button></div>
      <div class="training-detail-meta">
        <div class="training-detail-box"><div class="training-detail-label">Jumlah Pelatihan</div><div class="training-detail-value">${history.length}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Pelatihan Terakhir</div><div class="training-detail-value">${history[0] ? esc(formatDateRange(history[0])) : '—'}</div></div>
      </div>
      <div class="training-list-title">Riwayat Pelatihan</div>
      <div class="training-history-list">${history.length ? history.map((training) => {
        const participant = participantFor(training, employeeId);
        return `<button class="training-history-item training-history-button" type="button" data-participant-detail="${esc(employeeId)}" data-participant-training="${esc(training.id)}"><span><strong>${esc(training.name)}</strong><small>${esc(formatDateRange(training))} · ${esc(training.organizer || '—')}</small></span>${statusBadge(participant.status)}</button>`;
      }).join('') : '<div class="training-participant-empty">Belum ada pelatihan pada periode ini.</div>'}</div>
    </aside>`;
  }

  function participantTrainingDrawer(training, employeeId) {
    const identity = employeeIdentity(employeeId);
    const participant = participantFor(training, employeeId);
    const documents = participant.documents || [];
    return `<div class="training-drawer-backdrop" data-action="close-drawer"></div><aside class="training-drawer" role="dialog" aria-modal="true" aria-label="Detail pelatihan pegawai">
      <div class="training-drawer-head"><div><h3>${esc(training.name || 'Pelatihan')}</h3><p>${esc(identity.name || 'Pegawai')} · Detail relasi pegawai dan pelatihan.</p></div><button class="icon-btn" type="button" data-action="close-drawer" aria-label="Tutup">×</button></div>
      <div class="training-detail-meta">
        <div class="training-detail-box"><div class="training-detail-label">Penyelenggara</div><div class="training-detail-value">${esc(training.organizer || '—')}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Periode</div><div class="training-detail-value">${esc(formatDateRange(training))}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Status Peserta</div><div class="training-detail-value">${statusBadge(participant.status)}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Pegawai</div><div class="training-detail-value">${esc(identity.name || '—')}${identity.nip ? `<span class="training-sub">NIP ${esc(identity.nip)}</span>` : ''}</div></div>
      </div>
      <div class="training-list-title">Catatan Peserta</div>
      <div class="training-note-box">${participant.note ? esc(participant.note) : '<span>Belum ada catatan peserta.</span>'}</div>
      <div class="training-document-head"><div><div class="training-list-title">Dokumen Pelatihan</div><p>Metadata dokumen melekat pada pegawai dan pelatihan. File fisik belum disimpan oleh modul ini.</p></div><button class="btn btn-secondary btn-sm" type="button" data-add-training-document="${esc(training.id)}" data-employee-id="${esc(employeeId)}">+ Tambah Dokumen</button></div>
      <div class="training-history-list">${documents.length ? documents.map((doc) => `<div class="training-document-row"><div><strong>${esc(doc.name || doc.type || 'Dokumen')}</strong><span>${esc(doc.type || 'Dokumen Lainnya')}${doc.number ? ` · No. ${esc(doc.number)}` : ''}${doc.note ? `<br>${esc(doc.note)}` : ''}</span></div><div class="training-row-actions"><button class="btn btn-secondary btn-sm" type="button" data-edit-training-document="${esc(doc.id)}" data-training-id="${esc(training.id)}" data-employee-id="${esc(employeeId)}">Edit</button><button class="btn btn-danger btn-sm" type="button" data-delete-training-document="${esc(doc.id)}" data-training-id="${esc(training.id)}" data-employee-id="${esc(employeeId)}">Hapus</button></div></div>`).join('') : '<div class="training-participant-empty">Belum ada metadata dokumen pelatihan.</div>'}</div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-training-edit="${esc(training.id)}">Edit Status/Catatan</button></div>
    </aside>`;
  }

  function trainingModal() {
    if (!state.modal) return '';
    const training = state.modal.training || {};
    const participants = state.modal.participants || new Map();
    const employees = activeEmployees();
    return `<div class="training-modal-backdrop" data-action="close-modal"><div class="training-modal training-modal-wide" role="dialog" aria-modal="true" aria-label="${training.id ? 'Edit' : 'Tambah'} pelatihan" data-modal-panel>
      <div class="training-modal-head"><div><h3>${training.id ? 'Edit Data Pelatihan' : 'Tambah Data Pelatihan'}</h3><p>Isi informasi pelatihan, pilih peserta, lalu tetapkan status dan catatan yang spesifik untuk masing-masing peserta.</p></div><button class="icon-btn" type="button" data-action="close-modal" aria-label="Tutup">×</button></div>
      <div class="training-modal-grid">
        <div class="field span-2"><label for="training-name">Nama Pelatihan</label><input id="training-name" type="text" value="${esc(training.name || '')}" autocomplete="off" required></div>
        <div class="field span-2"><label for="training-organizer">Penyelenggara</label><input id="training-organizer" type="text" value="${esc(training.organizer || '')}" autocomplete="off" required></div>
        <div class="field"><label for="training-start-date">Tanggal Mulai</label><input id="training-start-date" type="date" value="${esc(training.startDate || '')}" required></div>
        <div class="field"><label for="training-end-date">Tanggal Selesai <span class="training-optional">(opsional)</span></label><input id="training-end-date" type="date" value="${esc(training.endDate || '')}"></div>
      </div>
      <div class="training-participants">
        <div class="training-participant-head"><strong>Peserta Pelatihan</strong><span><span data-selected-count>${participants.size}</span> pegawai dipilih</span></div>
        <div class="training-participant-search"><div class="search"><input id="participant-search" type="search" placeholder="Cari nama pegawai..." autocomplete="off"></div></div>
        <div class="training-participant-list training-participant-list-rich" data-participant-list>
          ${employees.length ? employees.map((employee) => {
            const participant = participants.get(employee.id);
            const selected = Boolean(participant);
            return `<div class="training-participant-option training-participant-rich" data-participant-row data-search-name="${esc(employee.name.toLowerCase())}">
              <label class="training-participant-identity"><input type="checkbox" data-participant-id="${esc(employee.id)}" ${selected ? 'checked' : ''}><span><strong>${esc(employee.name)}</strong><small>${employee.nip ? `NIP ${esc(employee.nip)}` : 'NIP belum tersedia'}</small></span></label>
              <div class="training-participant-fields ${selected ? '' : 'hidden'}" data-participant-fields="${esc(employee.id)}">
                <div class="field"><label>Status</label><select data-participant-status-select="${esc(employee.id)}">${TRAINING_STATUSES.map((status) => `<option value="${esc(status)}" ${(participant?.status || 'Diusulkan') === status ? 'selected' : ''}>${esc(status)}</option>`).join('')}</select></div>
                <div class="field"><label>Catatan</label><input type="text" data-participant-note="${esc(employee.id)}" value="${esc(participant?.note || '')}" placeholder="Opsional"></div>
              </div>
            </div>`;
          }).join('') : '<div class="training-participant-empty">Master pegawai belum tersedia.</div>'}
        </div>
      </div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-training">Simpan</button></div>
    </div></div>`;
  }

  function trainingDocumentModal() {
    if (!state.documentModal) return '';
    const data = state.documentModal.document || {};
    return `<div class="training-modal-backdrop" data-action="close-document-modal"><div class="training-modal training-document-modal" role="dialog" aria-modal="true" data-document-modal-panel>
      <div class="training-modal-head"><div><h3>${data.id ? 'Edit Metadata Dokumen' : 'Tambah Metadata Dokumen'}</h3><p>Metadata saja. Modul Pelatihan belum menyimpan file binary/attachment.</p></div><button class="icon-btn" type="button" data-action="close-document-modal">×</button></div>
      <div class="training-modal-grid">
        <div class="field"><label for="training-doc-type">Jenis Dokumen</label><select id="training-doc-type">${DOCUMENT_TYPES.map((type) => `<option value="${esc(type)}" ${(data.type || 'Sertifikat') === type ? 'selected' : ''}>${esc(type)}</option>`).join('')}</select></div>
        <div class="field"><label for="training-doc-number">Nomor Dokumen</label><input id="training-doc-number" type="text" value="${esc(data.number || '')}" placeholder="Opsional"></div>
        <div class="field span-2"><label for="training-doc-name">Nama Dokumen</label><input id="training-doc-name" type="text" value="${esc(data.name || '')}" placeholder="Contoh: Sertifikat Pelatihan Data"></div>
        <div class="field span-2"><label for="training-doc-note">Keterangan</label><textarea id="training-doc-note" rows="3" placeholder="Opsional">${esc(data.note || '')}</textarea></div>
      </div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-document-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-training-document">Simpan</button></div>
    </div></div>`;
  }

  function overlayMarkup() {
    let drawer = '';
    if (state.drawer?.type === 'training') drawer = trainingDrawer(state.trainings.find((item) => item.id === state.drawer.id) || {});
    else if (state.drawer?.type === 'employee') drawer = employeeDrawer(state.drawer.id);
    else if (state.drawer?.type === 'participant') {
      const training = state.trainings.find((item) => item.id === state.drawer.trainingId) || {};
      drawer = participantTrainingDrawer(training, state.drawer.employeeId);
    }
    return `${drawer}${trainingModal()}${trainingDocumentModal()}${state.toast ? `<div class="training-toast">${esc(state.toast)}</div>` : ''}`;
  }

  function render() {
    if (state.loading) {
      app.innerHTML = window.AppShell.render({
        module: 'training', view: state.view, viewLabel: viewLabel(),
        content: '<div class="card training-loading"><strong>Memuat data pelatihan</strong>Membaca master pegawai dan arsip pelatihan dari Cloud Firestore...</div>'
      });
      return;
    }

    const mainView = state.view === 'dashboard' ? dashboardView() : state.view === 'pegawai' ? employeeView() : trainingView();
    const content = `${state.error ? `<div class="alert alert-warning"><div class="alert-title">Data belum dapat dimuat sempurna</div>${esc(state.error)}</div>` : ''}${mainView}`;
    app.innerHTML = window.AppShell.render({
      module: 'training',
      view: state.view,
      viewLabel: viewLabel(),
      content,
      overlays: overlayMarkup()
    });
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
      console.warn('Audit pelatihan gagal ditulis:', error);
    }
  }

  async function loadData() {
    state.loading = true;
    state.error = '';
    render();
    try {
      const [masterSnap, directorySnap, trainingSnap] = await Promise.all([
        db().collection('masterEmployees').get(),
        db().collection('masterDirectory').get(),
        db().collection('trainings').get()
      ]);
      state.employees = masterSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      state.directory = new Map(directorySnap.docs.map((doc) => [doc.id, { id: doc.id, ...doc.data() }]));
      state.trainings = trainingSnap.docs.map((doc) => normalizeTraining(doc.data(), doc.id));
    } catch (error) {
      state.error = error.message || String(error);
    } finally {
      state.loading = false;
      render();
    }
  }

  function openAddModal() {
    state.drawer = null;
    state.modal = { training: {}, participants: new Map() };
    render();
  }

  function openEditModal(id) {
    const training = state.trainings.find((item) => item.id === id);
    if (!training) return;
    state.drawer = null;
    state.modal = {
      training: { ...training },
      participants: new Map((training.participants || []).map((item) => [item.employeeId, { ...item, documents: (item.documents || []).map((doc) => ({ ...doc })) }]))
    };
    render();
  }

  function showToast(message) {
    state.toast = String(message || '');
    render();
    clearTimeout(window.__trainingToastTimer);
    window.__trainingToastTimer = setTimeout(() => {
      state.toast = '';
      render();
    }, 3200);
  }

  async function saveTraining(button) {
    if (!state.modal) return;
    const name = String(document.getElementById('training-name')?.value || '').trim();
    const organizer = String(document.getElementById('training-organizer')?.value || '').trim();
    const startDate = String(document.getElementById('training-start-date')?.value || '').trim();
    const endDate = String(document.getElementById('training-end-date')?.value || '').trim();
    const participantIds = [...state.modal.participants.keys()];
    const participants = participantIds.map((employeeId) => {
      const existing = state.modal.participants.get(employeeId) || normalizeParticipant({}, employeeId);
      const status = String(document.querySelector(`[data-participant-status-select="${CSS.escape(employeeId)}"]`)?.value || existing.status || 'Diusulkan');
      const note = String(document.querySelector(`[data-participant-note="${CSS.escape(employeeId)}"]`)?.value || '').trim();
      return normalizeParticipant({ ...existing, employeeId, status, note }, employeeId);
    });

    if (!name || !organizer || !startDate) {
      alert('Nama pelatihan, penyelenggara, dan tanggal mulai wajib diisi.');
      return;
    }
    if (endDate && endDate < startDate) {
      alert('Tanggal selesai tidak boleh lebih awal dari tanggal mulai.');
      return;
    }
    if (!participantIds.length) {
      alert('Pilih minimal satu pegawai sebagai peserta pelatihan.');
      return;
    }

    const existingId = state.modal.training?.id || '';
    const ref = existingId ? db().collection('trainings').doc(existingId) : db().collection('trainings').doc();
    button.disabled = true;
    button.textContent = 'Menyimpan...';
    try {
      const payload = {
        name,
        organizer,
        startDate,
        endDate: endDate || '',
        participantIds,
        participants,
        schemaVersion: 2,
        updatedAt: serverTimestamp(),
        updatedBy: user()?.email || ''
      };
      if (!existingId) {
        payload.createdAt = serverTimestamp();
        payload.createdBy = user()?.email || '';
      }
      await ref.set(payload, { merge: true });
      await audit(existingId ? 'UPDATE_TRAINING' : 'CREATE_TRAINING', `trainings/${ref.id}`, `${name} · ${participantIds.length} peserta`);
      state.modal = null;
      await loadData();
      showToast(existingId ? 'Data pelatihan berhasil diperbarui.' : 'Data pelatihan berhasil ditambahkan.');
    } catch (error) {
      button.disabled = false;
      button.textContent = 'Simpan';
      alert(`Data gagal disimpan: ${error.message || error}`);
    }
  }

  async function deleteTraining(id) {
    const training = state.trainings.find((item) => item.id === id);
    if (!training) return;
    const ok = window.confirm(`Hapus arsip pelatihan "${training.name}"? Data status, catatan, dan metadata dokumen peserta pada pelatihan ini juga akan terhapus.`);
    if (!ok) return;
    try {
      await db().collection('trainings').doc(id).delete();
      await audit('DELETE_TRAINING', `trainings/${id}`, `${training.name} · ${(training.participantIds || []).length} peserta`);
      state.drawer = null;
      state.modal = null;
      await loadData();
      showToast('Data pelatihan berhasil dihapus.');
    } catch (error) {
      alert(`Data gagal dihapus: ${error.message || error}`);
    }
  }

  function openDocumentModal(trainingId, employeeId, documentId) {
    const training = state.trainings.find((item) => item.id === trainingId);
    if (!training) return;
    const participant = participantFor(training, employeeId);
    const document = documentId ? (participant.documents || []).find((item) => item.id === documentId) : null;
    state.documentModal = { trainingId, employeeId, document: document ? { ...document } : {} };
    render();
  }

  async function saveTrainingDocument(button) {
    const ctx = state.documentModal;
    if (!ctx) return;
    const training = state.trainings.find((item) => item.id === ctx.trainingId);
    if (!training) return;
    const type = String(document.getElementById('training-doc-type')?.value || 'Dokumen Lainnya');
    const name = String(document.getElementById('training-doc-name')?.value || '').trim();
    const number = String(document.getElementById('training-doc-number')?.value || '').trim();
    const note = String(document.getElementById('training-doc-note')?.value || '').trim();
    if (!name) {
      alert('Nama dokumen wajib diisi.');
      return;
    }

    const participants = (training.participants || []).map((item) => ({ ...item, documents: (item.documents || []).map((doc) => ({ ...doc })) }));
    const index = participants.findIndex((item) => item.employeeId === ctx.employeeId);
    if (index < 0) return;
    const docs = participants[index].documents || [];
    const existingId = String(ctx.document?.id || '');
    const id = existingId || `doc-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const payload = { id, type, name, number, note };
    const docIndex = docs.findIndex((item) => item.id === id);
    if (docIndex >= 0) docs[docIndex] = payload;
    else docs.push(payload);
    participants[index].documents = docs;

    button.disabled = true;
    try {
      await db().collection('trainings').doc(training.id).set({
        participants,
        participantIds: training.participantIds || [],
        schemaVersion: 2,
        updatedAt: serverTimestamp(),
        updatedBy: user()?.email || ''
      }, { merge: true });
      await audit(existingId ? 'UPDATE_TRAINING_DOCUMENT' : 'CREATE_TRAINING_DOCUMENT', `trainings/${training.id}`, `${employeeName(ctx.employeeId)} · ${name}`);
      state.documentModal = null;
      await loadData();
      state.drawer = { type: 'participant', trainingId: training.id, employeeId: ctx.employeeId };
      showToast(existingId ? 'Metadata dokumen diperbarui.' : 'Metadata dokumen ditambahkan.');
    } catch (error) {
      button.disabled = false;
      alert(`Metadata dokumen gagal disimpan: ${error.message || error}`);
    }
  }

  async function deleteTrainingDocument(trainingId, employeeId, documentId) {
    const training = state.trainings.find((item) => item.id === trainingId);
    if (!training) return;
    const participant = participantFor(training, employeeId);
    const target = (participant.documents || []).find((item) => item.id === documentId);
    if (!target) return;
    if (!window.confirm(`Hapus metadata dokumen "${target.name || target.type}"?`)) return;
    try {
      const participants = (training.participants || []).map((item) => item.employeeId === employeeId
        ? { ...item, documents: (item.documents || []).filter((doc) => doc.id !== documentId) }
        : { ...item, documents: (item.documents || []).map((doc) => ({ ...doc })) });
      await db().collection('trainings').doc(trainingId).set({ participants, schemaVersion: 2, updatedAt: serverTimestamp(), updatedBy: user()?.email || '' }, { merge: true });
      await audit('DELETE_TRAINING_DOCUMENT', `trainings/${trainingId}`, `${employeeName(employeeId)} · ${target.name || target.type}`);
      await loadData();
      state.drawer = { type: 'participant', trainingId, employeeId };
      showToast('Metadata dokumen dihapus.');
    } catch (error) {
      alert(`Metadata dokumen gagal dihapus: ${error.message || error}`);
    }
  }

  function updateFiltersFromDom() {
    state.filters.search = String(document.getElementById('training-search')?.value || '');
    state.filters.startMonth = String(document.getElementById('training-start-month')?.value || '');
    state.filters.endMonth = String(document.getElementById('training-end-month')?.value || '');
    if (state.filters.startMonth && state.filters.endMonth && state.filters.startMonth > state.filters.endMonth) {
      state.filters.endMonth = state.filters.startMonth;
    }
    render();
  }

  app.addEventListener('input', (event) => {
    if (event.target?.id === 'training-search') {
      state.filters.search = String(event.target.value || '');
      render();
      const input = document.getElementById('training-search');
      if (input) { input.focus(); input.setSelectionRange(input.value.length, input.value.length); }
      return;
    }
    if (event.target?.id === 'participant-search') {
      const q = String(event.target.value || '').trim().toLowerCase();
      app.querySelectorAll('[data-participant-row]').forEach((row) => {
        row.style.display = !q || String(row.dataset.searchName || '').includes(q) ? '' : 'none';
      });
      return;
    }
    const noteEl = event.target?.closest?.('[data-participant-note]');
    if (noteEl && state.modal) {
      const id = String(noteEl.dataset.participantNote || '');
      const participant = state.modal.participants.get(id);
      if (participant) participant.note = String(noteEl.value || '');
    }
  });

  app.addEventListener('change', (event) => {
    if (event.target?.id === 'training-start-month' || event.target?.id === 'training-end-month') {
      updateFiltersFromDom();
      return;
    }
    if (event.target?.matches?.('[data-participant-id]') && state.modal) {
      const id = String(event.target.dataset.participantId || '');
      if (event.target.checked) {
        if (!state.modal.participants.has(id)) state.modal.participants.set(id, normalizeParticipant({}, id));
      } else {
        state.modal.participants.delete(id);
      }
      const count = app.querySelector('[data-selected-count]');
      if (count) count.textContent = String(state.modal.participants.size);
      const fields = app.querySelector(`[data-participant-fields="${CSS.escape(id)}"]`);
      fields?.classList.toggle('hidden', !event.target.checked);
      return;
    }
    if (event.target?.matches?.('[data-participant-status-select]') && state.modal) {
      const id = String(event.target.dataset.participantStatusSelect || '');
      const participant = state.modal.participants.get(id);
      if (participant) participant.status = String(event.target.value || 'Diusulkan');
    }
  });

  app.addEventListener('click', (event) => {
    const actionEl = event.target.closest?.('[data-action]');
    const action = actionEl?.dataset.action;

    if (action === 'close-modal' && actionEl?.classList.contains('training-modal-backdrop') && event.target !== actionEl) return;
    if (action === 'close-document-modal' && actionEl?.classList.contains('training-modal-backdrop') && event.target !== actionEl) return;
    if (action === 'reset-filter') {
      state.filters = { search: '', startMonth: `${now.getFullYear()}-01`, endMonth: currentMonth };
      render();
      return;
    }
    if (action === 'add-training') { openAddModal(); return; }
    if (action === 'close-modal') { state.modal = null; render(); return; }
    if (action === 'close-document-modal') { state.documentModal = null; render(); return; }
    if (action === 'close-drawer') { state.drawer = null; render(); return; }
    if (action === 'save-training') { saveTraining(event.target.closest('[data-action]')); return; }
    if (action === 'save-training-document') { saveTrainingDocument(event.target.closest('[data-action]')); return; }

    const employeeDetail = event.target.closest?.('[data-employee-detail]')?.dataset.employeeDetail;
    if (employeeDetail) { state.drawer = { type: 'employee', id: employeeDetail }; render(); return; }

    const trainingDetail = event.target.closest?.('[data-training-detail]')?.dataset.trainingDetail;
    if (trainingDetail) { state.drawer = { type: 'training', id: trainingDetail }; render(); return; }

    const participantDetailEl = event.target.closest?.('[data-participant-detail]');
    if (participantDetailEl) {
      state.drawer = { type: 'participant', trainingId: participantDetailEl.dataset.participantTraining, employeeId: participantDetailEl.dataset.participantDetail };
      render();
      return;
    }

    const addDoc = event.target.closest?.('[data-add-training-document]');
    if (addDoc) { openDocumentModal(addDoc.dataset.addTrainingDocument, addDoc.dataset.employeeId, ''); return; }

    const editDoc = event.target.closest?.('[data-edit-training-document]');
    if (editDoc) { openDocumentModal(editDoc.dataset.trainingId, editDoc.dataset.employeeId, editDoc.dataset.editTrainingDocument); return; }

    const deleteDoc = event.target.closest?.('[data-delete-training-document]');
    if (deleteDoc) { deleteTrainingDocument(deleteDoc.dataset.trainingId, deleteDoc.dataset.employeeId, deleteDoc.dataset.deleteTrainingDocument); return; }

    const trainingEdit = event.target.closest?.('[data-training-edit]')?.dataset.trainingEdit;
    if (trainingEdit) { openEditModal(trainingEdit); return; }

    const trainingDelete = event.target.closest?.('[data-training-delete]')?.dataset.trainingDelete;
    if (trainingDelete) deleteTraining(trainingDelete);
  });

  window.addEventListener('hashchange', () => {
    state.view = viewFromHash();
    state.filters.search = '';
    state.drawer = null;
    state.modal = null;
    state.documentModal = null;
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
