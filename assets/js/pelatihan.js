(function () {
  'use strict';

  const app = document.getElementById('app');
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  const state = {
    loading: true,
    error: '',
    toast: '',
    view: 'pegawai',
    employees: [],
    directory: new Map(),
    trainings: [],
    filters: {
      search: '',
      startMonth: `${now.getFullYear()}-01`,
      endMonth: currentMonth
    },
    modal: null,
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
    return value === 'pelatihan' ? 'pelatihan' : 'pegawai';
  }

  function viewFromHash() {
    return normalizeView(String(location.hash || '').replace(/^#/, '').toLowerCase());
  }

  function employeeIdentity(id) {
    const row = state.directory.get(String(id || '')) || {};
    return {
      name: String(row.name || '').trim(),
      nip: String(row.nip || '').replace(/\D/g, '')
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

  function filteredTrainings() {
    const q = state.filters.search.trim().toLowerCase();
    const rows = trainingsInPeriod();
    if (!q) return rows;
    return rows.filter((training) => {
      const participantNames = (training.participantIds || []).map(employeeName).join(' ');
      return [training.name, training.organizer, participantNames]
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
    return rows.filter((employee) => employee.name.toLowerCase().includes(q));
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
    const placeholder = state.view === 'pegawai' ? 'Cari nama pegawai...' : 'Cari pelatihan, penyelenggara, atau peserta...';
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
      ${metric('Sudah Pelatihan', trained, periodLabel())}
      ${metric('Belum Pelatihan', Math.max(0, employees.length - trained), periodLabel())}
      ${metric('Keikutsertaan', participation, 'Total riwayat pada periode')}
    </div>`;
  }

  function trainingMetrics() {
    const rows = trainingsInPeriod();
    const participantIds = new Set(rows.flatMap((item) => item.participantIds || []));
    const participation = rows.reduce((sum, item) => sum + (item.participantIds || []).length, 0);
    const organizers = new Set(rows.map((item) => String(item.organizer || '').trim()).filter(Boolean));
    return `<div class="training-metrics">
      ${metric('Pelatihan Tercatat', rows.length, periodLabel())}
      ${metric('Pegawai Terlibat', participantIds.size, 'Peserta unik pada periode')}
      ${metric('Keikutsertaan', participation, 'Total peserta seluruh pelatihan')}
      ${metric('Penyelenggara', organizers.size, 'Penyelenggara unik')}
    </div>`;
  }

  function employeeTable() {
    const rows = filteredEmployees();
    return `<section class="card training-card">
      <div class="training-card-head"><div><h2>Data Pegawai</h2><p>Klik nama pegawai untuk melihat riwayat pelatihan pada rentang bulan yang dipilih.</p></div><span class="card-subtitle">${rows.length} pegawai</span></div>
      <div class="training-table-wrap"><table class="data-table training-table"><thead><tr><th>Nama Pegawai</th><th>Jumlah Pelatihan</th><th>Pelatihan Terakhir</th><th>Tanggal Terakhir</th><th>Aksi</th></tr></thead><tbody>
        ${rows.length ? rows.map((employee) => {
          const history = employeeHistory(employee.id);
          const latest = history[0] || null;
          return `<tr>
            <td><button class="training-name-button" type="button" data-employee-detail="${esc(employee.id)}">${esc(employee.name)}</button></td>
            <td><span class="training-count ${history.length ? '' : 'zero'}">${history.length}</span></td>
            <td>${latest ? `${esc(latest.name)}<span class="training-sub">${esc(latest.organizer || '—')}</span>` : '<span class="training-sub">Belum ada pelatihan pada periode ini</span>'}</td>
            <td class="nowrap">${latest ? esc(formatDateRange(latest)) : '—'}</td>
            <td><button class="btn btn-secondary btn-sm" type="button" data-employee-detail="${esc(employee.id)}">Detail</button></td>
          </tr>`;
        }).join('') : '<tr><td colspan="5"><div class="empty-state"><strong>Pegawai tidak ditemukan</strong>Ubah kata kunci atau rentang bulan yang digunakan.</div></td></tr>'}
      </tbody></table></div>
      <div class="training-period-note">Periode aktif: ${esc(periodLabel())}. Pelatihan yang melintasi batas bulan tetap dihitung apabila tanggalnya beririsan dengan periode.</div>
    </section>`;
  }

  function trainingTable() {
    const rows = filteredTrainings();
    return `<section class="card training-card">
      <div class="training-card-head"><div><h2>Data Pelatihan</h2><p>Satu pelatihan dicatat satu kali dan dapat memiliki beberapa peserta pegawai.</p></div><span class="card-subtitle">${rows.length} pelatihan</span></div>
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
      <div class="training-head"><div><h1>Data Pelatihan</h1><p>Pencatatan dan arsip pelatihan yang telah diikuti oleh pegawai Pusat Data dan Informasi.</p></div><div class="training-head-actions"><a class="btn btn-secondary" href="pelatihan.html#pegawai">Data Pegawai</a><button class="btn btn-primary" type="button" data-action="add-training">+ Tambah Pelatihan</button></div></div>
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
      <div class="training-drawer-head"><div><h3>${esc(training.name || 'Pelatihan')}</h3><p>Detail arsip pelatihan dan daftar pegawai yang mengikuti.</p></div><button class="icon-btn" type="button" data-action="close-drawer" aria-label="Tutup">×</button></div>
      <div class="training-detail-meta">
        <div class="training-detail-box"><div class="training-detail-label">Tanggal</div><div class="training-detail-value">${esc(formatDateRange(training))}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Penyelenggara</div><div class="training-detail-value">${esc(training.organizer || '—')}</div></div>
      </div>
      <div class="training-list-title">Peserta (${participants.length})</div>
      <div class="training-history-list">${participants.length ? participants.map((id) => `<div class="training-person-row"><strong>${esc(employeeName(id))}</strong></div>`).join('') : '<div class="training-participant-empty">Belum ada peserta tercatat.</div>'}</div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-training-edit="${esc(training.id)}">Edit</button><button class="btn btn-danger" type="button" data-training-delete="${esc(training.id)}">Hapus</button></div>
    </aside>`;
  }

  function employeeDrawer(employeeId) {
    const name = employeeName(employeeId) || 'Nama pegawai belum tersedia';
    const history = employeeHistory(employeeId);
    return `<div class="training-drawer-backdrop" data-action="close-drawer"></div><aside class="training-drawer" role="dialog" aria-modal="true" aria-label="Riwayat pelatihan pegawai">
      <div class="training-drawer-head"><div><h3>${esc(name)}</h3><p>Riwayat pelatihan pada ${esc(periodLabel())}.</p></div><button class="icon-btn" type="button" data-action="close-drawer" aria-label="Tutup">×</button></div>
      <div class="training-detail-meta">
        <div class="training-detail-box"><div class="training-detail-label">Jumlah Pelatihan</div><div class="training-detail-value">${history.length}</div></div>
        <div class="training-detail-box"><div class="training-detail-label">Pelatihan Terakhir</div><div class="training-detail-value">${history[0] ? esc(formatDateRange(history[0])) : '—'}</div></div>
      </div>
      <div class="training-list-title">Riwayat Pelatihan</div>
      <div class="training-history-list">${history.length ? history.map((training) => `<div class="training-history-item"><strong>${esc(training.name)}</strong><span>${esc(formatDateRange(training))} · ${esc(training.organizer || '—')}</span></div>`).join('') : '<div class="training-participant-empty">Belum ada pelatihan pada periode ini.</div>'}</div>
    </aside>`;
  }

  function trainingModal() {
    if (!state.modal) return '';
    const training = state.modal.training || {};
    const selected = state.modal.selected || new Set();
    const employees = activeEmployees();
    return `<div class="training-modal-backdrop" data-action="close-modal"><div class="training-modal" role="dialog" aria-modal="true" aria-label="${training.id ? 'Edit' : 'Tambah'} pelatihan" data-modal-panel>
      <div class="training-modal-head"><div><h3>${training.id ? 'Edit Data Pelatihan' : 'Tambah Data Pelatihan'}</h3><p>Isi informasi utama pelatihan dan pilih satu atau beberapa pegawai sebagai peserta.</p></div><button class="icon-btn" type="button" data-action="close-modal" aria-label="Tutup">×</button></div>
      <div class="training-modal-grid">
        <div class="field span-2"><label for="training-name">Nama Pelatihan</label><input id="training-name" type="text" value="${esc(training.name || '')}" autocomplete="off" required></div>
        <div class="field span-2"><label for="training-organizer">Penyelenggara</label><input id="training-organizer" type="text" value="${esc(training.organizer || '')}" autocomplete="off" required></div>
        <div class="field"><label for="training-start-date">Tanggal Mulai</label><input id="training-start-date" type="date" value="${esc(training.startDate || '')}" required></div>
        <div class="field"><label for="training-end-date">Tanggal Selesai <span style="text-transform:none;font-weight:520;color:var(--muted)">(opsional)</span></label><input id="training-end-date" type="date" value="${esc(training.endDate || '')}"></div>
      </div>
      <div class="training-participants">
        <div class="training-participant-head"><strong>Peserta Pelatihan</strong><span><span data-selected-count>${selected.size}</span> pegawai dipilih</span></div>
        <div class="training-participant-search"><div class="search"><input id="participant-search" type="search" placeholder="Cari nama pegawai..." autocomplete="off"></div></div>
        <div class="training-participant-list" data-participant-list>
          ${employees.length ? employees.map((employee) => `<label class="training-participant-option" data-participant-row data-search-name="${esc(employee.name.toLowerCase())}"><input type="checkbox" data-participant-id="${esc(employee.id)}" ${selected.has(employee.id) ? 'checked' : ''}><strong>${esc(employee.name)}</strong><span class="${selected.has(employee.id) ? 'training-selected' : ''}" data-participant-status>${selected.has(employee.id) ? 'Dipilih' : ''}</span></label>`).join('') : '<div class="training-participant-empty">Master pegawai belum tersedia.</div>'}
        </div>
      </div>
      <div class="training-modal-actions"><button class="btn btn-secondary" type="button" data-action="close-modal">Batal</button><button class="btn btn-primary" type="button" data-action="save-training">Simpan</button></div>
    </div></div>`;
  }

  function overlayMarkup() {
    const drawer = state.drawer?.type === 'training'
      ? trainingDrawer(state.trainings.find((item) => item.id === state.drawer.id) || {})
      : state.drawer?.type === 'employee' ? employeeDrawer(state.drawer.id) : '';
    return `${drawer}${trainingModal()}${state.toast ? `<div class="training-toast">${esc(state.toast)}</div>` : ''}`;
  }

  function render() {
    if (state.loading) {
      app.innerHTML = window.AppShell.render({
        module: 'training', view: state.view, viewLabel: state.view === 'pegawai' ? 'Data Pegawai' : 'Data Pelatihan',
        content: '<div class="card training-loading"><strong>Memuat data pelatihan</strong>Membaca master pegawai dan arsip pelatihan dari Cloud Firestore...</div>'
      });
      return;
    }

    const content = `${state.error ? `<div class="alert alert-warning"><div class="alert-title">Data belum dapat dimuat sempurna</div>${esc(state.error)}</div>` : ''}${state.view === 'pegawai' ? employeeView() : trainingView()}`;
    app.innerHTML = window.AppShell.render({
      module: 'training',
      view: state.view,
      viewLabel: state.view === 'pegawai' ? 'Data Pegawai' : 'Data Pelatihan',
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
      state.trainings = trainingSnap.docs.map((doc) => {
        const data = doc.data() || {};
        return {
          ...data,
          id: doc.id,
          name: String(data.name || ''),
          organizer: String(data.organizer || ''),
          startDate: String(data.startDate || ''),
          endDate: String(data.endDate || ''),
          participantIds: [...new Set((Array.isArray(data.participantIds) ? data.participantIds : []).map(String).filter(Boolean))]
        };
      });
    } catch (error) {
      state.error = error.message || String(error);
    } finally {
      state.loading = false;
      render();
    }
  }

  function openAddModal() {
    state.drawer = null;
    state.modal = { training: {}, selected: new Set() };
    render();
  }

  function openEditModal(id) {
    const training = state.trainings.find((item) => item.id === id);
    if (!training) return;
    state.drawer = null;
    state.modal = { training: { ...training }, selected: new Set(training.participantIds || []) };
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
    const participantIds = [...(state.modal.selected || new Set())];

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
        schemaVersion: 1,
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
    const ok = window.confirm(`Hapus arsip pelatihan "${training.name}"? Data peserta pada pelatihan ini juga akan terhapus dari riwayat.`);
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
    }
  });

  app.addEventListener('change', (event) => {
    if (event.target?.id === 'training-start-month' || event.target?.id === 'training-end-month') {
      updateFiltersFromDom();
      return;
    }
    if (event.target?.matches?.('[data-participant-id]') && state.modal) {
      const id = String(event.target.dataset.participantId || '');
      if (event.target.checked) state.modal.selected.add(id);
      else state.modal.selected.delete(id);
      const count = app.querySelector('[data-selected-count]');
      if (count) count.textContent = String(state.modal.selected.size);
      const status = event.target.closest('[data-participant-row]')?.querySelector('[data-participant-status]');
      if (status) {
        status.textContent = event.target.checked ? 'Dipilih' : '';
        status.classList.toggle('training-selected', event.target.checked);
      }
    }
  });

  app.addEventListener('click', (event) => {
    const actionEl = event.target.closest?.('[data-action]');
    const action = actionEl?.dataset.action;

    // The modal backdrop wraps the modal panel. Without this guard, clicking
    // any input inside the modal finds the backdrop via closest('[data-action]')
    // and is incorrectly interpreted as a close-modal action.
    if (action === 'close-modal' && actionEl?.classList.contains('training-modal-backdrop') && event.target !== actionEl) {
      return;
    }
    if (action === 'reset-filter') {
      state.filters = { search: '', startMonth: `${now.getFullYear()}-01`, endMonth: currentMonth };
      render();
      return;
    }
    if (action === 'add-training') { openAddModal(); return; }
    if (action === 'close-modal') { state.modal = null; render(); return; }
    if (action === 'close-drawer') { state.drawer = null; render(); return; }
    if (action === 'save-training') { saveTraining(event.target.closest('[data-action]')); return; }

    const employeeDetail = event.target.closest?.('[data-employee-detail]')?.dataset.employeeDetail;
    if (employeeDetail) { state.drawer = { type: 'employee', id: employeeDetail }; render(); return; }

    const trainingDetail = event.target.closest?.('[data-training-detail]')?.dataset.trainingDetail;
    if (trainingDetail) { state.drawer = { type: 'training', id: trainingDetail }; render(); return; }

    const trainingEdit = event.target.closest?.('[data-training-edit]')?.dataset.trainingEdit;
    if (trainingEdit) { openEditModal(trainingEdit); return; }

    const trainingDelete = event.target.closest?.('[data-training-delete]')?.dataset.trainingDelete;
    if (trainingDelete) { deleteTraining(trainingDelete); }
  });

  window.addEventListener('hashchange', () => {
    state.view = viewFromHash();
    state.filters.search = '';
    state.drawer = null;
    state.modal = null;
    render();
  });

  async function init() {
    await window.FirebaseClient.requireAdmin();
    state.view = viewFromHash();
    if (!location.hash) history.replaceState(null, '', '#pegawai');
    await loadData();
  }

  init();
})();
