/* Monitoring konflik dokumen dan status presensi; seluruh tindakan bersifat baca-saja. */
(function () {
  'use strict';
  const app = document.getElementById('app');
  const E = window.AttendanceRegisterEngine;
  if (!app || !E) return;

  const state = {
    year: new Date().getFullYear(), month: 0, kind: 'all', search: '',
    loading: true, error: '', attendanceError: '', importMessage: '', importError: '',
    importing: false, employees: new Map(), sources: [], snapshots: [], imported: [],
    attendance: [], issues: []
  };
  const KIND = {
    offsite_tanpa_dokumen: 'Offsite tanpa Cuti / Surat Tugas',
    tidak_hadir_tanpa_dokumen: 'Tidak Hadir tanpa Cuti / Surat Tugas',
    cuti_surat: 'Cuti dan Surat Tugas',
    surat_ganda: 'Surat Tugas ganda',
    cuti_ganda: 'Cuti ganda'
  };
  const esc = value => String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const dateText = value => E.iso(value)
    ? new Date(`${value}T00:00:00Z`).toLocaleDateString('id-ID', {
      day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC'
    }) : '-';
  const periodText = (start, end) => dateText(start) + (end !== start ? ` - ${dateText(end)}` : '');
  const employeeName = id => state.employees.get(String(id))?.name
    || state.employees.get(String(id))?.fullName || String(id);
  const moduleFor = type => type === 'cuti' ? 'cuti' : 'surat';
  const labelFor = type => type === 'cuti' ? 'Data Cuti' : 'Data Surat Tugas';
  const canOpen = type => window.UserAccess?.canView?.(moduleFor(type), 'data') === true;

  function sourceUrl(source) {
    const page = source.type === 'cuti' ? 'cuti.html' : 'surat-tugas.html';
    const year = source.startDate.slice(0, 4);
    return `${page}?record=${encodeURIComponent(source.id)}&year=${encodeURIComponent(year)}#data`;
  }
  function rebuild() {
    state.attendance = E.makeAttendance(state.snapshots, state.imported);
    state.issues = [...E.conflicts(state.sources), ...E.attendanceIssues(state.attendance, state.sources)]
      .sort((a, b) => b.startDate.localeCompare(a.startDate) || a.employeeId.localeCompare(b.employeeId));
  }
  function inSelection(item) {
    if (state.month > 0) return E.inPeriod(item, state.year, state.month);
    return item.startDate <= `${state.year}-12-31` && item.endDate >= `${state.year}-01-01`;
  }
  function rows() {
    const query = state.search.toLocaleLowerCase('id-ID').trim();
    return state.issues.filter(item => {
      if (!inSelection(item)) return false;
      if (state.kind !== 'all' && item.kind !== state.kind) return false;
      if (!query) return true;
      return [employeeName(item.employeeId), KIND[item.kind], item.startDate, item.endDate,
        item.status, item.attendanceSource,
        ...item.sources.flatMap(source => [source.number, source.activity, source.label])]
        .join(' ').toLocaleLowerCase('id-ID').includes(query);
    });
  }
  function sourceMarkup(source, index) {
    const heading = source.type === 'cuti' ? 'Cuti' : 'Surat Tugas';
    const href = canOpen(source.type) ? `<a class="rg-source-link" href="${esc(sourceUrl(source))}">
      Buka ${esc(labelFor(source.type))} <span aria-hidden="true">&#8599;</span></a>` :
      '<small class="rg-muted">Akses ke menu sumber tidak tersedia</small>';
    return `<div class="rg-source">
      <div class="rg-source-heading"><span class="rg-source-number">${index + 1}</span><b>${esc(heading)}</b></div>
      <strong>${esc(source.number || '(Tanpa nomor dokumen)')}</strong>
      <small>${esc(source.activity || '')}</small>
      <small>${periodText(source.startDate, source.endDate)}</small>
      ${href}
    </div>`;
  }
  function attendanceStatus(row) {
    if (row.status) return `<span class="rg-presence ${row.kind.startsWith('offsite_') || row.kind.startsWith('tidak_hadir_') ? 'rg-presence-warn' : ''}">
      ${esc(row.normalizedStatus || row.status)}</span><small class="rg-status-detail">${esc(row.attendanceSource || '')}</small>`;
    const matching = state.attendance.filter(day => day.employeeId === row.employeeId
      && day.date >= row.startDate && day.date <= row.endDate);
    if (!matching.length) return '<span class="rg-muted">Tidak ada data presensi</span>';
    const statusLabels = [...new Set(matching.map(day => day.normalizedStatus || day.status).filter(Boolean))];
    return `<span class="rg-presence">${esc(statusLabels.slice(0, 4).join(', '))}${statusLabels.length > 4 ? ', ...' : ''}</span>`;
  }
  function missingDocumentMarkup(row) {
    const links = [['cuti', 'cuti.html#data', 'Periksa Data Cuti'],
      ['surat', 'surat-tugas.html#data', 'Periksa Data Surat Tugas']]
      .filter(([kind]) => canOpen(kind))
      .map(([, url, label]) => `<a class="rg-source-link" href="${esc(url)}">${esc(label)}</a>`).join('');
    return `<div class="rg-missing"><strong>Dokumen pendukung belum ditemukan</strong>
      <small>Belum terdapat Cuti atau Surat Tugas untuk pegawai dan tanggal yang sama.</small>
      <small>Periksa kelengkapan data sumber dan kesesuaian status presensi. Tanda ini bukan keputusan ketidakhadiran yang sah/tidak sah.</small>
      ${links ? `<div class="rg-missing-links">${links}</div>` : ''}</div>`;
  }
  function render() {
    const current = rows();
    const all = state.issues.filter(inSelection);
    const counts = E.statusCounts(state.attendance, state.year, state.month);
    const statusCount = Object.values(counts).reduce((sum, value) => sum + value, 0);
    const yearOptions = [state.year - 2, state.year - 1, state.year, state.year + 1];
    const months = ['Semua bulan', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const controls = `<div class="rg-filters">
      <label>Tahun <select id="rg-year">${yearOptions.map(year =>
        `<option value="${year}" ${year === state.year ? 'selected' : ''}>${year}</option>`).join('')}</select></label>
      <label>Bulan <select id="rg-month">${months.map((name, month) =>
        `<option value="${month}" ${month === state.month ? 'selected' : ''}>${name}</option>`).join('')}</select></label>
      <label>Jenis temuan <select id="rg-kind"><option value="all">Semua jenis</option>${Object.entries(KIND).map(([key, label]) =>
        `<option value="${key}" ${key === state.kind ? 'selected' : ''}>${esc(label)}</option>`).join('')}</select></label>
      <label class="rg-search-label">Cari pegawai atau nomor dokumen <input type="search" id="rg-search" value="${esc(state.search)}" placeholder="Nama, status, nomor surat, jenis cuti..."></label>
      <button type="button" class="btn btn-secondary" data-refresh>Perbarui data</button>
    </div>`;
    const upload = `<div class="rg-import"><div><b>Impor Excel presensi (opsional)</b>
      <small>Gunakan berkas laporan presensi XLS/XLSX untuk memeriksa data lama yang belum tersinkron. File diperiksa pada perangkat dan tidak disimpan oleh menu ini.</small></div>
      <label class="btn btn-secondary rg-file-label" for="rg-files">${state.importing ? 'Membaca berkas...' : 'Pilih berkas presensi'}
        <input type="file" id="rg-files" multiple accept=".xls,.xlsx" ${state.importing ? 'disabled' : ''}></label></div>`;
    const stats = `<div class="rg-status-summary"><b>Status presensi terbaca (${statusCount} tanggal-pegawai):</b>
      ${E.STATUS.map(status => `<span class="rg-stat"><strong>${esc(status === 'tidak_hadir' ? 'Tidak Hadir' : status)}</strong> ${counts[status]}</span>`).join('')}</div>`;
    const table = `<div class="rg-table-wrap"><table class="rg-table"><thead><tr>
      <th>Pegawai</th><th>Jenis temuan</th><th>Status presensi</th><th>Tanggal</th><th>Dokumen terkait dan tindak lanjut</th>
      </tr></thead><tbody>${current.length ? current.map(row => `<tr>
        <td><strong>${esc(employeeName(row.employeeId))}</strong></td>
        <td><span class="rg-kind ${row.sources.length ? '' : 'rg-kind-warn'}">${esc(KIND[row.kind] || row.kind)}</span></td>
        <td>${attendanceStatus(row)}</td>
        <td class="rg-date">${periodText(row.startDate, row.endDate)}</td>
        <td>${row.sources.length ? `<div class="rg-source-list">${row.sources.map(sourceMarkup).join('')}</div>` : missingDocumentMarkup(row)}</td>
      </tr>`).join('') : `<tr><td colspan="5" class="rg-empty">${state.sources.length || state.attendance.length ?
        'Tidak ada konflik atau status presensi tanpa dokumen yang sesuai dengan filter.' :
        'Belum ada dokumen Cuti / Surat Tugas atau data presensi pada periode yang diperiksa.'}</td></tr>`}</tbody></table></div>`;

    const content = `<div class="rg">
      <div class="rg-lead"><span class="rg-overline">PUSDATIN - TATA USAHA</span>
        <h1>Monitoring Konflik Kehadiran</h1>
        <p>Memeriksa irisan Cuti dan Surat Tugas serta status WFO, WFH, Libur, Dinas, Offsite, dan Tidak Hadir. Offsite atau Tidak Hadir tanpa dokumen Cuti / Surat Tugas pada tanggal yang sama ditandai untuk pemeriksaan administrasi.</p></div>
      <div class="rg-notice"><strong>Tindak lanjut pada dokumen sumber.</strong>
        Pemantauan ini tidak mengubah absensi, Cuti, maupun Surat Tugas. Untuk status Offsite / Tidak Hadir, pastikan kesesuaian tanggal, pegawai, dan kelengkapan dokumen.</div>
      ${state.error ? `<div class="rg-alert rg-alert-warn">${esc(state.error)}</div>` : ''}
      ${state.attendanceError ? `<div class="rg-alert rg-alert-warn">${esc(state.attendanceError)}</div>` : ''}
      ${state.importError ? `<div class="rg-alert rg-alert-warn">${esc(state.importError)}</div>` : ''}
      ${state.importMessage ? `<div class="rg-alert rg-alert-ok">${esc(state.importMessage)}</div>` : ''}
      <section class="card rg-card">
        <div class="rg-summary"><strong>${state.loading ? '...' : all.length}</strong>
          <div><b>Potensi temuan pada periode terpilih</b><span>${state.loading ? 'Memuat data...' :
            `${all.filter(row => row.kind === 'offsite_tanpa_dokumen').length} Offsite tanpa dokumen, ${all.filter(row => row.kind === 'tidak_hadir_tanpa_dokumen').length} Tidak Hadir tanpa dokumen; ${all.filter(row => row.kind === 'cuti_surat').length} Cuti-Surat Tugas, ${all.filter(row => row.kind === 'surat_ganda').length} Surat Tugas ganda, ${all.filter(row => row.kind === 'cuti_ganda').length} Cuti ganda`}</span></div>
        </div>
        ${controls}
        ${upload}
        ${stats}
        ${state.loading ? '<div class="rg-empty">Memuat data Cuti, Surat Tugas, dan status presensi...</div>' :
          state.error ? '<div class="rg-empty">Data dokumen tidak berhasil dimuat. Pilih Perbarui data untuk mencoba kembali.</div>' : table}
      </section>
      <p class="rg-footnote">Status presensi otomatis bersumber dari ringkasan proses Tukin / Uang Makan yang tersimpan setelah pembaruan aplikasi ini, atau dari Excel yang diimpor pada sesi aktif. WFO, WFH, Libur, dan Dinas dibaca sebagai konteks, bukan otomatis dianggap bermasalah. Penanda hanya mengindikasikan perlunya pemeriksaan data, bukan penetapan ketidakabsahan presensi.</p>
    </div>`;
    app.innerHTML = window.AppShell.render({
      module: 'register', view: 'cases', viewLabel: 'Monitoring Konflik', content
    });
  }

  async function load() {
    state.loading = true;
    state.error = '';
    state.attendanceError = '';
    render();
    try {
      const years = [state.year - 1, state.year, state.year + 1];
      const db = window.FirebaseClient.getDb();
      const [bundle, leaveSnapshot, assignmentSnapshot] = await Promise.all([
        window.MasterDataService.getBundle(false),
        db.collection('leaveRecords').where('leaveYear', 'in', years).get(),
        db.collection('assignmentRecords').where('assignmentYear', 'in', years).get()
      ]);
      state.employees = new Map(bundle.employees.map(person => [String(person.id),
        {...person, ...(bundle.directory.get(String(person.id)) || {})}]));
      state.sources = E.makeSources(
        leaveSnapshot.docs.map(doc => ({...doc.data(), id: doc.id})),
        assignmentSnapshot.docs.map(doc => ({...doc.data(), id: doc.id}))
      );
      try {
        const attendanceSnapshot = await db.collection('attendanceStatusRecords')
          .where('year', '==', state.year).get();
        state.snapshots = attendanceSnapshot.docs.map(doc => ({...doc.data(), id: doc.id}));
      } catch (error) {
        state.snapshots = [];
        state.attendanceError = 'Status presensi tersimpan belum dapat dibaca. Periksa izin Firestore untuk attendanceStatusRecords atau impor file Excel. Detail: '
          + (error.message || error);
      }
      rebuild();
    } catch (error) {
      state.error = `Data tidak dapat dimuat: ${error.message || error}. Periksa koneksi dan izin baca dokumen.`;
      state.sources = [];
      state.snapshots = [];
      state.issues = [];
      state.attendance = [];
    } finally {
      state.loading = false;
      render();
    }
  }
  async function importAttendance(files) {
    if (!files?.length || state.importing) return;
    state.importing = true;
    state.importError = '';
    state.importMessage = '';
    render();
    try {
      await window.ResourceLoader.ensure('xlsx');
      const result = await window.AttendanceImport.parseFiles(files, state.year);
      if (result.rows.length) {
        state.imported.push(...result.rows);
        rebuild();
        state.importMessage = `${result.names.length} berkas berhasil dibaca, ${result.rows.length} status tanggal-pegawai ditemukan untuk tahun ${state.year}.`;
      }
      if (result.errors.length) state.importError = result.errors.join(' | ');
      if (!result.rows.length && !result.errors.length) state.importError = 'Tidak ada data presensi yang dapat dibaca.';
    } catch (error) {
      state.importError = `Impor presensi gagal: ${error.message || error}`;
    } finally {
      state.importing = false;
      render();
    }
  }
  app.addEventListener('click', event => {
    if (event.target.closest('[data-refresh]')) load();
  });
  app.addEventListener('change', event => {
    if (event.target.id === 'rg-year') {
      state.year = Number(event.target.value);
      state.imported = [];
      state.importMessage = '';
      state.importError = '';
      load();
    }
    if (event.target.id === 'rg-month') {state.month = Number(event.target.value); render();}
    if (event.target.id === 'rg-kind') {state.kind = event.target.value; render();}
    if (event.target.id === 'rg-files') importAttendance(Array.from(event.target.files || []));
  });
  app.addEventListener('input', event => {
    if (event.target.id !== 'rg-search') return;
    state.search = event.target.value;
    const selection = event.target.selectionStart;
    render();
    const field = document.getElementById('rg-search');
    if (field) {
      field.focus();
      if (selection != null && field.setSelectionRange) field.setSelectionRange(selection, selection);
    }
  });
  (async () => {
    try {
      if (!await window.UserAccess.requireModule('register')) return;
      await load();
    } catch (error) {
      state.loading = false;
      state.error = error.message || String(error);
      render();
    }
  })();
})();
