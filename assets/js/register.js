/* Monitoring konflik kehadiran: hanya membaca Cuti/Surat Tugas, tanpa alur persetujuan. */
(function () {
  'use strict';
  const app = document.getElementById('app');
  const E = window.AttendanceRegisterEngine;
  if (!app || !E) return;

  const state = {
    year: new Date().getFullYear(), month: 0, kind: 'all', search: '',
    loading: true, error: '', employees: new Map(), sources: []
  };
  const KIND = {
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
  const periodText = (start, end) => dateText(start) + (end !== start ? ` – ${dateText(end)}` : '');
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

  function rows() {
    return E.conflicts(state.sources).filter(item => {
      if (state.month > 0) {
        if (!E.inPeriod(item, state.year, state.month)) return false;
      } else if (item.endDate < `${state.year}-01-01` || item.startDate > `${state.year}-12-31`) {
        return false;
      }
      if (state.kind !== 'all' && item.kind !== state.kind) return false;
      const query = state.search.toLocaleLowerCase('id-ID').trim();
      if (!query) return true;
      return [employeeName(item.employeeId), KIND[item.kind], item.startDate, item.endDate,
        ...item.sources.flatMap(source => [source.number, source.activity, source.label])]
        .join(' ').toLocaleLowerCase('id-ID').includes(query);
    });
  }

  function sourceMarkup(source, index) {
    const heading = source.type === 'cuti' ? 'Cuti' : 'Surat Tugas';
    const href = canOpen(source.type) ? `<a class="rg-source-link" href="${esc(sourceUrl(source))}">
      Buka ${esc(labelFor(source.type))} <span aria-hidden="true">↗</span></a>` :
      '<small class="rg-muted">Akses ke menu sumber tidak tersedia</small>';
    return `<div class="rg-source">
      <div class="rg-source-heading"><span class="rg-source-number">${index + 1}</span><b>${esc(heading)}</b></div>
      <strong>${esc(source.number || '(Tanpa nomor dokumen)')}</strong>
      <small>${esc(source.activity || '')}</small>
      <small>${periodText(source.startDate, source.endDate)}</small>
      ${href}
    </div>`;
  }

  function render() {
    const current = rows();
    const all = E.conflicts(state.sources).filter(item =>
      state.month > 0 ? E.inPeriod(item, state.year, state.month)
        : item.startDate <= `${state.year}-12-31` && item.endDate >= `${state.year}-01-01`);
    const yearOptions = [state.year - 2, state.year - 1, state.year, state.year + 1];
    const months = ['Semua bulan', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const controls = `<div class="rg-filters">
      <label>Tahun <select id="rg-year">${yearOptions.map(year =>
        `<option value="${year}" ${year === state.year ? 'selected' : ''}>${year}</option>`).join('')}</select></label>
      <label>Bulan <select id="rg-month">${months.map((name, month) =>
        `<option value="${month}" ${month === state.month ? 'selected' : ''}>${name}</option>`).join('')}</select></label>
      <label>Jenis irisan <select id="rg-kind"><option value="all">Semua jenis</option>${Object.entries(KIND).map(([key, label]) =>
        `<option value="${key}" ${key === state.kind ? 'selected' : ''}>${esc(label)}</option>`).join('')}</select></label>
      <label class="rg-search-label">Cari pegawai atau nomor dokumen <input type="search" id="rg-search" value="${esc(state.search)}" placeholder="Nama, nomor surat, jenis cuti…"></label>
      <button type="button" class="btn btn-secondary" data-refresh>Perbarui data</button>
    </div>`;
    const table = `<div class="rg-table-wrap"><table class="rg-table"><thead><tr>
      <th>Pegawai</th><th>Jenis irisan</th><th>Tanggal beririsan</th><th>Dokumen terkait dan tindakan</th>
      </tr></thead><tbody>${current.length ? current.map(row => `<tr>
        <td><strong>${esc(employeeName(row.employeeId))}</strong></td>
        <td><span class="rg-kind">${esc(KIND[row.kind] || row.kind)}</span></td>
        <td class="rg-date">${periodText(row.startDate, row.endDate)}</td>
        <td><div class="rg-source-list">${row.sources.map(sourceMarkup).join('')}</div></td>
      </tr>`).join('') : `<tr><td colspan="4" class="rg-empty">${state.sources.length ?
        'Tidak ada konflik atau irisan yang sesuai dengan filter.' :
        'Tidak ada dokumen Cuti atau Surat Tugas pada periode yang diperiksa.'}</td></tr>`}</tbody></table></div>`;

    const content = `<div class="rg">
      <div class="rg-lead"><span class="rg-overline">PUSDATIN · TATA USAHA</span>
        <h1>Monitoring Konflik Kehadiran</h1>
        <p>Menampilkan potensi irisan antara dokumen Cuti dan Surat Tugas. Tidak ada verifikasi, rekonsiliasi, atau perubahan dokumen dari menu ini.</p></div>
      <div class="rg-notice"><strong>Tindak lanjut langsung di dokumen sumber.</strong>
        Jika terdapat data yang perlu dikoreksi, buka Data Cuti atau Data Surat Tugas melalui tautan pada baris konflik.
        Daftar akan menyesuaikan setelah perubahan disimpan dan data monitoring diperbarui.</div>
      ${state.error ? `<div class="rg-alert rg-alert-warn">${esc(state.error)}</div>` : ''}
      <section class="card rg-card">
        <div class="rg-summary"><strong>${state.loading ? '…' : all.length}</strong>
          <div><b>Potensi konflik pada periode terpilih</b><span>${state.loading ? 'Memuat data…' :
            `${all.filter(row => row.kind === 'cuti_surat').length} Cuti–Surat Tugas · ${all.filter(row => row.kind === 'surat_ganda').length} Surat Tugas ganda · ${all.filter(row => row.kind === 'cuti_ganda').length} Cuti ganda`}</span></div>
        </div>
        ${controls}
        ${state.loading ? '<div class="rg-empty">Memuat data Cuti dan Surat Tugas…</div>' :
          state.error ? '<div class="rg-empty">Data tidak berhasil dimuat. Pilih Perbarui data untuk mencoba kembali.</div>' : table}
      </section>
      <p class="rg-footnote">Irisan tanggal adalah indikasi administratif, bukan penetapan bahwa dokumen tidak sah. Penggunaan dokumen untuk Tukin, Uang Makan, dan Lembur mengikuti data serta aturan pada modul masing-masing.</p>
    </div>`;
    app.innerHTML = window.AppShell.render({
      module: 'register', view: 'cases', viewLabel: 'Monitoring Konflik', content
    });
  }

  async function load() {
    state.loading = true;
    state.error = '';
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
    } catch (error) {
      state.error = `Data tidak dapat dimuat: ${error.message || error}. Periksa koneksi dan izin baca dokumen.`;
      state.sources = [];
    } finally {
      state.loading = false;
      render();
    }
  }

  app.addEventListener('click', event => {
    if (event.target.closest('[data-refresh]')) load();
  });
  app.addEventListener('change', event => {
    if (event.target.id === 'rg-year') {state.year = Number(event.target.value); load();}
    if (event.target.id === 'rg-month') {state.month = Number(event.target.value); render();}
    if (event.target.id === 'rg-kind') {state.kind = event.target.value; render();}
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
