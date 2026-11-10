/* Impor presensi baca-saja untuk Monitoring Konflik (tidak menyimpan file). */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceImport = api;
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';
  const pad = value => String(value).padStart(2, '0');
  function key(y, m, d) {
    const date = new Date(y, m - 1, d, 12);
    return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d
      ? `${y}-${pad(m)}-${pad(d)}` : '';
  }
  function dateKey(raw, displayed) {
    if (raw instanceof Date && !Number.isNaN(raw.getTime()))
      return key(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
    if (typeof raw === 'number' && Number.isFinite(raw)) {
      const result = window.XLSX?.SSF?.parse_date_code(raw);
      return result ? key(result.y, result.m, result.d) : '';
    }
    const value = String(displayed || raw || '').trim();
    let match = value.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)$/);
    if (match) return key(+match[1], +match[2], +match[3]);
    match = value.match(/^([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})$/);
    return match ? key(+match[3], +match[2], +match[1]) : '';
  }
  function identity(value, fileName) {
    const full = String(value || '').trim();
    let name = full, nip = '';
    const match = full.match(/^(.*?)\s*[-\u2013]\s*(\d{8,})\s*$/) || full.match(/^(.*?)\s+(\d{8,})\s*$/);
    if (match) {name = match[1].trim(); nip = match[2];}
    if (!name) name = String(fileName || '').replace(/\.(xlsx?|xls)$/i, '')
      .replace(/^laporan_presensi_/i, '').replace(/_\d{8}_\d{8}$/i, '').replace(/_/g, ' ').trim();
    return {name, nip};
  }
  async function parseFile(file, year) {
    if (!window.XLSX) throw new Error('Pustaka Excel belum tersedia.');
    if (!/\.xlsx?$/i.test(file?.name || '') && !/\.xls$/i.test(file?.name || ''))
      throw new Error('Format berkas harus XLS atau XLSX.');
    const workbook = window.XLSX.read(await file.arrayBuffer(), {type:'array', cellDates:true, cellText:true});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error('Sheet presensi tidak ditemukan.');
    const raw = window.XLSX.utils.sheet_to_json(sheet, {header:1, raw:true, defval:''});
    const display = window.XLSX.utils.sheet_to_json(sheet, {header:1, raw:false, defval:''});
    const headerIndex = display.slice(0, 20).findIndex(row => {
      const labels = (row || []).map(value => String(value || '').trim().toLowerCase());
      return labels.includes('tanggal') && labels.includes('masuk') && labels.includes('keluar');
    });
    if (headerIndex < 0) throw new Error('Header Tanggal/Masuk/Keluar tidak ditemukan.');
    const header = display[headerIndex].map(value => String(value || '').trim().toLowerCase());
    const dateColumn = header.indexOf('tanggal');
    const statusColumn = ['status', 'keterangan', 'status presensi', 'status kehadiran']
      .map(label => header.indexOf(label)).find(index => index >= 0) ?? 21;
    const identityText = String(sheet.A2?.w ?? sheet.A2?.v ?? display[1]?.[0] ?? '');
    const person = identity(identityText, file.name);
    if (!person.name) throw new Error('Nama pegawai tidak ditemukan pada baris identitas.');
    const master = await window.MasterDataService.findByIdentity(person);
    if (!master?.id) throw new Error(`Pegawai ${person.name} belum cocok dengan Master Pegawai.`);
    const rows = [];
    for (let index = headerIndex + 1; index < raw.length; index++) {
      const displayed = display[index] || [], values = raw[index] || [];
      if (String(displayed[0] || '').trim().toUpperCase() === 'TOTAL') break;
      const date = dateKey(values[dateColumn], displayed[dateColumn]);
      if (!date || Number(date.slice(0, 4)) !== Number(year)) continue;
      const status = String(displayed[statusColumn] || values[statusColumn] || '').trim().slice(0, 160);
      if (!status) continue;
      rows.push({employeeId:String(master.id), date, status, source:`Impor Excel: ${file.name}`,
        updatedAt:new Date().toISOString()});
    }
    if (!rows.length) throw new Error(`Tidak ada status presensi tahun ${year} yang dapat dibaca.`);
    return rows;
  }
  async function parseFiles(files, year) {
    const rows = [], errors = [], names = [];
    for (const file of Array.from(files || [])) {
      try {rows.push(...await parseFile(file, year)); names.push(file.name);}
      catch (error) {errors.push(`${file.name || 'Berkas'}: ${error.message || error}`);}
    }
    return {rows, errors, names};
  }
  return Object.freeze({identity, dateKey, parseFile, parseFiles});
});
