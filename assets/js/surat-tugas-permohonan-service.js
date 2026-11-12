(function (root, factory) {
  'use strict';
  const service = factory();
  if (typeof module === 'object' && module.exports) module.exports = service;
  if (root) root.STChangeRequests = service;
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';

  function normalizedHeader(value) {
    return String(value ?? '').toLowerCase().normalize('NFKD')
      .replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
  }
  const COLUMN_NAMES = {
    id: ['id', 'responseid'],
    start: ['starttime', 'waktumulai'],
    completed: ['completiontime', 'waktuselesai', 'waktupengiriman'],
    email: ['email', 'alamatemail'],
    account: ['name', 'namaakun'],
    name: ['nama', 'namapegawai', 'namapemohon'],
    type: ['jenis', 'jenispermohonan'],
    letter: ['nomorsuratperihal', 'nomorsurat', 'perihal'],
    description: ['deksripsi', 'deskripsi', 'keterangan', 'uraian'],
    status: ['status', 'statuspermohonan']
  };

  function toExcelDate(value) {
    if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 90000) {
      const date = new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000));
      return { display: `${String(date.getUTCDate()).padStart(2,'0')}/${String(date.getUTCMonth()+1).padStart(2,'0')}/${date.getUTCFullYear()} ${String(date.getUTCHours()).padStart(2,'0')}:${String(date.getUTCMinutes()).padStart(2,'0')}`, sort: date.getTime() };
    }
    const raw = String(value ?? '').trim();
    if (!raw) return { display: '', sort: 0 };
    // Excel Online dapat mengirimkan string ISO, teks Excel, atau nomor seri.
    const numeric = Number(raw);
    if (/^\d{5}(\.\d+)?$/.test(raw)) return toExcelDate(numeric);
    const parsed = Date.parse(raw);
    return { display: raw, sort: Number.isNaN(parsed) ? 0 : parsed };
  }

  function pickHeader(values) {
    for (let rowIndex = 0; rowIndex < Math.min(values.length, 40); rowIndex += 1) {
      const cells = values[rowIndex] || [];
      const header = {};
      for (const [field, aliases] of Object.entries(COLUMN_NAMES)) {
        const index = cells.findIndex(v => aliases.includes(normalizedHeader(v)));
        if (index >= 0) header[field] = index;
      }
      if (Number.isInteger(header.id) && Number.isInteger(header.completed)
          && Number.isInteger(header.name) && Number.isInteger(header.letter)) {
        return { rowIndex, header };
      }
    }
    return null;
  }

  function parseWorksheet(range) {
    const values = Array.isArray(range?.values) ? range.values : [];
    const display = Array.isArray(range?.text) ? range.text : [];
    const layout = pickHeader(values);
    if (!layout) return null;
    const rows = [];
    const ids = new Set();
    for (let i = layout.rowIndex + 1; i < values.length; i += 1) {
      const row = values[i] || [];
      const formatted = display[i] || [];
      const field = name => row[layout.header[name]];
      const str = name => String(field(name) ?? '').trim();
      const id = str('id');
      const name = str('name');
      const letter = str('letter');
      // Draft respons Microsoft Forms (tidak ada Completion time) bukan tiket masuk.
      if (!id || !str('completed') || !name || !letter || ids.has(id)) continue;
      ids.add(id);
      const completedValue = field('completed');
      const completed = toExcelDate(completedValue);
      const status = str('status');
      rows.push({
        id, name, letter, type: str('type') || 'Permohonan',
        description: str('description'), status: status || 'Belum ditangani',
        completed: completed.display || String(formatted[layout.header.completed] || ''),
        completedSort: completed.sort,
        email: str('email').toLowerCase() === 'anonymous' ? '' : str('email'),
        account: str('account')
      });
    }
    return rows.sort((a,b) => b.completedSort - a.completedSort || Number(b.id) - Number(a.id) || b.id.localeCompare(a.id));
  }

  function shareIdFromUrl(url) {
    const value = String(url || '').trim();
    if (!/^https:\/\//i.test(value)) throw new Error('URL Excel Microsoft Forms belum dikonfigurasi.');
    const bytes = new TextEncoder().encode(value);
    let str = '';
    for (const byte of bytes) str += String.fromCharCode(byte);
    return `u!${btoa(str).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_')}`;
  }

  function graphError(status, code, detail) {
    if (status === 401) return 'Sesi Microsoft kedaluwarsa. Hubungkan kembali akun Microsoft lalu coba Refresh.';
    if (status === 403) return 'Akses ke Excel ditolak (403). Pastikan akun Microsoft Anda memiliki akses ke file dan izin Files.ReadWrite telah disetujui.';
    if (status === 404) return 'File Excel tidak ditemukan (404). Periksa URL SharePoint dan apakah tautannya masih aktif.';
    if (status === 429) return 'Layanan Microsoft sedang membatasi permintaan (429). Coba Refresh kembali nanti.';
    return `Microsoft Graph ${status}${code ? ` (${code})` : ''}: ${detail || 'Permintaan gagal'}`;
  }

  // Convert worksheet-relative column offsets to an Excel A1 reference.
  function columnLetters(index) {
    if (!Number.isInteger(index) || index < 0 || index >= 16384) throw new Error('Kolom Excel tidak valid.');
    let value = index + 1;
    let letters = '';
    while (value) {
      value -= 1;
      letters = String.fromCharCode(65 + value % 26) + letters;
      value = Math.floor(value / 26);
    }
    return letters;
  }

  function columnNumber(letters) {
    return String(letters).toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  }

  function findStatusCell(range, responseId) {
    const values = Array.isArray(range?.values) ? range.values : [];
    const layout = pickHeader(values);
    if (!layout) throw new Error('Kolom Id, Completion time, Nama, atau Nomor Surat/Perihal tidak ditemukan.');
    if (!Number.isInteger(layout.header.status)) throw new Error('Kolom Status tidak ditemukan pada worksheet Excel.');
    // usedRange may begin anywhere, not necessarily A1; rowIndex/columnIndex are 0-based.
    const addr = String(range?.address || '').match(/!\$?([A-Z]{1,3})\$?(\d+)/i);
    const startRow = Number.isInteger(range?.rowIndex) && range.rowIndex >= 0
      ? range.rowIndex : addr ? Number(addr[2]) - 1 : null;
    const startCol = Number.isInteger(range?.columnIndex) && range.columnIndex >= 0
      ? range.columnIndex : addr ? columnNumber(addr[1]) : null;
    if (startRow === null || startCol === null) {
      throw new Error('Posisi baris Excel tidak tersedia dari Microsoft Graph. Perubahan dibatalkan untuk mencegah salah sel.');
    }
    const matches = [];
    for (let i = layout.rowIndex + 1; i < values.length; i += 1) {
      const row = values[i] || [];
      if (String(row[layout.header.id] ?? '').trim() !== String(responseId).trim()) continue;
      if (!String(row[layout.header.completed] ?? '').trim()) continue;
      matches.push({
        address: `${columnLetters(startCol + layout.header.status)}${startRow + i + 1}`,
        status: String(row[layout.header.status] ?? '').trim()
      });
    }
    if (matches.length !== 1) throw new Error(matches.length ? 'Terdapat duplikasi ID respons di Excel; pembaruan dibatalkan.' : 'ID respons tidak ditemukan di Excel. Klik Refresh sebelum mencoba kembali.');
    return matches[0];
  }

  async function graphClient() {
    const auth = window.MicrosoftAuth;
    if (!auth?.acquireToken) throw new Error('Autentikasi Microsoft tidak tersedia. Jalankan npm install dan npm run build sebelum deployment.');
    const token = await auth.acquireToken(true);
    const apiBase = String(window.MICROSOFT_CONFIG?.graphBaseUrl || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
    async function request(method, path, payload, extraHeaders = {}) {
      const response = await fetch(`${apiBase}${path}`, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          ...(payload === undefined ? {} : { 'Content-Type': 'application/json' }),
          ...extraHeaders
        },
        ...(payload === undefined ? {} : { body: JSON.stringify(payload) })
      });
      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(graphError(response.status, body?.error?.code, body?.error?.message));
      }
      return response.status === 204 ? {} : response.json();
    }
    return {
      get(path, headers = {}) { return request('GET', path, undefined, headers); },
      patch(path, data) { return request('PATCH', path, data); }
    };
  }

  async function workbookContext(client) {
    const cfg = window.ST_CHANGE_REQUESTS_CONFIG || {};
    const share = shareIdFromUrl(cfg.shareUrl);
    const item = await client.get(`/shares/${share}/driveItem?$select=id,name,webUrl,file,parentReference`, { Prefer: 'redeemSharingLink' });
    const driveId = item?.remoteItem?.parentReference?.driveId || item?.parentReference?.driveId;
    const itemId = item?.remoteItem?.id || item?.id;
    if (!driveId || !itemId) throw new Error('Microsoft Graph tidak memberikan driveId/itemId untuk file Excel ini. Pastikan tautan mengarah ke file Excel yang dapat diakses.');
    const prefix = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;
    const sheets = await client.get(`${prefix}/worksheets?$select=id,name`);
    if (!Array.isArray(sheets?.value) || !sheets.value.length) throw new Error('Workbook Excel tidak mempunyai worksheet yang dapat dibaca.');
    return { cfg, item, prefix, sheets: sheets.value };
  }

  async function findRequestSheet(client, book) {
    for (const sheet of book.sheets) {
      const range = await client.get(`${book.prefix}/worksheets/${encodeURIComponent(sheet.id)}/usedRange(valuesOnly=true)`);
      const records = parseWorksheet(range);
      if (records !== null) return { sheet, range, records };
    }
    throw new Error('Kolom Excel tidak sesuai. Diperlukan Id, Completion time, Nama, dan Nomor Surat/ Perihal.');
  }

  async function fetchRequests() {
    const client = await graphClient();
    const book = await workbookContext(client);
    const { sheet, records } = await findRequestSheet(client, book);
    return { rows: records, worksheet: sheet.name || '', fileName: book.item.name || 'Respons Microsoft Forms', webUrl: book.item.webUrl || book.cfg.shareUrl };
  }

  // Called only on explicit user confirmation. Re-read the latest range and resolve
  // the target from Forms' response ID, never from the old table's row position.
  async function completeRequest(responseId) {
    const id = String(responseId || '').trim();
    if (!id) throw new Error('ID respons permohonan tidak valid.');
    const client = await graphClient();
    const book = await workbookContext(client);
    const { sheet, range } = await findRequestSheet(client, book);
    const cell = findStatusCell(range, id);
    if (/^(selesai)$/i.test(cell.status)) return { id, status: 'Selesai', alreadyDone: true };
    if (/^(ditolak|dibatalkan|batal|closed|resolved)$/i.test(cell.status)) {
      throw new Error(`Status permohonan ini telah berubah menjadi "${cell.status}" di Excel. Klik Refresh untuk memperbarui tampilan.`);
    }
    const encodedAddress = encodeURIComponent(cell.address);
    const path = `${book.prefix}/worksheets/${encodeURIComponent(sheet.id)}/range(address='${encodedAddress}')`;
    const result = await client.patch(path, { values: [['Selesai']] });
    if (Array.isArray(result?.values) && String(result.values?.[0]?.[0] || '').trim() !== 'Selesai') {
      throw new Error('Microsoft Graph tidak mengonfirmasi nilai Selesai pada sel Status. Periksa Excel lalu klik Refresh.');
    }
    return { id, status: 'Selesai', alreadyDone: false, address: cell.address };
  }

  return Object.freeze({ fetchRequests, completeRequest, parseWorksheet, findStatusCell, toExcelDate, shareIdFromUrl, normalizedHeader });
});
