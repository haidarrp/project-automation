/* Register Administrasi Kehadiran: logika murni, tanpa dependensi UI/Firestore. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceRegisterEngine = api;
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';
  const iso = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) && !Number.isNaN(Date.parse(`${v}T00:00:00Z`)) && new Date(`${v}T00:00:00Z`).toISOString().slice(0,10) === v;
  const keyFor = s => `${s.type}:${s.id}:${s.employeeId}`;
  function overlaps(a, b) { return iso(a.startDate) && iso(a.endDate) && iso(b.startDate) && iso(b.endDate) && a.startDate <= b.endDate && b.startDate <= a.endDate; }
  function hash(text) { let n = 2166136261; for (const ch of String(text)) { n ^= ch.charCodeAt(0); n = Math.imul(n, 16777619); } return (n >>> 0).toString(16).padStart(8, '0'); }
  function fingerprint(row) {
    return hash(JSON.stringify([row.type, row.id, row.employeeId, row.startDate, row.endDate, row.number, row.activity, row.updatedAt || '', row.documentDate || '', row.type==='surat'?row.accountable!==false:'']));
  }
  function makeSources(leaves, assignments) {
    const list = [];
    for (const row of leaves || []) {
      const employeeId = String(row.employeeId || '');
      if (!employeeId || !iso(row.startDate) || !iso(row.endDate) || row.endDate < row.startDate) continue;
      const item = { type:'cuti', id:String(row.id || ''), employeeId, startDate:row.startDate, endDate:row.endDate,
        number:String(row.documentNumber || ''), activity:String(row.leaveType || 'Cuti'), documentDate:String(row.documentDate || ''),
        updatedAt:timestamp(row.updatedAt), label:`Cuti ${row.documentNumber || row.leaveType || row.id || ''}` };
      item.key = keyFor(item); item.fingerprint = fingerprint(item); list.push(item);
    }
    for (const row of assignments || []) {
      if (!iso(row.startDate) || !iso(row.endDate) || row.endDate < row.startDate) continue;
      const seen = new Set();
      for (const person of row.employees || []) {
        const employeeId = String(person.employeeId || person.id || '');
        if (!employeeId || seen.has(employeeId)) continue;
        seen.add(employeeId);
        const item = { type:'surat', id:String(row.id || ''), employeeId, startDate:row.startDate, endDate:row.endDate,
          number:String(row.letterNumber || ''), activity:String(row.activity || 'Surat tugas'), documentDate:String(row.letterDate || ''),
          updatedAt:timestamp(row.updatedAt), accountable:row.accountable!==false, label:`ST ${row.letterNumber || row.id || ''}` };
        item.key = keyFor(item); item.fingerprint = fingerprint(item); list.push(item);
      }
    }
    return list;
  }
  function timestamp(v) { if (!v) return ''; if (typeof v.toDate === 'function') return v.toDate().toISOString(); if (v.seconds != null) return `${v.seconds}:${v.nanoseconds || 0}`; return String(v); }
  function conflicts(sources) {
    const byEmployee = new Map(), found = [];
    for (const source of sources || []) {
      if (!byEmployee.has(source.employeeId)) byEmployee.set(source.employeeId, []);
      byEmployee.get(source.employeeId).push(source);
    }
    for (const [employeeId, rows] of byEmployee) {
      rows.sort((a,b) => a.startDate.localeCompare(b.startDate) || a.key.localeCompare(b.key));
      for (let i=0;i<rows.length;i++) for (let j=i+1;j<rows.length;j++) {
        if (rows[j].startDate > rows[i].endDate) break;
        if (!overlaps(rows[i], rows[j])) continue;
        const a=rows[i], b=rows[j];
        const isMixed = a.type !== b.type;
        const kind = isMixed ? 'cuti_surat' : a.type === 'surat' ? 'surat_ganda' : 'cuti_ganda';
        const pair=[a.key,b.key].sort();
        const key=`conflict:${employeeId}:${pair.join('|')}`;
        found.push({ key, kind, employeeId, sources:[a,b], startDate:a.startDate > b.startDate ? a.startDate : b.startDate,
          endDate:a.endDate < b.endDate ? a.endDate : b.endDate,
          fingerprint:hash([a.fingerprint,b.fingerprint].sort().join('|')) });
      }
    }
    return found.sort((a,b)=>b.startDate.localeCompare(a.startDate));
  }
  function inPeriod(source, year, month) {
    const start = `${year}-${String(month).padStart(2,'0')}-01`;
    const last = new Date(Number(year),Number(month),0).getDate();
    const end = `${year}-${String(month).padStart(2,'0')}-${last}`;
    return overlaps(source, {startDate:start,endDate:end});
  }
  function reconciliationFingerprint(sources, records, year, month) {
    const part=(sources || []).filter(s=>inPeriod(s,year,month)).map(s=>`${s.key}/${s.fingerprint}`).sort();
    const active=(records || []).filter(r=>r.startDate && inPeriod(r,year,month)).map(r=>`${r.key}/${r.fingerprint}/${r.workflow || ''}/${r.realization || ''}/${r.selection || ''}`).sort();
    return hash(JSON.stringify([year,month,part,active]));
  }
  // Status pada kolom Keterangan (presensi) dinormalisasi terpisah dari status dokumen.
  const STATUS = Object.freeze(['WFO', 'WFH', 'Libur', 'Dinas', 'Offsite', 'tidak_hadir']);
  function normalizeStatus(value) {
    const status = String(value == null ? '' : value).normalize('NFKC').toLowerCase().trim()
      .replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ');
    if (!status) return '';
    if (/^(off\s?site|kerja di luar kantor)(?:\b|$)/.test(status)) return 'Offsite';
    if (/^(tidak hadir|tanpa keterangan|tanpa kabar|alpa|alpha|absen)(?:\b|$)/.test(status)) return 'tidak_hadir';
    if (/^(libur|tanggal merah|hari libur|holiday)(?:\b|$)/.test(status)) return 'Libur';
    if (/^(dinas|perjalanan dinas|dinas luar|tugas luar)(?:\b|$)/.test(status)) return 'Dinas';
    if (/^(wfh|work from home|bekerja dari rumah)(?:\b|$)/.test(status)) return 'WFH';
    if (/^(wfo|work from office|bekerja dari kantor)(?:\b|$)/.test(status)) return 'WFO';
    return '';
  }
  // Snapshot presensi mempunyai satu dokumen per pegawai/bulan. File yang baru
  // diunggah dalam sesi monitoring didahulukan atas snapshot tersimpan.
  function makeAttendance(snapshotRows = [], importedRows = []) {
    const chosen = new Map();
    function add(row, date, record, priority) {
      const employeeId = String(row.employeeId || '').trim();
      const status = String(record?.status || record?.workMode || '').trim();
      if (!employeeId || !iso(date) || !status) return;
      const normalized = normalizeStatus(status);
      const item = {employeeId, date, startDate: date, endDate: date, status,
        normalizedStatus: normalized, source: row.source === 'tukin' ? 'Proses Tukin' : row.source === 'uangMakan' ? 'Proses Uang Makan' : String(row.source || 'Presensi'),
        updatedAt: String(row.updatedAt || ''), priority};
      const key = `${employeeId}:${date}`, old = chosen.get(key);
      if (!old || priority > old.priority || (priority === old.priority && item.updatedAt > old.updatedAt)) chosen.set(key, item);
    }
    for (const [rows, priority] of [[snapshotRows, 0], [importedRows, 1]]) {
      for (const row of rows || []) {
        if (row?.days && typeof row.days === 'object') {
          for (const [date, record] of Object.entries(row.days)) add(row, date, record, priority);
        } else if (row) add(row, String(row.date || row.key || ''), row, priority);
      }
    }
    return [...chosen.values()].sort((a, b) => b.date.localeCompare(a.date) || a.employeeId.localeCompare(b.employeeId));
  }
  function attendanceIssues(attendance, sources) {
    const documents = new Map();
    for (const doc of sources || []) {
      if (!documents.has(doc.employeeId)) documents.set(doc.employeeId, []);
      documents.get(doc.employeeId).push(doc);
    }
    const result = [];
    for (const day of attendance || []) {
      if (!['Offsite', 'tidak_hadir'].includes(day.normalizedStatus || normalizeStatus(day.status))) continue;
      const matched = (documents.get(day.employeeId) || []).filter(doc =>
        doc.startDate <= day.date && day.date <= doc.endDate);
      if (matched.length) continue;
      const kind = (day.normalizedStatus || normalizeStatus(day.status)) === 'Offsite'
        ? 'offsite_tanpa_dokumen' : 'tidak_hadir_tanpa_dokumen';
      result.push({key:`presensi:${day.employeeId}:${day.date}:${kind}`, kind,
        employeeId:day.employeeId, startDate:day.date, endDate:day.date,
        status:day.status, normalizedStatus:day.normalizedStatus || normalizeStatus(day.status),
        attendanceSource:day.source, sources:[],
        fingerprint:hash([day.employeeId, day.date, day.status, day.updatedAt].join('|'))});
    }
    return result.sort((a, b) => b.startDate.localeCompare(a.startDate) || a.employeeId.localeCompare(b.employeeId));
  }
  function statusCounts(attendance, year, month = 0) {
    const counts = Object.fromEntries(STATUS.map(status => [status, 0]));
    for (const day of attendance || []) {
      if (day.date.slice(0, 4) !== String(year)) continue;
      if (month && Number(day.date.slice(5, 7)) !== Number(month)) continue;
      const status = day.normalizedStatus || normalizeStatus(day.status);
      if (Object.hasOwn(counts, status)) counts[status]++;
    }
    return counts;
  }
  return Object.freeze({iso, overlaps, hash, fingerprint, makeSources, conflicts, inPeriod,
    reconciliationFingerprint, STATUS, normalizeStatus, makeAttendance, attendanceIssues, statusCounts});
});
