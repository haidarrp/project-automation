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
  return Object.freeze({iso, overlaps, hash, fingerprint, makeSources, conflicts, inPeriod, reconciliationFingerprint});
});
