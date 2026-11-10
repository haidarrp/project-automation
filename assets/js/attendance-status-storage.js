/* Ringkasan status presensi untuk Monitoring Konflik, terpisah dari nominal keuangan. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceStatusStorage = api;
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';
  const COLLECTION = 'attendanceStatusRecords';
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  const idValid = value => Boolean(value) && !String(value).includes('/');
  const iso = value => datePattern.test(String(value || '')) &&
    !Number.isNaN(Date.parse(`${value}T00:00:00Z`)) &&
    new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
  function prepare(source, runId, ownerUid, employees, recordsFor, coverageMonths = []) {
    const grouped = new Map();
    if (!idValid(runId)) throw new Error('ID riwayat presensi tidak valid.');
    let unmatched = 0;
    for (const employee of employees || []) {
      const employeeId = String(source === 'tukin' ? employee.masterEmployeeId || '' : employee.id || '');
      if (!idValid(employeeId)) { unmatched++; continue; }
      function ensureMonth(year, month) {
        // Pisahkan per proses agar Tukin (periode 11-10) dan Uang Makan
        // (bulan kalender) tidak saling menghapus sebagian hari.
        const docId = `${year}-${String(month).padStart(2, '0')}_${employeeId}_${source}_${runId}`;
        if (!grouped.has(docId)) grouped.set(docId, {employeeId, year, month, source,
          sourceRunId: String(runId || ''), ownerUid: String(ownerUid || ''), days: {}});
        return grouped.get(docId);
      }
      // Buat dokumen kosong untuk bulan yang diproses sehingga status lama
      // tidak bertahan jika semua baris status pada bulan tersebut dihapus.
      coverageMonths.forEach(([year, month]) => ensureMonth(year, month));
      const days = recordsFor(employee) || {};
      for (const [date, record] of Object.entries(days)) {
        if (!iso(date)) continue;
        const status = String(record?.status || record?.workMode || '').trim().slice(0, 160);
        if (!status) continue;
        const year = Number(date.slice(0, 4)), month = Number(date.slice(5, 7));
        ensureMonth(year, month).days[date] = {status};
      }
    }
    return {records: [...grouped.entries()].map(([id, row]) => ({id, ...row})), unmatched};
  }
  function fromTukin(run) {
    const period = run?.period || {};
    const year = Number(period.year), month = Number(period.month);
    const covered = Number.isInteger(year) && Number.isInteger(month) && month >= 1 && month <= 12
      ? [new Date(year, month - 3, 11), new Date(year, month - 2, 10)]
        .map(date => [date.getFullYear(), date.getMonth() + 1]) : [];
    const months = [...new Map(covered.map(pair => [pair.join('-'), pair])).values()];
    return prepare('tukin', run?.id, run?.ownerUid, run?.employees, employee => employee.records, months);
  }
  function fromUangMakan(run, employees) {
    const year = Number(run?.year), month = Number(run?.month);
    const covered = Number.isInteger(year) && Number.isInteger(month) && month >= 1 && month <= 12
      ? [[year, month]] : [];
    return prepare('uangMakan', run?.id, run?.ownerUid, employees, employee => employee.days, covered);
  }
  async function persist(result) {
    const db = window.FirebaseClient?.getDb?.();
    const uid = window.FirebaseClient?.getCurrentUser?.()?.uid;
    if (!db || !uid) throw new Error('Sesi Firebase tidak tersedia untuk sinkronisasi status presensi.');
    const records = result.records || [];
    const updatedAt = new Date().toISOString();
    for (let offset = 0; offset < records.length; offset += 300) {
      const batch = db.batch();
      records.slice(offset, offset + 300).forEach(row => {
        const {id, ...data} = row;
        batch.set(db.collection(COLLECTION).doc(id), {...data, updatedAt, updatedByUid: uid});
      });
      await batch.commit();
    }
    return {written: records.length, unmatched: result.unmatched};
  }
  async function removeForRun(source, runId) {
    const db = window.FirebaseClient?.getDb?.();
    if (!db || !runId || !['tukin', 'uangMakan'].includes(source)) return 0;
    const snapshot = await db.collection(COLLECTION).where('sourceRunId', '==', String(runId)).get();
    const docs = snapshot.docs.filter(doc => doc.data()?.source === source);
    for (let offset = 0; offset < docs.length; offset += 300) {
      const batch = db.batch();
      docs.slice(offset, offset + 300).forEach(doc => batch.delete(doc.ref));
      await batch.commit();
    }
    return docs.length;
  }
  return Object.freeze({COLLECTION, fromTukin, fromUangMakan, removeForRun,
    saveFromTukin: run => persist(fromTukin(run)),
    saveFromUangMakan: (run, employees) => persist(fromUangMakan(run, employees))});
});
