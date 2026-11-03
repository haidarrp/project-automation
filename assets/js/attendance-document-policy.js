/* Kebijakan dokumen keuangan: membaca sumber Cuti/Surat Tugas secara langsung.
 * Register sekarang hanya memonitor irisan. Pilihan kasus lama tidak lagi berlaku.
 */
(function (root, factory) {
  const engine = root?.AttendanceRegisterEngine || (typeof require === 'function' ? require('./register-engine.js') : null);
  const api = factory(engine);
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceDocumentPolicy = api;
})(typeof window !== 'undefined' ? window : null, function (E) {
  'use strict';
  function build(leaves = [], assignments = []) {
    if (!E) throw new Error('Mesin pemeriksaan dokumen belum tersedia.');
    const sources = E.makeSources(leaves, assignments);
    const currentAssignments = new Map(assignments.map(row => [String(row.id), row]));
    const signature = E.hash(JSON.stringify(sources.map(source =>
      `${source.key}/${source.fingerprint}`).sort()));
    return {
      // Field kompatibilitas untuk modul Tukin, Uang Makan, dan Lembur.
      excluded: new Set(), stale: [], sources, signature,
      allows() { return true; },
      selectionReason() { return ''; },
      isAccountable(assignment) {
        return (currentAssignments.get(String(assignment?.id)) || assignment)?.accountable !== false;
      },
      exportable(type, record) {
        return type !== 'surat' || this.isAccountable(record);
      },
      reason(type, record) {
        return type === 'surat' && !this.isAccountable(record)
          ? 'Surat Tugas tidak dipertanggungjawabkan' : '';
      }
    };
  }
  async function fetch(years) {
    const db = window.FirebaseClient?.getDb?.();
    if (!db) throw new Error('Firestore belum siap untuk memuat dokumen.');
    const validYears = [...new Set(years.map(Number).filter(Number.isInteger))];
    const [leaves, assignments] = await Promise.all([
      Promise.all(validYears.map(year => db.collection('leaveRecords').where('leaveYear', '==', year).get())),
      Promise.all(validYears.map(year => db.collection('assignmentRecords').where('assignmentYear', '==', year).get()))
    ]);
    const unpack = snapshots => snapshots.flatMap(snapshot => snapshot.docs.map(doc =>
      ({...doc.data(), id: doc.id})));
    return build(unpack(leaves), unpack(assignments));
  }
  return Object.freeze({build, fetch});
});
