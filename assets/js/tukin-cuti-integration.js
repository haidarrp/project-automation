(function () {
  'use strict';

  const rules = window.TukinRules;
  const master = window.MasterDataService;

  function db() {
    const value = window.FirebaseClient?.getDb?.();
    if (!value) throw new Error('Firestore belum siap untuk membaca data Cuti.');
    return value;
  }

  function dateKey(value) {
    if (!value) return '';
    if (value instanceof Date) return rules.dateKey(value);
    return String(value).slice(0, 10);
  }

  function periodYears(period) {
    const range = rules.attendancePeriod(period);
    return [...new Set([range.start.getFullYear(), range.end.getFullYear()])];
  }

  function normalizeLeave(doc) {
    const row = doc.data() || {};
    return {
      id: doc.id,
      employeeId: String(row.employeeId || ''),
      leaveType: String(row.leaveType || 'Cuti'),
      startDate: String(row.startDate || ''),
      endDate: String(row.endDate || row.startDate || ''),
      leaveYear: Number(row.leaveYear || String(row.startDate || '').slice(0, 4) || 0),
      documentNumber: String(row.documentNumber || '')
    };
  }

  async function fetchLeaves(period) {
    const years = periodYears(period);
    const snapshots = await Promise.all(
      years.map((year) => db().collection('leaveRecords').where('leaveYear', '==', year).get())
    );
    const byId = new Map();
    snapshots.forEach((snapshot) => {
      snapshot.docs.forEach((doc) => byId.set(doc.id, normalizeLeave(doc)));
    });

    const range = rules.attendancePeriod(period);
    const start = rules.dateKey(range.start);
    const end = rules.dateKey(range.end);
    return [...byId.values()]
      .filter((leave) => leave.employeeId && leave.startDate && leave.endDate)
      .filter((leave) => leave.startDate <= end && leave.endDate >= start)
      .sort((a, b) => a.startDate.localeCompare(b.startDate) || a.endDate.localeCompare(b.endDate));
  }

  async function resolveMasterEmployeeId(employee) {
    if (employee?.masterEmployeeId) return String(employee.masterEmployeeId);
    if (!master?.findByIdentity) return '';
    const found = await master.findByIdentity({ name: employee?.name || '', nip: employee?.nip || '' });
    if (!found?.id) return '';
    employee.masterEmployeeId = String(found.id);
    return employee.masterEmployeeId;
  }

  function autoCutiAdjustment(record) {
    return record?.adjustmentSource === 'cuti' || record?.cutiAdjustment?.automatic === true;
  }

  function hasProtectedManualAdjustment(record) {
    if (record?.adjustedPercent === null || record?.adjustedPercent === undefined) return false;
    return !autoCutiAdjustment(record);
  }

  function clearCutiAdjustment(record) {
    if (!autoCutiAdjustment(record)) return false;
    record.adjustedPercent = null;
    record.adjustmentNote = '';
    record.adjustmentSource = '';
    record.cutiAdjustment = null;
    return true;
  }

  function leaveNote(leave) {
    const period = leave.startDate === leave.endDate
      ? leave.startDate
      : `${leave.startDate} s.d. ${leave.endDate}`;
    const number = leave.documentNumber ? `, dokumen ${leave.documentNumber}` : '';
    return `Penyesuaian otomatis menjadi 0% berdasarkan data ${leave.leaveType} pada modul Cuti (${period}${number}).`;
  }

  function applyLeaveToRecord(record, leave) {
    record.adjustedPercent = 0;
    record.adjustmentSource = 'cuti';
    record.adjustmentNote = leaveNote(leave);
    record.cutiAdjustment = {
      automatic: true,
      leaveId: leave.id,
      leaveType: leave.leaveType,
      startDate: leave.startDate,
      endDate: leave.endDate,
      documentNumber: leave.documentNumber || ''
    };
  }

  async function apply(employees, period) {
    const list = Array.isArray(employees) ? employees : [];
    const leaves = await fetchLeaves(period);
    const byEmployee = new Map();
    leaves.forEach((leave) => {
      if (!byEmployee.has(leave.employeeId)) byEmployee.set(leave.employeeId, []);
      byEmployee.get(leave.employeeId).push(leave);
    });

    let adjustedRecords = 0;
    let clearedRecords = 0;
    let preservedManualRecords = 0;
    let matchedEmployees = 0;
    let unresolvedEmployees = 0;

    for (const employee of list) {
      const masterEmployeeId = await resolveMasterEmployeeId(employee);
      if (!masterEmployeeId) {
        unresolvedEmployees += 1;
        continue;
      }
      const employeeLeaves = byEmployee.get(masterEmployeeId) || [];
      if (employeeLeaves.length) matchedEmployees += 1;

      for (const record of Object.values(employee.records || {})) {
        const key = dateKey(record.key || record.date);
        const matchingLeave = employeeLeaves.find((leave) => leave.startDate <= key && leave.endDate >= key) || null;

        if (matchingLeave && record.needsVerification) {
          if (hasProtectedManualAdjustment(record)) {
            preservedManualRecords += 1;
            continue;
          }
          applyLeaveToRecord(record, matchingLeave);
          adjustedRecords += 1;
          continue;
        }

        if (clearCutiAdjustment(record)) clearedRecords += 1;
      }
    }

    return {
      status: 'applied',
      leaveRecords: leaves.length,
      adjustedRecords,
      clearedRecords,
      matchedEmployees,
      unresolvedEmployees,
      preservedManualRecords
    };
  }

  window.TukinCutiIntegration = Object.freeze({ apply, fetchLeaves });
})();
