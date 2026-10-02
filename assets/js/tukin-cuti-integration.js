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

  function normalizeLeaveDocument(raw, index) {
    const row = raw && typeof raw === 'object' ? raw : {};
    const file = row.file && typeof row.file === 'object' ? { ...row.file } : null;
    if (file) delete file.file;
    return {
      id: String(row.id || `doc-${index + 1}`),
      type: String(row.type || 'Dokumen Cuti'),
      name: String(row.name || ''),
      number: String(row.number || ''),
      note: String(row.note || ''),
      file
    };
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
      documentNumber: String(row.documentNumber || ''),
      documents: (Array.isArray(row.documents) ? row.documents : []).map(normalizeLeaveDocument)
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

  function isAutoCutiEvidence(item) {
    return item?.cutiEvidence?.automatic === true || String(item?.id || '').startsWith('cuti-evidence-');
  }

  function cutiEvidenceId(leave, document, index) {
    const leaveId = String(leave?.id || 'cuti').replace(/[^A-Za-z0-9_-]/g, '-');
    const documentId = String(document?.id || `doc-${index + 1}`).replace(/[^A-Za-z0-9_-]/g, '-');
    return `cuti-evidence-${leaveId}-${documentId}`;
  }

  function cutiEvidenceFromDocument(leave, document, index) {
    const source = document?.file;
    if (!source?.driveId || !source?.itemId) return null;
    const name = String(source.name || source.remoteName || document.name || `Bukti Cuti ${index + 1}`);
    return {
      id: cutiEvidenceId(leave, document, index),
      name,
      size: Number(source.size || source.sourceSize || 0),
      type: String(source.type || ''),
      lastModified: Number(source.lastModified || source.sourceLastModified || 0),
      file: null,
      provider: 'cuti',
      // Sebelum file disalin ke folder run Tukin, tautan ini tetap membuka sumber CUTI.
      webUrl: String(source.webUrl || ''),
      localOnly: true,
      cutiEvidence: {
        automatic: true,
        copiedToTukin: false,
        leaveId: String(leave.id || ''),
        documentId: String(document.id || ''),
        documentType: String(document.type || ''),
        documentName: String(document.name || name),
        documentNumber: String(document.number || ''),
        sourceDriveId: String(source.driveId || ''),
        sourceItemId: String(source.itemId || ''),
        sourceWebUrl: String(source.webUrl || ''),
        sourceName: name,
        sourceETag: String(source.eTag || ''),
        sourceCTag: String(source.cTag || '')
      }
    };
  }

  function sameCutiSource(existing, desired) {
    const current = existing?.cutiEvidence || {};
    const next = desired?.cutiEvidence || {};
    if (!current.sourceDriveId || !current.sourceItemId) return false;
    if (current.sourceDriveId !== next.sourceDriveId || current.sourceItemId !== next.sourceItemId) return false;
    if (current.sourceETag && next.sourceETag && current.sourceETag !== next.sourceETag) return false;
    if (current.sourceCTag && next.sourceCTag && current.sourceCTag !== next.sourceCTag) return false;
    return true;
  }

  function mergeCutiEvidence(record, leave) {
    const evidence = Array.isArray(record.evidence) ? record.evidence : [];
    const manualEvidence = evidence.filter((item) => !isAutoCutiEvidence(item));
    const existingAuto = evidence.filter(isAutoCutiEvidence);
    const desiredBySource = new Map();
    (leave.documents || [])
      .map((document, index) => cutiEvidenceFromDocument(leave, document, index))
      .filter(Boolean)
      .forEach((item) => {
        const meta = item.cutiEvidence || {};
        const sourceKey = meta.sourceDriveId && meta.sourceItemId
          ? `${meta.sourceDriveId}:${meta.sourceItemId}`
          : item.id;
        if (!desiredBySource.has(sourceKey)) desiredBySource.set(sourceKey, item);
      });
    const desired = [...desiredBySource.values()];
    const removed = [];
    const nextAuto = [];

    for (const target of desired) {
      const existing = existingAuto.find((item) => item.id === target.id) || null;
      if (existing && sameCutiSource(existing, target)) {
        const copied = Boolean(existing?.driveId && existing?.itemId && existing?.cutiEvidence?.copiedToTukin);
        nextAuto.push({
          ...existing,
          name: target.name || existing.name,
          size: Number(target.size || existing.size || 0),
          type: target.type || existing.type || '',
          lastModified: Number(target.lastModified || existing.lastModified || 0),
          cutiEvidence: {
            ...existing.cutiEvidence,
            ...target.cutiEvidence,
            copiedToTukin: copied
          }
        });
      } else {
        if (existing?.driveId && existing?.itemId && existing?.cutiEvidence?.copiedToTukin) removed.push(existing);
        nextAuto.push(target);
      }
    }

    const desiredIds = new Set(desired.map((item) => item.id));
    for (const existing of existingAuto) {
      if (desiredIds.has(existing.id)) continue;
      if (existing?.driveId && existing?.itemId && existing?.cutiEvidence?.copiedToTukin) removed.push(existing);
    }

    record.evidence = [...manualEvidence, ...nextAuto];
    return { linked: nextAuto.length, removed };
  }

  function detachCutiEvidence(record) {
    const evidence = Array.isArray(record.evidence) ? record.evidence : [];
    const removed = evidence.filter((item) => isAutoCutiEvidence(item) && item?.driveId && item?.itemId && item?.cutiEvidence?.copiedToTukin);
    record.evidence = evidence.filter((item) => !isAutoCutiEvidence(item));
    return removed;
  }

  function clearCutiAdjustment(record) {
    if (!autoCutiAdjustment(record)) return { cleared: false, removed: [] };
    const removed = detachCutiEvidence(record);
    record.adjustedPercent = null;
    record.adjustmentNote = '';
    record.adjustmentSource = '';
    record.cutiAdjustment = null;
    return { cleared: true, removed };
  }

  function leaveNote(leave) {
    const period = leave.startDate === leave.endDate
      ? leave.startDate
      : `${leave.startDate} s.d. ${leave.endDate}`;
    const number = leave.documentNumber ? `, dokumen ${leave.documentNumber}` : '';
    return `Penyesuaian otomatis menjadi 0% berdasarkan data ${leave.leaveType} pada modul Cuti (${period}${number}).`;
  }

  function applyLeaveToRecord(record, leave, includeEvidence) {
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

    // Bukti dukung cuti hanya ditempelkan pada satu tanggal kanonis dalam satu
    // rentang cuti. Seluruh tanggal tetap memperoleh penyesuaian 0%, tetapi file
    // yang sama tidak lagi ikut tersalin berulang kali ke folder/ZIP Tukin.
    if (includeEvidence) return mergeCutiEvidence(record, leave);
    return { linked: 0, removed: detachCutiEvidence(record) };
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
    let evidenceFiles = 0;
    const removedEvidence = [];

    for (const employee of list) {
      const masterEmployeeId = await resolveMasterEmployeeId(employee);
      if (!masterEmployeeId) {
        unresolvedEmployees += 1;
        continue;
      }
      const employeeLeaves = byEmployee.get(masterEmployeeId) || [];
      if (employeeLeaves.length) matchedEmployees += 1;

      const records = Object.values(employee.records || {}).sort((a, b) => {
        return dateKey(a.key || a.date).localeCompare(dateKey(b.key || b.date));
      });

      // Tentukan satu record pemilik bukti untuk setiap leave record. Record paling
      // awal dalam rentang yang memang perlu diverifikasi dipilih sebagai pemilik.
      // Ini sekaligus memigrasikan data lama yang sebelumnya menempelkan bukti yang
      // sama pada setiap tanggal cuti.
      const evidenceOwnerByLeaveId = new Map();
      for (const leave of employeeLeaves) {
        const owner = records.find((record) => {
          const key = dateKey(record.key || record.date);
          return key >= leave.startDate && key <= leave.endDate && record.needsVerification && !hasProtectedManualAdjustment(record);
        }) || null;
        if (owner) evidenceOwnerByLeaveId.set(leave.id, owner);
      }

      for (const record of records) {
        const key = dateKey(record.key || record.date);
        const matchingLeave = employeeLeaves.find((leave) => leave.startDate <= key && leave.endDate >= key) || null;

        if (matchingLeave && record.needsVerification) {
          if (hasProtectedManualAdjustment(record)) {
            preservedManualRecords += 1;
            continue;
          }
          const includeEvidence = evidenceOwnerByLeaveId.get(matchingLeave.id) === record;
          const evidenceResult = applyLeaveToRecord(record, matchingLeave, includeEvidence);
          evidenceFiles += Number(evidenceResult.linked || 0);
          removedEvidence.push(...(evidenceResult.removed || []));
          adjustedRecords += 1;
          continue;
        }

        const clearResult = clearCutiAdjustment(record);
        if (clearResult.cleared) clearedRecords += 1;
        removedEvidence.push(...(clearResult.removed || []));
      }
    }

    return {
      status: 'applied',
      leaveRecords: leaves.length,
      adjustedRecords,
      clearedRecords,
      matchedEmployees,
      unresolvedEmployees,
      preservedManualRecords,
      evidenceFiles,
      removedEvidence
    };
  }

  window.TukinCutiIntegration = Object.freeze({ apply, fetchLeaves });
})();
