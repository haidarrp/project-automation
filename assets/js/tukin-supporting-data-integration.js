(function () {
  'use strict';

  const rules = window.TukinRules;
  const master = window.MasterDataService;
  const ASSIGNMENT_TYPES = Object.freeze({
    official_travel: 'Dinas Luar / Perjalanan Dinas',
    work_arrangement: 'Pengaturan Kerja',
    other: 'Lainnya'
  });

  function normalizeAssignmentType(value, legacyAdjustmentType) {
    const key = String(value || '').trim();
    if (Object.prototype.hasOwnProperty.call(ASSIGNMENT_TYPES, key)) return key;
    if (!key && String(legacyAdjustmentType || '') === 'official_duty') return 'official_travel';
    return 'official_travel';
  }

  function assignmentTypeLabel(value) {
    return ASSIGNMENT_TYPES[normalizeAssignmentType(value)] || ASSIGNMENT_TYPES.official_travel;
  }

  function assignmentUsesAutoZero(assignment) {
    return normalizeAssignmentType(assignment?.assignmentType, assignment?.adjustmentType) === 'official_travel';
  }

  function db() {
    const value = window.FirebaseClient?.getDb?.();
    if (!value) throw new Error('Firestore belum siap untuk membaca data pendukung Tukin.');
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

  function normalizeDocument(raw, index, fallbackType) {
    const row = raw && typeof raw === 'object' ? raw : {};
    const file = row.file && typeof row.file === 'object' ? { ...row.file } : null;
    if (file) delete file.file;
    return {
      id: String(row.id || `doc-${index + 1}`),
      type: String(row.type || fallbackType || 'Dokumen'),
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
      documents: (Array.isArray(row.documents) ? row.documents : []).map((item, index) => normalizeDocument(item, index, 'Dokumen Cuti'))
    };
  }

  function normalizeAssignment(doc) {
    const row = doc.data() || {};
    return {
      id: doc.id,
      letterNumber: String(row.letterNumber || ''),
      letterDate: String(row.letterDate || ''),
      startDate: String(row.startDate || ''),
      endDate: String(row.endDate || row.startDate || ''),
      assignmentYear: Number(row.assignmentYear || String(row.startDate || '').slice(0, 4) || 0),
      activity: String(row.activity || ''),
      location: String(row.location || ''),
      assignmentType: normalizeAssignmentType(row.assignmentType, row.adjustmentType),
      adjustmentType: String(row.adjustmentType || ''),
      employeeIds: (Array.isArray(row.employees) ? row.employees : [])
        .map((employee) => String(employee?.employeeId || employee?.id || ''))
        .filter(Boolean),
      documents: (Array.isArray(row.documents) ? row.documents : []).map((item, index) => normalizeDocument(item, index, 'Surat Tugas'))
    };
  }

  async function fetchLeaves(period) {
    const years = periodYears(period);
    const snapshots = await Promise.all(years.map((year) => db().collection('leaveRecords').where('leaveYear', '==', year).get()));
    const byId = new Map();
    snapshots.forEach((snapshot) => snapshot.docs.forEach((doc) => byId.set(doc.id, normalizeLeave(doc))));
    const range = rules.attendancePeriod(period);
    const start = rules.dateKey(range.start);
    const end = rules.dateKey(range.end);
    return [...byId.values()]
      .filter((leave) => leave.employeeId && leave.startDate && leave.endDate)
      .filter((leave) => leave.startDate <= end && leave.endDate >= start)
      .sort((a, b) => a.startDate.localeCompare(b.startDate) || a.endDate.localeCompare(b.endDate));
  }

  async function fetchAssignments(period) {
    const years = periodYears(period);
    const snapshots = await Promise.all(years.map((year) => db().collection('assignmentRecords').where('assignmentYear', '==', year).get()));
    const byId = new Map();
    snapshots.forEach((snapshot) => snapshot.docs.forEach((doc) => byId.set(doc.id, normalizeAssignment(doc))));
    const range = rules.attendancePeriod(period);
    const start = rules.dateKey(range.start);
    const end = rules.dateKey(range.end);
    return [...byId.values()]
      .filter((assignment) => assignment.employeeIds.length && assignment.startDate && assignment.endDate)
      .filter((assignment) => assignment.startDate <= end && assignment.endDate >= start)
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

  function isAutomaticAdjustment(record) {
    const source = String(record?.adjustmentSource || '');
    if (source === 'manual') return false;
    return ['cuti', 'surat_tugas', 'conflict'].includes(source) ||
      record?.cutiAdjustment?.adjustmentApplied === true ||
      (record?.cutiAdjustment?.automatic === true && source === 'cuti') ||
      record?.assignmentAdjustment?.adjustmentApplied === true ||
      (record?.assignmentAdjustment?.automatic === true && source === 'surat_tugas') ||
      record?.adjustmentConflict?.automatic === true;
  }

  function hasAutomaticSupport(record) {
    return Boolean(
      record?.cutiAdjustment?.automatic === true ||
      record?.assignmentAdjustment?.linkedAutomatically === true ||
      record?.assignmentAdjustment?.automatic === true ||
      record?.adjustmentConflict?.automatic === true ||
      ['cuti', 'surat_tugas', 'conflict'].includes(String(record?.adjustmentSource || ''))
    );
  }

  function hasProtectedManualAdjustment(record) {
    if (record?.adjustedPercent === null || record?.adjustedPercent === undefined) return false;
    return !isAutomaticAdjustment(record);
  }

  function evidenceMeta(item) {
    if (item?.cutiEvidence?.automatic) return { type: 'cuti', meta: item.cutiEvidence };
    if (item?.assignmentEvidence?.automatic) return { type: 'surat_tugas', meta: item.assignmentEvidence };
    return null;
  }

  function isAutomaticEvidence(item) {
    return Boolean(evidenceMeta(item));
  }

  function evidenceSourceKey(item) {
    const info = evidenceMeta(item);
    if (!info) return '';
    const meta = info.meta || {};
    if (meta.sourceDriveId && meta.sourceItemId) return `source:${meta.sourceDriveId}:${meta.sourceItemId}`;
    if (info.type === 'cuti' && meta.leaveId && meta.documentId) return `leave:${meta.leaveId}:${meta.documentId}`;
    if (info.type === 'surat_tugas' && meta.assignmentId && meta.documentId) return `assignment:${meta.assignmentId}:${meta.documentId}`;
    return item?.id ? `id:${item.id}` : '';
  }

  function sourceEvidenceId(prefix, sourceId, documentId, index) {
    const safeSource = String(sourceId || prefix).replace(/[^A-Za-z0-9_-]/g, '-');
    const safeDocument = String(documentId || `doc-${index + 1}`).replace(/[^A-Za-z0-9_-]/g, '-');
    return `${prefix}-evidence-${safeSource}-${safeDocument}`;
  }

  function cutiEvidenceFromDocument(leave, document, index) {
    const source = document?.file;
    if (!source?.driveId || !source?.itemId) return null;
    const name = String(source.name || source.remoteName || document.name || `Bukti Cuti ${index + 1}`);
    return {
      id: sourceEvidenceId('cuti', leave.id, document.id, index),
      name,
      size: Number(source.size || source.sourceSize || 0),
      type: String(source.type || ''),
      lastModified: Number(source.lastModified || source.sourceLastModified || 0),
      file: null,
      provider: 'cuti',
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

  function assignmentEvidenceFromDocument(assignment, document, index) {
    const source = document?.file;
    if (!source?.driveId || !source?.itemId) return null;
    const name = String(source.name || source.remoteName || document.name || `Surat Tugas ${index + 1}`);
    return {
      id: sourceEvidenceId('assignment', assignment.id, document.id, index),
      name,
      size: Number(source.size || source.sourceSize || 0),
      type: String(source.type || ''),
      lastModified: Number(source.lastModified || source.sourceLastModified || 0),
      file: null,
      provider: 'assignment',
      webUrl: String(source.webUrl || ''),
      localOnly: true,
      assignmentEvidence: {
        automatic: true,
        copiedToTukin: false,
        assignmentId: String(assignment.id || ''),
        documentId: String(document.id || ''),
        letterNumber: String(assignment.letterNumber || ''),
        activity: String(assignment.activity || ''),
        assignmentType: normalizeAssignmentType(assignment.assignmentType, assignment.adjustmentType),
        assignmentTypeLabel: assignmentTypeLabel(assignment.assignmentType),
        autoZero: assignmentUsesAutoZero(assignment),
        startDate: String(assignment.startDate || ''),
        endDate: String(assignment.endDate || ''),
        sourceDriveId: String(source.driveId || ''),
        sourceItemId: String(source.itemId || ''),
        sourceWebUrl: String(source.webUrl || ''),
        sourceName: name,
        sourceETag: String(source.eTag || ''),
        sourceCTag: String(source.cTag || '')
      }
    };
  }

  function sameSource(existing, desired) {
    const current = evidenceMeta(existing)?.meta || {};
    const next = evidenceMeta(desired)?.meta || {};
    if (!current.sourceDriveId || !current.sourceItemId) return false;
    if (current.sourceDriveId !== next.sourceDriveId || current.sourceItemId !== next.sourceItemId) return false;
    if (current.sourceETag && next.sourceETag && current.sourceETag !== next.sourceETag) return false;
    if (current.sourceCTag && next.sourceCTag && current.sourceCTag !== next.sourceCTag) return false;
    return true;
  }

  function mergeAutomaticEvidence(record, desiredItems) {
    const evidence = Array.isArray(record.evidence) ? record.evidence : [];
    const manualEvidence = evidence.filter((item) => !isAutomaticEvidence(item));
    const existingAuto = evidence.filter(isAutomaticEvidence);
    const desiredBySource = new Map();
    (desiredItems || []).filter(Boolean).forEach((item) => {
      const key = evidenceSourceKey(item) || item.id;
      if (!desiredBySource.has(key)) desiredBySource.set(key, item);
    });
    const desired = [...desiredBySource.values()];
    const nextAuto = [];
    const removed = [];

    for (const target of desired) {
      const targetKey = evidenceSourceKey(target);
      const existing = existingAuto.find((item) => item.id === target.id || (targetKey && evidenceSourceKey(item) === targetKey)) || null;
      if (existing && sameSource(existing, target)) {
        const currentInfo = evidenceMeta(existing);
        const targetInfo = evidenceMeta(target);
        const copied = Boolean(existing?.driveId && existing?.itemId && currentInfo?.meta?.copiedToTukin);
        const merged = {
          ...existing,
          name: target.name || existing.name,
          size: Number(target.size || existing.size || 0),
          type: target.type || existing.type || '',
          lastModified: Number(target.lastModified || existing.lastModified || 0)
        };
        if (targetInfo?.type === 'cuti') merged.cutiEvidence = { ...existing.cutiEvidence, ...target.cutiEvidence, copiedToTukin: copied };
        if (targetInfo?.type === 'surat_tugas') merged.assignmentEvidence = { ...existing.assignmentEvidence, ...target.assignmentEvidence, copiedToTukin: copied };
        nextAuto.push(merged);
      } else {
        if (existing?.driveId && existing?.itemId && evidenceMeta(existing)?.meta?.copiedToTukin) removed.push(existing);
        nextAuto.push(target);
      }
    }

    const desiredKeys = new Set(desired.map((item) => evidenceSourceKey(item) || item.id));
    for (const existing of existingAuto) {
      const key = evidenceSourceKey(existing) || existing.id;
      if (desiredKeys.has(key)) continue;
      if (existing?.driveId && existing?.itemId && evidenceMeta(existing)?.meta?.copiedToTukin) removed.push(existing);
    }

    record.evidence = [...manualEvidence, ...nextAuto];
    return { linked: nextAuto.length, removed };
  }

  function detachAutomaticEvidence(record) {
    const evidence = Array.isArray(record.evidence) ? record.evidence : [];
    const removed = evidence.filter((item) => isAutomaticEvidence(item) && item?.driveId && item?.itemId && evidenceMeta(item)?.meta?.copiedToTukin);
    record.evidence = evidence.filter((item) => !isAutomaticEvidence(item));
    return removed;
  }

  function clearAutomaticAdjustment(record) {
    if (!hasAutomaticSupport(record)) return { cleared: false, removed: [] };
    const removed = detachAutomaticEvidence(record);
    const protectedManual = hasProtectedManualAdjustment(record);
    if (!protectedManual) {
      if (isAutomaticAdjustment(record)) record.adjustedPercent = null;
      record.adjustmentNote = '';
      if (record.adjustmentSource !== 'manual') record.adjustmentSource = '';
    }
    record.cutiAdjustment = null;
    record.assignmentAdjustment = null;
    record.adjustmentConflict = null;
    return { cleared: true, removed };
  }

  function leaveMeta(leave, options) {
    const opts = options || {};
    return {
      automatic: true,
      linkedAutomatically: true,
      adjustmentApplied: Boolean(opts.adjustmentApplied),
      informational: Boolean(opts.informational),
      leaveId: leave.id,
      leaveType: leave.leaveType,
      startDate: leave.startDate,
      endDate: leave.endDate,
      documentNumber: leave.documentNumber || ''
    };
  }

  function assignmentMeta(assignments, options) {
    const opts = options || {};
    return {
      automatic: Boolean(opts.adjustmentApplied),
      linkedAutomatically: true,
      adjustmentApplied: Boolean(opts.adjustmentApplied),
      informational: Boolean(opts.informational),
      assignments: (assignments || []).map((assignment) => ({
        assignmentId: assignment.id,
        letterNumber: assignment.letterNumber || '',
        letterDate: assignment.letterDate || '',
        activity: assignment.activity || '',
        location: assignment.location || '',
        assignmentType: normalizeAssignmentType(assignment.assignmentType, assignment.adjustmentType),
        assignmentTypeLabel: assignmentTypeLabel(assignment.assignmentType),
        autoZero: assignmentUsesAutoZero(assignment),
        startDate: assignment.startDate,
        endDate: assignment.endDate
      }))
    };
  }

  function leaveNote(leave) {
    const period = leave.startDate === leave.endDate ? leave.startDate : `${leave.startDate} s.d. ${leave.endDate}`;
    const number = leave.documentNumber ? `, dokumen ${leave.documentNumber}` : '';
    return `Penyesuaian otomatis menjadi 0% berdasarkan data ${leave.leaveType} pada modul Cuti (${period}${number}).`;
  }

  function leaveInformationNote(leave) {
    const period = leave.startDate === leave.endDate ? leave.startDate : `${leave.startDate} s.d. ${leave.endDate}`;
    const number = leave.documentNumber ? `, dokumen ${leave.documentNumber}` : '';
    return `Data ${leave.leaveType} pada modul Cuti (${period}${number}) tersinkron sebagai data pendukung. Hasil presensi otomatis pada tanggal ini sudah 0%, sehingga tidak diperlukan penyesuaian persentase.`;
  }

  function assignmentNote(assignments) {
    const items = assignments || [];
    if (!items.length) return 'Penyesuaian otomatis menjadi 0% berdasarkan Surat Tugas Dinas Luar / Perjalanan Dinas.';
    const numbers = items.map((item) => item.letterNumber || item.id).filter(Boolean).join(', ');
    const first = items[0];
    const period = first.startDate === first.endDate ? first.startDate : `${first.startDate} s.d. ${first.endDate}`;
    return `Penyesuaian otomatis menjadi 0% berdasarkan Surat Tugas ${numbers} (${period}) dengan jenis Dinas Luar / Perjalanan Dinas.`;
  }

  function assignmentAutoZeroInformationNote(assignments) {
    const items = (assignments || []).filter(assignmentUsesAutoZero);
    if (!items.length) return 'Data Surat Tugas tersinkron sebagai data pendukung.';
    const numbers = items.map((item) => item.letterNumber || item.id).filter(Boolean).join(', ');
    const first = items[0];
    const period = first.startDate === first.endDate ? first.startDate : `${first.startDate} s.d. ${first.endDate}`;
    return `Surat Tugas ${numbers} (${period}) dengan jenis Dinas Luar / Perjalanan Dinas tersinkron sebagai data pendukung. Hasil presensi otomatis pada tanggal ini sudah 0%, sehingga tidak diperlukan penyesuaian persentase.`;
  }

  function assignmentInformationNote(assignments) {
    const items = assignments || [];
    if (!items.length) return 'Ditemukan data Surat Tugas. Jenis penugasan ini tidak mengubah persentase secara otomatis.';
    const numbers = items.map((item) => item.letterNumber || item.id).filter(Boolean).join(', ');
    const types = [...new Set(items.map((item) => assignmentTypeLabel(item.assignmentType)))].join(', ');
    const first = items[0];
    const period = first.startDate === first.endDate ? first.startDate : `${first.startDate} s.d. ${first.endDate}`;
    return `Ditemukan Surat Tugas ${numbers} (${period}) dengan jenis ${types}. Tidak ada penyesuaian persentase otomatis; koreksi dapat dilakukan manual pada Proses Tukin bila diperlukan.`;
  }

  function conflictNote(leave, assignments) {
    const assignmentNumbers = (assignments || []).map((item) => item.letterNumber || item.id).filter(Boolean).join(', ');
    return `Konflik data: tanggal ini tercatat sebagai ${leave.leaveType || 'Cuti'} pada modul Cuti dan juga tercakup Surat Tugas ${assignmentNumbers || ''}. Penyesuaian otomatis tidak diterapkan dan perlu diverifikasi.`;
  }

  function desiredLeaveEvidence(leave, include) {
    if (!include) return [];
    return (leave.documents || []).map((document, index) => cutiEvidenceFromDocument(leave, document, index)).filter(Boolean);
  }

  function desiredAssignmentEvidence(assignments) {
    const desired = [];
    (assignments || []).forEach((assignment) => {
      (assignment.documents || []).forEach((document, index) => {
        const item = assignmentEvidenceFromDocument(assignment, document, index);
        if (item) desired.push(item);
      });
    });
    return desired;
  }

  function applyLeave(record, leave, includeEvidence) {
    record.adjustedPercent = 0;
    record.adjustmentSource = 'cuti';
    record.adjustmentNote = leaveNote(leave);
    record.cutiAdjustment = leaveMeta(leave, { adjustmentApplied: true, informational: false });
    record.assignmentAdjustment = null;
    record.adjustmentConflict = null;
    return mergeAutomaticEvidence(record, desiredLeaveEvidence(leave, includeEvidence));
  }

  function applyLeaveInformation(record, leave, includeEvidence) {
    record.adjustedPercent = null;
    record.adjustmentSource = '';
    record.adjustmentNote = leaveInformationNote(leave);
    record.cutiAdjustment = leaveMeta(leave, { adjustmentApplied: false, informational: true });
    record.assignmentAdjustment = null;
    record.adjustmentConflict = null;
    return mergeAutomaticEvidence(record, desiredLeaveEvidence(leave, includeEvidence));
  }

  function applyAssignments(record, assignments, evidenceAssignments) {
    record.adjustedPercent = 0;
    record.adjustmentSource = 'surat_tugas';
    record.adjustmentNote = assignmentNote(assignments.filter(assignmentUsesAutoZero));
    record.cutiAdjustment = null;
    record.assignmentAdjustment = assignmentMeta(assignments, { adjustmentApplied: true, informational: false });
    record.adjustmentConflict = null;
    return mergeAutomaticEvidence(record, desiredAssignmentEvidence(evidenceAssignments));
  }

  function applyAssignmentSupport(record, assignments, evidenceAssignments) {
    record.adjustedPercent = null;
    record.adjustmentSource = '';
    record.adjustmentNote = assignmentAutoZeroInformationNote(assignments);
    record.cutiAdjustment = null;
    record.assignmentAdjustment = assignmentMeta(assignments, { adjustmentApplied: false, informational: true });
    record.adjustmentConflict = null;
    return mergeAutomaticEvidence(record, desiredAssignmentEvidence(evidenceAssignments));
  }

  function applyAssignmentInformation(record, assignments, evidenceAssignments) {
    record.adjustedPercent = null;
    record.adjustmentSource = '';
    record.adjustmentNote = assignmentInformationNote(assignments);
    record.cutiAdjustment = null;
    record.assignmentAdjustment = assignmentMeta(assignments, { adjustmentApplied: false, informational: true });
    record.adjustmentConflict = null;
    return mergeAutomaticEvidence(record, desiredAssignmentEvidence(evidenceAssignments));
  }

  function applyConflict(record, leave, assignments) {
    const evidenceResult = mergeAutomaticEvidence(record, []);
    record.adjustedPercent = null;
    record.adjustmentSource = 'conflict';
    record.adjustmentNote = conflictNote(leave, assignments);
    record.cutiAdjustment = leaveMeta(leave, { adjustmentApplied: false, informational: true });
    record.assignmentAdjustment = assignmentMeta(assignments, { adjustmentApplied: false, informational: true });
    record.adjustmentConflict = {
      automatic: true,
      type: 'cuti_vs_surat_tugas',
      leaveId: leave.id,
      assignmentIds: assignments.map((item) => item.id)
    };
    return evidenceResult;
  }

  function attachSupportToManual(record, leave, assignments, leaveEvidenceOwner, assignmentOwnerById) {
    const desired = [];
    if (leave && !assignments.length && leaveEvidenceOwner === record) {
      desired.push(...desiredLeaveEvidence(leave, true));
    }
    if (!leave && assignments.length) {
      const evidenceAssignments = assignments.filter((assignment) => assignmentOwnerById.get(assignment.id) === record);
      desired.push(...desiredAssignmentEvidence(evidenceAssignments));
    }
    record.cutiAdjustment = leave ? leaveMeta(leave, { adjustmentApplied: false, informational: true }) : null;
    record.assignmentAdjustment = assignments.length ? assignmentMeta(assignments, { adjustmentApplied: false, informational: true }) : null;
    record.adjustmentConflict = leave && assignments.length
      ? {
          automatic: true,
          type: 'cuti_vs_surat_tugas',
          leaveId: leave.id,
          assignmentIds: assignments.map((item) => item.id)
        }
      : null;
    return mergeAutomaticEvidence(record, desired);
  }

  function sourceMatchesDate(source, key) {
    return Boolean(source && source.startDate <= key && source.endDate >= key);
  }

  async function apply(employees, period) {
    const list = Array.isArray(employees) ? employees : [];
    const [leaves, assignments] = await Promise.all([fetchLeaves(period), fetchAssignments(period)]);
    const leavesByEmployee = new Map();
    leaves.forEach((leave) => {
      if (!leavesByEmployee.has(leave.employeeId)) leavesByEmployee.set(leave.employeeId, []);
      leavesByEmployee.get(leave.employeeId).push(leave);
    });
    const assignmentsByEmployee = new Map();
    assignments.forEach((assignment) => {
      assignment.employeeIds.forEach((employeeId) => {
        if (!assignmentsByEmployee.has(employeeId)) assignmentsByEmployee.set(employeeId, []);
        assignmentsByEmployee.get(employeeId).push(assignment);
      });
    });

    let adjustedRecords = 0;
    let cutiAdjustedRecords = 0;
    let assignmentAdjustedRecords = 0;
    let assignmentInformationalRecords = 0;
    let conflictRecords = 0;
    let clearedRecords = 0;
    let preservedManualRecords = 0;
    let matchedEmployees = 0;
    let unresolvedEmployees = 0;
    let evidenceFiles = 0;
    let linkedRecords = 0;
    let cutiLinkedRecords = 0;
    let assignmentLinkedRecords = 0;
    const removedEvidence = [];

    for (const employee of list) {
      const masterEmployeeId = await resolveMasterEmployeeId(employee);
      if (!masterEmployeeId) {
        unresolvedEmployees += 1;
        continue;
      }
      const employeeLeaves = leavesByEmployee.get(masterEmployeeId) || [];
      const employeeAssignments = assignmentsByEmployee.get(masterEmployeeId) || [];
      if (employeeLeaves.length || employeeAssignments.length) matchedEmployees += 1;

      const records = Object.values(employee.records || {}).sort((a, b) => dateKey(a.key || a.date).localeCompare(dateKey(b.key || b.date)));
      const leaveOwnerById = new Map();
      const assignmentOwnerById = new Map();

      for (const leave of employeeLeaves) {
        const owner = records.find((record) => {
          const key = dateKey(record.key || record.date);
          const assignmentsOnDate = employeeAssignments.filter((assignment) => sourceMatchesDate(assignment, key));
          return sourceMatchesDate(leave, key) && !assignmentsOnDate.length;
        }) || null;
        if (owner) leaveOwnerById.set(leave.id, owner);
      }

      for (const assignment of employeeAssignments) {
        const owner = records.find((record) => {
          const key = dateKey(record.key || record.date);
          const leaveOnDate = employeeLeaves.find((leave) => sourceMatchesDate(leave, key));
          return sourceMatchesDate(assignment, key) && !leaveOnDate;
        }) || null;
        if (owner) assignmentOwnerById.set(assignment.id, owner);
      }

      for (const record of records) {
        const key = dateKey(record.key || record.date);
        const matchingLeave = employeeLeaves.find((leave) => sourceMatchesDate(leave, key)) || null;
        const matchingAssignments = employeeAssignments.filter((assignment) => sourceMatchesDate(assignment, key));

        if (matchingLeave || matchingAssignments.length) {
          linkedRecords += 1;
          if (matchingLeave) cutiLinkedRecords += 1;
          if (matchingAssignments.length) assignmentLinkedRecords += 1;

          let evidenceResult = { linked: 0, removed: [] };
          if (hasProtectedManualAdjustment(record)) {
            preservedManualRecords += 1;
            evidenceResult = attachSupportToManual(
              record,
              matchingLeave,
              matchingAssignments,
              matchingLeave ? leaveOwnerById.get(matchingLeave.id) : null,
              assignmentOwnerById
            );
          } else {
            const autoZeroAssignments = matchingAssignments.filter(assignmentUsesAutoZero);
            if (matchingLeave && matchingAssignments.length) {
              evidenceResult = applyConflict(record, matchingLeave, matchingAssignments);
              conflictRecords += 1;
            } else if (matchingLeave) {
              const includeEvidence = leaveOwnerById.get(matchingLeave.id) === record;
              if (record.needsVerification) {
                evidenceResult = applyLeave(record, matchingLeave, includeEvidence);
                adjustedRecords += 1;
                cutiAdjustedRecords += 1;
              } else {
                evidenceResult = applyLeaveInformation(record, matchingLeave, includeEvidence);
              }
            } else {
              const evidenceAssignments = matchingAssignments.filter((assignment) => assignmentOwnerById.get(assignment.id) === record);
              if (record.needsVerification && autoZeroAssignments.length) {
                evidenceResult = applyAssignments(record, matchingAssignments, evidenceAssignments);
                adjustedRecords += 1;
                assignmentAdjustedRecords += 1;
              } else if (autoZeroAssignments.length) {
                evidenceResult = applyAssignmentSupport(record, matchingAssignments, evidenceAssignments);
              } else {
                evidenceResult = applyAssignmentInformation(record, matchingAssignments, evidenceAssignments);
                assignmentInformationalRecords += 1;
              }
            }
          }
          evidenceFiles += Number(evidenceResult.linked || 0);
          removedEvidence.push(...(evidenceResult.removed || []));
          continue;
        }

        const clearResult = clearAutomaticAdjustment(record);
        if (clearResult.cleared) clearedRecords += 1;
        removedEvidence.push(...(clearResult.removed || []));
      }
    }

    return {
      status: 'applied',
      leaveRecords: leaves.length,
      assignmentRecords: assignments.length,
      adjustedRecords,
      cutiAdjustedRecords,
      assignmentAdjustedRecords,
      assignmentInformationalRecords,
      conflictRecords,
      clearedRecords,
      matchedEmployees,
      unresolvedEmployees,
      preservedManualRecords,
      evidenceFiles,
      linkedRecords,
      cutiLinkedRecords,
      assignmentLinkedRecords,
      removedEvidence
    };
  }

  window.TukinSupportingDataIntegration = Object.freeze({ apply, fetchLeaves, fetchAssignments });
})();
