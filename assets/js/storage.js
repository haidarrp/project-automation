(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;
  const client = window.FirebaseClient;
  const common = window.FirebaseStorageCommon;
  const HISTORY_LIMIT = Number(window.FIREBASE_APP_SETTINGS?.historyLimit || cfg.HISTORY_LIMIT || 24);

  function runsCollection() {
    return client.userCollection('lemburRuns');
  }

  function normalizeMetadata(item) {
    if (!item || typeof item !== 'object') return null;
    return {
      id: item.id,
      period: item.period,
      processedAt: item.processedAt || null,
      updatedAt: item.updatedAt || null,
      summary: item.summary || {},
      holidays: rules.normalizeHolidays(item.holidays || [])
    };
  }

  async function saveRun(run) {
    if (!run || !run.id) throw new Error('Data riwayat Lembur tidak valid.');
    const existing = await getMetadata(run.id);
    const processedAt = run.processedAt || existing?.processedAt || new Date().toISOString();
    const updatedAt = run.updatedAt || null;
    const serializedEmployees = rules.serializeEmployees(run.employees || []);
    const ref = runsCollection().doc(run.id);

    await ref.set(common.sanitize({
      id: run.id,
      period: run.period,
      processedAt,
      updatedAt,
      sortAt: updatedAt || processedAt,
      summary: run.summary || rules.summarize(run.employees || []),
      holidays: rules.normalizeHolidays(run.holidays || []),
      employeeCount: serializedEmployees.length,
      schemaVersion: 2
    }));

    await common.replaceSubcollection(ref, 'employees', serializedEmployees.map((employee, index) => ({
      id: `employee-${String(index + 1).padStart(4, '0')}`,
      data: common.sanitize({ order: index, ...employee })
    })));

    await trimHistory();
    return normalizeMetadata({ ...run, processedAt, updatedAt });
  }

  async function getMetadata(id) {
    if (!id) return null;
    const snap = await runsCollection().doc(id).get();
    return snap.exists ? normalizeMetadata(snap.data()) : null;
  }

  async function listHistory() {
    await migrateLegacyOnce();
    const snap = await runsCollection().orderBy('sortAt', 'desc').limit(HISTORY_LIMIT).get();
    return snap.docs.map((doc) => normalizeMetadata(doc.data())).filter(Boolean);
  }

  async function getRun(id) {
    if (!id) return null;
    const ref = runsCollection().doc(id);
    const [metaSnap, employeesSnap] = await Promise.all([
      ref.get(),
      ref.collection('employees').orderBy('order', 'asc').get()
    ]);
    if (!metaSnap.exists) return null;
    const meta = metaSnap.data();
    const serialized = employeesSnap.docs.map((doc) => {
      const data = { ...doc.data() };
      delete data.order;
      return data;
    });
    return {
      ...normalizeMetadata(meta),
      employees: rules.hydrateEmployees(serialized)
    };
  }

  async function deleteRun(id) {
    if (!id) return false;
    const ref = runsCollection().doc(id);
    const snap = await ref.get();
    if (!snap.exists) return false;
    await common.deleteRunWithChildren(ref, ['employees']);
    return true;
  }

  async function clearHistory() {
    const snap = await runsCollection().get();
    for (const doc of snap.docs) await common.deleteRunWithChildren(doc.ref, ['employees']);
  }

  async function trimHistory() {
    const snap = await runsCollection().orderBy('sortAt', 'desc').get();
    const excess = snap.docs.slice(HISTORY_LIMIT);
    for (const doc of excess) await common.deleteRunWithChildren(doc.ref, ['employees']);
  }

  function legacyRuns() {
    try {
      const raw = localStorage.getItem(cfg.STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed.history) ? parsed.history : [];
    } catch (_) {
      return [];
    }
  }

  async function migrateLegacyOnce() {
    const user = client.getCurrentUser();
    if (!user) return;
    const marker = `${cfg.STORAGE_KEY}:firebase-migrated:${user.uid}`;
    if (localStorage.getItem(marker) === '1') return;
    const legacy = legacyRuns();
    for (const item of legacy.slice(0, HISTORY_LIMIT)) {
      if (!item?.id) continue;
      const exists = await runsCollection().doc(item.id).get();
      if (exists.exists) continue;
      const hydrated = {
        ...item,
        holidays: rules.normalizeHolidays(item.holidays || []),
        employees: rules.hydrateEmployees(item.employees || [])
      };
      await saveRun(hydrated);
    }
    localStorage.setItem(marker, '1');
  }

  window.AppStorage = Object.freeze({
    saveRun,
    listHistory,
    getRun,
    deleteRun,
    clearHistory,
    migrateLegacyOnce
  });
})();
