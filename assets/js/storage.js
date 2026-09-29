(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;
  const client = window.FirebaseClient;
  const common = window.FirebaseStorageCommon;
  const HISTORY_LIMIT = Number(window.FIREBASE_APP_SETTINGS?.historyLimit || cfg.HISTORY_LIMIT || 24);
  const SHARED_COLLECTION = 'lemburRuns';
  const LEGACY_COLLECTION = 'lemburRuns';
  const DELETION_COLLECTION = 'lemburRunDeletions';

  function runsCollection() {
    return client.sharedCollection(SHARED_COLLECTION);
  }

  function legacyRunsCollection() {
    return client.userCollection(LEGACY_COLLECTION);
  }

  function deletionCollection() {
    return client.sharedCollection(DELETION_COLLECTION);
  }

  async function isDeletedRun(id) {
    if (!id) return false;
    try {
      const snap = await deletionCollection().doc(id).get();
      return snap.exists;
    } catch (error) {
      console.warn('Status penghapusan riwayat Lembur tidak dapat diperiksa:', error);
      return false;
    }
  }

  async function markDeletedRun(id, meta) {
    if (!id) return;
    const currentActor = actor();
    await deletionCollection().doc(id).set(common.sanitize({
      id,
      ownerUid: meta?.ownerUid || currentActor.uid,
      ownerEmail: meta?.ownerEmail || currentActor.email,
      deletedByUid: currentActor.uid,
      deletedByEmail: currentActor.email,
      deletedAt: new Date().toISOString(),
      schemaVersion: 1
    }));
  }

  function deleteLegacyLocalRun(id) {
    try {
      const raw = localStorage.getItem(cfg.STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed.history)) return;
      const next = parsed.history.filter((item) => item?.id !== id);
      if (next.length === parsed.history.length) return;
      parsed.history = next;
      localStorage.setItem(cfg.STORAGE_KEY, JSON.stringify(parsed));
    } catch (error) {
      console.warn('Riwayat Lembur lokal lama gagal dibersihkan:', error);
    }
  }

  async function deleteLegacyCloudRun(id) {
    try {
      const legacyRef = legacyRunsCollection().doc(id);
      const legacySnap = await legacyRef.get();
      if (legacySnap.exists) await common.deleteRunWithChildren(legacyRef, ['employees']);
    } catch (error) {
      console.warn('Riwayat Lembur privat lama gagal dibersihkan:', error);
    }
  }

  function actor() {
    const user = client.getCurrentUser();
    return {
      uid: user?.uid || '',
      email: user?.email || '',
      name: user?.displayName || ''
    };
  }

  function normalizeMetadata(item) {
    if (!item || typeof item !== 'object') return null;
    return {
      id: item.id,
      period: item.period,
      processedAt: item.processedAt || null,
      updatedAt: item.updatedAt || null,
      summary: item.summary || {},
      holidays: rules.normalizeHolidays(item.holidays || []),
      ownerUid: item.ownerUid || '',
      ownerEmail: item.ownerEmail || '',
      ownerName: item.ownerName || '',
      updatedByUid: item.updatedByUid || '',
      updatedByEmail: item.updatedByEmail || ''
    };
  }

  function canManage(run) {
    const user = client.getCurrentUser();
    if (!user || !run) return false;
    return client.isAdmin() || (run.ownerUid && run.ownerUid === user.uid);
  }

  async function saveRun(run) {
    if (!run || !run.id) throw new Error('Data riwayat Lembur tidak valid.');

    const existing = await getMetadata(run.id);
    if (existing && !canManage(existing)) {
      throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat diubah oleh pembuat atau administrator.');
    }

    const currentActor = actor();
    const processedAt = run.processedAt || existing?.processedAt || new Date().toISOString();
    const updatedAt = run.updatedAt || null;
    const serializedEmployees = rules.serializeEmployees(run.employees || []);
    const ref = runsCollection().doc(run.id);
    const ownerUid = existing?.ownerUid || currentActor.uid;
    const ownerEmail = existing?.ownerEmail || currentActor.email;
    const ownerName = existing?.ownerName || currentActor.name;

    await ref.set(common.sanitize({
      id: run.id,
      period: run.period,
      processedAt,
      updatedAt,
      sortAt: updatedAt || processedAt,
      summary: run.summary || rules.summarize(run.employees || []),
      holidays: rules.normalizeHolidays(run.holidays || []),
      employeeCount: serializedEmployees.length,
      ownerUid,
      ownerEmail,
      ownerName,
      updatedByUid: currentActor.uid,
      updatedByEmail: currentActor.email,
      schemaVersion: 3
    }));

    await common.replaceSubcollection(ref, 'employees', serializedEmployees.map((employee, index) => ({
      id: `employee-${String(index + 1).padStart(4, '0')}`,
      data: common.sanitize({ order: index, ...employee })
    })));

    return normalizeMetadata({
      ...run,
      processedAt,
      updatedAt,
      ownerUid,
      ownerEmail,
      ownerName,
      updatedByUid: currentActor.uid,
      updatedByEmail: currentActor.email
    });
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
    const meta = snap.exists ? normalizeMetadata(snap.data()) : null;

    if (meta && !canManage(meta)) {
      throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat dihapus oleh pembuat atau administrator.');
    }

    // Catat tombstone SEBELUM menghapus dokumen utama. Dengan demikian,
    // data legacy di localStorage / users/{uid}/lemburRuns tidak dapat
    // dimigrasikan kembali pada login atau perangkat berikutnya.
    await markDeletedRun(id, meta);

    if (snap.exists) await common.deleteRunWithChildren(ref, ['employees']);

    // Bersihkan salinan legacy milik akun yang sedang login bila ada.
    // Untuk penghapusan oleh admin atas data akun lain, tombstone global
    // tetap cukup untuk mencegah data lama muncul kembali.
    deleteLegacyLocalRun(id);
    await deleteLegacyCloudRun(id);
    return Boolean(snap.exists);
  }

  async function clearHistory() {
    const snap = await runsCollection().get();
    for (const doc of snap.docs) {
      const meta = normalizeMetadata(doc.data());
      if (canManage(meta)) await deleteRun(doc.id);
    }
  }

  function legacyLocalRuns() {
    try {
      const raw = localStorage.getItem(cfg.STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed.history) ? parsed.history : [];
    } catch (_) {
      return [];
    }
  }

  async function getLegacyCloudRun(id) {
    const ref = legacyRunsCollection().doc(id);
    const [metaSnap, employeesSnap] = await Promise.all([
      ref.get(),
      ref.collection('employees').orderBy('order', 'asc').get()
    ]);
    if (!metaSnap.exists) return null;
    const meta = metaSnap.data();
    const employees = employeesSnap.docs.map((doc) => {
      const data = { ...doc.data() };
      delete data.order;
      return data;
    });
    return {
      id: meta.id || id,
      period: meta.period,
      processedAt: meta.processedAt || null,
      updatedAt: meta.updatedAt || null,
      summary: meta.summary || {},
      holidays: rules.normalizeHolidays(meta.holidays || []),
      employees: rules.hydrateEmployees(employees)
    };
  }

  async function migrateLegacyCloud() {
    let snap;
    try {
      snap = await legacyRunsCollection().get();
    } catch (error) {
      console.warn('Riwayat Lembur privat lama tidak dapat dibaca untuk migrasi:', error);
      return;
    }

    for (const doc of snap.docs) {
      if (await isDeletedRun(doc.id)) continue;
      const target = runsCollection().doc(doc.id);
      const exists = await target.get();
      if (exists.exists) continue;
      const legacy = await getLegacyCloudRun(doc.id);
      if (legacy) await saveRun(legacy);
    }
  }

  async function migrateLegacyLocal() {
    const legacy = legacyLocalRuns();
    for (const item of legacy.slice(0, HISTORY_LIMIT)) {
      if (!item?.id) continue;
      if (await isDeletedRun(item.id)) continue;
      const target = runsCollection().doc(item.id);
      const exists = await target.get();
      if (exists.exists) continue;
      const hydrated = {
        ...item,
        holidays: rules.normalizeHolidays(item.holidays || []),
        employees: rules.hydrateEmployees(item.employees || [])
      };
      await saveRun(hydrated);
    }
  }

  async function migrateLegacyOnce() {
    const user = client.getCurrentUser();
    if (!user) return;
    const marker = `${cfg.STORAGE_KEY}:shared-history-migrated-v1:${user.uid}`;
    if (localStorage.getItem(marker) === '1') return;

    await migrateLegacyCloud();
    await migrateLegacyLocal();
    localStorage.setItem(marker, '1');
  }


  function subscribeHistory(onChange, onError) {
    return runsCollection().orderBy('sortAt', 'desc').limit(HISTORY_LIMIT).onSnapshot((snap) => {
      const history = snap.docs.map((doc) => normalizeMetadata(doc.data())).filter(Boolean);
      if (typeof onChange === 'function') onChange(history);
    }, (error) => {
      if (typeof onError === 'function') onError(error);
      else console.warn('Sinkronisasi realtime riwayat Lembur terputus:', error);
    });
  }

  window.AppStorage = Object.freeze({
    saveRun,
    listHistory,
    getRun,
    deleteRun,
    clearHistory,
    migrateLegacyOnce,
    canManage,
    subscribeHistory
  });
})();
