(function () {
  'use strict';

  const DB_NAME = 'pusdatin-pkp-generator';
  const DB_VERSION = 1;
  const STORE_NAME = 'tukinRuns';
  const HISTORY_LIMIT = Number(window.FIREBASE_APP_SETTINGS?.historyLimit || 24);
  const client = window.FirebaseClient;
  const common = window.FirebaseStorageCommon;
  const rules = window.TukinRules;

  function cloudCollection() {
    return client.userCollection('tukinRuns');
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!('indexedDB' in window)) {
        reject(new Error('Browser tidak mendukung IndexedDB.'));
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error('Gagal membuka penyimpanan lokal Tukin.'));
    });
  }

  function txRequest(mode, action) {
    return openDb().then((db) => new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, mode);
      const store = tx.objectStore(STORE_NAME);
      let request;
      try { request = action(store); }
      catch (error) { db.close(); reject(error); return; }
      tx.oncomplete = () => { db.close(); resolve(request && 'result' in request ? request.result : undefined); };
      tx.onerror = () => { const error = tx.error || request?.error || new Error('Operasi penyimpanan lokal gagal.'); db.close(); reject(error); };
      tx.onabort = () => { const error = tx.error || new Error('Operasi penyimpanan lokal dibatalkan.'); db.close(); reject(error); };
    }));
  }

  async function localListFull() {
    const all = await txRequest('readonly', (store) => store.getAll());
    return Array.isArray(all) ? all : [];
  }

  async function localGet(id) {
    if (!id) return null;
    return (await txRequest('readonly', (store) => store.get(id))) || null;
  }

  async function localSave(run) {
    await txRequest('readwrite', (store) => store.put(run));
  }

  async function localDelete(id) {
    await txRequest('readwrite', (store) => store.delete(id));
  }

  function summaryOnly(run) {
    return {
      id: run.id,
      period: run.period,
      settings: run.settings || {},
      processedAt: run.processedAt,
      updatedAt: run.updatedAt || null,
      generatedName: run.generatedName || '',
      summary: run.summary || {},
      cloudOnly: Boolean(run.cloudOnly)
    };
  }

  function serializeEmployeeForCloud(employee, index) {
    const safe = common.sanitize(employee, { omitKeys: new Set(['sourceFiles', 'file']) });
    safe.order = index;
    safe.sourceFileMeta = (employee.sourceFiles || []).map((file) => ({
      name: file?.name || '',
      size: Number(file?.size || 0),
      type: file?.type || '',
      lastModified: Number(file?.lastModified || 0),
      localOnly: true
    }));
    safe.sourceFiles = [];
    Object.entries(safe.records || {}).forEach(([key, record]) => {
      record.date = key;
      record.evidence = (record.evidence || []).map((item) => ({
        id: item?.id || '',
        name: item?.name || item?.file?.name || '',
        size: Number(item?.size || item?.file?.size || 0),
        type: item?.type || item?.file?.type || '',
        localOnly: true
      }));
    });
    return safe;
  }

  function hydrateCloudEmployee(data) {
    const employee = { ...data };
    delete employee.order;
    employee.sourceFiles = [];
    employee.records = Object.fromEntries(Object.entries(employee.records || {}).map(([key, record]) => [key, {
      ...record,
      date: rules.dateFromKey(key),
      evidence: (record.evidence || []).map((item) => ({ ...item, file: null, localOnly: true }))
    }]));
    return employee;
  }

  async function cloudSave(run) {
    const ref = cloudCollection().doc(run.id);
    const processedAt = run.processedAt || new Date().toISOString();
    const updatedAt = run.updatedAt || null;
    const employees = Array.isArray(run.employees) ? run.employees : [];
    const hasLocalBinary = employees.some((employee) =>
      (employee.sourceFiles || []).length || Object.values(employee.records || {}).some((record) => (record.evidence || []).some((item) => item?.file))
    );

    await ref.set(common.sanitize({
      id: run.id,
      period: run.period,
      settings: run.settings || {},
      validationResults: run.validationResults || [],
      summary: run.summary || {},
      generatedName: run.generatedName || '',
      processedAt,
      updatedAt,
      sortAt: updatedAt || processedAt,
      employeeCount: employees.length,
      hasLocalBinary,
      schemaVersion: 2
    }));

    await common.replaceSubcollection(ref, 'employees', employees.map((employee, index) => ({
      id: `employee-${String(index + 1).padStart(4, '0')}`,
      data: serializeEmployeeForCloud(employee, index)
    })));
  }

  async function cloudList() {
    const snap = await cloudCollection().orderBy('sortAt', 'desc').limit(HISTORY_LIMIT).get();
    return snap.docs.map((doc) => summaryOnly(doc.data()));
  }

  async function cloudGet(id) {
    const ref = cloudCollection().doc(id);
    const [metaSnap, empSnap] = await Promise.all([
      ref.get(),
      ref.collection('employees').orderBy('order', 'asc').get()
    ]);
    if (!metaSnap.exists) return null;
    const meta = metaSnap.data();
    return {
      ...meta,
      employees: empSnap.docs.map((doc) => hydrateCloudEmployee(doc.data())),
      cloudOnly: true,
      binaryFilesAvailable: false
    };
  }

  async function listRuns() {
    await migrateLegacyOnce();
    try {
      return await cloudList();
    } catch (cloudError) {
      console.warn('Firestore tidak dapat dibaca, menggunakan cache lokal Tukin:', cloudError);
      return (await localListFull())
        .sort((a, b) => common.latestTimestamp(b) - common.latestTimestamp(a))
        .slice(0, HISTORY_LIMIT)
        .map(summaryOnly);
    }
  }

  async function getRun(id) {
    if (!id) return null;
    let local = null;
    try { local = await localGet(id); }
    catch (error) { console.warn('Riwayat lokal tidak dapat dibaca:', error); }

    try {
      const cloudMeta = await cloudCollection().doc(id).get();
      if (!cloudMeta.exists) return local ? { ...local, cloudOnly: false, binaryFilesAvailable: true } : null;
      const cloudData = cloudMeta.data();
      const localTime = common.latestTimestamp(local);
      const cloudTime = common.latestTimestamp(cloudData);
      if (local && localTime >= cloudTime) return { ...local, cloudOnly: false, binaryFilesAvailable: true };
      return cloudGet(id);
    } catch (cloudError) {
      if (local) return { ...local, cloudOnly: false, binaryFilesAvailable: true };
      throw cloudError;
    }
  }

  async function saveRun(run) {
    if (!run || !run.id) throw new Error('Data riwayat Tukin tidak valid.');
    // Full binary tetap lokal agar Spark tidak memerlukan Cloud Storage.
    await localSave(run);
    await cloudSave(run);
    await trimHistory();
    return run;
  }

  async function deleteRun(id) {
    if (!id) return false;
    try { await localDelete(id); } catch (error) { console.warn(error); }
    const ref = cloudCollection().doc(id);
    const snap = await ref.get();
    if (snap.exists) await common.deleteRunWithChildren(ref, ['employees']);
    return true;
  }

  async function trimHistory() {
    try {
      const local = (await localListFull()).sort((a, b) => common.latestTimestamp(b) - common.latestTimestamp(a));
      for (const run of local.slice(HISTORY_LIMIT)) await localDelete(run.id);
    } catch (error) { console.warn('Trim lokal gagal:', error); }

    const snap = await cloudCollection().orderBy('sortAt', 'desc').get();
    for (const doc of snap.docs.slice(HISTORY_LIMIT)) await common.deleteRunWithChildren(doc.ref, ['employees']);
  }

  async function migrateLegacyOnce() {
    const user = client.getCurrentUser();
    if (!user) return;
    const marker = `tukin-firebase-migrated:${user.uid}`;
    if (localStorage.getItem(marker) === '1') return;
    let local = [];
    try { local = await localListFull(); } catch (error) { console.warn(error); }
    for (const run of local.slice(0, HISTORY_LIMIT)) {
      if (!run?.id) continue;
      const ref = cloudCollection().doc(run.id);
      const existing = await ref.get();
      if (!existing.exists) await cloudSave(run);
    }
    localStorage.setItem(marker, '1');
  }

  window.TukinStorage = Object.freeze({ listRuns, getRun, saveRun, deleteRun, migrateLegacyOnce });
})();
