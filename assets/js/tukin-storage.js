(function () {
  'use strict';

  const DB_NAME = 'pusdatin-pkp-generator';
  const DB_VERSION = 1;
  const STORE_NAME = 'tukinRuns';
  const HISTORY_LIMIT = Number(window.FIREBASE_APP_SETTINGS?.historyLimit || 24);
  const SHARED_COLLECTION = 'tukinRuns';
  const LEGACY_COLLECTION = 'tukinRuns';
  const DELETION_COLLECTION = 'tukinRunDeletions';
  const client = window.FirebaseClient;
  const common = window.FirebaseStorageCommon;
  const rules = window.TukinRules;

  function cloudCollection() {
    return client.sharedCollection(SHARED_COLLECTION);
  }

  function legacyCloudCollection() {
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
      console.warn('Status penghapusan riwayat Tukin tidak dapat diperiksa:', error);
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

  async function deleteLegacyCloudRun(id, ownerUid) {
    try {
      const currentUser = client.getCurrentUser();
      const targetUid = ownerUid || currentUser?.uid || '';
      if (!targetUid) return;
      if (targetUid !== currentUser?.uid && !client.isAdmin()) return;
      const legacyRef = client.getDb().collection('users').doc(targetUid).collection(LEGACY_COLLECTION).doc(id);
      const legacySnap = await legacyRef.get();
      if (legacySnap.exists) await common.deleteRunWithChildren(legacyRef, ['employees']);
    } catch (error) {
      console.warn('Riwayat Tukin privat lama gagal dibersihkan:', error);
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

  function canManage(run) {
    const user = client.getCurrentUser();
    if (!user || !run) return false;
    return client.isAdmin() || (run.ownerUid && run.ownerUid === user.uid);
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

  function summaryOnly(run, documentId) {
    return {
      id: documentId || run.id,
      period: run.period,
      settings: run.settings || {},
      processedAt: run.processedAt,
      updatedAt: run.updatedAt || null,
      generatedName: run.generatedName || '',
      summary: run.summary || {},
      ownerUid: run.ownerUid || '',
      ownerEmail: run.ownerEmail || '',
      ownerName: run.ownerName || '',
      updatedByUid: run.updatedByUid || '',
      updatedByEmail: run.updatedByEmail || '',
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
    const existingSnap = await ref.get();
    const existing = existingSnap.exists ? existingSnap.data() : null;
    if (existing && !canManage(existing)) {
      throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat diubah oleh pembuat atau administrator.');
    }

    const currentActor = actor();
    const processedAt = run.processedAt || existing?.processedAt || new Date().toISOString();
    const updatedAt = run.updatedAt || null;
    const employees = Array.isArray(run.employees) ? run.employees : [];
    const hasLocalBinary = employees.some((employee) =>
      (employee.sourceFiles || []).length || Object.values(employee.records || {}).some((record) => (record.evidence || []).some((item) => item?.file))
    );
    const ownerUid = existing?.ownerUid || currentActor.uid;
    const ownerEmail = existing?.ownerEmail || currentActor.email;
    const ownerName = existing?.ownerName || currentActor.name;

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
      ownerUid,
      ownerEmail,
      ownerName,
      updatedByUid: currentActor.uid,
      updatedByEmail: currentActor.email,
      schemaVersion: 3
    }));

    await common.replaceSubcollection(ref, 'employees', employees.map((employee, index) => ({
      id: `employee-${String(index + 1).padStart(4, '0')}`,
      data: serializeEmployeeForCloud(employee, index)
    })));

    return {
      ...run,
      processedAt,
      updatedAt,
      ownerUid,
      ownerEmail,
      ownerName,
      updatedByUid: currentActor.uid,
      updatedByEmail: currentActor.email
    };
  }

  async function cloudList() {
    const snap = await cloudCollection().orderBy('sortAt', 'desc').limit(HISTORY_LIMIT).get();
    return snap.docs.map((doc) => summaryOnly(doc.data(), doc.id));
  }

  async function cloudGetFrom(collection, id, cloudOnly) {
    const ref = collection.doc(id);
    const [metaSnap, empSnap] = await Promise.all([
      ref.get(),
      ref.collection('employees').orderBy('order', 'asc').get()
    ]);
    if (!metaSnap.exists) return null;
    const meta = metaSnap.data();
    return {
      ...meta,
      id: metaSnap.id,
      employees: empSnap.docs.map((doc) => hydrateCloudEmployee(doc.data())),
      cloudOnly: Boolean(cloudOnly),
      binaryFilesAvailable: false
    };
  }

  async function cloudGet(id) {
    return cloudGetFrom(cloudCollection(), id, true);
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
        .map((run) => summaryOnly(run));
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
      if (local && localTime >= cloudTime) {
        return {
          ...local,
          ownerUid: cloudData.ownerUid || local.ownerUid || '',
          ownerEmail: cloudData.ownerEmail || local.ownerEmail || '',
          ownerName: cloudData.ownerName || local.ownerName || '',
          updatedByUid: cloudData.updatedByUid || local.updatedByUid || '',
          updatedByEmail: cloudData.updatedByEmail || local.updatedByEmail || '',
          cloudOnly: false,
          binaryFilesAvailable: true
        };
      }
      return cloudGet(id);
    } catch (cloudError) {
      if (local) return { ...local, cloudOnly: false, binaryFilesAvailable: true };
      throw cloudError;
    }
  }

  async function saveRun(run) {
    if (!run || !run.id) throw new Error('Data riwayat Tukin tidak valid.');
    const enriched = await cloudSave(run);
    await localSave(enriched);
    await trimLocalHistory();
    return enriched;
  }

  async function resolveRunDocuments(id) {
    const found = new Map();
    const directRef = cloudCollection().doc(id);
    const directSnap = await directRef.get({ source: 'server' });
    if (directSnap.exists) found.set(directSnap.id, directSnap);

    const byStoredId = await cloudCollection().where('id', '==', id).limit(10).get({ source: 'server' });
    byStoredId.docs.forEach((doc) => found.set(doc.id, doc));
    return [...found.values()];
  }

  async function deleteRun(id) {
    if (!id) throw new Error('ID riwayat Tukin tidak valid.');

    const documents = await resolveRunDocuments(id);
    if (!documents.length) {
      throw new Error(`Dokumen Tukin dengan ID ${id} tidak ditemukan di Firestore server.`);
    }

    for (const snap of documents) {
      const meta = { ...snap.data(), id: snap.id };
      if (!canManage(meta)) {
        throw new Error('Riwayat ini dibuat oleh akun lain dan hanya dapat dihapus oleh pembuat atau administrator.');
      }
    }

    for (const snap of documents) {
      const raw = snap.data() || {};
      const canonicalId = snap.id;
      const logicalId = raw.id || canonicalId;
      const ownerUid = raw.ownerUid || '';

      await markDeletedRun(canonicalId, { ...raw, id: canonicalId });
      if (logicalId !== canonicalId) await markDeletedRun(logicalId, { ...raw, id: logicalId });

      // Hapus cloud lebih dulu; cache lokal baru dibersihkan setelah server
      // mengonfirmasi parent document benar-benar hilang.
      await common.deleteRunWithChildren(snap.ref, ['employees']);

      try { await localDelete(logicalId); } catch (error) { console.warn(error); }
      if (canonicalId !== logicalId) {
        try { await localDelete(canonicalId); } catch (error) { console.warn(error); }
      }
      await deleteLegacyCloudRun(canonicalId, ownerUid);
      if (logicalId !== canonicalId) await deleteLegacyCloudRun(logicalId, ownerUid);
    }

    const leftovers = await cloudCollection().where('id', '==', id).limit(10).get({ source: 'server' });
    if (!leftovers.empty) {
      throw new Error(`Penghapusan belum tuntas: masih ada ${leftovers.size} dokumen Tukin dengan ID ${id} di Firestore.`);
    }
    return true;
  }

  async function trimLocalHistory() {
    try {
      const local = (await localListFull()).sort((a, b) => common.latestTimestamp(b) - common.latestTimestamp(a));
      for (const run of local.slice(HISTORY_LIMIT)) await localDelete(run.id);
    } catch (error) { console.warn('Trim lokal gagal:', error); }
  }

  async function migrateLegacyCloud() {
    let snap;
    try {
      snap = await legacyCloudCollection().get();
    } catch (error) {
      console.warn('Riwayat Tukin privat lama tidak dapat dibaca untuk migrasi:', error);
      return;
    }

    for (const doc of snap.docs) {
      if (await isDeletedRun(doc.id)) continue;
      const target = cloudCollection().doc(doc.id);
      const exists = await target.get();
      if (exists.exists) continue;
      const legacy = await cloudGetFrom(legacyCloudCollection(), doc.id, true);
      if (legacy) await cloudSave(legacy);
    }
  }

  async function migrateLegacyLocal() {
    let local = [];
    try { local = await localListFull(); } catch (error) { console.warn(error); }
    for (const run of local.slice(0, HISTORY_LIMIT)) {
      if (!run?.id) continue;
      if (await isDeletedRun(run.id)) continue;
      const ref = cloudCollection().doc(run.id);
      const existing = await ref.get();
      if (!existing.exists) await cloudSave(run);
    }
  }

  async function migrateLegacyOnce() {
    const user = client.getCurrentUser();
    if (!user) return;
    const marker = `tukin-shared-history-migrated-v1:${user.uid}`;
    if (localStorage.getItem(marker) === '1') return;

    await migrateLegacyCloud();
    await migrateLegacyLocal();
    localStorage.setItem(marker, '1');
  }


  function subscribeRuns(onChange, onError) {
    return cloudCollection().orderBy('sortAt', 'desc').limit(HISTORY_LIMIT).onSnapshot((snap) => {
      const history = snap.docs.map((doc) => summaryOnly(doc.data(), doc.id));
      if (typeof onChange === 'function') onChange(history);
    }, (error) => {
      if (typeof onError === 'function') onError(error);
      else console.warn('Sinkronisasi realtime riwayat Tukin terputus:', error);
    });
  }

  window.TukinStorage = Object.freeze({ listRuns, getRun, saveRun, deleteRun, migrateLegacyOnce, canManage, subscribeRuns });
})();
