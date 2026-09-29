(function () {
  'use strict';

  const BATCH_SIZE = 400;

  function isFileLike(value) {
    return (typeof File !== 'undefined' && value instanceof File) ||
      (typeof Blob !== 'undefined' && value instanceof Blob);
  }

  function sanitize(value, options) {
    const opts = options || {};
    if (value === undefined || typeof value === 'function') return undefined;
    if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
    if (value instanceof Date) return value.toISOString();
    if (isFileLike(value)) return opts.fileMetadata
      ? { name: value.name || '', size: Number(value.size || 0), type: value.type || '', lastModified: Number(value.lastModified || 0), localOnly: true }
      : undefined;
    if (Array.isArray(value)) return value.map((item) => sanitize(item, opts)).filter((item) => item !== undefined);
    if (typeof value === 'object') {
      const result = {};
      Object.entries(value).forEach(([key, item]) => {
        if (opts.omitKeys && opts.omitKeys.has(key)) return;
        const safe = sanitize(item, opts);
        if (safe !== undefined) result[key] = safe;
      });
      return result;
    }
    return String(value);
  }

  async function commitOperations(db, operations) {
    for (let offset = 0; offset < operations.length; offset += BATCH_SIZE) {
      const batch = db.batch();
      operations.slice(offset, offset + BATCH_SIZE).forEach((op) => {
        if (op.type === 'set') batch.set(op.ref, op.data, op.options || {});
        if (op.type === 'delete') batch.delete(op.ref);
      });
      await batch.commit();
    }
  }

  async function replaceSubcollection(parentRef, collectionName, documents) {
    const db = window.FirebaseClient.getDb();
    const col = parentRef.collection(collectionName);
    const old = await col.get();
    const deleteOps = old.docs.map((doc) => ({ type: 'delete', ref: doc.ref }));
    if (deleteOps.length) await commitOperations(db, deleteOps);

    const setOps = (documents || []).map((item, index) => ({
      type: 'set',
      ref: col.doc(item.id || `item-${String(index + 1).padStart(4, '0')}`),
      data: item.data
    }));
    if (setOps.length) await commitOperations(db, setOps);
  }

  async function deleteSubcollection(parentRef, collectionName) {
    const db = window.FirebaseClient.getDb();
    const snap = await parentRef.collection(collectionName).get();
    const operations = snap.docs.map((doc) => ({ type: 'delete', ref: doc.ref }));
    if (operations.length) await commitOperations(db, operations);
  }

  async function deleteRunWithChildren(parentRef, childCollections) {
    for (const name of childCollections || []) await deleteSubcollection(parentRef, name);
    await parentRef.delete();

    // Jangan anggap delete berhasil hanya karena cache lokal sudah berubah.
    // Baca ulang langsung dari server agar UI hanya menyatakan sukses setelah
    // dokumen induk benar-benar hilang dari Firestore.
    const verification = await parentRef.get({ source: 'server' });
    if (verification.exists) {
      throw new Error(`Hard delete Firestore gagal: dokumen ${parentRef.path} masih ada di server.`);
    }
    return true;
  }

  function latestTimestamp(run) {
    return new Date(run?.updatedAt || run?.processedAt || 0).getTime() || 0;
  }

  window.FirebaseStorageCommon = Object.freeze({
    sanitize,
    replaceSubcollection,
    deleteSubcollection,
    deleteRunWithChildren,
    latestTimestamp
  });
})();
