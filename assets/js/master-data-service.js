(function () {
  'use strict';

  const COLLECTION = 'masterEmployees';
  const CACHE_TTL_MS = 5 * 60 * 1000;
  let cache = null;
  let cacheAt = 0;
  let lookup = null;

  function db() {
    const value = window.FirebaseClient?.getDb?.();
    if (!value) throw new Error('Firestore belum siap. Pastikan pengguna sudah login.');
    return value;
  }

  function normalizeName(name) {
    return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  async function sha256(text) {
    const input = new TextEncoder().encode(String(text || ''));
    const digest = await crypto.subtle.digest('SHA-256', input);
    return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
  }

  async function hashIdentity(identity) {
    const nip = String(identity?.nip || '').replace(/\D/g, '');
    const normalizedName = normalizeName(identity?.name || '');
    const [nipHash, nameHash] = await Promise.all([
      nip ? sha256(nip) : Promise.resolve(''),
      normalizedName ? sha256(normalizedName) : Promise.resolve('')
    ]);
    return { nip, normalizedName, nipHash, nameHash };
  }

  function nameHashes(item) {
    const values = Array.isArray(item?.nameHashes) ? item.nameHashes : [];
    if (item?.nameHash) values.push(item.nameHash);
    return [...new Set(values.map(String).filter(Boolean))];
  }

  function buildLookup(items) {
    const byNip = new Map();
    const byName = new Map();
    items.forEach((item) => {
      if (!item || item.active === false) return;
      if (item.nipHash) byNip.set(String(item.nipHash), item);
      nameHashes(item).forEach((hash) => byName.set(hash, item));
    });
    return { byNip, byName };
  }

  async function getEmployees(force) {
    const now = Date.now();
    if (!force && cache && now - cacheAt < CACHE_TTL_MS) return cache;
    const snap = await db().collection(COLLECTION).get();
    cache = snap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    cacheAt = now;
    lookup = buildLookup(cache);
    return cache;
  }

  async function getLookup(force) {
    await getEmployees(force);
    if (!lookup) lookup = buildLookup(cache || []);
    return lookup;
  }

  async function findByIdentity(identity) {
    const hashes = await hashIdentity(identity || {});
    const maps = await getLookup(false);
    if (hashes.nipHash && maps.byNip.has(hashes.nipHash)) return maps.byNip.get(hashes.nipHash);
    if (hashes.nameHash && maps.byName.has(hashes.nameHash)) return maps.byName.get(hashes.nameHash);
    return null;
  }

  function clearCache() {
    cache = null;
    lookup = null;
    cacheAt = 0;
  }

  async function stats() {
    const items = await getEmployees(false);
    return {
      total: items.length,
      active: items.filter((item) => item.active !== false).length,
      inactive: items.filter((item) => item.active === false).length
    };
  }

  window.MasterDataService = Object.freeze({
    collectionName: COLLECTION,
    normalizeName,
    sha256,
    hashIdentity,
    nameHashes,
    getEmployees,
    getLookup,
    findByIdentity,
    clearCache,
    stats
  });
})();
