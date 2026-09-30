(function () {
  'use strict';

  const EMPLOYEE_COLLECTION = 'masterEmployees';
  const DIRECTORY_COLLECTION = 'masterDirectory';
  const CACHE_TTL_MS = 5 * 60 * 1000;

  let employeeCache = null;
  let employeeCacheAt = 0;
  let employeeLookup = null;
  let directoryCache = null;
  let directoryCacheAt = 0;
  let bundlePromise = null;

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
    const values = Array.isArray(item?.nameHashes) ? [...item.nameHashes] : [];
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

  function fresh(cacheAt) {
    return cacheAt > 0 && Date.now() - cacheAt < CACHE_TTL_MS;
  }

  async function getEmployees(force) {
    if (!force && employeeCache && fresh(employeeCacheAt)) return employeeCache;
    const snap = await db().collection(EMPLOYEE_COLLECTION).get();
    employeeCache = snap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    employeeCacheAt = Date.now();
    employeeLookup = buildLookup(employeeCache);
    return employeeCache;
  }

  async function getDirectory(force) {
    if (!force && directoryCache && fresh(directoryCacheAt)) return directoryCache;
    const snap = await db().collection(DIRECTORY_COLLECTION).get();
    directoryCache = new Map(snap.docs.map((doc) => [doc.id, { id: doc.id, ...doc.data() }]));
    directoryCacheAt = Date.now();
    return directoryCache;
  }

  async function getBundle(force) {
    if (!force && employeeCache && directoryCache && fresh(employeeCacheAt) && fresh(directoryCacheAt)) {
      return { employees: employeeCache, directory: directoryCache };
    }
    if (!force && bundlePromise) return bundlePromise;
    bundlePromise = Promise.all([getEmployees(force), getDirectory(force)])
      .then(([employees, directory]) => ({ employees, directory }))
      .finally(() => { bundlePromise = null; });
    return bundlePromise;
  }

  async function getLookup(force) {
    await getEmployees(force);
    if (!employeeLookup) employeeLookup = buildLookup(employeeCache || []);
    return employeeLookup;
  }

  async function findByIdentity(identity) {
    const hashes = await hashIdentity(identity || {});
    const maps = await getLookup(false);
    if (hashes.nipHash && maps.byNip.has(hashes.nipHash)) return maps.byNip.get(hashes.nipHash);
    if (hashes.nameHash && maps.byName.has(hashes.nameHash)) return maps.byName.get(hashes.nameHash);
    return null;
  }

  function clearCache() {
    employeeCache = null;
    employeeLookup = null;
    employeeCacheAt = 0;
    directoryCache = null;
    directoryCacheAt = 0;
    bundlePromise = null;
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
    collectionName: EMPLOYEE_COLLECTION,
    directoryCollectionName: DIRECTORY_COLLECTION,
    normalizeName,
    sha256,
    hashIdentity,
    nameHashes,
    getEmployees,
    getDirectory,
    getBundle,
    getLookup,
    findByIdentity,
    clearCache,
    stats
  });
})();
