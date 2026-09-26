(function () {
  'use strict';

  const ALLOWED = Object.freeze({
    lembur: Object.freeze([
      'ORGANIZATION_UNIT', 'WORK_UNIT', 'CITY',
      'RESPONSIBLE_TITLE_1', 'RESPONSIBLE_TITLE_2',
      'RESPONSIBLE_NAME', 'RESPONSIBLE_NIP', 'SPKL_ADDRESS'
    ]),
    tukin: Object.freeze([
      'SATKER', 'DEFAULT_TUKIN', 'SKP_SCORE', 'DEFAULT_SKP_DEDUCTION',
      'UNIT_WORK', 'UNIT_ORGANIZATION', 'UNIT_LABEL'
    ])
  });

  let loaded = false;
  let snapshots = { lembur: null, tukin: null };

  function db() {
    const value = window.FirebaseClient?.getDb?.();
    if (!value) throw new Error('Firestore belum siap. Pastikan pengguna sudah login.');
    return value;
  }

  function apply(target, source, allowedKeys) {
    if (!target || !source) return;
    allowedKeys.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(source, key)) target[key] = source[key];
    });
  }

  async function loadAndApply(force) {
    if (loaded && !force) return snapshots;
    const refs = ALLOWED;
    const [lemburSnap, tukinSnap] = await Promise.all([
      db().collection('appSettings').doc('lembur').get(),
      db().collection('appSettings').doc('tukin').get()
    ]);
    snapshots = {
      lembur: lemburSnap.exists ? lemburSnap.data() : null,
      tukin: tukinSnap.exists ? tukinSnap.data() : null
    };
    apply(window.APP_CONFIG, snapshots.lembur, refs.lembur);
    apply(window.TUKIN_CONFIG, snapshots.tukin, refs.tukin);
    loaded = true;
    return snapshots;
  }

  function missing(moduleName) {
    if (moduleName === 'lembur') {
      const cfg = window.APP_CONFIG || {};
      return ['RESPONSIBLE_NAME', 'RESPONSIBLE_NIP', 'SPKL_ADDRESS'].filter((key) => !String(cfg[key] || '').trim());
    }
    if (moduleName === 'tukin') {
      const cfg = window.TUKIN_CONFIG || {};
      return ['SATKER', 'DEFAULT_TUKIN'].filter((key) => !cfg[key]);
    }
    return [];
  }

  function allowedKeys(moduleName) {
    return [...(ALLOWED[moduleName] || [])];
  }

  function getSnapshots() {
    return snapshots;
  }

  window.AppSettingsService = Object.freeze({ loadAndApply, missing, allowedKeys, getSnapshots });
})();
