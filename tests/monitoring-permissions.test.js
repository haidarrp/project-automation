'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

test('izin dashboard/reconcile Register lama bermigrasi untuk akses monitoring baca-saja', () => {
  const window = {FirebaseClient:{}};
  const document = {readyState:'loading',addEventListener(){}};
  class MutationObserver {observe(){}}
  vm.runInNewContext(fs.readFileSync('assets/js/user-access.js','utf8'),
    {window,document,MutationObserver,console});
  const convert = raw => window.UserAccess.normalizeGrants({register:raw}).register;
  assert.equal(convert({'dashboard.view':true})['cases.view'],true);
  assert.equal(convert({'reconcile.view':true})['cases.view'],true);
  assert.equal(convert({'cases.view':true})['cases.view'],true);
  assert.equal(convert({})['cases.view'],undefined);
  assert.equal(convert({'cases.view':true,'cases.verify':true})['cases.verify'],undefined);
  assert.equal(window.UserAccess.MOD.register.sections.length,1);
});

test('navigasi Register hanya satu menu dan arsip verifikasi tidak dapat ditulis', () => {
  const shell=fs.readFileSync('assets/js/ui-shell.js','utf8');
  const rule=fs.readFileSync('firebase/firestore.rules','utf8');
  const registerConfig=shell.split('register: {')[1].split('},')[0];
  assert.match(registerConfig,/Monitoring Konflik/);
  assert.match(registerConfig,/view:'cases'/);
  assert.doesNotMatch(registerConfig,/view:'reconcile'|view:'history'|view:'dashboard'/);
  for(const collection of ['attendanceCases','attendanceCaseLogs','attendanceReconciliations']) {
    const block=rule.split(`match /${collection}/`)[1].split('    }')[0];
    assert.match(block,/allow create, update, delete: if false;/);
  }
});
