(function () {
  'use strict';

  const config = window.FIREBASE_CONFIG || {};
  const settings = window.FIREBASE_APP_SETTINGS || {};
  let app = null;
  let auth = null;
  let db = null;
  let currentUser = null;
  let adminUser = false;
  let authReadyPromise = null;
  let configured = false;

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function configLooksValid() {
    const required = ['apiKey', 'authDomain', 'projectId', 'appId'];
    return required.every((key) => {
      const value = String(config[key] || '').trim();
      return value && !value.startsWith('GANTI_');
    });
  }

  function emailAllowed(user) {
    if (!user || !user.email) return false;
    const domain = String(settings.allowedEmailDomain || '').trim().toLowerCase();
    if (!domain) return true;
    return String(user.email).toLowerCase().endsWith(`@${domain}`);
  }

  function accessAllowed(user) {
    return Boolean(user && user.emailVerified && emailAllowed(user));
  }

  function initialize() {
    if (configured) return;
    configured = true;
    if (!configLooksValid()) return;
    if (!window.firebase) throw new Error('Firebase SDK tidak termuat. Periksa koneksi internet/CDN.');

    app = firebase.apps && firebase.apps.length ? firebase.app() : firebase.initializeApp(config);
    auth = firebase.auth();
    db = firebase.firestore();

    // GitHub Pages berjalan melalui HTTPS. Long polling fallback membantu pada jaringan kantor/proxy
    // yang membatasi WebChannel, tanpa mengubah model keamanan aplikasi.
    try {
      db.settings({ experimentalAutoDetectLongPolling: true });
    } catch (_) {
      // Beberapa versi SDK tidak memerlukan/menolak settings setelah Firestore mulai dipakai.
    }

    authReadyPromise = new Promise((resolve) => {
      const unsubscribe = auth.onAuthStateChanged(async (user) => {
        currentUser = user || null;
        if (user && accessAllowed(user)) {
          try { await ensureUserProfile(user); } catch (error) { console.warn('Profil pengguna gagal diperbarui:', error); }
          await refreshAdminStatus(user);
        } else {
          adminUser = false;
        }
        unsubscribe();
        resolve(currentUser);
      });
    });
  }

  async function ensureUserProfile(user) {
    if (!db || !user) return;
    await db.collection('users').doc(user.uid).set({
      uid: user.uid,
      email: user.email || '',
      displayName: user.displayName || '',
      photoURL: user.photoURL || '',
      lastLoginAt: new Date().toISOString()
    }, { merge: true });
  }

  async function refreshAdminStatus(user) {
    adminUser = false;
    if (!db || !user || !accessAllowed(user)) return false;
    try {
      const snap = await db.collection('admins').doc(user.uid).get();
      adminUser = Boolean(snap.exists && snap.data()?.active === true);
    } catch (error) {
      console.warn('Status admin tidak dapat diperiksa:', error);
      adminUser = false;
    }
    return adminUser;
  }

  function renderAdminRequired() {
    const root = document.getElementById('app') || document.body;
    root.innerHTML = `<section class="firebase-auth-page"><div class="firebase-auth-card firebase-setup-card">
      <img src="assets/img/logo-pkp.png" alt="Kementerian PKP" class="firebase-auth-logo">
      <div class="firebase-auth-kicker">Akses Administrasi</div>
      <h1>Akses admin diperlukan</h1>
      <p>Akun ini dapat menggunakan aplikasi, tetapi belum terdaftar sebagai administrator master data.</p>
      <div class="firebase-login-form"><a class="btn btn-primary" href="tukin.html#dashboard">Kembali ke Dashboard Tukin</a></div>
    </div></section>`;
  }

  function renderSetupRequired() {
    const root = document.getElementById('app') || document.body;
    root.innerHTML = `<section class="firebase-auth-page"><div class="firebase-auth-card firebase-setup-card">
      <img src="assets/img/logo-pkp.png" alt="Kementerian PKP" class="firebase-auth-logo">
      <h1>Firebase belum dikonfigurasi</h1>
      <p>Isi <code>assets/js/firebase-config.js</code> menggunakan konfigurasi Web App dari Firebase Console, kemudian deploy ulang ke GitHub Pages.</p>
      <div class="firebase-auth-note">Petunjuk lengkap tersedia pada <strong>README_FIREBASE.md</strong> di root project.</div>
    </div></section>`;
  }

  function authErrorMessage(error) {
    const code = String(error?.code || '');
    const map = {
      'auth/invalid-credential': 'Email atau password tidak sesuai.',
      'auth/user-not-found': 'Akun tidak ditemukan.',
      'auth/wrong-password': 'Password tidak sesuai.',
      'auth/too-many-requests': 'Terlalu banyak percobaan login. Coba kembali beberapa saat lagi.',
      'auth/popup-closed-by-user': 'Jendela login Google ditutup sebelum proses selesai.',
      'auth/popup-blocked': 'Popup login diblokir browser. Izinkan popup untuk situs ini.',
      'auth/network-request-failed': 'Koneksi ke Firebase gagal. Periksa jaringan internet.',
      'auth/unauthorized-domain': 'Domain GitHub Pages belum ditambahkan ke Authorized domains di Firebase Authentication.'
    };
    return map[code] || error?.message || 'Autentikasi gagal.';
  }

  function loginMarkup(message) {
    const emailForm = settings.enableEmailPasswordSignIn !== false
      ? `<form class="firebase-login-form" data-firebase-email-form>
          <div class="field"><label for="firebase-email">Email</label><input id="firebase-email" type="email" autocomplete="username" required placeholder="nama@pkp.go.id"></div>
          <div class="field"><label for="firebase-password">Password</label><input id="firebase-password" type="password" autocomplete="current-password" required></div>
          <button class="btn btn-primary" type="submit">Masuk</button>
          <button class="firebase-link-button" type="button" data-firebase-reset>Lupa password?</button>
        </form>`
      : '';
    return `<section class="firebase-auth-page"><div class="firebase-auth-card">
      <img src="assets/img/logo-pkp.png" alt="Kementerian PKP" class="firebase-auth-logo">
      <div class="firebase-auth-kicker">Generator Dokumen Pusdatin</div>
      <h1>Selamat Datang</h1>
      <p>Gunakan akun pegawai atau akun institusi Anda</p>
      ${message ? `<div class="alert alert-danger firebase-auth-message">${esc(message)}</div>` : ''}
      ${emailForm}
    </div></section>`;
  }

  function renderVerification(user, message) {
    const root = document.getElementById('app') || document.body;
    root.innerHTML = `<section class="firebase-auth-page"><div class="firebase-auth-card">
      <img src="assets/img/logo-pkp.png" alt="Kementerian PKP" class="firebase-auth-logo">
      <div class="firebase-auth-kicker">Verifikasi akun</div>
      <h1>Verifikasi email diperlukan</h1>
      <p>Akun <strong>${esc(user?.email || '')}</strong> sudah login, tetapi alamat email belum terverifikasi. Verifikasi diperlukan sebelum Firestore dapat diakses.</p>
      ${message ? `<div class="alert alert-warning firebase-auth-message">${esc(message)}</div>` : ''}
      <div class="firebase-login-form">
        <button class="btn btn-primary" type="button" data-firebase-send-verification>Kirim email verifikasi</button>
        <button class="btn btn-secondary" type="button" data-firebase-check-verification>Saya sudah verifikasi</button>
        <button class="firebase-link-button" type="button" data-firebase-verification-signout>Gunakan akun lain</button>
      </div>
    </div></section>`;

    root.querySelector('[data-firebase-send-verification]')?.addEventListener('click', async () => {
      try {
        await user.sendEmailVerification();
        renderVerification(user, 'Email verifikasi telah dikirim. Buka tautan pada email tersebut, lalu kembali ke halaman ini.');
      } catch (error) {
        renderVerification(user, authErrorMessage(error));
      }
    });
    root.querySelector('[data-firebase-check-verification]')?.addEventListener('click', async () => {
      try {
        await user.reload();
        if (auth.currentUser?.emailVerified) location.reload();
        else renderVerification(auth.currentUser || user, 'Status email masih belum terverifikasi.');
      } catch (error) {
        renderVerification(user, authErrorMessage(error));
      }
    });
    root.querySelector('[data-firebase-verification-signout]')?.addEventListener('click', async () => {
      await auth.signOut();
      renderLogin('');
    });
  }

  function renderLogin(message) {
    const root = document.getElementById('app') || document.body;
    root.innerHTML = loginMarkup(message || '');

    root.querySelector('[data-firebase-google]')?.addEventListener('click', async () => {
      try {
        const provider = new firebase.auth.GoogleAuthProvider();
        provider.setCustomParameters({ prompt: 'select_account' });
        const result = await auth.signInWithPopup(provider);
        if (!emailAllowed(result.user)) {
          const email = result.user?.email || '';
          await auth.signOut();
          renderLogin(`Akun ${email} tidak termasuk domain yang diizinkan.`);
        } else if (!result.user?.emailVerified) {
          renderVerification(result.user);
        }
      } catch (error) {
        renderLogin(authErrorMessage(error));
      }
    });

    root.querySelector('[data-firebase-email-form]')?.addEventListener('submit', async (event) => {
      event.preventDefault();
      const email = String(root.querySelector('#firebase-email')?.value || '').trim();
      const password = String(root.querySelector('#firebase-password')?.value || '');
      try {
        const result = await auth.signInWithEmailAndPassword(email, password);
        if (!emailAllowed(result.user)) {
          await auth.signOut();
          renderLogin(`Akun ${email} tidak termasuk domain yang diizinkan.`);
        } else if (!result.user?.emailVerified) {
          renderVerification(result.user);
        }
      } catch (error) {
        renderLogin(authErrorMessage(error));
      }
    });

    root.querySelector('[data-firebase-reset]')?.addEventListener('click', async () => {
      const email = String(root.querySelector('#firebase-email')?.value || '').trim();
      if (!email) {
        renderLogin('Isi alamat email terlebih dahulu untuk mengirim tautan reset password.');
        return;
      }
      try {
        await auth.sendPasswordResetEmail(email);
        renderLogin(`Tautan reset password telah dikirim ke ${email}.`);
      } catch (error) {
        renderLogin(authErrorMessage(error));
      }
    });
  }

  async function requireAuth() {
    initialize();
    if (!configLooksValid()) {
      renderSetupRequired();
      return new Promise(() => {});
    }
    const firstUser = await authReadyPromise;
    if (firstUser && accessAllowed(firstUser)) {
      currentUser = firstUser;
      return firstUser;
    }
    if (firstUser && !emailAllowed(firstUser)) {
      await auth.signOut();
      renderLogin('Akun yang digunakan tidak termasuk domain yang diizinkan.');
    } else if (firstUser && !firstUser.emailVerified) {
      renderVerification(firstUser);
    } else {
      renderLogin('');
    }

    return new Promise((resolve) => {
      auth.onAuthStateChanged(async (user) => {
        currentUser = user || null;
        if (!user) return;
        if (!emailAllowed(user)) {
          const email = user.email || '';
          await auth.signOut();
          renderLogin(`Akun ${email} tidak termasuk domain yang diizinkan.`);
          return;
        }
        if (!user.emailVerified) {
          renderVerification(user);
          return;
        }
        try { await ensureUserProfile(user); } catch (error) { console.warn(error); }
        await refreshAdminStatus(user);
        resolve(user);
      });
    });
  }

  async function signOut() {
    initialize();
    adminUser = false;
    if (auth) await auth.signOut();
    location.replace('index.html');
  }

  async function requireAdmin() {
    const user = await requireAuth();
    if (!adminUser) await refreshAdminStatus(user);
    if (!adminUser) {
      renderAdminRequired();
      return new Promise(() => {});
    }
    return user;
  }

  function userRoot() {
    if (!currentUser) throw new Error('Pengguna belum login.');
    return db.collection('users').doc(currentUser.uid);
  }

  function userCollection(name) {
    return userRoot().collection(name);
  }

  function sharedCollection(name) {
    if (!currentUser) throw new Error('Pengguna belum login.');
    return db.collection(name);
  }

  function getCurrentUser() { return currentUser; }
  function getDb() { return db; }
  function getAuth() { return auth; }
  function isAdmin() { return adminUser; }

  window.FirebaseClient = Object.freeze({
    initialize,
    requireAuth,
    requireAdmin,
    signOut,
    userRoot,
    userCollection,
    sharedCollection,
    getCurrentUser,
    getDb,
    getAuth,
    isAdmin,
    refreshAdminStatus,
    emailAllowed,
    accessAllowed,
    configLooksValid
  });
})();
