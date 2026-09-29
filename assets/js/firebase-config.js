(function () {
  'use strict';

  // Salin nilai ini dari Firebase Console -> Project settings -> Your apps -> Web app.
  // Firebase web config bukan secret. Keamanan data tetap ditentukan oleh Authentication
  // dan Firestore Security Rules.
  window.FIREBASE_CONFIG = Object.freeze({
    apiKey: 'AIzaSyC5mV7TT-Vt8ii8eC0iN40aqC7emnB37sw',
    authDomain: 'administration-project-cf8a1.firebaseapp.com',
    projectId: 'administration-project-cf8a1',
    storageBucket: 'administration-project-cf8a1.firebasestorage.app',
    messagingSenderId: '935707499710',
    appId: '1:935707499710:web:147ef5e06d218345c612ef'
  });

  window.FIREBASE_APP_SETTINGS = Object.freeze({
    // Kosongkan ('') jika tidak ingin membatasi domain email di sisi UI.
    // Firestore Rules pada folder /firebase/firestore.rules tetap menjadi pengaman utama.
    allowedEmailDomain: 'pkp.go.id',

    // Google login dapat dimatikan jika organisasi hanya memakai Email/Password.
    enableGoogleSignIn: false,
    enableEmailPasswordSignIn: true,

    // Riwayat Tukin dan Lembur menggunakan koleksi bersama lintas akun terverifikasi.
    dataScope: 'shared',

    // Maksimum item riwayat terbaru yang dimuat ke tabel/listener per modul.
    // Dokumen cloud yang lebih lama tidak dihapus otomatis.
    historyLimit: 24
  });
})();
