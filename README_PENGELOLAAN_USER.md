# Pengelolaan User: Hak Akses Granular

## Ruang lingkup

Versi ini menambahkan halaman **Administrasi > Pengelolaan User** (`pengguna.html`), hanya untuk akun yang telah didaftarkan sebagai Admin di koleksi Firestore `admins` (`active: true`).

Admin dapat:

- Melihat daftar akun terdaftar yang pernah login (koleksi `users`), status, peran, dan izin masing-masing.
- Mencari pengguna berdasarkan nama/email, menyaring Admin/User, serta memuat ulang daftar.
- Mengaktifkan/menonaktifkan akses keseluruhan pengguna **nonadmin**.
- Mencentang hak akses per modul, per submenu, dan per tindakan (lihat, tambah/input, ubah/edit, hapus, unduh/cetak, dan tindakan khusus).
- Menggunakan preset **Hanya Lihat**, **Lihat + Input**, **Semua Izin**, atau **Nonaktifkan Modul**.
- Menyimpan semua perubahan secara atomik dalam transaksi Firestore dan mencatatnya pada `adminAudit`. Revisi digunakan untuk mencegah perubahan saling menimpa.

Izin per tindakan yang dapat diatur:

| Modul | Submenu / tindakan |
| --- | --- |
| Cuti | Dashboard (lihat); Data Cuti (lihat/input/edit/hapus); Kalender (lihat); Data Pegawai (lihat); Saldo Cuti (lihat/edit saldo); Laporan (lihat/cetak); Bukti Dukung (lihat/unduh/tambah/edit/hapus) |
| Surat Tugas | Dashboard; Data Surat Tugas (lihat/input/edit/hapus); Berkas Surat Tugas (lihat/unggah/unduh); Kalender |
| Pelatihan Pegawai | Dashboard; Data Pegawai; Data Pelatihan (lihat/input/edit/hapus); Status & Catatan Peserta (lihat/edit); Metadata Dokumen (lihat/tambah/edit/hapus) |
| Tunjangan Kinerja | Dashboard; Proses (lihat/unggah/hitung/verifikasi/bukti/simpan); Parameter (lihat/tanggal merah/Ramadan); Riwayat (lihat/unduh/edit/hapus) |
| Lembur | Dashboard; Proses (lihat/unggah/hitung/koreksi/tanggal merah/buat dokumen); Riwayat (lihat/unduh/edit/hapus) |

## Langkah deployment (PENTING)

1. **Cadangkan data Firestore dan file proyek sebelum deployment.** Uji dahulu di proyek Firebase *staging* yang terpisah.
2. Unggah file proyek ke GitHub Pages atau hosting HTTPS yang sama seperti sebelumnya.
3. **Deploy `firebase/firestore.rules` ke Firebase project yang sedang digunakan**. Contoh melalui Firebase CLI pada root proyek setelah memilih proyek yang benar:
   ```bash
   firebase login
   firebase use <FIREBASE_PROJECT_ID>
   firebase deploy --only firestore:rules
   ```
   File `firebase.json` sudah menunjuk ke `firebase/firestore.rules`. **Deploy GitHub Pages saja tidak memperbarui Firestore Security Rules.** Tanpa Rules baru, pengguna nonadmin tidak dapat menyimpan data modul Cuti/Surat Tugas/Pelatihan atau membaca daftar izin.
4. Pastikan dokumen Admin awal sudah tersedia di Firestore `admins/{UID}`, dengan field `active: true`. Status Admin tidak dapat diberikan oleh aplikasi, demi keamanan.
5. Login sebagai Admin > Administrasi > Pengelolaan User. Cari pengguna yang sudah pernah login, pilih **Kelola Akses**, tentukan izin, lalu klik **Simpan Perubahan**.
6. Keluar/masuk kembali dengan akun pengguna dan uji setiap skenario: hanya lihat, input tanpa edit/hapus, edit saldo tanpa edit data, pengguna nonaktif, dan akses yang ditolak. Gunakan lingkungan staging terlebih dahulu.

## Model data

- `users/{uid}`: profil minimal yang dicatat saat login (nama/email/login terakhir).
- `admins/{uid}`: daftar Admin resmi (dikelola manual Firebase Console; tidak dapat diubah klien).
- `userPermissions/{uid}`: `uid`, `active`, `grants: {module: {"section.action": true}}`, `revision`, `updatedAt`, `updatedByUid`, `updatedByEmail`, dan `schemaVersion`.
- `adminAudit/{id}`: catatan perubahan izin bersifat *append-only*.

**Kompatibilitas akun lama:** Pengguna yang belum mempunyai dokumen `userPermissions` tetap memperoleh hak fitur Tukin/Lembur lama. Modul Cuti, Surat Tugas, dan Pelatihan perlu diaktifkan Admin. Begitu Admin menyimpan izin khusus pengguna, seluruh haknya mengikuti daftar checkbox yang tersimpan. Admin tetap memiliki akses penuh.

**Aturan akses bertingkat:** Menghapus izin `section.view` otomatis menghapus tindakan dari section tersebut, dan menghapus izin submenu induk akan menghapus tindakan dari submenu turunannya. Tombol yang tidak berizin disembunyikan. Tindakan simpan/hapus juga divalidasi pada kode modul, dan aturan Firestore memeriksa operasi baca/tulis pada koleksi.

## Batasan keamanan yang perlu dipahami

- **Firestore Security Rules harus dipasang dan diuji.** Tanpa deploy rules, pengaturan dari web hanya membatasi antarmuka, bukan akses langsung ke Firestore.
- **Microsoft SharePoint / OneDrive memiliki model izin terpisah.** Checkbox "Lihat/Unduh/Unggah" dalam aplikasi tidak otomatis mengubah izin Microsoft 365. Akun yang masih memiliki akses langsung ke folder SharePoint mungkin tetap bisa membaca/mengunggah file dari luar aplikasi. Untuk pembatasan ketat, perlu kontrol izin Microsoft Graph/SharePoint pada sisi server.
- **Izin Cuti: Lihat Dokumen vs Lihat Data**: informasi metadata dokumen secara historis melekat pada `leaveRecords`. Pengguna yang mendapat akses baca koleksi data bisa membaca metadata melalui API Firestore. UI sudah menyembunyikan panel dokumen tanpa izinnya; pemisahan metadata memerlukan migrasi skema.
- **Izin Pelatihan: Edit Status Peserta vs Edit Metadata Dokumen**: keduanya berada di array `participants` dalam dokumen Firestore yang sama. UI menolak perubahan yang tidak berizin, tetapi Firestore Rules hanya bisa membatasi perubahan field `participants` secara keseluruhan. Untuk enforcement per elemen/tipe aksi yang sepenuhnya independen, disarankan memigrasikan peserta dan dokumen ke subkoleksi terpisah atau menggunakan backend tepercaya.
- **Aturan edit saldo** memisahkan pengubahan jatah dan sisa bawaan dari pembaruan penggunaan otomatis akibat input/edit/hapus cuti. Penghitungan penggunaan tetap disinkronkan agar data konsisten.
- **Tukin/Lembur:** hak perubahan riwayat tetap dibatasi pemilik data atau Admin, selaras alur penyimpanan bersama yang sudah ada.
- Pengguna yang baru dibuat tetapi **belum pernah login** tidak akan muncul sampai profilnya tersimpan pada koleksi `users`.

## Berkas utama

- `pengguna.html`
- `assets/js/pengguna.js`, `assets/css/pengguna.css`
- `assets/js/user-access.js`
- `assets/js/ui-shell.js`
- `assets/js/cuti.js`, `surat-tugas.js`, `pelatihan.js`, `tukin-app.js`, `app.js`
- `firebase/firestore.rules`

## Build Microsoft MSAL

GitHub Actions proyek sudah menjalankan `npm install` + `npm run build` untuk menghasilkan `assets/vendor/msal-browser.bundle.js` dan `msal-redirect-bridge.bundle.js`. Jika menjalankan secara lokal, jalankan kedua perintah tersebut sebelum membuka modul yang memakai Microsoft 365.
