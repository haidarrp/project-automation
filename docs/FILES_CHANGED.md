# File yang Ditambahkan / Diubah

## File baru

- `.github/workflows/pages.yml` — build dependency MSAL dan deploy GitHub Pages.
- `README.md` — ringkasan implementasi.
- `package.json` — dependency MSAL Browser + esbuild.
- `src/msal-global.js` — entry bundler MSAL Browser.
- `src/msal-redirect-entry.js` — entry redirect bridge MSAL v5.
- `msal-redirect.html` — redirect URI khusus autentikasi Microsoft.
- `assets/js/microsoft-config.js` — konfigurasi publik Entra/SharePoint.
- `assets/js/ms-auth.js` — login/token delegated Microsoft Graph.
- `assets/js/sharepoint-storage.js` — resolve shared folder, create folder, upload, metadata, download, delete.
- `docs/SHAREPOINT_SETUP.md` — petunjuk setup lengkap.
- `docs/FILES_CHANGED.md` — file ini.

## File diubah

- `tukin.html` — memuat bundle MSAL dan modul SharePoint.
- `assets/js/tukin-app.js` — sinkronisasi SharePoint, UI koneksi, link file, reload Excel remote, delete cleanup.
- `assets/js/tukin-storage.js` — simpan/hydrate pointer SharePoint pada Firestore.
- `assets/js/tukin-generator.js` — ambil file SharePoint saat membuat ZIP.
- `sw.js` — menghindari cache lama pada bundle/config autentikasi Microsoft.
- `.gitignore` — mengabaikan `node_modules` dan artifact Pages lokal.

## Tidak diubah

Rules Firestore tetap mengikuti model akses sebelumnya: semua akun instansi yang valid dapat membaca riwayat bersama, tetapi perubahan/hapus riwayat dibatasi kepada pembuat atau admin aplikasi.
