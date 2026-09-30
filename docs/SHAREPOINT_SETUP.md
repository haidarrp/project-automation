# Setup SharePoint untuk Modul Tukin

Dokumen ini adalah langkah implementasi untuk menyimpan **file Excel absensi asli**, **bukti dukung**, dan **rekap Tukin** ke SharePoint/OneDrive for Business. Firestore menyimpan data perhitungan dan pointer file; IndexedDB hanya menjadi cache lokal.

## A. Arsitektur akhir

```text
Browser / Aplikasi Tukin
  |-- Firebase Auth + Firestore
  |     `-- data perhitungan + metadata/pointer file
  |
  `-- Microsoft Entra ID + Microsoft Graph
        `-- SharePoint/OneDrive for Business
              `-- file Excel asli, bukti dukung, rekap
```

SharePoint diperlakukan sebagai **source of truth** untuk file setelah file memiliki `driveId` dan `itemId`. Jika file Excel diedit melalui Excel Online, aplikasi dapat mengambil versi SharePoint terbaru dan melakukan parsing/perhitungan ulang.

---

## 1. Prasyarat

Siapkan:

- akun Microsoft 365 Kementerian PKP;
- hak **Edit** ke folder SharePoint/OneDrive for Business yang menjadi root;
- hak membuat **App Registration** di Microsoft Entra ID, atau bantuan administrator tenant;
- repository GitHub;
- GitHub Pages atau web hosting HTTPS lain.

Folder root yang saat ini dikonfigurasi:

```text
https://kemenpkp-my.sharepoint.com/:f:/g/personal/haidar_rasyid_pkp_go_id/IgBUU-kkbmhFQZoyFaBIzSPNAeWMrJrwlTFe7Kb29SrVtpo?e=cPWVIx
```

> Catatan tata kelola: URL tersebut berada pada area `-my.sharepoint.com/personal/...`, yaitu OneDrive for Business milik akun tertentu. Kode tetap dapat menggunakannya. Untuk penggunaan produksi jangka panjang, lebih baik root dipindahkan ke **SharePoint Site/Document Library milik unit kerja** agar keberlangsungan arsip tidak bergantung pada satu akun pegawai.

---

## 2. Buat Microsoft Entra ID App Registration

1. Buka **Microsoft Entra admin center**.
2. Masuk ke **Identity > Applications > App registrations**.
3. Pilih **New registration**.
4. Nama yang disarankan:

   ```text
   Pusdatin PKP - Tukin SharePoint
   ```

5. Pada **Supported account types**, pilih:

   ```text
   Accounts in this organizational directory only (Single tenant)
   ```

6. Selesaikan pembuatan.
7. Pada halaman **Overview**, catat:
   - **Application (client) ID**;
   - **Directory (tenant) ID**.

Aplikasi ini adalah **Single Page Application/public client**. Jangan membuat atau menaruh `client secret`, certificate private key, password, atau credential lain di JavaScript/GitHub.

---

## 3. Tambahkan Redirect URI SPA

Pada App Registration:

1. Pilih **Authentication**.
2. Pilih **Add a platform**.
3. Pilih **Single-page application (SPA)**.
4. Tambahkan redirect URI produksi.

Contoh jika GitHub Pages:

```text
https://USERNAME.github.io/NAMA-REPOSITORY/msal-redirect.html
```

Contoh apabila repository adalah:

```text
https://github.com/kemenpkp/project-automation
```

maka pola Pages-nya biasanya:

```text
https://kemenpkp.github.io/project-automation/msal-redirect.html
```

Gunakan URL aktual repository Anda.

Untuk pengujian lokal, bila diperlukan tambahkan:

```text
http://localhost:5500/msal-redirect.html
```

atau port yang benar-benar digunakan web server lokal.

**Redirect URI harus sama persis**: protokol, domain, port, path, dan trailing slash bila ada.

### Pengaturan yang tidak diperlukan

- Jangan mengaktifkan implicit grant hanya untuk integrasi ini.
- Jangan memasukkan client secret ke SPA.
- Platform yang dipakai harus **SPA**, bukan `Web`.

MSAL Browser v5 menggunakan halaman redirect bridge khusus. File proyek yang digunakan adalah:

```text
msal-redirect.html
src/msal-redirect-entry.js
```

---

## 4. Tambahkan Microsoft Graph delegated permission

Pada App Registration:

1. Buka **API permissions**.
2. Pilih **Add a permission**.
3. Pilih **Microsoft Graph**.
4. Pilih **Delegated permissions**.
5. Tambahkan:

   ```text
   Files.ReadWrite
   ```

6. Bila tenant mewajibkan persetujuan administrator, minta administrator melakukan **Grant admin consent**.

Kode menggunakan delegated access: operasi Graph berjalan atas identitas pengguna Microsoft yang sedang login dan tetap tunduk pada hak akses file/folder pengguna tersebut.

---

## 5. Atur hak akses folder SharePoint

Pada folder root SharePoint yang diberikan:

1. Buka **Manage access**.
2. Berikan hak **Can edit/Edit** kepada akun yang menggunakan modul Tukin.
3. Untuk pengelolaan lebih baik, gunakan **Microsoft 365 Group/Security Group** daripada membagikan satu per satu.
4. Hindari konfigurasi **Anyone with the link** untuk data internal pegawai. Gunakan akun/grup organisasi yang terkontrol.

Contoh:

```text
Pusdatin-Tukin-Admin  -> Edit
```

Perlu dibedakan:

- hak **file** mengikuti SharePoint;
- hak **mengubah data perhitungan/riwayat di aplikasi** mengikuti Firestore. Pada rules proyek saat ini, pembuat riwayat atau administrator aplikasi yang dapat mengedit riwayat.

---

## 6. Isi konfigurasi Microsoft pada proyek

Edit:

```text
assets/js/microsoft-config.js
```

Ganti:

```js
tenantId: 'REPLACE_WITH_TENANT_ID',
clientId: 'REPLACE_WITH_CLIENT_ID',
```

menjadi, misalnya:

```js
tenantId: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
clientId: 'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy',
```

`shareUrl` telah diisi dengan link folder yang diberikan.

Konfigurasi lain yang tersedia:

```js
rootFolderName: 'TUKIN',
scopes: ['Files.ReadWrite'],
allowedEmailDomain: 'pkp.go.id',
requireFirebaseEmailMatch: false
```

Jika akun Microsoft harus sama persis dengan akun Firebase, ubah:

```js
requireFirebaseEmailMatch: true
```

`tenantId` dan `clientId` adalah identifier publik SPA, tetapi **client secret tidak boleh ada** di file ini.

---

## 7. Install dependency dan build

Dari root repository:

```bash
npm install
npm run build
```

Perintah build menghasilkan:

```text
assets/vendor/msal-browser.bundle.js
assets/vendor/msal-redirect-bridge.bundle.js
```

Dependency saat implementasi ini dipin ke:

```text
@azure/msal-browser 5.23.0
```

Jika `assets/vendor/...` belum ada saat membuka HTML secara lokal, autentikasi Microsoft tidak akan bekerja. GitHub Actions yang disediakan melakukan build otomatis sebelum deploy Pages.

---

## 8. Push ke GitHub

Jika folder belum menjadi repository Git:

```bash
git init
git add .
git commit -m "Add SharePoint storage for Tukin files"
git branch -M main
git remote add origin https://github.com/ORGANISASI/NAMA-REPOSITORY.git
git push -u origin main
```

Jika repository sudah ada:

```bash
git add .
git commit -m "Add SharePoint storage for Tukin files"
git push origin main
```

File utama yang baru/berubah:

```text
.github/workflows/pages.yml
package.json
src/msal-global.js
src/msal-redirect-entry.js
msal-redirect.html
assets/js/microsoft-config.js
assets/js/ms-auth.js
assets/js/sharepoint-storage.js
assets/js/tukin-app.js
assets/js/tukin-generator.js
assets/js/tukin-storage.js
tukin.html
sw.js
docs/SHAREPOINT_SETUP.md
README.md
```

---

## 9. Aktifkan GitHub Pages

Repository GitHub:

1. Buka **Settings > Pages**.
2. Pada **Build and deployment**, pilih:

   ```text
   Source: GitHub Actions
   ```

3. Push ke branch `main`.
4. Buka tab **Actions**.
5. Pastikan workflow **Deploy GitHub Pages** berhasil.
6. Buka URL Pages hasil deploy.

Workflow akan menjalankan:

```text
npm install
npm run build
upload artifact
Deploy Pages
```

Setelah mengetahui URL final Pages, pastikan URL `msal-redirect.html` pada Entra App Registration sama persis dengan URL deploy.

---

## 10. Struktur folder otomatis

Saat satu proses Tukin disimpan, aplikasi membuat struktur:

```text
TUKIN/
  2026/
    2026-09_September/
      Runs/
        Run_<id>/
          Rekap/
            (September) Rekap Potongan Tunjangan Kinerja CPNS 2026.xlsx
          Pegawai/
            199001012020121001_Nama Pegawai/
              Absensi/
                Absensi_2026-09_199001012020121001_Nama_Pegawai.xlsx
              Bukti_Dukung/
                2026-07-15_ab12cd34_Surat Tugas.pdf
```

Setiap run memiliki folder sendiri supaya file antarproses tidak saling menimpa.

Firestore menyimpan pointer seperti:

```text
driveId
itemId
webUrl
eTag
remoteName
remotePath
uploadedAt
```

File biner tidak dimasukkan ke Firestore.

---

## 11. Uji pertama

Lakukan urutan berikut:

1. Login ke aplikasi seperti biasa melalui Firebase.
2. Buka modul **Tunjangan Kinerja**.
3. Pastikan panel menampilkan **Hubungkan SharePoint**.
4. Klik **Hubungkan SharePoint**.
5. Login memakai akun Microsoft 365 Kementerian PKP.
6. Bila muncul consent, setujui `Files.ReadWrite` atau minta admin tenant memberi consent sesuai kebijakan organisasi.
7. Upload file Excel absensi.
8. Jalankan **Hitung & Validasi**.
9. Tambahkan koreksi/bukti dukung bila diperlukan.
10. Klik **Generate ZIP Final**.
11. Periksa folder SharePoint.
12. Pastikan tersedia:
    - file Excel asli;
    - bukti dukung;
    - rekap Tukin.
13. Periksa Firestore; dokumen employee harus berisi pointer SharePoint, bukan file biner.

---

## 12. Uji lintas perangkat / akun

### Perangkat kedua

1. Buka aplikasi pada browser/perangkat lain.
2. Login Firebase.
3. Klik **Hubungkan SharePoint**.
4. Login Microsoft 365.
5. Buka **Riwayat Tukin**.
6. Pilih run yang sudah dibuat.
7. Klik **SharePoint** atau **Generate Ulang**.

Generate ulang mengambil file yang memiliki pointer SharePoint dari penyimpanan remote, bukan bergantung pada IndexedDB perangkat pertama. Jika `eTag` file absensi menunjukkan bahwa sumber Excel telah berubah sejak perhitungan terakhir, aplikasi akan menahan Generate Ulang dan meminta pembuat/admin memuat ulang absensi terlebih dahulu agar rekap tidak berbeda dengan file sumber.

### Akun Microsoft lain

Akun lain dapat membuka/mengedit file bila:

- akun tersebut mempunyai hak **Edit** pada root SharePoint; dan
- kebijakan tenant mengizinkan consent/penggunaan App Registration tersebut.

Riwayat aplikasi tetap mengikuti Firestore rules. Secara default, pengguna lain dapat membaca riwayat bersama, tetapi perubahan perhitungan melalui UI hanya oleh pembuat run atau admin aplikasi.

---

## 13. Workflow edit Excel Online lalu hitung ulang

Setelah run pernah digenerate:

1. Buka **Riwayat Tukin**.
2. Klik link SharePoint atau buka detail file absensi.
3. Buka file `.xlsx` di **Excel Online**.
4. Lakukan perubahan lalu simpan.
5. Kembali ke aplikasi.
6. Pembuat run/admin pilih **Verifikasi/Edit**.
7. Pada layar validasi klik:

   ```text
   Muat Ulang Absensi dari SharePoint
   ```

8. Aplikasi mengunduh file versi SharePoint terbaru.
9. Aplikasi melakukan parsing dan perhitungan ulang.
10. Koreksi manual dan bukti dukung lama dipertahankan berdasarkan tanggal yang masih ada.
11. Klik **Simpan & Generate Ulang**.
12. Setelah tersimpan, `eTag`/pointer Firestore kembali merepresentasikan file sumber yang dipakai pada proses terbaru.

Catatan: mengedit Excel Online **tidak otomatis** mengubah angka Firestore sampai tombol **Muat Ulang Absensi dari SharePoint** dan **Simpan & Generate Ulang** dijalankan.

---

## 14. Riwayat lama sebelum integrasi SharePoint

Riwayat versi lama dapat mempunyai kondisi:

```text
Firestore -> metadata/perhitungan ada
SharePoint -> file belum ada
IndexedDB perangkat lama -> mungkin masih menyimpan file asli
```

Jika file lama masih tersedia di IndexedDB pada perangkat asal:

1. buka riwayat dari perangkat asal;
2. pilih **Verifikasi/Edit**;
3. klik **Simpan & Generate Ulang**;
4. file lokal akan disinkronkan ke SharePoint dan pointer baru disimpan ke Firestore.

Jika file fisik sudah hilang dari IndexedDB dan tidak pernah diunggah ke SharePoint, file tersebut tidak dapat direkonstruksi hanya dari metadata Firestore.

---

## 15. Troubleshooting

### `AADSTS50011` / redirect URI mismatch

Periksa bahwa URL `msal-redirect.html` pada Entra App Registration **identik** dengan URL aplikasi.

### `consent_required` / perlu admin approval

Tenant membatasi consent. Administrator Entra perlu memberikan consent untuk delegated permission yang digunakan.

### `403 Access denied`

Periksa:

- akun Microsoft yang login;
- hak **Edit** ke folder root;
- sharing link masih aktif;
- delegated permission `Files.ReadWrite` tersedia/consented.

### `SharePoint belum dikonfigurasi`

Isi `tenantId` dan `clientId` pada `assets/js/microsoft-config.js`, lalu deploy ulang.

### `msal-browser.bundle.js` 404

Jalankan:

```bash
npm install
npm run build
```

atau pastikan workflow GitHub Pages sukses.

### Popup login tidak muncul

Klik tombol **Hubungkan SharePoint** secara langsung dan pastikan browser tidak memblokir popup untuk domain aplikasi.

### File SharePoint dapat dibuka tetapi angka Tukin belum berubah

Perubahan Excel Online tidak langsung menghitung ulang data Firestore. Jalankan **Muat Ulang Absensi dari SharePoint**, verifikasi, kemudian **Simpan & Generate Ulang**.

---

## 16. Catatan keamanan

- Tidak ada client secret di browser/repository.
- Access token ditangani MSAL dan tidak disimpan ke Firestore.
- Hak file tetap dikontrol SharePoint.
- Gunakan link/grup internal, bukan `Anyone with the link`.
- Jangan menaruh bukti dukung pada lokasi yang dapat diakses pihak di luar kewenangan.
- Pertimbangkan retensi, klasifikasi informasi, audit akses, dan kebijakan pelindungan data pribadi organisasi karena file presensi/bukti dukung dapat memuat data pegawai.
- Untuk produksi jangka panjang, gunakan SharePoint Site/Document Library institusional sebagai root, bukan OneDrive pribadi pegawai.

---

## 17. Checklist go-live

- [ ] App Registration single-tenant dibuat.
- [ ] `Application (client) ID` dicatat.
- [ ] `Directory (tenant) ID` dicatat.
- [ ] Platform **SPA** dibuat.
- [ ] Redirect URI `msal-redirect.html` sesuai URL produksi.
- [ ] Delegated Graph permission `Files.ReadWrite` ditambahkan.
- [ ] Consent sesuai kebijakan tenant selesai.
- [ ] Root SharePoint diberikan hak Edit ke grup/pengguna yang benar.
- [ ] `tenantId` dan `clientId` diisi.
- [ ] `npm run build` berhasil atau GitHub Action berhasil.
- [ ] Login Microsoft dari aplikasi berhasil.
- [ ] Generate Tukin mengunggah Excel, bukti, dan rekap.
- [ ] Riwayat dapat dibuka dari perangkat kedua.
- [ ] Excel Online dapat diedit dan dimuat ulang ke aplikasi.
- [ ] Generate ulang menggunakan versi SharePoint terbaru.
- [ ] Riwayat lama yang masih diperlukan telah dimigrasikan dari perangkat yang masih mempunyai cache file.
