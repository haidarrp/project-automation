(function () {
  'use strict';

  // Konfigurasi ini adalah konfigurasi PUBLIC untuk Single Page Application (SPA).
  // Jangan pernah menaruh client secret, certificate, atau password di file ini.
  window.MICROSOFT_CONFIG = Object.freeze({
    // Isi dari Microsoft Entra ID -> App registrations -> Overview.
    tenantId: 'REPLACE_WITH_TENANT_ID',
    clientId: 'REPLACE_WITH_CLIENT_ID',

    // Folder SharePoint/OneDrive for Business yang diberikan sebagai root penyimpanan.
    shareUrl: 'https://kemenpkp-my.sharepoint.com/:f:/g/personal/haidar_rasyid_pkp_go_id/IgBUU-kkbmhFQZoyFaBIzSPNAeWMrJrwlTFe7Kb29SrVtpo?e=cPWVIx',

    // Aplikasi membuat folder ini di bawah folder yang dibagikan di atas.
    rootFolderName: 'TUKIN',

    // Least-privileged delegated permission yang digunakan untuk operasi file/folder.
    scopes: Object.freeze(['Files.ReadWrite']),

    // Pembatasan akun Microsoft yang boleh dipakai pada UI.
    allowedEmailDomain: 'pkp.go.id',

    // Jika true, alamat Microsoft harus sama persis dengan akun Firebase yang login.
    // Set false bila organisasi memakai alias/UPN berbeda dari email aplikasi.
    requireFirebaseEmailMatch: false,

    // MSAL Browser v5 memakai redirect bridge khusus untuk popup/silent flow.
    // Kosong berarti otomatis menunjuk ke msal-redirect.html pada origin aplikasi.
    redirectUri: '',

    graphBaseUrl: 'https://graph.microsoft.com/v1.0',

    // PUT /content mendukung file sampai 250 MB. Aplikasi menolak file lebih besar.
    maxSimpleUploadBytes: 250 * 1024 * 1024
  });
})();
