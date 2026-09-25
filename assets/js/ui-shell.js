(function () {
  'use strict';

  const ICONS = Object.freeze({
    home: '<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M9 21v-6h6v6"/>',
    tukin: '<path d="M5 4h14v16H5z"/><path d="M8 8h8M8 12h8M8 16h5"/>',
    overtime: '<circle cx="12" cy="12" r="8"/><path d="M12 7v5l3 2"/>',
    history: '<path d="M4 6h16"/><path d="M4 12h16"/><path d="M4 18h10"/>',
    process: '<path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M5 18v3h14v-3"/>'
  });

  function icon(name) {
    return `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[name] || ICONS.home}</svg>`;
  }

  function moduleGroup(module, activeModule, activeView) {
    const isTukin = module === 'tukin';
    const active = activeModule === module;
    const label = isTukin ? 'Tunjangan Kinerja' : 'Lembur';
    const processLabel = isTukin ? 'Proses Tukin' : 'Proses Lembur';
    const file = isTukin ? 'tukin.html' : 'lembur.html';
    return `<div class="nav-group ${active ? 'active-group' : ''}">
      <div class="nav-group-title"><span class="nav-icon">${icon(isTukin ? 'tukin' : 'overtime')}</span><span class="nav-label">${label}</span></div>
      <div class="nav-submenu">
        <a class="nav-sub-button ${active && activeView === 'process' ? 'active' : ''}" href="${file}#process"><span class="nav-sub-dot"></span><span>${processLabel}</span></a>
        <a class="nav-sub-button ${active && activeView === 'history' ? 'active' : ''}" href="${file}#history"><span class="nav-sub-dot"></span><span>Riwayat</span></a>
      </div>
    </div>`;
  }

  function navigation(activeModule, activeView) {
    return `<nav class="nav nav-modules" aria-label="Navigasi utama">
      <a class="nav-button nav-home ${activeModule === 'dashboard' ? 'active' : ''}" href="index.html"><span class="nav-icon">${icon('home')}</span><span class="nav-label">Dashboard</span></a>
      <div class="nav-groups" role="group" aria-label="Modul">
        ${moduleGroup('tukin', activeModule, activeView)}
        ${moduleGroup('lembur', activeModule, activeView)}
      </div>
    </nav>`;
  }

  function breadcrumb(module, viewLabel) {
    if (module === 'dashboard') return '<strong>Generator Dokumen</strong><span>/ Dashboard</span>';
    const moduleLabel = module === 'tukin' ? 'Tunjangan Kinerja' : 'Lembur';
    return `<strong>Generator Dokumen</strong><span>/ ${moduleLabel} / ${viewLabel || ''}</span>`;
  }

  function render(options) {
    const opts = options || {};
    const module = opts.module || 'dashboard';
    const view = opts.view || '';
    const content = opts.content || '';
    const overlays = opts.overlays || '';
    const user = window.FirebaseClient?.getCurrentUser?.();
    const displayName = user?.displayName || (user?.email ? user.email.split('@')[0] : 'Pengguna');
    const email = user?.email || '';
    const initials = String(displayName || 'PD').trim().split(/\s+/).slice(0, 2).map((part) => part[0] || '').join('').toUpperCase() || 'PD';
    const account = user ? `<div class="topbar-account"><div class="topbar-account-copy"><strong>${displayName.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</strong><span>${email.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</span></div><button class="btn btn-secondary btn-sm" type="button" onclick="window.FirebaseClient.signOut()">Keluar</button></div>` : '';
    return `<div class="app-shell">
      <aside class="sidebar">
        <div class="sidebar-brand"><a href="index.html" aria-label="Buka Dashboard"><img src="assets/img/logo-pkp.png" alt="Kementerian PKP"></a></div>
        ${navigation(module, view)}
        <div class="sidebar-footer"><div class="sidebar-avatar">${initials}</div><div class="sidebar-footer-copy"><strong>${displayName.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</strong>${email ? email.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : 'Kementerian PKP'}</div></div>
      </aside>
      <main class="main">
        <header class="topbar"><div class="topbar-label">${breadcrumb(module, opts.viewLabel)}</div>${account}</header>
        <section class="content">${content}</section>
      </main>
    </div>${overlays}`;
  }

  window.AppShell = Object.freeze({ render, navigation, icon });
})();
