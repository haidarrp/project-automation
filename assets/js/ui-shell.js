(function () {
  'use strict';

  const ICONS = Object.freeze({
    home: '<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M9 21v-6h6v6"/>',
    tukin: '<path d="M5 4h14v16H5z"/><path d="M8 8h8M8 12h8M8 16h5"/>',
    overtime: '<circle cx="12" cy="12" r="8"/><path d="M12 7v5l3 2"/>',
    training: '<path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v16H6.5A2.5 2.5 0 0 0 4 21.5z"/><path d="M4 5.5v16M8 7h8M8 11h6"/>',
    leave: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M8 3v4M16 3v4M4 10h16"/><path d="m9 15 2 2 4-4"/>',
    users: '<circle cx="9" cy="8" r="3"/><path d="M3.8 19c.7-3 2.4-4.5 5.2-4.5S13.5 16 14.2 19"/><circle cx="17" cy="9" r="2.2"/><path d="M15.7 14.8c2.6-.1 4.1 1.3 4.5 4.2"/>',
    history: '<path d="M4 6h16"/><path d="M4 12h16"/><path d="M4 18h10"/>',
    process: '<path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M5 18v3h14v-3"/>',
    chevron: '<path d="m9 6 6 6-6 6"/>',
    admin: '<circle cx="12" cy="8" r="3"/><path d="M5 20c.7-4 3-6 7-6s6.3 2 7 6"/><path d="M18.5 5.5 20 4m-1.5 6.5L20 12"/>'
  });

  const MODULES = Object.freeze({
    tukin: {
      label: 'Tunjangan Kinerja',
      icon: 'tukin',
      file: 'tukin.html',
      items: [
        { view: 'dashboard', label: 'Dashboard' },
        { view: 'process', label: 'Proses Tukin' },
        { view: 'history', label: 'Riwayat' }
      ]
    },
    lembur: {
      label: 'Lembur',
      icon: 'overtime',
      file: 'lembur.html',
      items: [
        { view: 'dashboard', label: 'Dashboard' },
        { view: 'process', label: 'Proses Lembur' },
        { view: 'history', label: 'Riwayat' }
      ]
    },
    training: {
      label: 'Pelatihan Pegawai',
      icon: 'training',
      file: 'pelatihan.html',
      items: [
        { view: 'dashboard', label: 'Dashboard' },
        { view: 'pegawai', label: 'Data Pegawai' },
        { view: 'pelatihan', label: 'Data Pelatihan' }
      ]
    },
    leave: {
      label: 'Cuti',
      icon: 'leave',
      file: 'cuti.html',
      items: [
        { view: 'dashboard', label: 'Dashboard' },
        { view: 'data', label: 'Data Cuti' },
        { view: 'kalender', label: 'Kalender' },
        { view: 'pegawai', label: 'Data Pegawai' },
        { view: 'saldo', label: 'Saldo Cuti' },
        { view: 'laporan', label: 'Laporan' }
      ]
    }
  });

  function icon(name) {
    return `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[name] || ICONS.home}</svg>`;
  }

  function moduleGroup(module, activeModule, activeView) {
    const config = MODULES[module];
    if (!config) return '';
    const active = activeModule === module;
    return `<div class="nav-group ${active ? 'active-group open' : ''}" data-nav-group="${module}">
      <button class="nav-group-title" type="button" aria-expanded="${active ? 'true' : 'false'}" onclick="window.AppShell.toggleModule(this)">
        <span class="nav-icon">${icon(config.icon)}</span>
        <span class="nav-label">${config.label}</span>
        <span class="nav-chevron">${icon('chevron')}</span>
      </button>
      <div class="nav-submenu">
        ${config.items.map((item) => `<a class="nav-sub-button ${active && activeView === item.view ? 'active' : ''}" href="${config.file}#${item.view}"><span class="nav-sub-dot"></span><span>${item.label}</span></a>`).join('')}
      </div>
    </div>`;
  }

  function toggleModule(button) {
    const group = button?.closest?.('.nav-group');
    if (!group) return;

    const shouldOpen = !group.classList.contains('open');
    const navGroups = group.parentElement?.querySelectorAll?.('.nav-group') || [];

    navGroups.forEach((item) => {
      item.classList.remove('open');
      const trigger = item.querySelector('.nav-group-title');
      if (trigger) trigger.setAttribute('aria-expanded', 'false');
    });

    if (shouldOpen) {
      group.classList.add('open');
      button.setAttribute('aria-expanded', 'true');
    }
  }

  function navigation(activeModule, activeView) {
    const isAdmin = Boolean(window.FirebaseClient?.isAdmin?.());
    const adminLink = isAdmin
      ? `<a class="nav-button nav-admin ${activeModule === 'admin' ? 'active' : ''}" href="admin.html"><span class="nav-icon">${icon('admin')}</span><span class="nav-label">Administrasi</span></a>`
      : '';
    const trainingGroup = isAdmin ? moduleGroup('training', activeModule, activeView) : '';
    const leaveGroup = isAdmin ? moduleGroup('leave', activeModule, activeView) : '';
    return `<nav class="nav nav-modules" aria-label="Navigasi utama">
      <div class="nav-groups" role="group" aria-label="Modul">
        ${moduleGroup('tukin', activeModule, activeView)}
        ${moduleGroup('lembur', activeModule, activeView)}
        ${trainingGroup}
        ${leaveGroup}
      </div>
      ${adminLink}
    </nav>`;
  }

  function breadcrumb(module, viewLabel) {
    if (module === 'dashboard') return '<strong>Generator Dokumen</strong><span>/ Dashboard</span>';
    if (module === 'admin') return '<strong>Generator Dokumen</strong><span>/ Administrasi / Master Data</span>';
    if (module === 'training') return `<strong>Pusdatin PKP</strong><span>/ Pelatihan Pegawai / ${viewLabel || ''}</span>`;
    if (module === 'leave') return `<strong>Pusdatin PKP</strong><span>/ Cuti / ${viewLabel || ''}</span>`;
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
        <div class="sidebar-brand"><a href="tukin.html#dashboard" aria-label="Buka Dashboard Tukin"><img src="assets/img/logo-pkp.png" alt="Kementerian PKP"></a></div>
        ${navigation(module, view)}
        <div class="sidebar-footer"><div class="sidebar-avatar">${initials}</div><div class="sidebar-footer-copy"><strong>${displayName.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</strong>${email ? email.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : 'Kementerian PKP'}</div></div>
      </aside>
      <main class="main">
        <header class="topbar"><div class="topbar-label">${breadcrumb(module, opts.viewLabel)}</div>${account}</header>
        <section class="content">${content}</section>
      </main>
    </div>${overlays}`;
  }

  window.AppShell = Object.freeze({ render, navigation, icon, toggleModule });
})();
