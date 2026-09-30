(function () {
  'use strict';

  const cfg = window.MICROSOFT_CONFIG || {};
  let instance = null;
  let initialized = false;
  let initializationPromise = null;

  function placeholder(value) {
    return !value || /^REPLACE_WITH_/i.test(String(value));
  }

  function isConfigured() {
    return Boolean(
      !placeholder(cfg.tenantId) &&
      !placeholder(cfg.clientId) &&
      cfg.shareUrl &&
      window.msal?.PublicClientApplication
    );
  }

  function redirectUri() {
    return cfg.redirectUri || new URL('msal-redirect.html', document.baseURI).href;
  }

  function accountEmail(account) {
    return String(
      account?.username ||
      account?.idTokenClaims?.preferred_username ||
      account?.idTokenClaims?.email ||
      ''
    ).trim().toLowerCase();
  }

  function validateAccount(account) {
    if (!account) throw new Error('Akun Microsoft tidak tersedia.');
    const email = accountEmail(account);
    const domain = String(cfg.allowedEmailDomain || '').trim().toLowerCase();
    if (domain && !email.endsWith(`@${domain}`)) {
      throw new Error(`Gunakan akun Microsoft organisasi dengan domain @${domain}.`);
    }

    if (cfg.requireFirebaseEmailMatch) {
      const firebaseEmail = String(window.FirebaseClient?.getCurrentUser?.()?.email || '').trim().toLowerCase();
      if (firebaseEmail && email && firebaseEmail !== email) {
        throw new Error(`Akun Microsoft (${email}) harus sama dengan akun aplikasi (${firebaseEmail}).`);
      }
    }
    return account;
  }

  async function initialize() {
    if (initialized) return true;
    if (initializationPromise) return initializationPromise;
    if (!isConfigured()) return false;

    initializationPromise = (async () => {
      instance = new window.msal.PublicClientApplication({
        auth: {
          clientId: cfg.clientId,
          authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
          redirectUri: redirectUri(),
          postLogoutRedirectUri: redirectUri(),
          navigateToLoginRequestUrl: true
        },
        cache: {
          cacheLocation: 'sessionStorage',
          storeAuthStateInCookie: false
        },
        system: {
          allowPlatformBroker: false
        }
      });

      await instance.initialize();
      const redirectResult = await instance.handleRedirectPromise();
      if (redirectResult?.account) instance.setActiveAccount(redirectResult.account);

      if (!instance.getActiveAccount()) {
        const accounts = instance.getAllAccounts();
        if (accounts.length === 1) instance.setActiveAccount(accounts[0]);
      }

      if (instance.getActiveAccount()) validateAccount(instance.getActiveAccount());
      initialized = true;
      return true;
    })().finally(() => {
      initializationPromise = null;
    });

    return initializationPromise;
  }

  async function connect() {
    if (!await initialize()) {
      throw new Error('Integrasi Microsoft belum dikonfigurasi. Isi tenantId dan clientId pada assets/js/microsoft-config.js lalu jalankan build MSAL.');
    }

    let account = instance.getActiveAccount();
    if (!account) {
      const result = await instance.loginPopup({
        scopes: Array.from(cfg.scopes || ['Files.ReadWrite']),
        prompt: 'select_account',
        redirectUri: redirectUri()
      });
      account = result.account;
      instance.setActiveAccount(account);
    }
    validateAccount(account);
    await acquireToken(true);
    return getStatus();
  }

  async function acquireToken(interactive) {
    if (!await initialize()) throw new Error('Integrasi Microsoft belum dikonfigurasi.');
    let account = instance.getActiveAccount();
    if (!account) {
      if (interactive) {
        await connect();
        account = instance.getActiveAccount();
      } else {
        throw new Error('SharePoint belum terhubung.');
      }
    }
    validateAccount(account);

    const request = {
      account,
      scopes: Array.from(cfg.scopes || ['Files.ReadWrite']),
      redirectUri: redirectUri()
    };

    try {
      const result = await instance.acquireTokenSilent(request);
      return result.accessToken;
    } catch (error) {
      const interactionRequired =
        error instanceof window.msal.InteractionRequiredAuthError ||
        ['interaction_required', 'login_required', 'consent_required'].includes(error?.errorCode);
      if (!interactive || !interactionRequired) throw error;
      const result = await instance.acquireTokenPopup(request);
      if (result.account) instance.setActiveAccount(result.account);
      return result.accessToken;
    }
  }

  async function disconnect() {
    if (!await initialize() || !instance.getActiveAccount()) return;
    const account = instance.getActiveAccount();
    instance.setActiveAccount(null);
    await instance.logoutPopup({ account, postLogoutRedirectUri: redirectUri() });
  }

  function getStatus() {
    const account = instance?.getActiveAccount?.() || null;
    return {
      configured: isConfigured(),
      initialized,
      connected: Boolean(account),
      email: accountEmail(account),
      name: account?.name || ''
    };
  }

  function getAccount() {
    return instance?.getActiveAccount?.() || null;
  }

  window.MicrosoftAuth = Object.freeze({
    initialize,
    connect,
    disconnect,
    acquireToken,
    getStatus,
    getAccount,
    isConfigured,
    redirectUri
  });
})();
