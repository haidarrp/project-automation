import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';

broadcastResponseToMainFrame().catch((error) => {
  console.error('MSAL redirect bridge gagal memproses respons autentikasi:', error);
});
