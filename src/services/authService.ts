import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from '@azure/msal-browser';
import { isMicrosoftAuthEnabled, loginRequest, msalConfig } from '../config/authConfig';

const MICROSOFT_AUTH_MISCONFIGURED_MESSAGE =
  'Microsoft authentication is not configured. Set VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID.';

const msalInstance = msalConfig ? new PublicClientApplication(msalConfig) : null;

let initializePromise: Promise<void> | null = null;

async function initializeAuth(): Promise<void> {
  if (!msalInstance) {
    return;
  }

  if (initializePromise !== null) {
    await initializePromise;
    return;
  }

  initializePromise = (async () => {
    await msalInstance.initialize();

    if (!isRunningInsidePopup()) {
      const redirectResult = await msalInstance.handleRedirectPromise();
      if (redirectResult?.account) {
        msalInstance.setActiveAccount(redirectResult.account);
        return;
      }
    }

    resolveActiveAccount();
  })()
    .catch((error) => {
      if (isNoTokenRequestCacheError(error)) {
        return;
      }

      initializePromise = null;
      throw error;
    });

  await initializePromise;
}

function getAccount(): AccountInfo | null {
  if (!msalInstance) {
    return null;
  }

  return resolveActiveAccount();
}

async function login(): Promise<AuthenticationResult> {
  if (!msalInstance) {
    throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
  }

  await initializeAuth();

  const loginResult = await msalInstance.loginPopup(loginRequest);
  if (loginResult.account) {
    msalInstance.setActiveAccount(loginResult.account);
  } else {
    resolveActiveAccount();
  }

  return loginResult;
}

async function logout(): Promise<void> {
  if (!msalInstance) {
    return;
  }

  await initializeAuth();
  await msalInstance.logoutPopup({
    mainWindowRedirectUri: getDefaultRedirectUri(),
  });
}

function resolveActiveAccount(): AccountInfo | null {
  if (!msalInstance) {
    return null;
  }

  const activeAccount = msalInstance.getActiveAccount();
  if (activeAccount) {
    return activeAccount;
  }

  const [firstAccount] = msalInstance.getAllAccounts();
  if (firstAccount) {
    msalInstance.setActiveAccount(firstAccount);
    return firstAccount;
  }

  return null;
}

function getDefaultRedirectUri(): string {
  if (globalThis.window === undefined) {
    return '/';
  }

  return globalThis.window.location.origin;
}

function isRunningInsidePopup(): boolean {
  if (globalThis.window === undefined) {
    return false;
  }

  return Boolean(globalThis.window.opener && globalThis.window.opener !== globalThis.window);
}

function isNoTokenRequestCacheError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  return error.message.includes('no_token_request_cache_error');
}

export const authService = {
  initializeAuth,
  getAccount,
  login,
  logout,
  isMicrosoftAuthEnabled,
};
