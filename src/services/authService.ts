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

  if (initializePromise) {
    await initializePromise;
    return;
  }

  initializePromise = (async () => {
    await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
      msalInstance.setActiveAccount(redirectResult.account);
      return;
    }

    resolveActiveAccount();
  })()
    .catch((error) => {
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
  if (typeof window === 'undefined') {
    return '/';
  }

  return window.location.origin;
}

export const authService = {
  initializeAuth,
  getAccount,
  login,
  logout,
  isMicrosoftAuthEnabled,
};
