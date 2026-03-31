import { PublicClientApplication, type AccountInfo } from '@azure/msal-browser';
import { isMicrosoftAuthEnabled, loginRequest, msalConfig } from '../config/authConfig';

const GRAPH_SCOPES = loginRequest.scopes;

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
  })().catch((error) => {
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

async function getAccountWithRetry(
  maxAttempts = 10,
  delayMs = 120
): Promise<AccountInfo | null> {
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const account = getAccount();
    if (account) {
      return account;
    }

    await sleep(delayMs);
  }

  return null;
}

/**
 * Full-page redirect login. Popup flow was unreliable when the return navigation
 * lost `window.opener`, which caused `handleRedirectPromise` to raise
 * `no_token_request_cache_error` in production.
 */
async function login(): Promise<void> {
  if (!msalInstance) {
    throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
  }

  await initializeAuth();
  await msalInstance.loginRedirect(loginRequest);
}

async function logout(): Promise<void> {
  if (!msalInstance) {
    return;
  }

  await initializeAuth();
  const postLogoutRedirectUri =
    globalThis.window === undefined ? '/login' : `${globalThis.window.location.origin}/login`;
  await msalInstance.logoutRedirect({ postLogoutRedirectUri });
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

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    globalThis.setTimeout(resolve, ms);
  });
}

/**
 * Token para Microsoft Graph (User.Read, Sites.*). Sin sesión MSAL o si falla
 * acquireTokenSilent, devuelve null (el caller puede usar API serverless).
 */
async function acquireGraphAccessToken(): Promise<string | null> {
  if (!msalInstance) {
    return null;
  }

  await initializeAuth();
  const account = getAccount();
  if (!account) {
    return null;
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account,
    });
    return result.accessToken ?? null;
  } catch {
    /* MSAL: sin caché, consentimiento interactivo requerido u otro error; el caller usa API o Graph más tarde. */
    return null;
  }
}

export const authService = {
  initializeAuth,
  getAccount,
  getAccountWithRetry,
  login,
  logout,
  isMicrosoftAuthEnabled,
  acquireGraphAccessToken,
};
