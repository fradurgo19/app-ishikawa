import { PublicClientApplication, type AccountInfo } from '@azure/msal-browser';
import { isMicrosoftAuthEnabled, loginRequest, msalConfig } from '../config/authConfig';

const MICROSOFT_AUTH_MISCONFIGURED_MESSAGE =
  'Microsoft authentication is not configured. Set VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID.';

const msalInstance = msalConfig ? new PublicClientApplication(msalConfig) : null;

let initializePromise: Promise<void> | null = null;

async function initializeAuth(): Promise<void> {
  if (!msalInstance) {
    return;
  }

  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H2',location:'authService.ts:initializeAuth:start',message:'initializeAuth invoked',data:{hasMsalInstance:Boolean(msalInstance),hasInitializePromise:initializePromise!==null,isPopupWindow:isRunningInsidePopup()},timestamp:Date.now()})}).catch(()=>{});
  // #endregion
  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run5',hypothesisId:'H9',location:'authService.ts:initializeAuth:configSnapshot',message:'MSAL runtime configuration snapshot',data:{cacheLocation:msalConfig?.cache?.cacheLocation ?? 'none',appRedirectUri:msalConfig?.auth?.redirectUri ?? 'none',loginRequestRedirectUri:loginRequest.redirectUri},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  if (initializePromise !== null) {
    await initializePromise;
    return;
  }

  initializePromise = (async () => {
    await msalInstance.initialize();

    // #region agent log
    fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run3',hypothesisId:'H7',location:'authService.ts:initializeAuth:beforeRedirectHandling',message:'Before handleRedirectPromise storage/hash snapshot',data:{hasHash:globalThis.window === undefined ? false : globalThis.window.location.hash.length > 0,hashHasCode:globalThis.window === undefined ? false : globalThis.window.location.hash.includes('code='),sessionStorageMsalKeys:countStorageKeys('sessionStorage','msal'),localStorageMsalKeys:countStorageKeys('localStorage','msal')},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    // #region agent log
    fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run5',hypothesisId:'H10',location:'authService.ts:initializeAuth:windowContext',message:'Window context before redirect handling',data:{isPopupWindow:isRunningInsidePopup(),hasWindowOpener:hasWindowOpener(),windowName:getWindowName(),path:getWindowPath(),referrer:getReferrerHost()},timestamp:Date.now()})}).catch(()=>{});
    // #endregion

    if (!isRunningInsidePopup()) {
      const redirectResult = await msalInstance.handleRedirectPromise();
      if (redirectResult?.account) {
        msalInstance.setActiveAccount(redirectResult.account);

        // #region agent log
        fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H3',location:'authService.ts:initializeAuth:redirectResult',message:'handleRedirectPromise returned account',data:{hasRedirectAccount:true,accountCacheCount:msalInstance.getAllAccounts().length},timestamp:Date.now()})}).catch(()=>{});
        // #endregion
        return;
      }
    }

    const activeAccount = resolveActiveAccount();
    // #region agent log
    fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run2',hypothesisId:'H6',location:'authService.ts:initializeAuth:postResolve',message:'Post-initialize account resolution',data:{hasActiveAccount:Boolean(activeAccount),accountCacheCount:msalInstance.getAllAccounts().length,isPopupWindow:isRunningInsidePopup()},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
  })()
    .catch((error) => {
      if (isNoTokenRequestCacheError(error)) {
        // #region agent log
        fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run3',hypothesisId:'H7',location:'authService.ts:initializeAuth:noTokenCacheError',message:'No token request cache error raised during redirect handling',data:{errorName:error instanceof Error ? error.name : 'unknown',errorMessage:error instanceof Error ? error.message : String(error),sessionStorageMsalKeys:countStorageKeys('sessionStorage','msal'),localStorageMsalKeys:countStorageKeys('localStorage','msal')},timestamp:Date.now()})}).catch(()=>{});
        // #endregion
        // #region agent log
        fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-run5',hypothesisId:'H11',location:'authService.ts:initializeAuth:noTokenWindowContext',message:'Window context when no token cache error is raised',data:{isPopupWindow:isRunningInsidePopup(),hasWindowOpener:hasWindowOpener(),windowName:getWindowName(),path:getWindowPath(),referrer:getReferrerHost()},timestamp:Date.now()})}).catch(()=>{});
        // #endregion
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
      // #region agent log
      fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H4',location:'authService.ts:getAccountWithRetry:resolved',message:'Account resolved during retry',data:{attempt:attempt+1,maxAttempts,hasAccount:true},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      return account;
    }

    await sleep(delayMs);
  }

  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H4',location:'authService.ts:getAccountWithRetry:exhausted',message:'Account retry exhausted',data:{maxAttempts,delayMs,hasAccount:false},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  return null;
}

/**
 * Full-page redirect login (not popup).
 * Production logs showed popup authorize URLs but return windows with no `window.opener`, so
 * `handleRedirectPromise` ran in the wrong interaction context and raised `no_token_request_cache_error`.
 * Redirect keeps PKCE/state in the same tab that started login. @see debug-34f201.log L19-L21.
 */
async function login(): Promise<void> {
  if (!msalInstance) {
    throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
  }

  await initializeAuth();

  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'post-fix-redirect',hypothesisId:'H12',location:'authService.ts:login:beforeRedirect',message:'Calling loginRedirect',data:{redirectUri:loginRequest.redirectUri,accountCacheCount:msalInstance.getAllAccounts().length},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  try {
    await msalInstance.loginRedirect(loginRequest);
  } catch (error) {
    // #region agent log
    fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'post-fix-redirect',hypothesisId:'H12',location:'authService.ts:login:redirectError',message:'loginRedirect rejected',data:{errorName:error instanceof Error ? error.name : 'unknown',errorMessage:error instanceof Error ? error.message : String(error)},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    throw error;
  }
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

function countStorageKeys(
  storageType: 'sessionStorage' | 'localStorage',
  prefix: string
): number {
  if (globalThis.window === undefined) {
    return -1;
  }

  const storage = storageType === 'sessionStorage'
    ? globalThis.window.sessionStorage
    : globalThis.window.localStorage;

  let total = 0;
  for (let index = 0; index < storage.length; index += 1) {
    const key = storage.key(index);
    if (typeof key === 'string' && key.includes(prefix)) {
      total += 1;
    }
  }

  return total;
}

function hasWindowOpener(): boolean {
  if (globalThis.window === undefined) {
    return false;
  }

  return Boolean(globalThis.window.opener);
}

function getWindowName(): string {
  if (globalThis.window === undefined) {
    return '';
  }

  return globalThis.window.name;
}

function getWindowPath(): string {
  if (globalThis.window === undefined) {
    return '';
  }

  return globalThis.window.location.pathname;
}

function getReferrerHost(): string {
  const referrer = globalThis.window?.document?.referrer;
  if (!referrer) {
    return '';
  }

  try {
    return new URL(referrer).hostname;
  } catch {
    return '';
  }
}

export const authService = {
  initializeAuth,
  getAccount,
  getAccountWithRetry,
  login,
  logout,
  isMicrosoftAuthEnabled,
};
