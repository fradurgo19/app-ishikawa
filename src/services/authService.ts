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

  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H2',location:'authService.ts:initializeAuth:start',message:'initializeAuth invoked',data:{hasMsalInstance:Boolean(msalInstance),hasInitializePromise:initializePromise!==null,isPopupWindow:isRunningInsidePopup()},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

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

        // #region agent log
        fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H3',location:'authService.ts:initializeAuth:redirectResult',message:'handleRedirectPromise returned account',data:{hasRedirectAccount:true,accountCacheCount:msalInstance.getAllAccounts().length},timestamp:Date.now()})}).catch(()=>{});
        // #endregion
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

async function login(): Promise<AuthenticationResult> {
  if (!msalInstance) {
    throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
  }

  await initializeAuth();

  const loginResult = await msalInstance.loginPopup(loginRequest);
  if (loginResult.account) {
    msalInstance.setActiveAccount(loginResult.account);
  } else {
    await getAccountWithRetry();
  }

  // #region agent log
  fetch('http://127.0.0.1:7840/ingest/2e8455b7-7021-4c1d-9cef-8f2a31248cb9',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'34f201'},body:JSON.stringify({sessionId:'34f201',runId:'msal-loop-pre',hypothesisId:'H3',location:'authService.ts:login:result',message:'loginPopup completed',data:{hasLoginResultAccount:Boolean(loginResult.account),accountCacheCount:msalInstance.getAllAccounts().length,isPopupWindow:isRunningInsidePopup()},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

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

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    globalThis.setTimeout(resolve, ms);
  });
}

export const authService = {
  initializeAuth,
  getAccount,
  getAccountWithRetry,
  login,
  logout,
  isMicrosoftAuthEnabled,
};
