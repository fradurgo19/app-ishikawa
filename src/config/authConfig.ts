import { BrowserCacheLocation, type Configuration } from '@azure/msal-browser';

const CLIENT_ID = normalizeEnvValue(import.meta.env.VITE_MSAL_CLIENT_ID || import.meta.env.VITE_CLIENT_ID);
const TENANT_ID = normalizeEnvValue(import.meta.env.VITE_MSAL_TENANT_ID || import.meta.env.VITE_TENANT_ID);
const REDIRECT_URI = normalizeEnvValue(
  import.meta.env.VITE_MSAL_REDIRECT_URI || import.meta.env.VITE_REDIRECT_URI
);
const POPUP_REDIRECT_URI = normalizeEnvValue(import.meta.env.VITE_MSAL_POPUP_REDIRECT_URI);
const COORDINATOR_EMAILS = parseCoordinatorEmails(import.meta.env.VITE_COORDINATOR_EMAILS);

export const isMicrosoftAuthEnabled = Boolean(CLIENT_ID && TENANT_ID);

export const msalConfig: Configuration | null = isMicrosoftAuthEnabled
  ? {
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        redirectUri: REDIRECT_URI || getDefaultRedirectUri(),
      },
      cache: {
        cacheLocation: BrowserCacheLocation.LocalStorage,
      },
    }
  : null;

const popupRedirectUri = POPUP_REDIRECT_URI || REDIRECT_URI || getDefaultRedirectUri();

export const loginRequest = {
  scopes: ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All'],
  prompt: 'select_account',
  redirectUri: popupRedirectUri,
};

export const coordinatorEmails = COORDINATOR_EMAILS;

const DELETE_RECORD_ALLOWED_EMAIL_RAW = normalizeEnvValue(import.meta.env.VITE_DELETE_RECORD_ALLOWED_EMAIL);

/** Solo esta cuenta puede eliminar registros en la tabla (validación duplicada en api/ishikawa DELETE). */
export const DELETE_RECORD_ALLOWED_EMAIL = (
  DELETE_RECORD_ALLOWED_EMAIL_RAW || 'jestrada@partequipos.com'
).toLowerCase();

export function canUserDeleteRecords(email: string | undefined): boolean {
  const normalized = (email ?? '').trim().toLowerCase();
  return normalized.length > 0 && normalized === DELETE_RECORD_ALLOWED_EMAIL;
}

function parseCoordinatorEmails(rawValue: string | undefined): string[] {
  const normalizedValue = normalizeEnvValue(rawValue);
  if (!normalizedValue) {
    return [];
  }

  return normalizedValue
    .split(',')
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

function normalizeEnvValue(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }

  return value.trim();
}

function getDefaultRedirectUri(): string {
  if (globalThis.window === undefined) {
    return '/';
  }

  return globalThis.window.location.origin;
}

/**
 * Ámbito delegado para SharePoint REST (`/_api/...`), p. ej. AttachmentFiles/add.
 * Debe coincidir con permisos delegados de la app en Azure (recurso SharePoint del tenant).
 */
export function buildSharePointResourceScope(siteUrl: string): string | null {
  const trimmed = typeof siteUrl === 'string' ? siteUrl.trim() : '';
  if (!trimmed) {
    return null;
  }
  try {
    return `https://${new URL(trimmed).hostname}/.default`;
  } catch {
    return null;
  }
}


