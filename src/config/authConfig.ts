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

const DEFAULT_DELETE_RECORD_ALLOWED_EMAILS = Object.freeze([
  'jestrada@partequipos.com',
  'analista.mantenimiento@partequipos.com',
]);

const DELETE_RECORD_ALLOWED_EMAILS_RAW = normalizeEnvValue(
  import.meta.env.VITE_DELETE_RECORD_ALLOWED_EMAILS ?? import.meta.env.VITE_DELETE_RECORD_ALLOWED_EMAIL
);

function parseCommaSeparatedLowerEmails(rawValue: string): string[] {
  return rawValue
    .split(',')
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

/**
 * Cuentas que pueden eliminar registros en la tabla (validación duplicada en api/ishikawa DELETE).
 * Sobrescribible con `VITE_DELETE_RECORD_ALLOWED_EMAILS` (lista separada por comas).
 */
export const deleteRecordAllowedEmails: readonly string[] = (() => {
  if (!DELETE_RECORD_ALLOWED_EMAILS_RAW) {
    return [...DEFAULT_DELETE_RECORD_ALLOWED_EMAILS];
  }
  const parsed = parseCommaSeparatedLowerEmails(DELETE_RECORD_ALLOWED_EMAILS_RAW);
  return parsed.length > 0 ? parsed : [...DEFAULT_DELETE_RECORD_ALLOWED_EMAILS];
})();

export function canUserDeleteRecords(email: string | undefined): boolean {
  const normalized = (email ?? '').trim().toLowerCase();
  return normalized.length > 0 && deleteRecordAllowedEmails.includes(normalized);
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


