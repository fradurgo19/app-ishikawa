/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_USE_MOCK_DATA?: string;
  /** Origen del API (p. ej. https://xxx.vercel.app). Vacío = mismas rutas relativas /api (producción o proxy Vite). */
  readonly VITE_API_BASE_URL?: string;
  readonly VITE_MSAL_CLIENT_ID?: string;
  readonly VITE_MSAL_TENANT_ID?: string;
  readonly VITE_MSAL_REDIRECT_URI?: string;
  readonly VITE_MSAL_POPUP_REDIRECT_URI?: string;
  readonly VITE_COORDINATOR_EMAILS?: string;
  readonly VITE_CLIENT_ID?: string;
  readonly VITE_TENANT_ID?: string;
  readonly VITE_REDIRECT_URI?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
