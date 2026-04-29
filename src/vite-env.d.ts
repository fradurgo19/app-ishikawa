/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_USE_MOCK_DATA?: string;
  /** Origen del API (p. ej. https://xxx.vercel.app). Vacío = mismas rutas relativas /api (producción o proxy Vite). */
  readonly VITE_API_BASE_URL?: string;
  /**
   * URL completa del sitio SharePoint (misma idea que SHAREPOINT_SITE_URL en servidor).
   * Con MSAL + nombre de lista, la app puede leer ítems vía Microsoft Graph en el navegador.
   */
  readonly VITE_SHAREPOINT_SITE_URL?: string;
  /** Nombre o displayName de la lista (prioridad sobre VITE_SHAREPOINT_LIST_TITLE). */
  readonly VITE_SHAREPOINT_LIST_NAME?: string;
  /** Alias opcional del nombre de lista (si VITE_SHAREPOINT_LIST_NAME está vacío). */
  readonly VITE_SHAREPOINT_LIST_TITLE?: string;
  /** `false` desactiva Graph y usa solo GET /api/ishikawa para diccionario y registros. */
  readonly VITE_USE_MICROSOFT_GRAPH_LIST?: string;
  /**
   * `true`: lectura/escritura de lista solo con token MSAL (Graph); sin reintento a /api/ishikawa al crear
   * si Graph falla; crear sin sesión muestra error explícito. Útil sin permisos de aplicación en el servidor.
   */
  readonly VITE_SHAREPOINT_DELEGATED_ONLY?: string;
  /** Nombres internos de columnas (alinear con Graph fields / lista SharePoint). */
  readonly VITE_SHAREPOINT_FIELD_TIPO_EQUIPO?: string;
  readonly VITE_SHAREPOINT_FIELD_BRAND?: string;
  readonly VITE_SHAREPOINT_FIELD_MODEL?: string;
  readonly VITE_SHAREPOINT_FIELD_SECTION?: string;
  readonly VITE_SHAREPOINT_FIELD_PROBLEM?: string;
  readonly VITE_SHAREPOINT_FIELD_ACTIVITY_TYPE?: string;
  readonly VITE_SHAREPOINT_FIELD_ACTIVITY?: string;
  readonly VITE_SHAREPOINT_FIELD_RESOURCE?: string;
  readonly VITE_SHAREPOINT_FIELD_TIME?: string;
  readonly VITE_SHAREPOINT_FIELD_CREATED_BY?: string;
  readonly VITE_SHAREPOINT_FIELD_ATTACHMENT_NAME?: string;
  readonly VITE_SHAREPOINT_FIELD_ATTACHMENT_URL?: string;
  readonly VITE_SHAREPOINT_FIELD_ATTACHMENT_TYPE?: string;
  readonly VITE_SHAREPOINT_FIELD_ATTACHMENT_SIZE?: string;
  readonly VITE_MSAL_CLIENT_ID?: string;
  readonly VITE_MSAL_TENANT_ID?: string;
  readonly VITE_MSAL_REDIRECT_URI?: string;
  readonly VITE_MSAL_POPUP_REDIRECT_URI?: string;
  readonly VITE_COORDINATOR_EMAILS?: string;
  /** Lista separada por comas; si falta, se usan jestrada y analista.mantenimiento @partequipos.com. */
  readonly VITE_DELETE_RECORD_ALLOWED_EMAILS?: string;
  /** Compatibilidad: mismo uso que VITE_DELETE_RECORD_ALLOWED_EMAILS. */
  readonly VITE_DELETE_RECORD_ALLOWED_EMAIL?: string;
  readonly VITE_CLIENT_ID?: string;
  readonly VITE_TENANT_ID?: string;
  readonly VITE_REDIRECT_URI?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
