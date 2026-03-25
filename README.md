# App Ishikawa - Deploy en Vercel con SharePoint

## Estado actual del proyecto

- Frontend: `Vite + React + TypeScript`.
- Ruteo SPA con `react-router-dom`.
- Estado y formularios con `React hooks` + `react-hook-form`.
- Antes de este ajuste, la capa de datos usaba solo `mockSharePointService`.
- Ahora existe integración real con SharePoint mediante funciones serverless de Vercel (`api/ishikawa.js`), preservando la misma interfaz de servicio usada por las pantallas.

## Arquitectura de datos aplicada

- **Frontend**: `src/services/sharePointService.ts`
  - Conserva métodos: `getBrands`, `getModels`, `getSections`, `getActivityTypes`, `getActivities`, `getRecords`, `createRecord`, `getKPIs`.
  - Consume `/api/ishikawa` y mantiene cache de diccionario para reducir llamadas repetidas.
  - Soporta `VITE_USE_MOCK_DATA=true` para desarrollo local sin backend.

- **Backend serverless en Vercel**:
  - `api/ishikawa.js`: endpoint principal para `records` y `dictionary`.
  - `api/_sharepoint.js`: autenticación OAuth2 (client credentials), lectura/escritura en SharePoint y transformación de datos.

## Variables de entorno requeridas

Usa `.env.example` como plantilla.

### Requeridas

- `SHAREPOINT_SITE_URL`  
  Ejemplo: `https://partequipos2.sharepoint.com/sites/servicioposventa`
- `SHAREPOINT_LIST_TITLE`  
  Ejemplo: `Ishikawa`
- `SHAREPOINT_TENANT_ID`
- `SHAREPOINT_CLIENT_ID`
- `SHAREPOINT_CLIENT_SECRET`

### Frontend

- `VITE_USE_MOCK_DATA=false` en producción.

### Mapeo de columnas (si aplica)

Si los nombres internos de columnas en SharePoint no coinciden con los defaults, define:

- `SHAREPOINT_FIELD_BRAND`
- `SHAREPOINT_FIELD_MODEL`
- `SHAREPOINT_FIELD_SECTION`
- `SHAREPOINT_FIELD_PROBLEM`
- `SHAREPOINT_FIELD_ACTIVITY_TYPE`
- `SHAREPOINT_FIELD_ACTIVITY`
- `SHAREPOINT_FIELD_RESOURCE`
- `SHAREPOINT_FIELD_TIME`
- `SHAREPOINT_FIELD_CREATED_BY`
- `SHAREPOINT_FIELD_ATTACHMENT_NAME`
- `SHAREPOINT_FIELD_ATTACHMENT_URL`
- `SHAREPOINT_FIELD_ATTACHMENT_TYPE`
- `SHAREPOINT_FIELD_ATTACHMENT_SIZE`

## Despliegue a Vercel

1. Instala dependencias:
   - `npm install`
2. Valida local:
   - `npm run lint`
   - `npm run typecheck`
   - `npm run build`
3. En Vercel, crea proyecto desde este repositorio/carpeta.
4. Carga las variables de entorno (Production).
5. Ejecuta deploy.
6. Valida endpoints:
   - `GET /api/ishikawa?resource=dictionary`
   - `GET /api/ishikawa?resource=records`

## Notas de seguridad

- Las credenciales de SharePoint/Azure AD viven solo en entorno serverless de Vercel.
- El frontend no expone `client_secret`.
- El endpoint devuelve mensajes controlados de error para facilitar diagnóstico sin filtrar secretos.
