/**
 * Nombres de columna en Graph `fields` / SharePoint (internal name).
 * Alinear con VITE_SHAREPOINT_FIELD_* y con la lista real (o field_N si Graph lo expone así).
 */
export interface ClientFieldMap {
  tipoEquipo: string;
  brand: string;
  model: string;
  section: string;
  problem: string;
  activityType: string;
  activity: string;
  resource: string;
  time: string;
  createdBy: string;
  attachmentName: string;
  attachmentUrl: string;
  attachmentType: string;
  attachmentSize: string;
}

export function getClientFieldMap(): ClientFieldMap {
  return {
    tipoEquipo: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_TIPO_EQUIPO) || 'TipoEquipo',
    brand: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_BRAND) || 'Marca',
    model: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_MODEL) || 'Modelo',
    section: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_SECTION) || 'Seccion',
    problem: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_PROBLEM) || 'Problema',
    activityType: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ACTIVITY_TYPE) || 'TipoActividad',
    activity: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ACTIVITY) || 'Actividad',
    resource: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_RESOURCE) || 'Recurso',
    time: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_TIME) || 'Tiempo',
    createdBy: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_CREATED_BY) || '',
    attachmentName: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ATTACHMENT_NAME) || '',
    attachmentUrl: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ATTACHMENT_URL) || '',
    attachmentType: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ATTACHMENT_TYPE) || '',
    attachmentSize: normalizeEnv(import.meta.env.VITE_SHAREPOINT_FIELD_ATTACHMENT_SIZE) || '',
  };
}

function normalizeEnv(value: string | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}
