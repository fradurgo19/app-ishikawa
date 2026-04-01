export interface User {
  id: string;
  username: string;
  role: 'basico' | 'coordinador';
  name: string;
}

export interface Brand {
  id: string;
  name: string;
}

export interface Model {
  id: string;
  name: string;
  brandId: string;
}

export interface Section {
  id: string;
  name: string;
  brandId: string;
  modelId: string;
}

export interface ActivityType {
  id: string;
  name: string;
}

export interface Activity {
  id: string;
  name: string;
  activityTypeId: string;
}

export interface Attachment {
  id: string;
  name: string;
  url: string;
  type: string;
  size: number;
}

/** Archivos en base64 para crear adjuntos nativos en la lista SharePoint (Attachments). */
export interface AttachmentFilePayload {
  name: string;
  contentType: string;
  contentBase64: string;
}

export interface MachineRecord {
  id: string;
  tipoEquipoId: string;
  brandId: string;
  modelId: string;
  sectionId: string;
  problem: string;
  activityTypeId: string;
  activityId: string;
  resource: string;
  /** Primer adjunto (compatibilidad con vistas que solo muestran uno). */
  attachment?: Attachment;
  /** Todos los adjuntos nativos cuando la fuente los expone (p. ej. AttachmentFiles en REST). */
  attachments?: Attachment[];
  time: number;
  createdBy: string;
  createdAt: string;
  updatedAt: string;
}

export interface FishboneNode {
  id: string;
  type: 'marca' | 'modelo' | 'seccion' | 'problema' | 'tipoActividad' | 'actividad' | 'recurso' | 'adjunto' | 'tiempo';
  label: string;
  children: FishboneNode[];
  expanded: boolean;
  data?: unknown;
}

export interface KPIData {
  totalTiposEquipo: number;
  totalMarcas: number;
  totalModelos: number;
  totalSecciones: number;
  totalRegistros: number;
}