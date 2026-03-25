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

export interface MachineRecord {
  id: string;
  brandId: string;
  modelId: string;
  sectionId: string;
  problem: string;
  activityTypeId: string;
  activityId: string;
  resource: string;
  attachment?: Attachment;
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
  totalMarcas: number;
  totalModelos: number;
  totalSecciones: number;
  totalRegistros: number;
}