import { Brand, Model, Section, ActivityType, Activity, MachineRecord, KPIData } from '../types';

// Datos simulados
const mockBrands: Brand[] = [
  { id: '1', name: 'Caterpillar' },
  { id: '2', name: 'Komatsu' },
  { id: '3', name: 'Hitachi' },
  { id: '4', name: 'Liebherr' },
];

const mockModels: Model[] = [
  { id: '1', name: '320D', brandId: '1' },
  { id: '2', name: '330C', brandId: '1' },
  { id: '3', name: 'PC200', brandId: '2' },
  { id: '4', name: 'PC300', brandId: '2' },
  { id: '5', name: 'ZX200', brandId: '3' },
  { id: '6', name: 'R920', brandId: '4' },
];

const mockSections: Section[] = [
  { id: '1', name: 'Motor', brandId: '1', modelId: '1' },
  { id: '2', name: 'Hidráulicos', brandId: '1', modelId: '1' },
  { id: '3', name: 'Transmisión', brandId: '1', modelId: '1' },
  { id: '4', name: 'Motor', brandId: '1', modelId: '2' },
  { id: '5', name: 'Sistema Eléctrico', brandId: '2', modelId: '3' },
  { id: '6', name: 'Sistema de Refrigeración', brandId: '3', modelId: '5' },
];

const mockActivityTypes: ActivityType[] = [
  { id: '1', name: 'Mantenimiento' },
  { id: '2', name: 'Reparación' },
  { id: '3', name: 'Inspección' },
  { id: '4', name: 'Reemplazo' },
];

const mockActivities: Activity[] = [
  { id: '1', name: 'Cambio de Aceite', activityTypeId: '1' },
  { id: '2', name: 'Reemplazo de Filtro', activityTypeId: '1' },
  { id: '3', name: 'Reparación de Componente', activityTypeId: '2' },
  { id: '4', name: 'Inspección Visual', activityTypeId: '3' },
  { id: '5', name: 'Reemplazo de Pieza', activityTypeId: '4' },
];

const mockRecords: MachineRecord[] = [
  {
    id: '1',
    brandId: '1',
    modelId: '1',
    sectionId: '1',
    problem: 'Sobrecalentamiento del motor',
    activityTypeId: '2',
    activityId: '3',
    resource: 'Manual del Motor - Sección 4.2',
    time: 120,
    createdBy: '1',
    createdAt: '2024-01-15T10:30:00Z',
    updatedAt: '2024-01-15T10:30:00Z',
  },
  {
    id: '2',
    brandId: '1',
    modelId: '1',
    sectionId: '2',
    problem: 'Pérdida de presión hidráulica',
    activityTypeId: '1',
    activityId: '1',
    resource: 'Guía del Sistema Hidráulico',
    time: 90,
    createdBy: '1',
    createdAt: '2024-01-16T14:15:00Z',
    updatedAt: '2024-01-16T14:15:00Z',
  },
  {
    id: '3',
    brandId: '2',
    modelId: '3',
    sectionId: '5',
    problem: 'Falla del sistema eléctrico',
    activityTypeId: '3',
    activityId: '4',
    resource: 'Manual de Solución de Problemas Eléctricos',
    time: 45,
    createdBy: '2',
    createdAt: '2024-01-17T09:00:00Z',
    updatedAt: '2024-01-17T09:00:00Z',
  },
];

class MockSharePointService {
  async getBrands(): Promise<Brand[]> {
    await this.delay(300);
    return mockBrands;
  }

  async getModels(brandId?: string): Promise<Model[]> {
    await this.delay(300);
    return brandId ? mockModels.filter(m => m.brandId === brandId) : mockModels;
  }

  async getSections(brandId?: string, modelId?: string): Promise<Section[]> {
    await this.delay(300);
    let filtered = mockSections;
    if (brandId) filtered = filtered.filter(s => s.brandId === brandId);
    if (modelId) filtered = filtered.filter(s => s.modelId === modelId);
    return filtered;
  }

  async getActivityTypes(): Promise<ActivityType[]> {
    await this.delay(300);
    return mockActivityTypes;
  }

  async getActivities(activityTypeId?: string): Promise<Activity[]> {
    await this.delay(300);
    return activityTypeId ? mockActivities.filter(a => a.activityTypeId === activityTypeId) : mockActivities;
  }

  async getRecords(filters?: Partial<MachineRecord>): Promise<MachineRecord[]> {
    await this.delay(500);
    let filtered = mockRecords;
    
    if (filters) {
      Object.entries(filters).forEach(([key, value]) => {
        if (value) {
          filtered = filtered.filter(record => 
            record[key as keyof MachineRecord] === value
          );
        }
      });
    }
    
    return filtered;
  }

  async createRecord(record: Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt'>): Promise<MachineRecord> {
    await this.delay(800);
    const newRecord: MachineRecord = {
      ...record,
      id: Date.now().toString(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    mockRecords.push(newRecord);
    return newRecord;
  }

  async getKPIs(): Promise<KPIData> {
    await this.delay(200);
    return {
      totalMarcas: mockBrands.length,
      totalModelos: mockModels.length,
      totalSecciones: mockSections.length,
      totalRegistros: mockRecords.length,
    };
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

export const sharePointService = new MockSharePointService();