import {
  Activity,
  ActivityType,
  Attachment,
  Brand,
  KPIData,
  MachineRecord,
  Model,
  Section,
} from '../types';
import { sharePointService as mockSharePointService } from './mockSharePointService';

type CreateRecordInput = Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt' | 'attachment'> & {
  attachment?: Attachment | string;
};

interface DictionaryResponse {
  brands: Brand[];
  models: Model[];
  sections: Section[];
  activityTypes: ActivityType[];
  activities: Activity[];
  kpis: KPIData;
}

interface RecordsResponse {
  records: MachineRecord[];
}

interface RecordResponse {
  record: MachineRecord;
}

interface SharePointDataService {
  getBrands: () => Promise<Brand[]>;
  getModels: (brandId?: string) => Promise<Model[]>;
  getSections: (brandId?: string, modelId?: string) => Promise<Section[]>;
  getActivityTypes: () => Promise<ActivityType[]>;
  getActivities: (activityTypeId?: string) => Promise<Activity[]>;
  getRecords: (filters?: Partial<MachineRecord>) => Promise<MachineRecord[]>;
  createRecord: (record: CreateRecordInput) => Promise<MachineRecord>;
  getKPIs: () => Promise<KPIData>;
}

const USE_MOCK_DATA = import.meta.env.VITE_USE_MOCK_DATA === 'true';

class LiveSharePointService implements SharePointDataService {
  private dictionaryCache: Promise<DictionaryResponse> | null = null;

  async getBrands(): Promise<Brand[]> {
    const dictionary = await this.getDictionary();
    return dictionary.brands;
  }

  async getModels(brandId?: string): Promise<Model[]> {
    const dictionary = await this.getDictionary();
    if (!brandId) {
      return dictionary.models;
    }
    return dictionary.models.filter((model) => model.brandId === brandId);
  }

  async getSections(brandId?: string, modelId?: string): Promise<Section[]> {
    const dictionary = await this.getDictionary();
    return dictionary.sections.filter((section) => {
      if (brandId && section.brandId !== brandId) {
        return false;
      }
      if (modelId && section.modelId !== modelId) {
        return false;
      }
      return true;
    });
  }

  async getActivityTypes(): Promise<ActivityType[]> {
    const dictionary = await this.getDictionary();
    return dictionary.activityTypes;
  }

  async getActivities(activityTypeId?: string): Promise<Activity[]> {
    const dictionary = await this.getDictionary();
    if (!activityTypeId) {
      return dictionary.activities;
    }
    return dictionary.activities.filter(
      (activity) => activity.activityTypeId === activityTypeId
    );
  }

  async getRecords(filters?: Partial<MachineRecord>): Promise<MachineRecord[]> {
    const queryParams = new URLSearchParams({ resource: 'records' });

    const allowedFilterKeys: Array<keyof MachineRecord> = [
      'tipoEquipoId',
      'brandId',
      'modelId',
      'sectionId',
      'problem',
      'activityTypeId',
      'activityId',
      'resource',
      'createdBy',
    ];

    allowedFilterKeys.forEach((key) => {
      const rawValue = filters?.[key];
      if (typeof rawValue === 'string' && rawValue.trim()) {
        queryParams.set(key, rawValue.trim());
      }
    });

    const response = await requestJson<RecordsResponse>(`/api/ishikawa?${queryParams.toString()}`);
    return response.records.map((record) => normalizeRecord(record));
  }

  async createRecord(record: CreateRecordInput): Promise<MachineRecord> {
    const payload = {
      record: normalizeCreateRecordInput(record),
    };

    const response = await requestJson<RecordResponse>('/api/ishikawa?resource=records', {
      method: 'POST',
      body: JSON.stringify(payload),
    });

    this.dictionaryCache = null;
    return normalizeRecord(response.record);
  }

  async getKPIs(): Promise<KPIData> {
    const dictionary = await this.getDictionary();
    return dictionary.kpis;
  }

  private async getDictionary(): Promise<DictionaryResponse> {
    this.dictionaryCache ??= requestJson<DictionaryResponse>('/api/ishikawa?resource=dictionary');

    try {
      return await this.dictionaryCache;
    } catch (error) {
      this.dictionaryCache = null;
      throw error;
    }
  }
}

const liveSharePointService = new LiveSharePointService();

const mockServiceAdapter: SharePointDataService = {
  getBrands: () => mockSharePointService.getBrands(),
  getModels: (brandId?: string) => mockSharePointService.getModels(brandId),
  getSections: (brandId?: string, modelId?: string) =>
    mockSharePointService.getSections(brandId, modelId),
  getActivityTypes: () => mockSharePointService.getActivityTypes(),
  getActivities: (activityTypeId?: string) => mockSharePointService.getActivities(activityTypeId),
  getRecords: (filters?: Partial<MachineRecord>) => mockSharePointService.getRecords(filters),
  createRecord: (record: CreateRecordInput) =>
    mockSharePointService.createRecord(normalizeCreateRecordInput(record)),
  getKPIs: () => mockSharePointService.getKPIs(),
};

export const sharePointService: SharePointDataService = USE_MOCK_DATA
  ? mockServiceAdapter
  : liveSharePointService;

async function requestJson<T>(url: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers);
  if (!headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const response = await fetch(url, {
    ...init,
    headers,
  });

  if (response.ok) {
    return (await response.json()) as T;
  }

  const message = await extractHttpError(response);
  throw new Error(message);
}

async function extractHttpError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as { message?: unknown };
    if (typeof body.message === 'string' && body.message.trim()) {
      return body.message;
    }
  } catch {
    return `HTTP ${response.status}: ${response.statusText}`;
  }

  return `HTTP ${response.status}: ${response.statusText}`;
}

function normalizeRecord(record: MachineRecord): MachineRecord {
  return {
    ...record,
    tipoEquipoId: normalizeText(record.tipoEquipoId),
    time: Number.isFinite(Number(record.time)) ? Number(record.time) : 0,
    attachment: normalizeAttachment(record.attachment),
  };
}

function normalizeCreateRecordInput(
  record: CreateRecordInput
): Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt'> {
  return {
    ...record,
    tipoEquipoId: normalizeText(record.tipoEquipoId),
    resource: normalizeText(record.resource),
    createdBy: normalizeText(record.createdBy),
    time: Number.isFinite(Number(record.time)) ? Number(record.time) : 0,
    attachment: normalizeAttachment(record.attachment),
  };
}

function normalizeAttachment(attachment?: Attachment | string): Attachment | undefined {
  if (!attachment) {
    return undefined;
  }

  if (typeof attachment === 'string') {
    const normalizedText = normalizeText(attachment);
    if (!normalizedText) {
      return undefined;
    }
    return {
      id: `attachment-${Date.now().toString()}`,
      name: normalizedText,
      url: normalizedText,
      type: 'text/plain',
      size: 0,
    };
  }

  const normalizedName = normalizeText(attachment.name) || 'Adjunto';
  const normalizedUrl = normalizeText(attachment.url);
  const normalizedType = normalizeText(attachment.type) || 'application/octet-stream';
  const normalizedSize = Number.isFinite(Number(attachment.size))
    ? Number(attachment.size)
    : 0;

  if (!normalizedName && !normalizedUrl) {
    return undefined;
  }

  return {
    id: normalizeText(attachment.id) || `attachment-${Date.now().toString()}`,
    name: normalizedName,
    url: normalizedUrl,
    type: normalizedType,
    size: normalizedSize,
  };
}

function normalizeText(value: unknown): string {
  if (typeof value === 'string') {
    return value.trim();
  }

  if (
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    typeof value === 'bigint'
  ) {
    return String(value);
  }

  return '';
}
