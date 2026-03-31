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

/** Opciones tal como en columnas Choice de SharePoint (texto del valor guardado). */
interface FieldChoiceOptions {
  section: string[];
  activityType: string[];
  activity: string[];
  tipoEquipo?: string[];
  brand?: string[];
  model?: string[];
}

interface DictionaryResponse {
  /** Valores de tipo de equipo vistos en registros + metadatos de lista (misma forma que Brand). */
  tiposEquipo?: Brand[];
  brands: Brand[];
  models: Model[];
  sections: Section[];
  activityTypes: ActivityType[];
  activities: Activity[];
  kpis: KPIData;
  /** Presente en API real: fusionar en cliente si el diccionario por registros queda incompleto. */
  fieldChoiceOptions?: FieldChoiceOptions;
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
  /** Nuevo registro: todas las opciones Choice de Sección, sin filtrar por marca/modelo. */
  getSectionOptionsForNewRecord: () => Promise<Section[]>;
  /** Nuevo registro: listas de texto desde columnas Choice SharePoint + registros (sin matriz local). */
  getNewRecordEquipmentSelectOptions: () => Promise<{
    tipos: string[];
    marcas: string[];
    modelos: string[];
  }>;
  /** Nuevo registro: todas las opciones Choice de Actividad, sin filtrar por tipo de actividad. */
  getActivityOptionsForNewRecord: () => Promise<Activity[]>;
  getActivityTypes: () => Promise<ActivityType[]>;
  getActivities: (activityTypeId?: string) => Promise<Activity[]>;
  getRecords: (filters?: Partial<MachineRecord>) => Promise<MachineRecord[]>;
  createRecord: (record: CreateRecordInput) => Promise<MachineRecord>;
  getKPIs: () => Promise<KPIData>;
  refreshDictionary?: () => Promise<void>;
}

const USE_MOCK_DATA = import.meta.env.VITE_USE_MOCK_DATA === 'true';

/** Origen absoluto del API (producción mismo sitio: vacío). En dev con Vite, suele bastar el proxy /api. */
function apiUrl(pathWithQuery: string): string {
  const raw = (import.meta.env.VITE_API_BASE_URL as string | undefined)?.trim() ?? '';
  const base = raw.replace(/\/$/, '');
  const path = pathWithQuery.startsWith('/') ? pathWithQuery : `/${pathWithQuery}`;
  return base ? `${base}${path}` : path;
}

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
    return dictionary.models.filter((model) => equalsIgnoreCase(model.brandId, brandId));
  }

  async getSections(brandId?: string, modelId?: string): Promise<Section[]> {
    const dictionary = await this.getDictionary();
    const brand = brandId?.trim() ?? '';
    const model = modelId?.trim() ?? '';

    const filtered = dictionary.sections.filter((section) => {
      if (brand && !equalsIgnoreCase(section.brandId, brand)) {
        return false;
      }
      if (model && section.modelId?.trim() && !equalsIgnoreCase(section.modelId, model)) {
        return false;
      }
      return true;
    });

    const raw = dictionary.fieldChoiceOptions?.section;
    if (!brand || !raw?.length) {
      return sortSections(filtered);
    }

    const seen = new Set(filtered.map((s) => s.id.toLowerCase()));
    const merged: Section[] = [...filtered];
    const targetModelId = model;
    for (const label of raw) {
      const id = label.trim();
      if (id && !seen.has(id.toLowerCase())) {
        merged.push({ id, name: id, brandId: brand, modelId: targetModelId });
        seen.add(id.toLowerCase());
      }
    }
    return sortSections(merged);
  }

  async getNewRecordEquipmentSelectOptions(): Promise<{
    tipos: string[];
    marcas: string[];
    modelos: string[];
  }> {
    const dictionary = await this.getDictionary();
    const fc = dictionary.fieldChoiceOptions;
    const tipos = mergeUniqueSortedStrings([
      ...(dictionary.tiposEquipo?.map((t) => t.id) ?? []),
      ...(fc?.tipoEquipo ?? []),
    ]);
    const marcas = mergeUniqueSortedStrings([
      ...dictionary.brands.map((b) => b.id),
      ...(fc?.brand ?? []),
    ]);
    const modelos = mergeUniqueSortedStrings([
      ...dictionary.models.map((m) => m.id),
      ...(fc?.model ?? []),
    ]);
    return { tipos, marcas, modelos };
  }

  async getSectionOptionsForNewRecord(): Promise<Section[]> {
    const dictionary = await this.getDictionary();
    const byKey = new Map<string, Section>();

    for (const s of dictionary.sections) {
      const key = s.id.trim().toLowerCase();
      if (key && !byKey.has(key)) {
        byKey.set(key, { ...s });
      }
    }

    const raw = dictionary.fieldChoiceOptions?.section ?? [];
    for (const label of raw) {
      const id = label.trim();
      if (!id) {
        continue;
      }
      const key = id.toLowerCase();
      if (!byKey.has(key)) {
        byKey.set(key, { id, name: id, brandId: '', modelId: '' });
      }
    }

    return sortSections(Array.from(byKey.values()));
  }

  async getActivityOptionsForNewRecord(): Promise<Activity[]> {
    const dictionary = await this.getDictionary();
    const byKey = new Map<string, Activity>();

    for (const activity of dictionary.activities) {
      const key = activity.id.trim().toLowerCase();
      if (key && !byKey.has(key)) {
        byKey.set(key, { ...activity });
      }
    }

    const raw = dictionary.fieldChoiceOptions?.activity ?? [];
    for (const label of raw) {
      const id = label.trim();
      if (!id) {
        continue;
      }
      const key = id.toLowerCase();
      if (!byKey.has(key)) {
        byKey.set(key, { id, name: id, activityTypeId: '' });
      }
    }

    return sortActivities(Array.from(byKey.values()));
  }

  async getActivityTypes(): Promise<ActivityType[]> {
    const dictionary = await this.getDictionary();
    const base = dictionary.activityTypes;
    const raw = dictionary.fieldChoiceOptions?.activityType;
    if (!raw?.length) {
      return sortActivityTypes(base);
    }

    const byKey = new Map<string, ActivityType>();
    base.forEach((t) => byKey.set(t.id.toLowerCase(), t));
    for (const label of raw) {
      const id = label.trim();
      if (id && !byKey.has(id.toLowerCase())) {
        byKey.set(id.toLowerCase(), { id, name: id });
      }
    }
    return sortActivityTypes(Array.from(byKey.values()));
  }

  async getActivities(activityTypeId?: string): Promise<Activity[]> {
    const dictionary = await this.getDictionary();
    const raw = dictionary.fieldChoiceOptions?.activity;

    if (!activityTypeId) {
      return sortActivities(dictionary.activities);
    }

    const fromDict = dictionary.activities.filter((activity) =>
      equalsIgnoreCase(activity.activityTypeId, activityTypeId)
    );
    if (!raw?.length) {
      return sortActivities(fromDict);
    }

    const seen = new Set(fromDict.map((a) => a.id.toLowerCase()));
    const merged: Activity[] = [...fromDict];
    for (const label of raw) {
      const id = label.trim();
      if (id && !seen.has(id.toLowerCase())) {
        merged.push({ id, name: id, activityTypeId });
        seen.add(id.toLowerCase());
      }
    }
    return sortActivities(merged);
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

    const response = await requestJson<RecordsResponse>(apiUrl(`/api/ishikawa?${queryParams.toString()}`));
    return response.records.map((record) => normalizeRecord(record));
  }

  async createRecord(record: CreateRecordInput): Promise<MachineRecord> {
    const payload = {
      record: normalizeCreateRecordInput(record),
    };

    const response = await requestJson<RecordResponse>(apiUrl('/api/ishikawa?resource=records'), {
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

  async refreshDictionary(): Promise<void> {
    this.dictionaryCache = null;
    await this.getDictionary();
  }

  private async getDictionary(): Promise<DictionaryResponse> {
    this.dictionaryCache ??= requestJson<DictionaryResponse>(apiUrl('/api/ishikawa?resource=dictionary'));

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
  getSectionOptionsForNewRecord: () => mockSharePointService.getSectionOptionsForNewRecord(),
  getNewRecordEquipmentSelectOptions: () =>
    mockSharePointService.getNewRecordEquipmentSelectOptions(),
  getActivityOptionsForNewRecord: () => mockSharePointService.getActivityOptionsForNewRecord(),
  getActivityTypes: () => mockSharePointService.getActivityTypes(),
  getActivities: (activityTypeId?: string) => mockSharePointService.getActivities(activityTypeId),
  getRecords: (filters?: Partial<MachineRecord>) => mockSharePointService.getRecords(filters),
  createRecord: (record: CreateRecordInput) =>
    mockSharePointService.createRecord(normalizeCreateRecordInput(record)),
  getKPIs: () => mockSharePointService.getKPIs(),
  refreshDictionary: async () => {},
};

export const sharePointService: SharePointDataService = USE_MOCK_DATA
  ? mockServiceAdapter
  : liveSharePointService;

function equalsIgnoreCase(a: string | undefined, b: string | undefined): boolean {
  return (a ?? '').trim().toLowerCase() === (b ?? '').trim().toLowerCase();
}

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

function mergeUniqueSortedStrings(values: string[]): string[] {
  const set = new Set<string>();
  for (const v of values) {
    const t = v?.trim();
    if (t) {
      set.add(t);
    }
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

function sortSections(sections: Section[]): Section[] {
  return [...sections].sort((a, b) => a.name.localeCompare(b.name, 'es'));
}

function sortActivityTypes(types: ActivityType[]): ActivityType[] {
  return [...types].sort((a, b) => a.name.localeCompare(b.name, 'es'));
}

function sortActivities(activities: Activity[]): Activity[] {
  return [...activities].sort((a, b) => a.name.localeCompare(b.name, 'es'));
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
