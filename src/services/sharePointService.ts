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
import { isMicrosoftAuthEnabled } from '../config/authConfig';
import { getClientFieldMap } from '../config/clientSharePointFieldMap';
import { filterMachineRecords } from '../utils/filterMachineRecords';
import { authService } from './authService';
import {
  createSharePointListItemViaMicrosoftGraph,
  fetchSharePointListViaMicrosoftGraph,
  loadGraphListItemAsMachineRecord,
  updateSharePointListItemViaMicrosoftGraph,
} from './microsoftGraphListService';
import type { GraphCreateRecordInput } from './microsoftGraphListService';
import { uploadListItemAttachmentRest } from './sharePointRestAttachments';
import { filesToAttachmentPayloads } from '../utils/attachmentFilePayload';
import { getDistinctModeloIdsFromMatrix } from '../data/equipmentMatrix';
import { sharePointService as mockSharePointService } from './mockSharePointService';

type CreateRecordInput = Omit<
  MachineRecord,
  'id' | 'createdAt' | 'updatedAt' | 'attachment' | 'attachments'
> & {
  attachment?: Attachment | string;
  /** Archivos locales: tras crear el ítem se suben por SharePoint REST con token del sitio (o fallback API en base64). */
  attachmentFiles?: File[];
};

/** Registro persistible sin id ni marcas de tiempo de servidor (Graph / POST /api). */
type MachineRecordWithoutMeta = Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt'>;

type CreateRecordFieldsInput = Omit<CreateRecordInput, 'attachmentFiles'>;

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

/** Opciones para modificar adjuntos nativos al actualizar (PATCH /api con credenciales de aplicación). */
export interface UpdateRecordAttachmentOptions {
  addFiles?: File[];
  removeAttachmentFileNames?: string[];
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
  /** Actualiza un ítem existente (Graph delegado o PATCH /api/ishikawa). */
  updateRecord: (
    record: MachineRecord,
    attachmentOptions?: UpdateRecordAttachmentOptions
  ) => Promise<MachineRecord>;
  getKPIs: () => Promise<KPIData>;
  refreshDictionary?: () => Promise<void>;
}

const USE_MOCK_DATA = import.meta.env.VITE_USE_MOCK_DATA === 'true';

function graphListEnvSiteUrl(): string {
  return normalizeViteEnvString(import.meta.env.VITE_SHAREPOINT_SITE_URL);
}

function graphListEnvListName(): string {
  return (
    normalizeViteEnvString(import.meta.env.VITE_SHAREPOINT_LIST_NAME) ||
    normalizeViteEnvString(import.meta.env.VITE_SHAREPOINT_LIST_TITLE)
  );
}

function normalizeViteEnvString(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

/**
 * Sin credenciales de aplicación en servidor: crear/leer lista solo con token delegado (MSAL).
 * Si falla Graph, no se reintenta POST /api/ishikawa; si no hay sesión, crear falla con mensaje claro.
 */
function isDelegatedOnlySharePointMode(): boolean {
  return import.meta.env.VITE_SHAREPOINT_DELEGATED_ONLY === 'true';
}

/** Origen absoluto del API (producción mismo sitio: vacío). En dev con Vite, suele bastar el proxy /api. */
function apiUrl(pathWithQuery: string): string {
  const raw = (import.meta.env.VITE_API_BASE_URL ?? '').trim();
  const base = raw.replace(/\/$/, '');
  const path = pathWithQuery.startsWith('/') ? pathWithQuery : `/${pathWithQuery}`;
  return base ? `${base}${path}` : path;
}

function isMicrosoftGraphListExplicitlyDisabled(): boolean {
  return import.meta.env.VITE_USE_MICROSOFT_GRAPH_LIST === 'false';
}

interface GraphListContext {
  siteUrl: string;
  listName: string;
}

interface PublicSharePointConfigShape {
  siteUrl: string;
  listTitle: string;
}

let publicSharePointConfigPromise: Promise<PublicSharePointConfigShape> | null = null;

async function loadPublicSharePointConfig(): Promise<PublicSharePointConfigShape> {
  try {
    const response = await fetch(apiUrl('/api/sharepoint-public-config'));
    if (!response.ok) {
      return { siteUrl: '', listTitle: '' };
    }
    const data = (await response.json()) as { siteUrl?: unknown; listTitle?: unknown };
    return {
      siteUrl: typeof data.siteUrl === 'string' ? data.siteUrl.trim() : '',
      listTitle: typeof data.listTitle === 'string' ? data.listTitle.trim() : '',
    };
  } catch {
    return { siteUrl: '', listTitle: '' };
  }
}

function getPublicSharePointConfig(): Promise<PublicSharePointConfigShape> {
  publicSharePointConfigPromise ??= loadPublicSharePointConfig();
  return publicSharePointConfigPromise;
}

/**
 * Contexto para Graph: VITE_SHAREPOINT_* en build, o respaldo desde /api/sharepoint-public-config
 * (lee SHAREPOINT_SITE_URL + SHAREPOINT_LIST_TITLE del servidor sin secretos).
 * Así el mismo despliegue en Vercel sirve para API serverless y para lectura delegada en el navegador.
 */
async function resolveGraphListContext(): Promise<GraphListContext | null> {
  if (isMicrosoftGraphListExplicitlyDisabled()) {
    return null;
  }
  if (!isMicrosoftAuthEnabled) {
    return null;
  }

  const viteSite = graphListEnvSiteUrl();
  const viteList = graphListEnvListName();
  let siteUrl = viteSite;
  let listName = viteList;

  if (!siteUrl || !listName) {
    const pub = await getPublicSharePointConfig();
    if (!siteUrl) {
      siteUrl = pub.siteUrl;
    }
    if (!listName) {
      listName = pub.listTitle;
    }
  }

  if (!siteUrl || !listName) {
    return null;
  }

  return { siteUrl, listName };
}

const ISHIKAWA_RECORDS_PATH = '/api/ishikawa?resource=records';

const RECORD_QUERY_FILTER_KEYS: Array<keyof MachineRecord> = [
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

const MSG_DELEGATED_SIGNIN_REQUIRED =
  'Inicie sesión con Microsoft para guardar registros. Este entorno usa solo permisos delegados.';

const POST_CREATE_ATTACHMENT_DELAY_MS = 500;

async function delayMs(ms: number): Promise<void> {
  await new Promise((resolve) => {
    globalThis.setTimeout(resolve, ms);
  });
}

function delegatedGraphCreateErrorMessage(graphError: unknown): string {
  const detail =
    graphError instanceof Error ? graphError.message : 'Error desconocido al crear en Graph';
  return `No se pudo crear el registro con Microsoft Graph (permisos delegados). ${detail}`;
}

function delegatedGraphUpdateErrorMessage(graphError: unknown): string {
  const detail =
    graphError instanceof Error ? graphError.message : 'Error desconocido al actualizar en Graph';
  return `No se pudo actualizar el registro con Microsoft Graph (permisos delegados). ${detail}`;
}

function machineRecordToGraphCreateInput(record: MachineRecord): GraphCreateRecordInput {
  return {
    tipoEquipoId: record.tipoEquipoId,
    brandId: record.brandId,
    modelId: record.modelId,
    sectionId: record.sectionId,
    problem: record.problem,
    activityTypeId: record.activityTypeId,
    activityId: record.activityId,
    resource: record.resource,
    time: record.time,
    createdBy: record.createdBy,
    attachment: record.attachment,
    attachments: record.attachments,
  };
}

class LiveSharePointService implements SharePointDataService {
  private dictionaryCache: Promise<DictionaryResponse> | null = null;

  /** Datos de lista cargados con Microsoft Graph (diccionario + registros en una sola carga). */
  private graphDataLoader: Promise<{ dictionary: DictionaryResponse; records: MachineRecord[] }> | null =
    null;

  private async loadSharePointListFromGraph(
    accessToken: string,
    ctx: GraphListContext
  ): Promise<{
    dictionary: DictionaryResponse;
    records: MachineRecord[];
  }> {
    let sharePointAccessToken: string | null = null;
    try {
      await authService.initializeAuth();
      sharePointAccessToken = await authService.acquireSharePointAccessToken(ctx.siteUrl);
    } catch {
      sharePointAccessToken = null;
    }

    const bundle = await fetchSharePointListViaMicrosoftGraph({
      siteUrl: ctx.siteUrl,
      listName: ctx.listName,
      fieldMap: getClientFieldMap(),
      accessToken: accessToken,
      sharePointAccessToken,
    });

    const dictionary: DictionaryResponse = {
      tiposEquipo: bundle.dictionary.tiposEquipo,
      brands: bundle.dictionary.brands,
      models: bundle.dictionary.models,
      sections: bundle.dictionary.sections,
      activityTypes: bundle.dictionary.activityTypes,
      activities: bundle.dictionary.activities,
      kpis: bundle.dictionary.kpis,
      fieldChoiceOptions: bundle.dictionary.fieldChoiceOptions,
    };

    return {
      dictionary,
      records: bundle.records.map((r) => normalizeRecord(r)),
    };
  }

  /**
   * Si el modo Graph aplica y hay token, devuelve diccionario + registros; si no, null (usar /api).
   */
  private invalidateCaches(): void {
    this.dictionaryCache = null;
    this.graphDataLoader = null;
  }

  /**
   * Persistencia vía Graph con token delegado. null = usar POST /api/ishikawa.
   * Lanza si modo delegado-only sin token o si Graph falla y no hay fallback.
   */
  private async tryPersistRecordViaMicrosoftGraph(
    normalized: MachineRecordWithoutMeta
  ): Promise<MachineRecord | null> {
    const ctx = await resolveGraphListContext();
    if (!ctx) {
      return null;
    }

    await authService.initializeAuth();
    await authService.getAccountWithRetry();
    const token = await authService.acquireGraphAccessToken();
    if (!token) {
      if (isDelegatedOnlySharePointMode()) {
        throw new Error(MSG_DELEGATED_SIGNIN_REQUIRED);
      }
      return null;
    }

    try {
      return await createSharePointListItemViaMicrosoftGraph({
        siteUrl: ctx.siteUrl,
        listName: ctx.listName,
        fieldMap: getClientFieldMap(),
        accessToken: token,
        record: normalized,
      });
    } catch (graphError) {
      if (isDelegatedOnlySharePointMode()) {
        throw new Error(delegatedGraphCreateErrorMessage(graphError));
      }
      return null;
    }
  }

  private async ensureGraphListData(): Promise<{
    dictionary: DictionaryResponse;
    records: MachineRecord[];
  } | null> {
    const ctx = await resolveGraphListContext();
    if (!ctx) {
      return null;
    }

    await authService.initializeAuth();
    await authService.getAccountWithRetry();
    const token = await authService.acquireGraphAccessToken();
    if (!token) {
      return null;
    }

    this.graphDataLoader ??= this.loadSharePointListFromGraph(token, ctx);

    try {
      return await this.graphDataLoader;
    } catch {
      this.graphDataLoader = null;
      return null;
    }
  }

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
      ...getDistinctModeloIdsFromMatrix(),
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
    const graphData = await this.ensureGraphListData();
    if (graphData) {
      return filterMachineRecords(graphData.records, filters ?? {});
    }

    const queryParams = new URLSearchParams({ resource: 'records' });

    RECORD_QUERY_FILTER_KEYS.forEach((key) => {
      const rawValue = filters?.[key];
      if (typeof rawValue === 'string' && rawValue.trim()) {
        queryParams.set(key, rawValue.trim());
      }
    });

    const response = await requestJson<RecordsResponse>(apiUrl(`/api/ishikawa?${queryParams.toString()}`));
    return response.records.map((record) => normalizeRecord(record));
  }

  /**
   * Sin token REST del sitio: mismo JSON + base64 que /api/ishikawa (credenciales de aplicación en servidor).
   */
  private async createRecordSendingAttachmentsViaApi(
    normalized: MachineRecordWithoutMeta,
    files: File[]
  ): Promise<MachineRecord> {
    const payloads = await filesToAttachmentPayloads(files);
    const response = await requestJson<RecordResponse>(apiUrl(ISHIKAWA_RECORDS_PATH), {
      method: 'POST',
      body: JSON.stringify({ record: normalized, attachmentFiles: payloads }),
    });
    this.invalidateCaches();
    return normalizeRecord(response.record);
  }

  private async createRecordWithoutAttachmentFiles(
    normalized: MachineRecordWithoutMeta
  ): Promise<MachineRecord> {
    const graphRecord = await this.tryPersistRecordViaMicrosoftGraph(normalized);
    if (graphRecord) {
      this.invalidateCaches();
      return normalizeRecord(graphRecord);
    }
    const response = await requestJson<RecordResponse>(apiUrl(ISHIKAWA_RECORDS_PATH), {
      method: 'POST',
      body: JSON.stringify({ record: normalized }),
    });
    this.invalidateCaches();
    return normalizeRecord(response.record);
  }

  /**
   * Con token: crear ítem (Graph o API), esperar, subir cada File por SharePoint REST (binario), como en VehicleFormReal.
   */
  private async createRecordThenUploadAttachmentsViaRest(
    normalized: MachineRecordWithoutMeta,
    files: File[],
    ctx: GraphListContext,
    sharePointRestToken: string
  ): Promise<MachineRecord> {
    const graphRecord = await this.tryPersistRecordViaMicrosoftGraph(normalized);
    let created: MachineRecord;
    if (graphRecord) {
      created = graphRecord;
    } else {
      const response = await requestJson<RecordResponse>(apiUrl(ISHIKAWA_RECORDS_PATH), {
        method: 'POST',
        body: JSON.stringify({ record: normalized }),
      });
      created = normalizeRecord(response.record);
    }

    await delayMs(POST_CREATE_ATTACHMENT_DELAY_MS);
    await this.uploadAttachmentsAfterCreate(created, files, ctx, sharePointRestToken);

    this.invalidateCaches();
    const graphAccess = await authService.acquireGraphAccessToken();
    if (graphAccess && ctx) {
      try {
        const refreshed = await loadGraphListItemAsMachineRecord({
          siteUrl: ctx.siteUrl,
          listName: ctx.listName,
          fieldMap: getClientFieldMap(),
          accessToken: graphAccess,
          itemId: created.id,
          sharePointAccessToken: sharePointRestToken,
        });
        return normalizeRecord(refreshed);
      } catch {
        /* Adjuntos ya en lista; la siguiente carga de datos los mostrará */
      }
    }
    return normalizeRecord(created);
  }

  private async uploadAttachmentsAfterCreate(
    created: MachineRecord,
    files: File[],
    ctx: GraphListContext,
    sharePointRestToken: string
  ): Promise<void> {
    for (const file of files) {
      try {
        await uploadListItemAttachmentRest({
          siteUrl: ctx.siteUrl,
          listTitle: ctx.listName,
          itemId: created.id,
          file,
          accessToken: sharePointRestToken,
        });
      } catch (uploadErr) {
        const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr);
        throw new Error(
          `El registro se creó (id ${created.id}) pero falló la subida de adjuntos por SharePoint REST. ${detail} Si el error es de red o CORS, configura el token del sitio en Azure o usa el envío por API sin sesión.`
        );
      }
    }
  }

  async createRecord(record: CreateRecordInput): Promise<MachineRecord> {
    const { attachmentFiles, ...recordFields } = record;
    const files = attachmentFiles;
    const normalized = normalizeCreateRecordInput(recordFields);
    const hasFiles = Boolean(files?.length);

    const ctx = await resolveGraphListContext();
    let sharePointRestToken: string | null = null;
    if (hasFiles && ctx) {
      await authService.initializeAuth();
      sharePointRestToken = await authService.acquireSharePointAccessToken(ctx.siteUrl);
    }
    const useClientRestForAttachments = Boolean(hasFiles && ctx && sharePointRestToken);

    if (hasFiles && !useClientRestForAttachments) {
      return this.createRecordSendingAttachmentsViaApi(normalized, files!);
    }

    if (!hasFiles) {
      return this.createRecordWithoutAttachmentFiles(normalized);
    }

    return this.createRecordThenUploadAttachmentsViaRest(normalized, files!, ctx!, sharePointRestToken!);
  }

  private async updateRecordViaApiWithAttachments(
    record: MachineRecord,
    attachmentOptions: UpdateRecordAttachmentOptions
  ): Promise<MachineRecord> {
    const id = record.id.trim();
    const normalizedFields = normalizeCreateRecordInput(
      machineRecordToGraphCreateInput(record) as CreateRecordFieldsInput
    );
    const addFiles = attachmentOptions.addFiles ?? [];
    const payloads = addFiles.length > 0 ? await filesToAttachmentPayloads(addFiles) : [];
    const removeNames = (attachmentOptions.removeAttachmentFileNames ?? [])
      .map((n) => n.trim())
      .filter(Boolean);

    const body: Record<string, unknown> = {
      record: { ...record, ...normalizedFields, id },
    };
    if (payloads.length > 0) {
      body.attachmentFiles = payloads;
    }
    if (removeNames.length > 0) {
      body.removeAttachmentFileNames = removeNames;
    }

    const response = await requestJson<RecordResponse>(apiUrl(ISHIKAWA_RECORDS_PATH), {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    this.invalidateCaches();
    return normalizeRecord(response.record);
  }

  async updateRecord(
    record: MachineRecord,
    attachmentOptions?: UpdateRecordAttachmentOptions
  ): Promise<MachineRecord> {
    const id = record.id?.trim();
    if (!id) {
      throw new Error('El registro debe tener id para actualizar.');
    }
    const hasAttachmentOps =
      Boolean(attachmentOptions?.addFiles?.length) ||
      Boolean(attachmentOptions?.removeAttachmentFileNames?.length);

    if (hasAttachmentOps && attachmentOptions) {
      return this.updateRecordViaApiWithAttachments(record, attachmentOptions);
    }

    const normalizedFields = normalizeCreateRecordInput(
      machineRecordToGraphCreateInput(record) as CreateRecordFieldsInput
    );

    const ctx = await resolveGraphListContext();
    await authService.initializeAuth();
    await authService.getAccountWithRetry();
    const token = await authService.acquireGraphAccessToken();

    if (ctx && token) {
      try {
        const updated = await updateSharePointListItemViaMicrosoftGraph({
          siteUrl: ctx.siteUrl,
          listName: ctx.listName,
          fieldMap: getClientFieldMap(),
          accessToken: token,
          itemId: id,
          record: normalizedFields,
        });
        this.invalidateCaches();
        return normalizeRecord(updated);
      } catch (graphError) {
        if (isDelegatedOnlySharePointMode()) {
          throw new Error(delegatedGraphUpdateErrorMessage(graphError));
        }
      }
    }

    const response = await requestJson<RecordResponse>(apiUrl(ISHIKAWA_RECORDS_PATH), {
      method: 'PATCH',
      body: JSON.stringify({ record: { ...record, ...normalizedFields, id } }),
    });
    this.invalidateCaches();
    return normalizeRecord(response.record);
  }

  async getKPIs(): Promise<KPIData> {
    const dictionary = await this.getDictionary();
    return dictionary.kpis;
  }

  async refreshDictionary(): Promise<void> {
    this.invalidateCaches();
    await this.getDictionary();
  }

  private async getDictionary(): Promise<DictionaryResponse> {
    const graphData = await this.ensureGraphListData();
    if (graphData) {
      return graphData.dictionary;
    }

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
    mockSharePointService.createRecord(toMockCreateRecordPayload(record)),
  updateRecord: (record, attachmentOptions) =>
    mockSharePointService.updateRecord(record, attachmentOptions),
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
  if (response.status === 413) {
    return 'El tamaño del envío supera el límite del servidor (p. ej. ~4,5 MB por solicitud en Vercel). Usa archivos más pequeños o menos adjuntos por registro.';
  }
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

function toMockCreateRecordPayload(
  record: CreateRecordInput
): MachineRecordWithoutMeta & { attachmentFiles?: File[] } {
  const { attachmentFiles, ...recordFields } = record;
  const normalized = normalizeCreateRecordInput(recordFields);
  return { ...normalized, attachmentFiles };
}

function normalizeRecord(record: MachineRecord): MachineRecord {
  const primary = normalizeAttachment(record.attachment);
  const listNorm =
    record.attachments
      ?.map((a) => normalizeAttachment(a))
      .filter((x): x is Attachment => Boolean(x)) ?? undefined;
  let mergedList: Attachment[] | undefined;
  if (listNorm?.length) {
    mergedList = listNorm;
  } else if (primary) {
    mergedList = [primary];
  } else {
    mergedList = undefined;
  }
  const first = primary ?? mergedList?.[0];

  return {
    ...record,
    tipoEquipoId: normalizeText(record.tipoEquipoId),
    time: Number.isFinite(Number(record.time)) ? Number(record.time) : 0,
    attachment: first,
    attachments: mergedList?.length ? mergedList : undefined,
  };
}

function normalizeCreateRecordInput(record: CreateRecordFieldsInput): MachineRecordWithoutMeta {
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
