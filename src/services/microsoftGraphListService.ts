import type { ClientFieldMap } from '../config/clientSharePointFieldMap';
import type { Attachment, MachineRecord } from '../types';
import {
  buildDictionaryFromRecords,
  mergeFieldChoiceOptionsFromRecordsAndDictionary,
  type DictionaryFromRecordsShape,
  type FieldChoiceOptionsShape,
} from '../utils/sharePointDictionaryFromRecords';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

const GRAPH_ERROR_BODY_MAX = 200;

interface GraphODataPage<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

interface GraphListItem {
  id: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  fields?: Record<string, unknown>;
}

interface GraphColumn {
  name: string;
  displayName?: string;
  choice?: { choices?: string[] };
}

export interface GraphListBundle {
  records: MachineRecord[];
  dictionary: DictionaryFromRecordsShape & { fieldChoiceOptions: FieldChoiceOptionsShape };
}

function graphHttpErrorMessage(status: number, bodyText: string): string {
  return `Microsoft Graph ${status}: ${bodyText.slice(0, GRAPH_ERROR_BODY_MAX)}`;
}

interface GraphJsonRequestOptions {
  method?: string;
  body?: unknown;
}

async function graphRequestJson<T>(
  url: string,
  accessToken: string,
  options: GraphJsonRequestOptions = {}
): Promise<T> {
  const method = options.method ?? 'GET';
  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json',
  };
  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
  }

  const response = await fetch(url, {
    method,
    headers,
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(graphHttpErrorMessage(response.status, text));
  }

  return (await response.json()) as T;
}

/** Misma forma que el payload del API (sin id ni fechas de servidor). */
export type GraphCreateRecordInput = Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt'>;

function requireNonEmpty(value: string, fieldName: string): string {
  const t = (value ?? '').trim();
  if (!t) {
    throw new Error(`Field "${fieldName}" is required`);
  }
  return t;
}

function normalizeTimeForGraph(raw: number): number {
  if (!Number.isFinite(raw) || raw < 0) {
    throw new Error('Field "time" must be a non-negative number');
  }
  return raw;
}

function setRequiredGraphListFields(
  fields: Record<string, string | number>,
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): void {
  fields[fieldMap.brand] = requireNonEmpty(record.brandId, 'brandId');
  fields[fieldMap.model] = requireNonEmpty(record.modelId, 'modelId');
  fields[fieldMap.section] = requireNonEmpty(record.sectionId, 'sectionId');
  fields[fieldMap.problem] = requireNonEmpty(record.problem, 'problem');
  fields[fieldMap.activityType] = requireNonEmpty(record.activityTypeId, 'activityTypeId');
  fields[fieldMap.activity] = requireNonEmpty(record.activityId, 'activityId');
}

function setOptionalTipoEquipoAndTime(
  fields: Record<string, string | number>,
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): void {
  if (fieldMap.tipoEquipo?.trim()) {
    fields[fieldMap.tipoEquipo] = requireNonEmpty(record.tipoEquipoId, 'tipoEquipoId');
  }
  if (fieldMap.time?.trim()) {
    fields[fieldMap.time] = normalizeTimeForGraph(record.time);
  }
}

function setOptionalResourceAndCreatedBy(
  fields: Record<string, string | number>,
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): void {
  const resource = (record.resource ?? '').trim();
  if (fieldMap.resource?.trim() && resource) {
    fields[fieldMap.resource] = resource;
  }
  const createdBy = (record.createdBy ?? '').trim();
  if (fieldMap.createdBy?.trim() && createdBy) {
    fields[fieldMap.createdBy] = createdBy;
  }
}

function putMappedColumn(
  fields: Record<string, string | number>,
  internalName: string | undefined,
  value: string | number
): void {
  if (internalName?.trim()) {
    fields[internalName] = value;
  }
}

function setOptionalAttachmentGraphFields(
  fields: Record<string, string | number>,
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): void {
  const att = record.attachment;
  if (!att) {
    return;
  }
  putMappedColumn(fields, fieldMap.attachmentName, att.name);
  putMappedColumn(fields, fieldMap.attachmentUrl, att.url);
  putMappedColumn(fields, fieldMap.attachmentType, att.type);
  putMappedColumn(fields, fieldMap.attachmentSize, att.size);
}

/**
 * Nombres internos de columna → valores para POST .../items con permisos delegados (Graph).
 * Alineado con buildRecordPayload del servidor (_sharepoint.js).
 */
export function buildGraphListItemFields(
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): Record<string, string | number> {
  const fields: Record<string, string | number> = {};
  setRequiredGraphListFields(fields, record, fieldMap);
  setOptionalTipoEquipoAndTime(fields, record, fieldMap);
  setOptionalResourceAndCreatedBy(fields, record, fieldMap);
  setOptionalAttachmentGraphFields(fields, record, fieldMap);
  return fields;
}

const GRAPH_ATTACHMENT_FETCH_CHUNK = 8;

interface GraphAttachmentApiRow {
  id?: string;
  name?: string;
  size?: number;
  contentType?: string;
}

async function fetchGraphListItemAttachments(
  siteId: string,
  listId: string,
  itemId: string,
  accessToken: string
): Promise<Attachment[]> {
  const url = `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/attachments`;
  const data = await graphRequestJson<{ value?: GraphAttachmentApiRow[] }>(url, accessToken);
  const rows = data.value || [];
  return rows.map((row, idx) => ({
    id: String(row.id ?? `att-${itemId}-${idx}`),
    name: row.name || 'Adjunto',
    url: '',
    type: row.contentType || 'application/octet-stream',
    size: Number(row.size) || 0,
  }));
}

async function buildGraphNativeAttachmentMap(
  siteId: string,
  listId: string,
  itemIds: string[],
  accessToken: string
): Promise<Map<string, Attachment[]>> {
  const map = new Map<string, Attachment[]>();
  for (let i = 0; i < itemIds.length; i += GRAPH_ATTACHMENT_FETCH_CHUNK) {
    const slice = itemIds.slice(i, i + GRAPH_ATTACHMENT_FETCH_CHUNK);
    await Promise.all(
      slice.map(async (id) => {
        try {
          const list = await fetchGraphListItemAttachments(siteId, listId, id, accessToken);
          if (list.length > 0) {
            map.set(id, list);
          }
        } catch {
          /* Algunos tenants o listas no exponen adjuntos vía Graph; se ignora. */
        }
      })
    );
  }
  return map;
}

export async function createSharePointListItemViaMicrosoftGraph(options: {
  siteUrl: string;
  listName: string;
  fieldMap: ClientFieldMap;
  accessToken: string;
  record: GraphCreateRecordInput;
}): Promise<MachineRecord> {
  const { siteUrl, listName, fieldMap, accessToken, record } = options;

  const siteId = await resolveGraphSiteId(siteUrl, accessToken);
  const listId = await resolveGraphListId(siteId, listName, accessToken);
  const fieldsPayload = buildGraphListItemFields(record, fieldMap);
  const createUrl = `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/items`;

  const created = await graphRequestJson<GraphListItem>(createUrl, accessToken, {
    method: 'POST',
    body: { fields: fieldsPayload },
  });

  const itemId = String(created.id);

  const expanded = await graphRequestJson<GraphListItem>(
    `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
    accessToken
  );
  const nativeList = await fetchGraphListItemAttachments(siteId, listId, itemId, accessToken).catch(
    () => [] as Attachment[]
  );
  return mapGraphListItemToMachineRecord(expanded, fieldMap, nativeList);
}

export async function resolveGraphSiteId(siteUrl: string, accessToken: string): Promise<string> {
  const normalized = siteUrl.replace(/\/$/, '');
  const url = new URL(normalized);
  const hostname = url.hostname;
  const path = url.pathname || '';
  const siteIdentifier = `${hostname}:${path}`;
  const encoded = encodeURIComponent(siteIdentifier);
  const data = await graphRequestJson<{ id: string }>(`${GRAPH_ROOT}/sites/${encoded}`, accessToken);
  if (!data.id) {
    throw new Error('Microsoft Graph: site id not returned');
  }
  return data.id;
}

type GraphListRef = { id: string; displayName?: string; name?: string };

export async function resolveGraphListId(
  siteId: string,
  listDisplayName: string,
  accessToken: string
): Promise<string> {
  const wanted = listDisplayName.trim().toLowerCase();
  const allLists: GraphListRef[] = [];
  let nextUrl: string | null = `${GRAPH_ROOT}/sites/${siteId}/lists`;

  while (nextUrl) {
    const page: GraphODataPage<GraphListRef> = await graphRequestJson(nextUrl, accessToken);
    allLists.push(...(page.value || []));
    nextUrl = page['@odata.nextLink'] ?? null;
  }

  const match = allLists.find(
    (l) => l.displayName?.toLowerCase() === wanted || l.name?.toLowerCase() === wanted
  );
  if (!match?.id) {
    throw new Error(`Microsoft Graph: list "${listDisplayName}" not found`);
  }
  return match.id;
}

async function fetchAllGraphColumns(
  siteId: string,
  listId: string,
  accessToken: string
): Promise<GraphColumn[]> {
  const all: GraphColumn[] = [];
  let nextUrl: string | null = `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/columns`;

  while (nextUrl) {
    const page: GraphODataPage<GraphColumn> = await graphRequestJson(nextUrl, accessToken);
    all.push(...(page.value || []));
    nextUrl = page['@odata.nextLink'] ?? null;
  }

  return all;
}

function extractChoiceOptionsFromColumns(
  columns: GraphColumn[],
  fieldMap: ClientFieldMap
): FieldChoiceOptionsShape {
  const byName = new Map(columns.map((c) => [c.name, c]));

  const choicesFor = (columnKey: keyof ClientFieldMap): string[] => {
    const internal = fieldMap[columnKey]?.trim();
    if (!internal) {
      return [];
    }
    const col = byName.get(internal);
    const raw = col?.choice?.choices;
    return Array.isArray(raw) ? [...raw] : [];
  };

  return {
    section: choicesFor('section'),
    activityType: choicesFor('activityType'),
    activity: choicesFor('activity'),
    tipoEquipo: choicesFor('tipoEquipo'),
    brand: choicesFor('brand'),
    model: choicesFor('model'),
  };
}

function textFromFieldScalar(value: unknown): string {
  if (typeof value === 'string') {
    return value.trim();
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  return '';
}

function textFromLookupField(value: object): string {
  if ('LookupValue' in value) {
    return String((value as { LookupValue?: string }).LookupValue ?? '').trim();
  }
  return '';
}

function fieldText(fields: Record<string, unknown>, key: string | undefined): string {
  if (!key) {
    return '';
  }
  const value = fields[key];
  if (value === null || value === undefined) {
    return '';
  }
  const scalar = textFromFieldScalar(value);
  if (scalar !== '' || typeof value === 'string') {
    return scalar;
  }
  if (typeof value === 'object') {
    return textFromLookupField(value);
  }
  return '';
}

function fieldNumber(fields: Record<string, unknown>, key: string | undefined): number {
  if (!key) {
    return 0;
  }
  const v = fields[key];
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function extractCustomAttachmentFromGraphFields(
  fields: Record<string, unknown>,
  fieldMap: ClientFieldMap,
  itemId: string
): Attachment | undefined {
  const attachmentName = fieldText(fields, fieldMap.attachmentName);
  const attachmentUrl = fieldText(fields, fieldMap.attachmentUrl);
  const attachmentType = fieldText(fields, fieldMap.attachmentType);
  const attachmentSize = fieldMap.attachmentSize ? fieldNumber(fields, fieldMap.attachmentSize) : 0;
  if (!attachmentName && !attachmentUrl) {
    return undefined;
  }
  return {
    id: `attachment-${itemId}-custom`,
    name: attachmentName || 'Adjunto',
    url: attachmentUrl,
    type: attachmentType || 'application/octet-stream',
    size: attachmentSize,
  };
}

export function mapGraphListItemToMachineRecord(
  item: GraphListItem,
  fieldMap: ClientFieldMap,
  nativeAttachments?: Attachment[]
): MachineRecord {
  const f = item.fields || {};
  const idStr = String(item.id);
  const createdBy = fieldMap.createdBy
    ? fieldText(f, fieldMap.createdBy)
    : '';

  const natives = nativeAttachments?.filter((a) => a?.name) ?? [];
  const custom = extractCustomAttachmentFromGraphFields(f, fieldMap, idStr);
  let resolvedList: Attachment[];
  if (natives.length > 0) {
    resolvedList = natives;
  } else if (custom) {
    resolvedList = [custom];
  } else {
    resolvedList = [];
  }
  const first = resolvedList[0];

  return {
    id: idStr,
    tipoEquipoId: fieldMap.tipoEquipo ? fieldText(f, fieldMap.tipoEquipo) : '',
    brandId: fieldText(f, fieldMap.brand),
    modelId: fieldText(f, fieldMap.model),
    sectionId: fieldText(f, fieldMap.section),
    problem: fieldText(f, fieldMap.problem),
    activityTypeId: fieldText(f, fieldMap.activityType),
    activityId: fieldText(f, fieldMap.activity),
    resource: fieldMap.resource ? fieldText(f, fieldMap.resource) : '',
    time: fieldMap.time ? fieldNumber(f, fieldMap.time) : 0,
    createdBy: createdBy || 'system',
    createdAt: item.createdDateTime || new Date().toISOString(),
    updatedAt: item.lastModifiedDateTime || new Date().toISOString(),
    ...(first ? { attachment: first, attachments: resolvedList } : {}),
  };
}

async function fetchAllGraphListItems(
  siteId: string,
  listId: string,
  accessToken: string
): Promise<GraphListItem[]> {
  const all: GraphListItem[] = [];
  let nextUrl: string | null =
    `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=5000&$orderby=id desc`;

  while (nextUrl) {
    const page: GraphODataPage<GraphListItem> = await graphRequestJson(nextUrl, accessToken);
    all.push(...(page.value || []));
    nextUrl = page['@odata.nextLink'] ?? null;
  }

  return all;
}

export async function fetchSharePointListViaMicrosoftGraph(options: {
  siteUrl: string;
  listName: string;
  fieldMap: ClientFieldMap;
  accessToken: string;
}): Promise<GraphListBundle> {
  const { siteUrl, listName, fieldMap, accessToken } = options;

  const siteId = await resolveGraphSiteId(siteUrl, accessToken);
  const listId = await resolveGraphListId(siteId, listName, accessToken);

  const [columns, items] = await Promise.all([
    fetchAllGraphColumns(siteId, listId, accessToken),
    fetchAllGraphListItems(siteId, listId, accessToken),
  ]);

  const itemIds = items.map((it) => String(it.id));
  const attachmentMap = await buildGraphNativeAttachmentMap(siteId, listId, itemIds, accessToken);

  const records = items.map((item) => {
    const id = String(item.id);
    return mapGraphListItemToMachineRecord(item, fieldMap, attachmentMap.get(id));
  });
  const baseDictionary = buildDictionaryFromRecords(records);
  const graphChoices = extractChoiceOptionsFromColumns(columns, fieldMap);
  const fieldChoiceOptions = mergeFieldChoiceOptionsFromRecordsAndDictionary(
    baseDictionary,
    records,
    graphChoices
  );

  return {
    records,
    dictionary: {
      ...baseDictionary,
      fieldChoiceOptions,
    },
  };
}
