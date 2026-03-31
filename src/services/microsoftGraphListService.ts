import type { ClientFieldMap } from '../config/clientSharePointFieldMap';
import type { MachineRecord } from '../types';
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
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
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

  const hasFields = created.fields && Object.keys(created.fields).length > 0;
  if (hasFields) {
    return mapGraphListItemToMachineRecord(created, fieldMap);
  }

  const itemId = String(created.id);
  const expanded = await graphRequestJson<GraphListItem>(
    `${GRAPH_ROOT}/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
    accessToken
  );
  return mapGraphListItemToMachineRecord(expanded, fieldMap);
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
    (l) =>
      (l.displayName && l.displayName.toLowerCase() === wanted) ||
      (l.name && l.name.toLowerCase() === wanted)
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

export function mapGraphListItemToMachineRecord(
  item: GraphListItem,
  fieldMap: ClientFieldMap
): MachineRecord {
  const f = item.fields || {};
  const createdBy = fieldMap.createdBy
    ? fieldText(f, fieldMap.createdBy)
    : '';

  return {
    id: String(item.id),
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

  const records = items.map((item) => mapGraphListItemToMachineRecord(item, fieldMap));
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
