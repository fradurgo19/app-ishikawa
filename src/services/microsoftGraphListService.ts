import type { ClientFieldMap } from '../config/clientSharePointFieldMap';
import type { MachineRecord } from '../types';
import {
  buildDictionaryFromRecords,
  mergeFieldChoiceOptionsFromRecordsAndDictionary,
  type DictionaryFromRecordsShape,
  type FieldChoiceOptionsShape,
} from '../utils/sharePointDictionaryFromRecords';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

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

async function graphFetchJson<T>(url: string, accessToken: string): Promise<T> {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Microsoft Graph ${response.status}: ${body.slice(0, 200)}`);
  }

  return (await response.json()) as T;
}

async function graphPostJson<T>(url: string, accessToken: string, body: unknown): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Microsoft Graph ${response.status}: ${text.slice(0, 200)}`);
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

/**
 * Nombres internos de columna → valores para POST .../items con permisos delegados (Graph).
 * Alineado con buildRecordPayload del servidor (_sharepoint.js).
 */
export function buildGraphListItemFields(
  record: GraphCreateRecordInput,
  fieldMap: ClientFieldMap
): Record<string, string | number> {
  const fields: Record<string, string | number> = {};

  fields[fieldMap.brand] = requireNonEmpty(record.brandId, 'brandId');
  fields[fieldMap.model] = requireNonEmpty(record.modelId, 'modelId');
  fields[fieldMap.section] = requireNonEmpty(record.sectionId, 'sectionId');
  fields[fieldMap.problem] = requireNonEmpty(record.problem, 'problem');
  fields[fieldMap.activityType] = requireNonEmpty(record.activityTypeId, 'activityTypeId');
  fields[fieldMap.activity] = requireNonEmpty(record.activityId, 'activityId');

  if (fieldMap.tipoEquipo?.trim()) {
    fields[fieldMap.tipoEquipo] = requireNonEmpty(record.tipoEquipoId, 'tipoEquipoId');
  }

  if (fieldMap.time?.trim()) {
    fields[fieldMap.time] = normalizeTimeForGraph(record.time);
  }

  const resource = (record.resource ?? '').trim();
  if (fieldMap.resource?.trim() && resource) {
    fields[fieldMap.resource] = resource;
  }

  const createdBy = (record.createdBy ?? '').trim();
  if (fieldMap.createdBy?.trim() && createdBy) {
    fields[fieldMap.createdBy] = createdBy;
  }

  const att = record.attachment;
  if (att) {
    if (fieldMap.attachmentName?.trim()) {
      fields[fieldMap.attachmentName] = att.name;
    }
    if (fieldMap.attachmentUrl?.trim()) {
      fields[fieldMap.attachmentUrl] = att.url;
    }
    if (fieldMap.attachmentType?.trim()) {
      fields[fieldMap.attachmentType] = att.type;
    }
    if (fieldMap.attachmentSize?.trim()) {
      fields[fieldMap.attachmentSize] = att.size;
    }
  }

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

  const created = await graphPostJson<GraphListItem>(createUrl, accessToken, { fields: fieldsPayload });

  const hasFields = created.fields && Object.keys(created.fields).length > 0;
  if (hasFields) {
    return mapGraphListItemToMachineRecord(created, fieldMap);
  }

  const itemId = String(created.id);
  const expanded = await graphFetchJson<GraphListItem>(
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
  const data = await graphFetchJson<{ id: string }>(`${GRAPH_ROOT}/sites/${encoded}`, accessToken);
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
    const page: GraphODataPage<GraphListRef> = await graphFetchJson(nextUrl, accessToken);
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
    const page: GraphODataPage<GraphColumn> = await graphFetchJson(nextUrl, accessToken);
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

function fieldText(fields: Record<string, unknown>, key: string | undefined): string {
  if (!key) {
    return '';
  }
  const v = fields[key];
  if (v === null || v === undefined) {
    return '';
  }
  if (typeof v === 'string') {
    return v.trim();
  }
  if (typeof v === 'number' || typeof v === 'boolean') {
    return String(v);
  }
  if (typeof v === 'object' && v !== null && 'LookupValue' in v) {
    return String((v as { LookupValue?: string }).LookupValue ?? '').trim();
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
    const page: GraphODataPage<GraphListItem> = await graphFetchJson(nextUrl, accessToken);
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
