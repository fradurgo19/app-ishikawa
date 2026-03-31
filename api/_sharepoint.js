const REQUIRED_ENV_VARS = Object.freeze([
  'SHAREPOINT_SITE_URL',
  'SHAREPOINT_LIST_TITLE',
  'SHAREPOINT_TENANT_ID',
  'SHAREPOINT_CLIENT_ID',
  'SHAREPOINT_CLIENT_SECRET',
]);

const DEFAULT_FIELD_MAP = Object.freeze({
  tipoEquipo: 'TipoEquipo',
  brand: 'Marca',
  model: 'Modelo',
  section: 'Seccion',
  problem: 'Problema',
  activityType: 'TipoActividad',
  activity: 'Actividad',
  resource: 'Recurso',
  time: 'Tiempo',
  createdBy: '',
  attachmentName: '',
  attachmentUrl: '',
  attachmentType: '',
  attachmentSize: '',
});

const MAX_PAGINATION_REQUESTS = 25;
const DEFAULT_PAGE_SIZE = 5000;

export function createHttpError(statusCode, message, details) {
  const error = new Error(message);
  error.statusCode = statusCode;
  if (details !== undefined) {
    error.details = details;
  }
  return error;
}

export function getSharePointConfig() {
  const missingEnvVars = REQUIRED_ENV_VARS.filter((key) => !process.env[key]);
  if (missingEnvVars.length > 0) {
    throw createHttpError(
      500,
      `Missing required environment variables: ${missingEnvVars.join(', ')}`
    );
  }

  let parsedSiteUrl;
  try {
    parsedSiteUrl = new URL(process.env.SHAREPOINT_SITE_URL);
  } catch (error) {
    throw createHttpError(500, 'SHAREPOINT_SITE_URL is not a valid URL', error);
  }

  const fieldMap = {
    tipoEquipo: process.env.SHAREPOINT_FIELD_TIPO_EQUIPO || DEFAULT_FIELD_MAP.tipoEquipo,
    brand: process.env.SHAREPOINT_FIELD_BRAND || DEFAULT_FIELD_MAP.brand,
    model: process.env.SHAREPOINT_FIELD_MODEL || DEFAULT_FIELD_MAP.model,
    section: process.env.SHAREPOINT_FIELD_SECTION || DEFAULT_FIELD_MAP.section,
    problem: process.env.SHAREPOINT_FIELD_PROBLEM || DEFAULT_FIELD_MAP.problem,
    activityType: process.env.SHAREPOINT_FIELD_ACTIVITY_TYPE || DEFAULT_FIELD_MAP.activityType,
    activity: process.env.SHAREPOINT_FIELD_ACTIVITY || DEFAULT_FIELD_MAP.activity,
    resource: process.env.SHAREPOINT_FIELD_RESOURCE || DEFAULT_FIELD_MAP.resource,
    time: process.env.SHAREPOINT_FIELD_TIME || DEFAULT_FIELD_MAP.time,
    createdBy: process.env.SHAREPOINT_FIELD_CREATED_BY || DEFAULT_FIELD_MAP.createdBy,
    attachmentName: process.env.SHAREPOINT_FIELD_ATTACHMENT_NAME || DEFAULT_FIELD_MAP.attachmentName,
    attachmentUrl: process.env.SHAREPOINT_FIELD_ATTACHMENT_URL || DEFAULT_FIELD_MAP.attachmentUrl,
    attachmentType: process.env.SHAREPOINT_FIELD_ATTACHMENT_TYPE || DEFAULT_FIELD_MAP.attachmentType,
    attachmentSize: process.env.SHAREPOINT_FIELD_ATTACHMENT_SIZE || DEFAULT_FIELD_MAP.attachmentSize,
  };

  const requiredFieldAliases = ['brand', 'model', 'section', 'problem', 'activityType', 'activity'];
  const missingRequiredFieldAliases = requiredFieldAliases.filter((alias) => !fieldMap[alias]);
  if (missingRequiredFieldAliases.length > 0) {
    throw createHttpError(
      500,
      `Missing SharePoint field mapping for: ${missingRequiredFieldAliases.join(', ')}`
    );
  }

  return {
    siteUrl: process.env.SHAREPOINT_SITE_URL.replace(/\/$/, ''),
    siteOrigin: parsedSiteUrl.origin,
    listTitle: getTextValue(process.env.SHAREPOINT_LIST_TITLE),
    tenantId: process.env.SHAREPOINT_TENANT_ID,
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
    tokenScope: `https://${parsedSiteUrl.hostname}/.default`,
    pageSize: normalizePositiveInteger(process.env.SHAREPOINT_PAGE_SIZE, DEFAULT_PAGE_SIZE),
    fieldMap,
  };
}

export async function fetchAllListItems(config) {
  const preferExpand =
    process.env.SHAREPOINT_LIST_ITEMS_EXPAND_ATTACHMENTS !== 'false';
  const allowFieldTextFallback = process.env.SHAREPOINT_LIST_ITEMS_FIELDTEXT_FALLBACK !== 'false';

  let items;
  try {
    if (preferExpand) {
      try {
        items = await fetchAllListItemsPaginated(config, true);
      } catch {
        items = await fetchAllListItemsPaginated(config, false);
      }
    } else {
      items = await fetchAllListItemsPaginated(config, false);
    }
  } catch (primaryErr) {
    if (!allowFieldTextFallback) {
      throw primaryErr;
    }
    try {
      return await fetchAllListItemsFieldValuesAsTextPaginated(config);
    } catch {
      throw primaryErr;
    }
  }

  if (allowFieldTextFallback && Array.isArray(items) && items.length === 0) {
    try {
      const viaText = await fetchAllListItemsFieldValuesAsTextPaginated(config);
      if (viaText.length > 0) {
        return viaText;
      }
    } catch {
      /* mantener [] si la lista está vacía o FVA también falla */
    }
  }

  return items;
}

async function fetchAllListItemsFieldValuesAsTextPaginated(config) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const baseItemsUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items`;
  const queryParams = new URLSearchParams();
  queryParams.set('$top', String(config.pageSize));
  queryParams.set('$select', 'Id,Created,Modified,AuthorId,Attachments');
  queryParams.set('$expand', 'FieldValuesAsText');

  const allItems = [];
  let nextUrl = `${baseItemsUrl}?${queryParams.toString()}`;
  let requestCount = 0;

  while (nextUrl && requestCount < MAX_PAGINATION_REQUESTS) {
    requestCount += 1;
    const responsePayload = await sharePointRequest(config, accessToken, {
      method: 'GET',
      url: nextUrl,
    });

    if (Array.isArray(responsePayload.value)) {
      allItems.push(...responsePayload.value);
    }

    const responseNextLink =
      responsePayload['@odata.nextLink'] ||
      responsePayload['odata.nextLink'] ||
      responsePayload.d?.__next ||
      null;
    nextUrl = typeof responseNextLink === 'string' ? responseNextLink : null;
  }

  if (requestCount >= MAX_PAGINATION_REQUESTS && nextUrl) {
    throw createHttpError(
      502,
      `Pagination limit reached (${MAX_PAGINATION_REQUESTS} requests). Narrow list size or increase limit.`
    );
  }

  return allItems;
}

async function fetchAllListItemsPaginated(config, expandAttachmentFiles) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const baseItemsUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items`;
  const schemaRows = await fetchListFieldsRows(config);
  const initialQueryParams = buildListItemsQueryParams(
    config.fieldMap,
    config.pageSize,
    expandAttachmentFiles,
    schemaRows
  );
  const initialUrl = `${baseItemsUrl}?${initialQueryParams.toString()}`;

  const allItems = [];
  let nextUrl = initialUrl;
  let requestCount = 0;

  while (nextUrl && requestCount < MAX_PAGINATION_REQUESTS) {
    requestCount += 1;
    const responsePayload = await sharePointRequest(config, accessToken, {
      method: 'GET',
      url: nextUrl,
    });

    if (Array.isArray(responsePayload.value)) {
      allItems.push(...responsePayload.value);
    }

    const responseNextLink =
      responsePayload['@odata.nextLink'] ||
      responsePayload['odata.nextLink'] ||
      responsePayload.d?.__next ||
      null;
    nextUrl = typeof responseNextLink === 'string' ? responseNextLink : null;
  }

  if (requestCount >= MAX_PAGINATION_REQUESTS && nextUrl) {
    throw createHttpError(
      502,
      `Pagination limit reached (${MAX_PAGINATION_REQUESTS} requests). Narrow list size or increase limit.`
    );
  }

  return allItems;
}

/**
 * Solo incluye en $select columnas que existen en /fields. Un InternalName erróneo en OData
 * rechaza toda la petición y en producción safeLoadMappedRecords devolvía [] sin registros visibles.
 *
 * @param {Array<{InternalName?: string}>} schemaRows filas de fetchListFieldsRows (vacío = sin filtrar).
 * @param {boolean} expandAttachmentFiles Si es false, no usa $expand=AttachmentFiles (evita 400/502 en listas donde el expand falla o no aplica).
 */
function buildListItemsQueryParams(fieldMap, pageSize, expandAttachmentFiles = true, schemaRows = []) {
  const selectFields = new Set(['Id', 'Created', 'Modified', 'AuthorId', 'Attachments']);

  const canonicalByLower = new Map();
  if (Array.isArray(schemaRows) && schemaRows.length > 0) {
    for (const row of schemaRows) {
      const internal = getTextValue(row.InternalName);
      if (internal) {
        canonicalByLower.set(internal.toLowerCase(), internal);
      }
    }
  }

  const mappedNames = Object.values(fieldMap)
    .map((v) => getTextValue(v))
    .filter(Boolean);

  if (canonicalByLower.size === 0) {
    mappedNames.forEach((name) => selectFields.add(name));
  } else {
    for (const name of mappedNames) {
      const canonical = canonicalByLower.get(name.toLowerCase());
      if (canonical) {
        selectFields.add(canonical);
      }
    }
  }

  if (expandAttachmentFiles) {
    selectFields.add('AttachmentFiles');
    selectFields.add('AttachmentFiles/FileName');
    selectFields.add('AttachmentFiles/ServerRelativeUrl');
    selectFields.add('AttachmentFiles/Length');
  }

  const selectedFieldList = Array.from(selectFields).filter(Boolean).join(',');

  const queryParams = new URLSearchParams();
  queryParams.set('$top', String(pageSize));
  queryParams.set('$select', selectedFieldList);
  if (expandAttachmentFiles) {
    queryParams.set('$expand', 'AttachmentFiles');
  }

  return queryParams;
}

export async function createListItem(config, payload) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const createUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items`;

  return sharePointRequest(config, accessToken, {
    method: 'POST',
    url: createUrl,
    body: payload,
  });
}

function getItemFieldText(item, internalName) {
  const key = getTextValue(internalName);
  if (!key || !item || typeof item !== 'object') {
    return '';
  }
  const direct = item[key];
  if (direct !== undefined && direct !== null) {
    const t = getTextValue(direct);
    if (t) {
      return t;
    }
  }
  const fvt = item.FieldValuesAsText || item.fieldValuesAsText;
  if (fvt && typeof fvt === 'object') {
    return getTextValue(fvt[key]);
  }
  return '';
}

function getItemFieldNumeric(item, internalName) {
  const key = getTextValue(internalName);
  if (!key || !item || typeof item !== 'object') {
    return 0;
  }
  const direct = item[key];
  if (direct !== undefined && direct !== null && direct !== '') {
    return getNumericValue(direct);
  }
  const fvt = item.FieldValuesAsText || item.fieldValuesAsText;
  if (fvt && typeof fvt === 'object' && fvt[key] !== undefined && fvt[key] !== null && fvt[key] !== '') {
    return getNumericValue(fvt[key]);
  }
  return 0;
}

export function mapListItemToMachineRecord(item, fieldMap, siteOrigin = '') {
  const nativeAttachment = extractNativeAttachment(item, siteOrigin);
  const customAttachment = extractCustomAttachment(item, fieldMap);
  const resolvedAttachment = nativeAttachment || customAttachment;

  const mappedRecord = {
    id: getTextValue(item.Id ?? item.ID ?? ''),
    tipoEquipoId: fieldMap.tipoEquipo ? getItemFieldText(item, fieldMap.tipoEquipo) : '',
    brandId: getItemFieldText(item, fieldMap.brand),
    modelId: getItemFieldText(item, fieldMap.model),
    sectionId: getItemFieldText(item, fieldMap.section),
    problem: getItemFieldText(item, fieldMap.problem),
    activityTypeId: getItemFieldText(item, fieldMap.activityType),
    activityId: getItemFieldText(item, fieldMap.activity),
    resource: fieldMap.resource ? getItemFieldText(item, fieldMap.resource) : '',
    time: fieldMap.time ? getItemFieldNumeric(item, fieldMap.time) : 0,
    createdBy: fieldMap.createdBy
      ? getTextValue(getItemFieldText(item, fieldMap.createdBy) || item.AuthorId || 'system')
      : getTextValue(item.AuthorId ?? 'system'),
    createdAt: toIsoString(item.Created),
    updatedAt: toIsoString(item.Modified),
  };

  if (resolvedAttachment) {
    mappedRecord.attachment = resolvedAttachment;
  }

  return mappedRecord;
}

export function buildRecordPayload(record, fieldMap) {
  const payload = {
    [fieldMap.brand]: normalizeRequiredText(record.brandId, 'brandId'),
    [fieldMap.model]: normalizeRequiredText(record.modelId, 'modelId'),
    [fieldMap.section]: normalizeRequiredText(record.sectionId, 'sectionId'),
    [fieldMap.problem]: normalizeRequiredText(record.problem, 'problem'),
    [fieldMap.activityType]: normalizeRequiredText(record.activityTypeId, 'activityTypeId'),
    [fieldMap.activity]: normalizeRequiredText(record.activityId, 'activityId'),
  };

  if (fieldMap.tipoEquipo) {
    payload[fieldMap.tipoEquipo] = normalizeRequiredText(record.tipoEquipoId, 'tipoEquipoId');
  }

  if (fieldMap.time) {
    payload[fieldMap.time] = normalizeTime(record.time);
  }

  const resource = getTextValue(record.resource);
  if (fieldMap.resource && resource) {
    payload[fieldMap.resource] = resource;
  }

  const createdBy = getTextValue(record.createdBy);
  if (fieldMap.createdBy && createdBy) {
    payload[fieldMap.createdBy] = createdBy;
  }

  const normalizedAttachment = normalizeAttachment(record.attachment);
  if (normalizedAttachment) {
    if (fieldMap.attachmentName) {
      payload[fieldMap.attachmentName] = normalizedAttachment.name;
    }
    if (fieldMap.attachmentUrl) {
      payload[fieldMap.attachmentUrl] = normalizedAttachment.url;
    }
    if (fieldMap.attachmentType) {
      payload[fieldMap.attachmentType] = normalizedAttachment.type;
    }
    if (fieldMap.attachmentSize) {
      payload[fieldMap.attachmentSize] = normalizedAttachment.size;
    }
  }

  return payload;
}

export function buildDictionaryFromRecords(records) {
  const uniqueTiposEquipo = new Set();
  const uniqueBrands = new Set();
  const uniqueModels = new Map();
  const uniqueSections = new Map();
  const uniqueActivityTypes = new Set();
  const uniqueActivities = new Map();

  records.forEach((record) => {
    if (record.tipoEquipoId) {
      uniqueTiposEquipo.add(record.tipoEquipoId);
    }

    if (record.brandId) {
      uniqueBrands.add(record.brandId);
    }

    if (record.brandId && record.modelId) {
      const modelKey = `${record.brandId}::${record.modelId}`;
      if (!uniqueModels.has(modelKey)) {
        uniqueModels.set(modelKey, {
          id: record.modelId,
          name: record.modelId,
          brandId: record.brandId,
        });
      }
    }

    if (record.brandId && record.modelId && record.sectionId) {
      const sectionKey = `${record.brandId}::${record.modelId}::${record.sectionId}`;
      if (!uniqueSections.has(sectionKey)) {
        uniqueSections.set(sectionKey, {
          id: record.sectionId,
          name: record.sectionId,
          brandId: record.brandId,
          modelId: record.modelId,
        });
      }
    }

    if (record.activityTypeId) {
      uniqueActivityTypes.add(record.activityTypeId);
    }

    if (record.activityTypeId && record.activityId) {
      const activityKey = `${record.activityTypeId}::${record.activityId}`;
      if (!uniqueActivities.has(activityKey)) {
        uniqueActivities.set(activityKey, {
          id: record.activityId,
          name: record.activityId,
          activityTypeId: record.activityTypeId,
        });
      }
    }
  });

  const sectionIdsSeen = new Set(
    Array.from(uniqueSections.values()).map((s) => getTextValue(s.id).toLowerCase())
  );
  records.forEach((record) => {
    const sid = getTextValue(record.sectionId);
    if (!sid) {
      return;
    }
    const lower = sid.toLowerCase();
    if (sectionIdsSeen.has(lower)) {
      return;
    }
    sectionIdsSeen.add(lower);
    uniqueSections.set(`flat:${lower}`, {
      id: sid,
      name: sid,
      brandId: getTextValue(record.brandId),
      modelId: getTextValue(record.modelId),
    });
  });

  const brands = Array.from(uniqueBrands)
    .map((value) => ({ id: value, name: value }))
    .sort((left, right) => left.name.localeCompare(right.name, 'es'));

  const models = Array.from(uniqueModels.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const sections = Array.from(uniqueSections.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const activityTypes = Array.from(uniqueActivityTypes)
    .map((value) => ({ id: value, name: value }))
    .sort((left, right) => left.name.localeCompare(right.name, 'es'));

  const activities = Array.from(uniqueActivities.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const tiposEquipo = Array.from(uniqueTiposEquipo)
    .map((value) => ({ id: value, name: value }))
    .sort((left, right) => left.name.localeCompare(right.name, 'es'));

  return {
    tiposEquipo,
    brands,
    models,
    sections,
    activityTypes,
    activities,
    kpis: {
      totalTiposEquipo: tiposEquipo.length,
      totalMarcas: brands.length,
      totalModelos: models.length,
      totalSecciones: sections.length,
      totalRegistros: records.length,
    },
  };
}

/**
 * Combina el diccionario derivado de registros con las opciones definidas en columnas
 * Choice/MultiChoice de la lista de SharePoint (valores que aún no aparecen en ningún ítem).
 */
export function mergeDictionaryWithColumnChoices(dictionary, columnChoices) {
  const tipoEquipoChoices = Array.isArray(columnChoices.tipoEquipoChoices)
    ? columnChoices.tipoEquipoChoices
    : [];
  const brandChoices = Array.isArray(columnChoices.brandChoices) ? columnChoices.brandChoices : [];
  const modelChoices = Array.isArray(columnChoices.modelChoices) ? columnChoices.modelChoices : [];
  const sectionChoices = Array.isArray(columnChoices.sectionChoices) ? columnChoices.sectionChoices : [];
  const activityTypeChoices = Array.isArray(columnChoices.activityTypeChoices)
    ? columnChoices.activityTypeChoices
    : [];
  const activityChoices = Array.isArray(columnChoices.activityChoices) ? columnChoices.activityChoices : [];

  const tiposMap = new Map();
  (dictionary.tiposEquipo || []).forEach((entry) => {
    tiposMap.set(choiceKey(entry.id), entry);
  });
  tipoEquipoChoices.forEach((raw) => {
    const label = getTextValue(raw);
    if (!label) {
      return;
    }
    const key = choiceKey(label);
    if (!tiposMap.has(key)) {
      tiposMap.set(key, { id: label, name: label });
    }
  });
  const tiposEquipo = Array.from(tiposMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const brandMap = new Map();
  dictionary.brands.forEach((entry) => {
    brandMap.set(choiceKey(entry.id), entry);
  });
  brandChoices.forEach((raw) => {
    const label = getTextValue(raw);
    if (!label) {
      return;
    }
    const key = choiceKey(label);
    if (!brandMap.has(key)) {
      brandMap.set(key, { id: label, name: label });
    }
  });
  const brands = Array.from(brandMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const modelMap = new Map();
  dictionary.models.forEach((entry) => {
    modelMap.set(choiceKey(entry.id), entry);
  });
  modelChoices.forEach((raw) => {
    const label = getTextValue(raw);
    if (!label) {
      return;
    }
    const key = choiceKey(label);
    if (!modelMap.has(key)) {
      modelMap.set(key, { id: label, name: label, brandId: '' });
    }
  });
  const models = Array.from(modelMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const mergeTypeMap = new Map();
  dictionary.activityTypes.forEach((entry) => {
    mergeTypeMap.set(choiceKey(entry.id), entry);
  });
  activityTypeChoices.forEach((raw) => {
    const label = getTextValue(raw);
    if (!label) {
      return;
    }
    const key = choiceKey(label);
    if (!mergeTypeMap.has(key)) {
      mergeTypeMap.set(key, { id: label, name: label });
    }
  });
  const activityTypes = Array.from(mergeTypeMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const sectionMap = new Map();
  dictionary.sections.forEach((section) => {
    sectionMap.set(`${choiceKey(section.brandId)}::${choiceKey(section.modelId)}::${choiceKey(section.id)}`, section);
  });
  dictionary.models.forEach((model) => {
    sectionChoices.forEach((raw) => {
      const label = getTextValue(raw);
      if (!label) {
        return;
      }
      const mapKey = `${choiceKey(model.brandId)}::${choiceKey(model.id)}::${choiceKey(label)}`;
      if (!sectionMap.has(mapKey)) {
        sectionMap.set(mapKey, {
          id: label,
          name: label,
          brandId: model.brandId,
          modelId: model.id,
        });
      }
    });
  });
  const sections = Array.from(sectionMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  const activityMap = new Map();
  dictionary.activities.forEach((activity) => {
    activityMap.set(`${choiceKey(activity.activityTypeId)}::${choiceKey(activity.id)}`, activity);
  });
  activityTypes.forEach((activityType) => {
    activityChoices.forEach((raw) => {
      const label = getTextValue(raw);
      if (!label) {
        return;
      }
      const mapKey = `${choiceKey(activityType.id)}::${choiceKey(label)}`;
      if (!activityMap.has(mapKey)) {
        activityMap.set(mapKey, {
          id: label,
          name: label,
          activityTypeId: activityType.id,
        });
      }
    });
  });
  const activities = Array.from(activityMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'es')
  );

  return {
    ...dictionary,
    tiposEquipo,
    brands,
    models,
    activityTypes,
    sections,
    activities,
    kpis: {
      ...dictionary.kpis,
      totalTiposEquipo: tiposEquipo.length,
      totalMarcas: brands.length,
      totalModelos: models.length,
      totalSecciones: sections.length,
    },
  };
}

function choiceKey(value) {
  return getTextValue(value).toLowerCase();
}

function extractChoicesFromFieldPayload(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  const direct = payload.Choices;
  if (Array.isArray(direct)) {
    return direct;
  }
  if (direct && typeof direct === 'object' && Array.isArray(direct.results)) {
    return direct.results;
  }

  if (Array.isArray(payload.results)) {
    return payload.results;
  }

  const legacyResults = payload.d?.Choices?.results ?? payload.Choices?.results;
  if (Array.isArray(legacyResults)) {
    return legacyResults;
  }

  return [];
}

function parseChoicesFromSchemaXml(schemaXml) {
  if (!schemaXml || typeof schemaXml !== 'string') {
    return [];
  }

  const choices = [];
  const choiceTag = /<CHOICE[^>]*>([^<]*)<\/CHOICE>/gi;
  let match = choiceTag.exec(schemaXml);
  while (match !== null) {
    const label = getTextValue(match[1]);
    if (label) {
      choices.push(label);
    }
    match = choiceTag.exec(schemaXml);
  }

  return choices;
}

function extractAllChoicesFromFieldObject(field) {
  if (!field || typeof field !== 'object') {
    return [];
  }

  const fromRest = extractChoicesFromFieldPayload(field);
  if (fromRest.length > 0) {
    return fromRest;
  }

  return parseChoicesFromSchemaXml(field.SchemaXml);
}

/**
 * Una sola petición a /fields (evita 3× escaneos en paralelo y reduce throttling 429 en SharePoint).
 */
async function fetchListFieldsRows(config) {
  try {
    const encodedList = escapeODataString(config.listTitle);
    const url = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedList}')/fields?$select=InternalName,Title,Choices,TypeAsString,SchemaXml&$top=500`;
    const accessToken = await getAccessToken(config);
    const payload = await sharePointRequest(config, accessToken, { method: 'GET', url });
    const rows = payload.value ?? payload.d?.results ?? [];
    return Array.isArray(rows) ? rows : [];
  } catch {
    return [];
  }
}

/**
 * Coincide nombre configurado (.env) con fila de esquema: InternalName o Title (sin distinguir acentos).
 * Así OData $select y escritura usan el InternalName real de la lista (p. ej. Sección vs Seccion).
 */
function normalizeFieldMatchKey(value) {
  return getTextValue(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{M}/gu, '');
}

function findListFieldRowByInternalOrTitle(rows, fieldKey) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return null;
  }
  const wanted = normalizeFieldMatchKey(fieldKey);
  if (!wanted) {
    return null;
  }
  return (
    rows.find((field) => {
      const internal = normalizeFieldMatchKey(field.InternalName);
      const title = normalizeFieldMatchKey(field.Title);
      return internal === wanted || title === wanted;
    }) ?? null
  );
}

function choicesForFieldFromRows(rows, fieldKey) {
  const match = findListFieldRowByInternalOrTitle(rows, fieldKey);
  if (!match) {
    return [];
  }
  return extractAllChoicesFromFieldObject(match);
}

const FIELD_MAP_KEYS_RESOLVABLE = Object.freeze([
  'tipoEquipo',
  'brand',
  'model',
  'section',
  'problem',
  'activityType',
  'activity',
  'resource',
  'time',
  'createdBy',
  'attachmentName',
  'attachmentUrl',
  'attachmentType',
  'attachmentSize',
]);

/**
 * Sustituye cada entrada del fieldMap por el InternalName real devuelto por /fields.
 * Evita $select OData inválido o campos vacíos cuando el .env usa título visible distinto del InternalName.
 */
export async function resolveFieldMapWithListSchema(config) {
  const base = config.fieldMap;
  const rows = await fetchListFieldsRows(config);
  if (!rows.length) {
    return base;
  }
  const resolved = { ...base };
  for (const key of FIELD_MAP_KEYS_RESOLVABLE) {
    const configured = getTextValue(resolved[key]);
    if (!configured) {
      continue;
    }
    const match = findListFieldRowByInternalOrTitle(rows, configured);
    if (match && getTextValue(match.InternalName)) {
      resolved[key] = getTextValue(match.InternalName);
    }
  }
  return resolved;
}

/**
 * Obtiene las opciones de un campo Choice/MultiChoice con varias rutas REST y SchemaXml como respaldo.
 */
export async function fetchListFieldChoicesRobust(config, internalName) {
  try {
    return await fetchListFieldChoicesRobustInner(config, internalName);
  } catch {
    return [];
  }
}

async function fetchListFieldChoicesRobustInner(config, internalName) {
  const trimmed = getTextValue(internalName);
  if (!trimmed) {
    return [];
  }

  const encodedList = escapeODataString(config.listTitle);
  const listBase = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedList}')`;

  const tryExtract = (payload) => {
    if (!payload || typeof payload !== 'object') {
      return [];
    }
    const direct = extractAllChoicesFromFieldObject(payload);
    if (direct.length > 0) {
      return direct;
    }
    const batch = payload.value ?? payload.d?.results;
    if (Array.isArray(batch)) {
      for (const field of batch) {
        const choices = extractAllChoicesFromFieldObject(field);
        if (choices.length > 0) {
          return choices;
        }
      }
    }
    return [];
  };

  let accessToken;
  const ensureToken = async () => {
    if (!accessToken) {
      accessToken = await getAccessToken(config);
    }
    return accessToken;
  };

  const tryUrl = async (url) => {
    try {
      const token = await ensureToken();
      const payload = await sharePointRequest(config, token, { method: 'GET', url });
      const choices = tryExtract(payload);
      return choices.length > 0 ? choices : [];
    } catch {
      return [];
    }
  };

  const encodedField = escapeODataString(trimmed);
  const byInternalNameUrl = `${listBase}/fields/getbyinternalnameortitle('${encodedField}')`;
  let result = await tryUrl(byInternalNameUrl);
  if (result.length > 0) {
    return result;
  }

  const odataFilter = `InternalName eq '${trimmed.replaceAll("'", "''")}'`;
  const filterUrl = `${listBase}/fields?$filter=${encodeURIComponent(odataFilter)}&$select=InternalName,Title,Choices,TypeAsString,SchemaXml&$top=5`;
  result = await tryUrl(filterUrl);
  if (result.length > 0) {
    return result;
  }

  const scanUrl = `${listBase}/fields?$select=InternalName,Title,Choices,TypeAsString,SchemaXml&$top=500`;
  try {
    const token = await ensureToken();
    const payload = await sharePointRequest(config, token, { method: 'GET', url: scanUrl });
    const rows = payload.value ?? payload.d?.results ?? [];
    if (!Array.isArray(rows)) {
      return [];
    }
    const wanted = trimmed.toLowerCase();
    const match = rows.find((field) => {
      const internal = getTextValue(field.InternalName).toLowerCase();
      const title = getTextValue(field.Title).toLowerCase();
      return internal === wanted || title === wanted;
    });
    if (match) {
      const choices = extractAllChoicesFromFieldObject(match);
      if (choices.length > 0) {
        return choices;
      }
    }
  } catch {
    return [];
  }

  return [];
}

/**
 * Asegura fieldChoiceOptions con metadatos SharePoint +  valor ya visto en ítems/diccionario
 * (si la API de campos no devuelve Choices o la lista falló parcialmente).
 */
export function mergeFieldChoiceOptionsFromRecordsAndDictionary(dictionary, records, fieldChoiceOptions) {
  const sectionSet = new Set(
    (fieldChoiceOptions.section || []).map((s) => getTextValue(s)).filter(Boolean)
  );
  const activityTypeSet = new Set(
    (fieldChoiceOptions.activityType || []).map((s) => getTextValue(s)).filter(Boolean)
  );
  const activitySet = new Set(
    (fieldChoiceOptions.activity || []).map((s) => getTextValue(s)).filter(Boolean)
  );
  const tipoEquipoSet = new Set(
    (fieldChoiceOptions.tipoEquipo || []).map((s) => getTextValue(s)).filter(Boolean)
  );
  const brandSet = new Set(
    (fieldChoiceOptions.brand || []).map((s) => getTextValue(s)).filter(Boolean)
  );
  const modelSet = new Set(
    (fieldChoiceOptions.model || []).map((s) => getTextValue(s)).filter(Boolean)
  );

  (dictionary.tiposEquipo || []).forEach((t) => {
    const id = getTextValue(t.id);
    if (id) {
      tipoEquipoSet.add(id);
    }
  });
  dictionary.brands.forEach((b) => {
    const id = getTextValue(b.id);
    if (id) {
      brandSet.add(id);
    }
  });
  dictionary.models.forEach((m) => {
    const id = getTextValue(m.id);
    if (id) {
      modelSet.add(id);
    }
  });

  dictionary.sections.forEach((s) => {
    const id = getTextValue(s.id);
    if (id) {
      sectionSet.add(id);
    }
  });
  dictionary.activityTypes.forEach((t) => {
    const id = getTextValue(t.id);
    if (id) {
      activityTypeSet.add(id);
    }
  });
  dictionary.activities.forEach((a) => {
    const id = getTextValue(a.id);
    if (id) {
      activitySet.add(id);
    }
  });

  (records || []).forEach((record) => {
    const sid = getTextValue(record.sectionId);
    if (sid) {
      sectionSet.add(sid);
    }
    const tid = getTextValue(record.activityTypeId);
    if (tid) {
      activityTypeSet.add(tid);
    }
    const aid = getTextValue(record.activityId);
    if (aid) {
      activitySet.add(aid);
    }
    const te = getTextValue(record.tipoEquipoId);
    if (te) {
      tipoEquipoSet.add(te);
    }
    const bid = getTextValue(record.brandId);
    if (bid) {
      brandSet.add(bid);
    }
    const mid = getTextValue(record.modelId);
    if (mid) {
      modelSet.add(mid);
    }
  });

  return {
    section: Array.from(sectionSet).sort((a, b) => a.localeCompare(b, 'es')),
    activityType: Array.from(activityTypeSet).sort((a, b) => a.localeCompare(b, 'es')),
    activity: Array.from(activitySet).sort((a, b) => a.localeCompare(b, 'es')),
    tipoEquipo: Array.from(tipoEquipoSet).sort((a, b) => a.localeCompare(b, 'es')),
    brand: Array.from(brandSet).sort((a, b) => a.localeCompare(b, 'es')),
    model: Array.from(modelSet).sort((a, b) => a.localeCompare(b, 'es')),
  };
}

export async function enrichDictionaryWithSharePointFieldChoices(config, dictionary, records = []) {
  try {
    const fieldMap = config.fieldMap;
    const rows = await fetchListFieldsRows(config);

    let tipoEquipoChoices = choicesForFieldFromRows(rows, fieldMap.tipoEquipo);
    let brandChoices = choicesForFieldFromRows(rows, fieldMap.brand);
    let modelChoices = choicesForFieldFromRows(rows, fieldMap.model);
    let sectionChoices = choicesForFieldFromRows(rows, fieldMap.section);
    let activityTypeChoices = choicesForFieldFromRows(rows, fieldMap.activityType);
    let activityChoices = choicesForFieldFromRows(rows, fieldMap.activity);

    if (tipoEquipoChoices.length === 0 && fieldMap.tipoEquipo) {
      tipoEquipoChoices = await fetchListFieldChoicesRobust(config, fieldMap.tipoEquipo);
    }
    if (brandChoices.length === 0 && fieldMap.brand) {
      brandChoices = await fetchListFieldChoicesRobust(config, fieldMap.brand);
    }
    if (modelChoices.length === 0 && fieldMap.model) {
      modelChoices = await fetchListFieldChoicesRobust(config, fieldMap.model);
    }
    if (sectionChoices.length === 0 && fieldMap.section) {
      sectionChoices = await fetchListFieldChoicesRobust(config, fieldMap.section);
    }
    if (activityTypeChoices.length === 0 && fieldMap.activityType) {
      activityTypeChoices = await fetchListFieldChoicesRobust(config, fieldMap.activityType);
    }
    if (activityChoices.length === 0 && fieldMap.activity) {
      activityChoices = await fetchListFieldChoicesRobust(config, fieldMap.activity);
    }

    let fieldChoiceOptions = {
      tipoEquipo: tipoEquipoChoices.map((c) => getTextValue(c)).filter(Boolean),
      brand: brandChoices.map((c) => getTextValue(c)).filter(Boolean),
      model: modelChoices.map((c) => getTextValue(c)).filter(Boolean),
      section: sectionChoices.map((c) => getTextValue(c)).filter(Boolean),
      activityType: activityTypeChoices.map((c) => getTextValue(c)).filter(Boolean),
      activity: activityChoices.map((c) => getTextValue(c)).filter(Boolean),
    };

    const merged = mergeDictionaryWithColumnChoices(dictionary, {
      tipoEquipoChoices,
      brandChoices,
      modelChoices,
      sectionChoices,
      activityTypeChoices,
      activityChoices,
    });

    fieldChoiceOptions = mergeFieldChoiceOptionsFromRecordsAndDictionary(merged, records, fieldChoiceOptions);

    return { ...merged, fieldChoiceOptions };
  } catch {
    const fieldChoiceOptions = mergeFieldChoiceOptionsFromRecordsAndDictionary(
      dictionary,
      records,
      { section: [], activityType: [], activity: [], tipoEquipo: [], brand: [], model: [] }
    );
    return {
      ...dictionary,
      fieldChoiceOptions,
    };
  }
}

const EXACT_MATCH_RECORD_KEYS = new Set([
  'brandId',
  'modelId',
  'sectionId',
  'tipoEquipoId',
  'activityTypeId',
  'activityId',
  'createdBy',
]);

export function filterRecords(records, filters) {
  const normalizedFilters = Object.entries(filters).filter(([, value]) =>
    Boolean(getTextValue(value))
  );

  if (normalizedFilters.length === 0) {
    return records;
  }

  return records.filter((record) =>
    normalizedFilters.every(([key, value]) => {
      const recordValue = getTextValue(record[key]);
      const filterValue = getTextValue(value);

      if (!recordValue || !filterValue) {
        return false;
      }

      if (EXACT_MATCH_RECORD_KEYS.has(key)) {
        return recordValue.toLowerCase() === filterValue.toLowerCase();
      }

      return recordValue.toLowerCase().includes(filterValue.toLowerCase());
    })
  );
}

function normalizePositiveInteger(rawValue, fallbackValue) {
  const parsed = Number.parseInt(rawValue ?? '', 10);
  if (Number.isFinite(parsed) && parsed > 0) {
    return parsed;
  }
  return fallbackValue;
}

async function getAccessToken(config) {
  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const tokenRequestBody = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: config.tokenScope,
  });

  const tokenResponse = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: tokenRequestBody.toString(),
  });

  const tokenPayload = await parseJsonResponse(tokenResponse);
  if (!tokenResponse.ok || !tokenPayload.access_token) {
    throw createHttpError(
      502,
      'Unable to obtain SharePoint access token',
      tokenPayload
    );
  }

  return tokenPayload.access_token;
}

async function sharePointRequest(config, accessToken, { method, url, body }) {
  const requestOptions = {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=nometadata',
    },
  };

  if (body !== undefined) {
    requestOptions.headers['Content-Type'] = 'application/json;odata=nometadata';
    requestOptions.body = JSON.stringify(body);
  }

  const response = await fetch(url, requestOptions);
  const responsePayload = await parseJsonResponse(response);

  if (!response.ok) {
    throw createHttpError(502, 'SharePoint request failed', {
      status: response.status,
      statusText: response.statusText,
      response: responsePayload,
      request: {
        method,
        url,
      },
      listTitle: config.listTitle,
    });
  }

  return responsePayload;
}

async function parseJsonResponse(response) {
  const responseText = await response.text();
  if (!responseText) {
    return {};
  }

  try {
    return JSON.parse(responseText);
  } catch (error) {
    throw createHttpError(502, 'Received non-JSON response from upstream service', {
      status: response.status,
      bodyPreview: responseText.slice(0, 500),
      error,
    });
  }
}

function normalizeTime(rawValue) {
  const numericValue = Number(rawValue);
  if (!Number.isFinite(numericValue) || numericValue < 0) {
    throw createHttpError(400, 'Field "time" must be a non-negative number');
  }
  return numericValue;
}

function normalizeRequiredText(rawValue, fieldName) {
  const normalizedValue = getTextValue(rawValue);
  if (!normalizedValue) {
    throw createHttpError(400, `Field "${fieldName}" is required`);
  }
  return normalizedValue;
}

function extractNativeAttachment(item, siteOrigin) {
  const attachmentFiles = Array.isArray(item.AttachmentFiles) ? item.AttachmentFiles : [];
  if (attachmentFiles.length === 0) {
    return null;
  }

  const firstAttachment = attachmentFiles[0];
  const fileName = getTextValue(firstAttachment.FileName) || 'Adjunto';
  const serverRelativeUrl = getTextValue(firstAttachment.ServerRelativeUrl);
  const attachmentUrl = toAbsoluteAttachmentUrl(serverRelativeUrl, siteOrigin);
  const attachmentSize = getNumericValue(firstAttachment.Length);

  return {
    id: `attachment-${getTextValue(item.Id ?? item.ID ?? Date.now().toString())}`,
    name: fileName,
    url: attachmentUrl,
    type: 'application/octet-stream',
    size: attachmentSize,
  };
}

function extractCustomAttachment(item, fieldMap) {
  const attachmentName = getItemFieldText(item, fieldMap.attachmentName);
  const attachmentUrl = getItemFieldText(item, fieldMap.attachmentUrl);
  const attachmentType = getItemFieldText(item, fieldMap.attachmentType);
  const attachmentSize = getItemFieldNumeric(item, fieldMap.attachmentSize);
  const hasAttachment = Boolean(attachmentName || attachmentUrl);

  if (!hasAttachment) {
    return null;
  }

  return {
    id: `attachment-${getTextValue(item.Id ?? item.ID ?? Date.now().toString())}`,
    name: attachmentName || 'Adjunto',
    url: attachmentUrl || '',
    type: attachmentType || 'application/octet-stream',
    size: attachmentSize,
  };
}

function normalizeAttachment(rawAttachment) {
  if (!rawAttachment) {
    return null;
  }

  if (typeof rawAttachment === 'string') {
    const attachmentText = getTextValue(rawAttachment);
    if (!attachmentText) {
      return null;
    }
    return {
      name: attachmentText,
      url: attachmentText,
      type: 'text/plain',
      size: 0,
    };
  }

  const attachmentName = getTextValue(rawAttachment.name);
  const attachmentUrl = getTextValue(rawAttachment.url);
  if (!attachmentName && !attachmentUrl) {
    return null;
  }

  return {
    name: attachmentName || 'Adjunto',
    url: attachmentUrl || '',
    type: getTextValue(rawAttachment.type) || 'application/octet-stream',
    size: getNumericValue(rawAttachment.size),
  };
}

function toAbsoluteAttachmentUrl(url, siteOrigin) {
  if (!url) {
    return '';
  }

  if (url.startsWith('http://') || url.startsWith('https://')) {
    return url;
  }

  if (url.startsWith('/') && siteOrigin) {
    return `${siteOrigin}${url}`;
  }

  return url;
}

function getTextValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

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

  if (Array.isArray(value)) {
    return value.map((entry) => getTextValue(entry)).filter(Boolean).join(', ');
  }

  if (typeof value === 'object') {
    if (typeof value.Value === 'string') {
      return value.Value.trim();
    }

    if (Array.isArray(value.results)) {
      return value.results.map((entry) => getTextValue(entry)).filter(Boolean).join(', ');
    }

    if (typeof value.LookupValue === 'string') {
      return value.LookupValue.trim();
    }
  }

  return '';
}

function getNumericValue(value) {
  const parsed = Number(value);
  if (Number.isFinite(parsed)) {
    return parsed;
  }
  return 0;
}

function toIsoString(value) {
  const parsedDate = new Date(value);
  if (Number.isNaN(parsedDate.getTime())) {
    return new Date().toISOString();
  }
  return parsedDate.toISOString();
}

function escapeODataString(value) {
  return String(value).replaceAll("'", "''");
}
