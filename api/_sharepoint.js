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
function buildCanonicalInternalNameMap(schemaRows) {
  const canonicalByLower = new Map();
  if (!Array.isArray(schemaRows) || schemaRows.length === 0) {
    return canonicalByLower;
  }
  for (const row of schemaRows) {
    const internal = getTextValue(row.InternalName);
    if (internal) {
      canonicalByLower.set(internal.toLowerCase(), internal);
    }
  }
  return canonicalByLower;
}

function mergeMappedFieldNamesIntoSelect(selectFields, fieldMap, canonicalByLower) {
  const mappedNames = Object.values(fieldMap)
    .map((v) => getTextValue(v))
    .filter(Boolean);

  if (canonicalByLower.size === 0) {
    mappedNames.forEach((name) => selectFields.add(name));
    return;
  }
  for (const name of mappedNames) {
    const canonical = canonicalByLower.get(name.toLowerCase());
    if (canonical) {
      selectFields.add(canonical);
    }
  }
}

function appendAttachmentSelectFieldNames(selectFields) {
  selectFields.add('AttachmentFiles');
  selectFields.add('AttachmentFiles/FileName');
  selectFields.add('AttachmentFiles/ServerRelativeUrl');
  selectFields.add('AttachmentFiles/Length');
}

function buildListItemsQueryParams(fieldMap, pageSize, expandAttachmentFiles = true, schemaRows = []) {
  const selectFields = new Set(['Id', 'Created', 'Modified', 'AuthorId', 'Attachments']);
  const canonicalByLower = buildCanonicalInternalNameMap(schemaRows);
  mergeMappedFieldNamesIntoSelect(selectFields, fieldMap, canonicalByLower);

  if (expandAttachmentFiles) {
    appendAttachmentSelectFieldNames(selectFields);
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

const MICROSOFT_GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
const MICROSOFT_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

async function getMicrosoftGraphAccessToken(config) {
  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const tokenRequestBody = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: MICROSOFT_GRAPH_SCOPE,
  });

  const tokenResponse = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: tokenRequestBody.toString(),
  });

  const tokenPayload = await parseJsonResponse(tokenResponse);
  if (!tokenResponse.ok || !tokenPayload.access_token) {
    throw createHttpError(502, 'Unable to obtain Microsoft Graph access token', tokenPayload);
  }

  return tokenPayload.access_token;
}

function graphSiteIdentifierFromSiteUrl(siteUrlRaw) {
  let parsed;
  try {
    parsed = new URL(siteUrlRaw);
  } catch {
    return '';
  }
  const host = getTextValue(parsed.hostname);
  const path = getTextValue(parsed.pathname).replace(/\/$/, '');
  if (!host) {
    return '';
  }
  return path ? `${host}:${path}` : host;
}

function graphListMatchesConfiguredTitle(graphList, wantedTitle) {
  const w = getTextValue(wantedTitle).toLowerCase();
  if (!w) {
    return false;
  }
  const display = getTextValue(graphList.displayName).toLowerCase();
  const name = getTextValue(graphList.name).toLowerCase();
  return display === w || name === w;
}

async function resolveGraphSiteId(config, graphToken) {
  const siteKey = graphSiteIdentifierFromSiteUrl(config.siteUrl);
  if (!siteKey) {
    throw createHttpError(500, 'Invalid site URL for Microsoft Graph');
  }
  const url = `${MICROSOFT_GRAPH_ROOT}/sites/${encodeURIComponent(siteKey)}`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${graphToken}` },
  });
  const data = await parseJsonResponse(response);
  if (!response.ok) {
    throw createHttpError(502, 'Microsoft Graph site resolution failed', {
      status: response.status,
      response: data,
      url,
    });
  }
  const id = getTextValue(data.id);
  if (!id) {
    throw createHttpError(502, 'Microsoft Graph returned no site id', data);
  }
  return id;
}

async function resolveGraphListIdForTitle(config, graphToken, siteGraphId) {
  const wanted = getTextValue(config.listTitle);
  let nextUrl = `${MICROSOFT_GRAPH_ROOT}/sites/${siteGraphId}/lists?$select=id,displayName,name&$top=200`;

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${graphToken}` },
    });
    const data = await parseJsonResponse(response);
    if (!response.ok) {
      throw createHttpError(502, 'Microsoft Graph list enumeration failed', {
        status: response.status,
        response: data,
      });
    }
    const lists = Array.isArray(data.value) ? data.value : [];
    const match = lists.find((list) => graphListMatchesConfiguredTitle(list, wanted));
    if (match && getTextValue(match.id)) {
      return getTextValue(match.id);
    }
    const link = data['@odata.nextLink'];
    nextUrl = typeof link === 'string' && link.trim() ? link.trim() : '';
    if (!nextUrl) {
      break;
    }
  }

  throw createHttpError(404, `Microsoft Graph: no list matching title "${wanted}"`);
}

const graphSiteListIdCache = new Map();
const GRAPH_SITE_LIST_CACHE_TTL_MS = 50 * 60 * 1000;

function graphSiteListCacheKey(config) {
  return `${getTextValue(config.siteUrl)}|${getTextValue(config.listTitle)}`;
}

async function resolveGraphSiteAndListIdsCached(config, graphToken) {
  const key = graphSiteListCacheKey(config);
  const now = Date.now();
  const cached = graphSiteListIdCache.get(key);
  if (cached && cached.expiresAt > now) {
    return { siteGraphId: cached.siteGraphId, listGraphId: cached.listGraphId };
  }
  const siteGraphId = await resolveGraphSiteId(config, graphToken);
  const listGraphId = await resolveGraphListIdForTitle(config, graphToken, siteGraphId);
  graphSiteListIdCache.set(key, {
    siteGraphId,
    listGraphId,
    expiresAt: now + GRAPH_SITE_LIST_CACHE_TTL_MS,
  });
  return { siteGraphId, listGraphId };
}

/**
 * Actualiza columnas del ítem con Graph (PATCH .../items/{id}/fields).
 * Misma idea que el cliente (Graph): muchos tenants aceptan Graph con app-only aunque REST MERGE devuelva 401.
 */
async function patchListItemFieldsViaMicrosoftGraph(config, itemId, fieldsPayload) {
  const graphToken = await getMicrosoftGraphAccessToken(config);
  const { siteGraphId, listGraphId } = await resolveGraphSiteAndListIdsCached(config, graphToken);
  const itemIdStr = getTextValue(itemId);
  if (!itemIdStr) {
    throw createHttpError(400, 'Invalid list item id for Microsoft Graph');
  }
  const url = `${MICROSOFT_GRAPH_ROOT}/sites/${siteGraphId}/lists/${listGraphId}/items/${encodeURIComponent(itemIdStr)}/fields`;
  const response = await fetch(url, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${graphToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(fieldsPayload),
  });
  const data = await parseJsonResponse(response);
  if (!response.ok) {
    throw createHttpError(502, 'Microsoft Graph list item update failed', {
      status: response.status,
      statusText: response.statusText,
      response: data,
      url,
    });
  }
  return data;
}

/**
 * InternalName OData del tipo de fila de la lista (p. ej. SP.Data.IshikawaListItem).
 * Sin esto, algunos inquilinos devuelven 401 en MERGE aunque el POST de alta funcione.
 */
async function tryGetListItemEntityTypeFullName(config, accessToken) {
  const encodedListTitle = escapeODataString(config.listTitle);
  const url = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')?$select=ListItemEntityTypeFullName`;
  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json;odata=nometadata',
      },
    });
    const responseText = await response.text();
    if (!response.ok || !responseText) {
      return '';
    }
    let data;
    try {
      data = JSON.parse(responseText);
    } catch {
      return '';
    }
    return getTextValue(data.ListItemEntityTypeFullName ?? data.d?.ListItemEntityTypeFullName);
  } catch {
    return '';
  }
}

function sharePointMergeFailureStatusAndMessage(upstreamStatus) {
  if (upstreamStatus === 400 || upstreamStatus === 404) {
    return {
      statusCode: upstreamStatus,
      message: 'SharePoint rejected the merge payload (invalid field or item).',
    };
  }
  if (upstreamStatus === 401 || upstreamStatus === 403) {
    return {
      statusCode: upstreamStatus,
      message: 'SharePoint REST denied the merge (401/403).',
    };
  }
  return {
    statusCode: 502,
    message: 'SharePoint merge request failed',
  };
}

function buildSharePointMergeRequestBody(useVerboseMetadata, entityType, payload) {
  if (useVerboseMetadata) {
    return { __metadata: { type: entityType }, ...payload };
  }
  return payload;
}

function buildSharePointMergeRequestHeaders(accessToken, useVerboseMetadata, digest) {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: useVerboseMetadata ? 'application/json;odata=verbose' : 'application/json;odata=nometadata',
    'Content-Type': useVerboseMetadata ? 'application/json' : 'application/json;odata=nometadata',
    'IF-MATCH': '*',
    'X-HTTP-Method': 'MERGE',
  };
  if (digest) {
    headers['X-RequestDigest'] = digest;
  }
  return headers;
}

/**
 * MERGE vía SharePoint REST (respaldo). El alta usa POST /items; aquí POST + X-HTTP-Method: MERGE.
 */
async function mergeListItemViaSharePointRest(config, itemId, payload) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const id = Number(itemId);
  const url = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${id})`;

  const entityType = await tryGetListItemEntityTypeFullName(config, accessToken);

  let digest;
  try {
    digest = await getSharePointFormDigest(config, accessToken);
  } catch {
    digest = undefined;
  }

  const useVerboseMetadata = Boolean(entityType);
  const bodyObject = buildSharePointMergeRequestBody(useVerboseMetadata, entityType, payload);
  const headers = buildSharePointMergeRequestHeaders(accessToken, useVerboseMetadata, digest);

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(bodyObject),
  });
  const responsePayload = await parseJsonResponse(response);
  if (!response.ok) {
    const upstreamStatus = response.status;
    const { statusCode, message } = sharePointMergeFailureStatusAndMessage(upstreamStatus);
    throw createHttpError(statusCode, message, {
      status: upstreamStatus,
      statusText: response.statusText,
      response: responsePayload,
      request: { method: 'MERGE', url },
      listTitle: config.listTitle,
    });
  }
  return responsePayload;
}

/**
 * Actualiza campos del ítem: primero Microsoft Graph (como en el flujo de lectura del SPA), luego MERGE REST.
 * Así se evita el 401 de MERGE en tenants donde el POST de alta y los adjuntos REST sí funcionan.
 */
export async function mergeListItem(config, itemId, payload) {
  const id = Number(itemId);
  if (!Number.isInteger(id) || id < 1) {
    throw createHttpError(400, 'Invalid list item id');
  }

  const fieldKeys = Object.keys(payload || {});
  if (fieldKeys.length === 0) {
    return {};
  }

  let graphError = null;
  try {
    await patchListItemFieldsViaMicrosoftGraph(config, id, payload);
    return {};
  } catch (err) {
    graphError = err;
  }

  try {
    return await mergeListItemViaSharePointRest(config, itemId, payload);
  } catch (restError) {
    throw createHttpError(502, 'Could not update list item fields. Microsoft Graph failed first; SharePoint REST MERGE also failed.', {
      graphAttempt: graphError?.details ?? { message: graphError?.message ?? String(graphError) },
      sharePointRestAttempt: restError?.details ?? { message: restError?.message ?? String(restError) },
    });
  }
}

export async function fetchListItemById(config, itemId, options = {}) {
  const expandAttachmentFiles = options.expandAttachmentFiles !== false;
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const id = Number(itemId);
  if (!Number.isInteger(id) || id < 1) {
    throw createHttpError(400, 'Invalid list item id');
  }
  const expand = expandAttachmentFiles ? '?$expand=AttachmentFiles' : '';
  const url = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${id})${expand}`;

  return sharePointRequest(config, accessToken, {
    method: 'GET',
    url,
  });
}

const MAX_NATIVE_ATTACHMENT_BYTES = 15 * 1024 * 1024;
const MAX_NATIVE_ATTACHMENTS_PER_ITEM = 20;

/**
 * Sube archivos a la columna nativa Attachments vía REST (AttachmentFiles/add).
 * @param {Array<{ name: string, contentType?: string, contentBase64: string }>} attachmentFiles
 */
export async function uploadListItemNativeAttachments(config, itemId, attachmentFiles) {
  if (!Array.isArray(attachmentFiles) || attachmentFiles.length === 0) {
    return;
  }
  if (attachmentFiles.length > MAX_NATIVE_ATTACHMENTS_PER_ITEM) {
    throw createHttpError(
      400,
      `At most ${MAX_NATIVE_ATTACHMENTS_PER_ITEM} attachments per request`
    );
  }

  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const id = Number(itemId);
  if (!Number.isInteger(id) || id < 1) {
    throw createHttpError(400, 'Invalid list item id');
  }

  let digest;
  try {
    digest = await getSharePointFormDigest(config, accessToken);
  } catch {
    digest = undefined;
  }

  for (const file of attachmentFiles) {
    const name = getTextValue(file?.name);
    const b64 = sanitizeAttachmentContentBase64(file?.contentBase64 ?? '');
    if (!name || !b64) {
      throw createHttpError(400, 'Each attachment must include non-empty name and contentBase64');
    }
    let buffer;
    try {
      buffer = Buffer.from(b64, 'base64');
    } catch {
      throw createHttpError(400, `Invalid base64 for attachment "${name}"`);
    }
    if (!buffer.length) {
      throw createHttpError(400, `Empty attachment content for "${name}"`);
    }
    if (buffer.length > MAX_NATIVE_ATTACHMENT_BYTES) {
      throw createHttpError(400, `Attachment "${name}" exceeds maximum size`);
    }
    const safeName = escapeODataString(name);
    const addUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${id})/AttachmentFiles/add(FileName='${safeName}')`;

    await sharePointBinaryRequest(config, accessToken, {
      method: 'POST',
      url: addUrl,
      body: buffer,
      digest,
    });
  }
}

/**
 * Elimina un archivo de la columna nativa Attachments del ítem (SharePoint REST DELETE).
 * Ignora 404 si el adjunto ya no existe.
 */
export async function deleteListItemNativeAttachment(config, itemId, fileName) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const id = Number(itemId);
  if (!Number.isInteger(id) || id < 1) {
    throw createHttpError(400, 'Invalid list item id');
  }
  const name = getTextValue(fileName);
  if (!name) {
    throw createHttpError(400, 'Attachment file name is required for delete');
  }
  const safeName = escapeODataString(name);
  const deleteUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${id})/AttachmentFiles('${safeName}')`;

  let digest;
  try {
    digest = await getSharePointFormDigest(config, accessToken);
  } catch {
    digest = undefined;
  }

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json;odata=nometadata',
    'IF-MATCH': '*',
  };
  if (digest) {
    headers['X-RequestDigest'] = digest;
  }

  const response = await fetch(deleteUrl, {
    method: 'DELETE',
    headers,
  });

  if (response.status === 404) {
    return;
  }

  if (!response.ok) {
    const responsePayload = await parseJsonResponse(response);
    throw createHttpError(502, 'SharePoint attachment delete failed', {
      status: response.status,
      statusText: response.statusText,
      response: responsePayload,
      fileName: name,
    });
  }
}

/**
 * Elimina un ítem de la lista (SharePoint REST DELETE → papelera de reciclaje).
 */
export async function deleteListItem(config, itemId) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const id = Number(itemId);
  if (!Number.isInteger(id) || id < 1) {
    throw createHttpError(400, 'Invalid list item id');
  }
  const deleteUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${id})`;

  let digest;
  try {
    digest = await getSharePointFormDigest(config, accessToken);
  } catch {
    digest = undefined;
  }

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json;odata=nometadata',
    'IF-MATCH': '*',
  };
  if (digest) {
    headers['X-RequestDigest'] = digest;
  }

  const response = await fetch(deleteUrl, {
    method: 'DELETE',
    headers,
  });

  if (response.status === 404) {
    throw createHttpError(404, 'List item not found');
  }

  if (!response.ok) {
    const responsePayload = await parseJsonResponse(response);
    throw createHttpError(502, 'SharePoint list item delete failed', {
      status: response.status,
      statusText: response.statusText,
      response: responsePayload,
      request: { method: 'DELETE', url: deleteUrl },
      listTitle: config.listTitle,
    });
  }
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
  const nativeAttachments = extractNativeAttachments(item, siteOrigin);
  const customAttachment = extractCustomAttachment(item, fieldMap);
  let resolvedList;
  if (nativeAttachments.length > 0) {
    resolvedList = nativeAttachments;
  } else if (customAttachment) {
    resolvedList = [customAttachment];
  } else {
    resolvedList = [];
  }
  const resolvedAttachment = resolvedList[0];

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

  if (resolvedList.length > 0) {
    mappedRecord.attachment = resolvedAttachment;
    mappedRecord.attachments = resolvedList;
  }

  return mappedRecord;
}

function putTipoEquipoOnPayload(payload, fieldMap, record) {
  if (!fieldMap.tipoEquipo) {
    return;
  }
  payload[fieldMap.tipoEquipo] = normalizeRequiredText(record.tipoEquipoId, 'tipoEquipoId');
}

function putTimeOnPayload(payload, fieldMap, record) {
  if (!fieldMap.time) {
    return;
  }
  payload[fieldMap.time] = normalizeTime(record.time);
}

function putResourceOnPayload(payload, fieldMap, record) {
  const resource = getTextValue(record.resource);
  if (!fieldMap.resource || !resource) {
    return;
  }
  payload[fieldMap.resource] = resource;
}

function putCreatedByOnPayloadForCreate(payload, fieldMap, record, isMerge) {
  if (isMerge) {
    return;
  }
  const createdBy = getTextValue(record.createdBy);
  if (!fieldMap.createdBy || !createdBy) {
    return;
  }
  payload[fieldMap.createdBy] = createdBy;
}

function putAttachmentColumnsOnPayload(payload, fieldMap, normalizedAttachment) {
  if (!normalizedAttachment) {
    return;
  }
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

/**
 * @param {{ isMerge?: boolean }} [options]
 * @param {boolean} [options.isMerge] — Actualización de ítem (MERGE). Omite columnas que SharePoint suele rechazar al actualizar (p. ej. autor / creado por).
 */
export function buildRecordPayload(record, fieldMap, options) {
  const isMerge = Boolean(options?.isMerge);
  const payload = {
    [fieldMap.brand]: normalizeRequiredText(record.brandId, 'brandId'),
    [fieldMap.model]: normalizeRequiredText(record.modelId, 'modelId'),
    [fieldMap.section]: normalizeRequiredText(record.sectionId, 'sectionId'),
    [fieldMap.problem]: normalizeRequiredText(record.problem, 'problem'),
    [fieldMap.activityType]: normalizeRequiredText(record.activityTypeId, 'activityTypeId'),
    [fieldMap.activity]: normalizeRequiredText(record.activityId, 'activityId'),
  };

  putTipoEquipoOnPayload(payload, fieldMap, record);
  putTimeOnPayload(payload, fieldMap, record);
  putResourceOnPayload(payload, fieldMap, record);
  putCreatedByOnPayloadForCreate(payload, fieldMap, record, isMerge);

  const normalizedAttachment = normalizeAttachment(record.attachment);
  putAttachmentColumnsOnPayload(payload, fieldMap, normalizedAttachment);

  return payload;
}

export function stripCustomAttachmentFieldsFromPayload(payload, fieldMap) {
  if (!payload || typeof payload !== 'object' || !fieldMap) {
    return;
  }
  const keys = [fieldMap.attachmentName, fieldMap.attachmentUrl, fieldMap.attachmentType, fieldMap.attachmentSize];
  for (const k of keys) {
    const key = getTextValue(k);
    if (key && Object.hasOwn(payload, key)) {
      delete payload[key];
    }
  }
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
    .replaceAll(/\p{M}/gu, '');
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

export function sanitizeAttachmentContentBase64(raw) {
  return typeof raw === 'string' ? raw.replaceAll(/\s/g, '') : '';
}

/** SharePoint Online suele exigir X-RequestDigest en POST que modifican contenido (p. ej. AttachmentFiles/add). */
async function getSharePointFormDigest(config, accessToken) {
  const url = `${config.siteUrl}/_api/contextinfo`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
    },
    body: '{}',
  });
  const responsePayload = await parseJsonResponse(response);
  if (!response.ok) {
    throw createHttpError(502, 'SharePoint form digest request failed', {
      status: response.status,
      statusText: response.statusText,
      response: responsePayload,
    });
  }
  const digest =
    getTextValue(responsePayload.FormDigestValue) ||
    getTextValue(responsePayload.d?.FormDigestValue);
  if (!digest) {
    throw createHttpError(502, 'SharePoint returned no FormDigestValue', responsePayload);
  }
  return digest;
}

async function sharePointBinaryRequest(config, accessToken, { method, url, body, digest }) {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/octet-stream',
  };
  if (digest) {
    headers['X-RequestDigest'] = digest;
  }
  const response = await fetch(url, {
    method,
    headers,
    body,
  });
  const responsePayload = await parseJsonResponse(response);

  if (!response.ok) {
    throw createHttpError(502, 'SharePoint binary request failed', {
      status: response.status,
      statusText: response.statusText,
      response: responsePayload,
      request: { method, url },
      listTitle: config.listTitle,
    });
  }

  return responsePayload;
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

function extractNativeAttachments(item, siteOrigin) {
  const attachmentFiles = Array.isArray(item.AttachmentFiles) ? item.AttachmentFiles : [];
  const itemId = getTextValue(item.Id ?? item.ID ?? Date.now().toString());

  return attachmentFiles.map((fileRow, index) => {
    const fileName = getTextValue(fileRow.FileName) || 'Adjunto';
    const serverRelativeUrl = getTextValue(fileRow.ServerRelativeUrl);
    const attachmentUrl = toAbsoluteAttachmentUrl(serverRelativeUrl, siteOrigin);
    const attachmentSize = getNumericValue(fileRow.Length);

    return {
      id: `attachment-${itemId}-${index}-${fileName}`,
      name: fileName,
      url: attachmentUrl,
      type: 'application/octet-stream',
      size: attachmentSize,
    };
  });
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
