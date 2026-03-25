const REQUIRED_ENV_VARS = Object.freeze([
  'SHAREPOINT_SITE_URL',
  'SHAREPOINT_LIST_TITLE',
  'SHAREPOINT_TENANT_ID',
  'SHAREPOINT_CLIENT_ID',
  'SHAREPOINT_CLIENT_SECRET',
]);

const DEFAULT_FIELD_MAP = Object.freeze({
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
    listTitle: process.env.SHAREPOINT_LIST_TITLE,
    tenantId: process.env.SHAREPOINT_TENANT_ID,
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
    tokenScope: `https://${parsedSiteUrl.hostname}/.default`,
    pageSize: normalizePositiveInteger(process.env.SHAREPOINT_PAGE_SIZE, DEFAULT_PAGE_SIZE),
    fieldMap,
  };
}

export async function fetchAllListItems(config) {
  const accessToken = await getAccessToken(config);
  const encodedListTitle = escapeODataString(config.listTitle);
  const baseItemsUrl = `${config.siteUrl}/_api/web/lists/GetByTitle('${encodedListTitle}')/items`;
  const initialQueryParams = buildListItemsQueryParams(config.fieldMap, config.pageSize);
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

function buildListItemsQueryParams(fieldMap, pageSize) {
  const selectFields = new Set([
    'Id',
    'Created',
    'Modified',
    'AuthorId',
    'Attachments',
    'AttachmentFiles',
    'AttachmentFiles/FileName',
    'AttachmentFiles/ServerRelativeUrl',
    'AttachmentFiles/Length',
    fieldMap.brand,
    fieldMap.model,
    fieldMap.section,
    fieldMap.problem,
    fieldMap.activityType,
    fieldMap.activity,
    fieldMap.resource,
    fieldMap.time,
    fieldMap.createdBy,
    fieldMap.attachmentName,
    fieldMap.attachmentUrl,
    fieldMap.attachmentType,
    fieldMap.attachmentSize,
  ]);

  const selectedFieldList = Array.from(selectFields).filter(Boolean).join(',');

  const queryParams = new URLSearchParams();
  queryParams.set('$top', String(pageSize));
  queryParams.set('$select', selectedFieldList);
  queryParams.set('$expand', 'AttachmentFiles');

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

export function mapListItemToMachineRecord(item, fieldMap, siteOrigin = '') {
  const nativeAttachment = extractNativeAttachment(item, siteOrigin);
  const customAttachment = extractCustomAttachment(item, fieldMap);
  const resolvedAttachment = nativeAttachment || customAttachment;

  const mappedRecord = {
    id: getTextValue(item.Id ?? item.ID ?? ''),
    brandId: getTextValue(item[fieldMap.brand]),
    modelId: getTextValue(item[fieldMap.model]),
    sectionId: getTextValue(item[fieldMap.section]),
    problem: getTextValue(item[fieldMap.problem]),
    activityTypeId: getTextValue(item[fieldMap.activityType]),
    activityId: getTextValue(item[fieldMap.activity]),
    resource: fieldMap.resource ? getTextValue(item[fieldMap.resource]) : '',
    time: fieldMap.time ? getNumericValue(item[fieldMap.time]) : 0,
    createdBy: fieldMap.createdBy
      ? getTextValue(item[fieldMap.createdBy] ?? item.AuthorId ?? 'system')
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
  const uniqueBrands = new Set();
  const uniqueModels = new Map();
  const uniqueSections = new Map();
  const uniqueActivityTypes = new Set();
  const uniqueActivities = new Map();

  records.forEach((record) => {
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

  return {
    brands,
    models,
    sections,
    activityTypes,
    activities,
    kpis: {
      totalMarcas: brands.length,
      totalModelos: models.length,
      totalSecciones: sections.length,
      totalRegistros: records.length,
    },
  };
}

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
  const attachmentName = getTextValue(item[fieldMap.attachmentName]);
  const attachmentUrl = getTextValue(item[fieldMap.attachmentUrl]);
  const attachmentType = getTextValue(item[fieldMap.attachmentType]);
  const attachmentSize = getNumericValue(item[fieldMap.attachmentSize]);
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
