import {
  buildDictionaryFromRecords,
  buildRecordPayload,
  createHttpError,
  createListItem,
  enrichDictionaryWithSharePointFieldChoices,
  fetchAllListItems,
  fetchListItemById,
  filterRecords,
  getSharePointConfig,
  mapListItemToMachineRecord,
  mergeFieldChoiceOptionsFromRecordsAndDictionary,
  mergeListItem,
  deleteListItem,
  resolveFieldMapWithListSchema,
  sanitizeAttachmentContentBase64,
  stripCustomAttachmentFieldsFromPayload,
  uploadListItemNativeAttachments,
  deleteListItemNativeAttachment,
} from './_sharepoint.js';

const ALLOWED_RESOURCES = Object.freeze(['records', 'dictionary']);

/** Debe coincidir con VITE_DELETE_RECORD_ALLOWED_EMAIL en el cliente (por defecto cuenta autorizada). */
const DELETE_RECORD_ALLOWED_EMAIL = (
  process.env.DELETE_RECORD_ALLOWED_EMAIL || 'jestrada@partequipos.com'
)
  .trim()
  .toLowerCase();
const ALLOWED_RECORD_FILTERS = Object.freeze([
  'tipoEquipoId',
  'brandId',
  'modelId',
  'sectionId',
  'problem',
  'activityTypeId',
  'activityId',
  'resource',
  'createdBy',
]);

/** Límite de ejecución en Vercel (planes que lo permitan; evita 502 por timeout en listas grandes). */
export const config = {
  maxDuration: 60,
};

async function writeDictionaryResponse(req, res, sharePointConfigResolved) {
  enforceMethod(req.method, ['GET']);
  const records = await safeLoadMappedRecords(sharePointConfigResolved);
  let dictionary = buildDictionaryFromRecords(records);
  try {
    dictionary = await enrichDictionaryWithSharePointFieldChoices(
      sharePointConfigResolved,
      dictionary,
      records
    );
  } catch {
    dictionary = {
      ...dictionary,
      fieldChoiceOptions: mergeFieldChoiceOptionsFromRecordsAndDictionary(
        dictionary,
        records,
        {
          section: [],
          activityType: [],
          activity: [],
          tipoEquipo: [],
          brand: [],
          model: [],
        }
      ),
    };
  }
  setDictionaryDiagnosticHeaders(res, records, dictionary);
  sendJson(res, 200, dictionary);
}

async function writeRecordsListResponse(req, res, sharePointConfigResolved) {
  const records = await safeLoadMappedRecords(sharePointConfigResolved);
  const filters = extractRecordFilters(req.query);
  const filteredRecords = filterRecords(records, filters);
  setRecordsListDiagnosticHeaders(res, records, filteredRecords);
  sendJson(res, 200, { records: filteredRecords });
}

/** SharePoint REST puede devolver Id en raíz o bajo d (formato clásico). */
function extractCreatedListItemId(createdItem) {
  if (!createdItem || typeof createdItem !== 'object') {
    return undefined;
  }
  return (
    createdItem.Id ??
    createdItem.ID ??
    createdItem.id ??
    createdItem.d?.Id ??
    createdItem.d?.ID ??
    createdItem.d?.id
  );
}

async function writeCreatedRecordResponse(req, res, sharePointConfigResolved) {
  const requestBody = parseRequestBody(req.body);
  const incomingRecord = requestBody.record;
  if (!incomingRecord || typeof incomingRecord !== 'object') {
    throw createHttpError(400, 'Request body must include a "record" object');
  }

  const attachmentFiles = normalizeAttachmentFilesFromRequest(requestBody.attachmentFiles);
  const payload = buildRecordPayload(incomingRecord, sharePointConfigResolved.fieldMap);
  if (attachmentFiles.length > 0) {
    stripCustomAttachmentFieldsFromPayload(payload, sharePointConfigResolved.fieldMap);
  }

  const createdItem = await createListItem(sharePointConfigResolved, payload);
  const itemId = extractCreatedListItemId(createdItem);
  if (attachmentFiles.length > 0) {
    if (itemId === undefined || itemId === null || itemId === '') {
      throw createHttpError(502, 'List item created without id; cannot upload attachments');
    }
    await uploadListItemNativeAttachments(sharePointConfigResolved, itemId, attachmentFiles);
  }

  const reloadedItem =
    attachmentFiles.length > 0
      ? await fetchListItemById(sharePointConfigResolved, itemId)
      : createdItem;
  const createdRecord = mapListItemToMachineRecord(
    reloadedItem,
    sharePointConfigResolved.fieldMap,
    sharePointConfigResolved.siteOrigin
  );
  sendJson(res, 201, { record: createdRecord });
}

function normalizeRemoveAttachmentFileNames(raw) {
  if (raw === undefined || raw === null) {
    return [];
  }
  if (!Array.isArray(raw)) {
    throw createHttpError(400, 'removeAttachmentFileNames must be an array when provided');
  }
  return raw
    .map((entry) => (typeof entry === 'string' ? entry.trim() : String(entry ?? '').trim()))
    .filter((name) => name.length > 0);
}

function normalizeDeleteRequesterEmail(headerValue) {
  if (typeof headerValue !== 'string') {
    return '';
  }
  return headerValue.trim().toLowerCase();
}

async function writeDeletedRecordResponse(req, res, sharePointConfigResolved) {
  enforceMethod(req.method, ['DELETE']);
  const rawId = getQueryValue(req.query.id);
  if (!rawId || String(rawId).trim() === '') {
    throw createHttpError(400, 'Query parameter "id" is required');
  }
  const requester = normalizeDeleteRequesterEmail(req.headers['x-ishikawa-delete-requested-by-email']);
  if (!requester || requester !== DELETE_RECORD_ALLOWED_EMAIL) {
    throw createHttpError(403, 'Eliminar registros no está permitido para esta cuenta.');
  }
  await deleteListItem(sharePointConfigResolved, rawId);
  sendJson(res, 200, { deleted: true, id: String(rawId) });
}

async function writeUpdatedRecordResponse(req, res, sharePointConfigResolved) {
  enforceMethod(req.method, ['PATCH']);
  const requestBody = parseRequestBody(req.body);
  const incomingRecord = requestBody.record;
  if (!incomingRecord || typeof incomingRecord !== 'object') {
    throw createHttpError(400, 'Request body must include a "record" object');
  }
  const rawId = incomingRecord.id;
  if (rawId === undefined || rawId === null || String(rawId).trim() === '') {
    throw createHttpError(400, 'Request body.record must include id');
  }
  const attachmentFiles = normalizeAttachmentFilesFromRequest(requestBody.attachmentFiles);
  const removeAttachmentFileNames = normalizeRemoveAttachmentFileNames(
    requestBody.removeAttachmentFileNames
  );

  const payload = buildRecordPayload(incomingRecord, sharePointConfigResolved.fieldMap, {
    isMerge: true,
  });
  const hasNativeAttachmentMutation =
    attachmentFiles.length > 0 || removeAttachmentFileNames.length > 0;
  if (hasNativeAttachmentMutation) {
    stripCustomAttachmentFieldsFromPayload(payload, sharePointConfigResolved.fieldMap);
  }
  await mergeListItem(sharePointConfigResolved, rawId, payload);

  for (const fileName of removeAttachmentFileNames) {
    await deleteListItemNativeAttachment(sharePointConfigResolved, rawId, fileName);
  }
  if (attachmentFiles.length > 0) {
    await uploadListItemNativeAttachments(sharePointConfigResolved, rawId, attachmentFiles);
  }

  const reloadedItem = await fetchListItemById(sharePointConfigResolved, rawId);
  const updatedRecord = mapListItemToMachineRecord(
    reloadedItem,
    sharePointConfigResolved.fieldMap,
    sharePointConfigResolved.siteOrigin
  );
  sendJson(res, 200, { record: updatedRecord });
}

export default async function handler(req, res) {
  setJsonHeaders(res);

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  try {
    const requestedResource = getQueryValue(req.query.resource);
    if (!ALLOWED_RESOURCES.includes(requestedResource)) {
      throw createHttpError(
        400,
        `Query parameter "resource" must be one of: ${ALLOWED_RESOURCES.join(', ')}`
      );
    }

    const sharePointConfig = getSharePointConfig();
    const sharePointConfigResolved = await withResolvedFieldMap(sharePointConfig);

    if (requestedResource === 'dictionary') {
      await writeDictionaryResponse(req, res, sharePointConfigResolved);
      return;
    }

    if (requestedResource === 'records' && req.method === 'GET') {
      await writeRecordsListResponse(req, res, sharePointConfigResolved);
      return;
    }

    if (requestedResource === 'records' && req.method === 'POST') {
      await writeCreatedRecordResponse(req, res, sharePointConfigResolved);
      return;
    }

    if (requestedResource === 'records' && req.method === 'PATCH') {
      await writeUpdatedRecordResponse(req, res, sharePointConfigResolved);
      return;
    }

    if (requestedResource === 'records' && req.method === 'DELETE') {
      await writeDeletedRecordResponse(req, res, sharePointConfigResolved);
      return;
    }

    throw createHttpError(405, `Method ${req.method} is not allowed for this resource`);
  } catch (error) {
    const statusCode = normalizeStatusCode(error.statusCode);
    sendJson(res, statusCode, {
      message: error.message || 'Unexpected error while processing request',
      details: error.details ?? null,
    });
  }
}

async function withResolvedFieldMap(sharePointConfig) {
  try {
    const fieldMap = await resolveFieldMapWithListSchema(sharePointConfig);
    return { ...sharePointConfig, fieldMap };
  } catch {
    return sharePointConfig;
  }
}

async function loadMappedRecords(sharePointConfig) {
  const listItems = await fetchAllListItems(sharePointConfig);
  return listItems.map((item) =>
    mapListItemToMachineRecord(item, sharePointConfig.fieldMap, sharePointConfig.siteOrigin)
  );
}

/**
 * Si la lectura OData de ítems falla (token, throttling, columnas, timeout), no tumba el API con 502:
 * el cliente recibe estructura vacía y puede seguir usando metadatos de columnas del enriquecimiento.
 */
async function safeLoadMappedRecords(sharePointConfig) {
  try {
    return await loadMappedRecords(sharePointConfig);
  } catch (error) {
    /* SharePoint list items failed; empty array avoids 502 and preserves API contract. */
    console.error('[ishikawa] SharePoint list items failed:', error?.message || error);
    return [];
  }
}

function extractRecordFilters(query) {
  const filters = {};

  ALLOWED_RECORD_FILTERS.forEach((key) => {
    const queryValue = getQueryValue(query[key]);
    if (queryValue) {
      filters[key] = queryValue;
    }
  });

  return filters;
}

function normalizeAttachmentFilesFromRequest(raw) {
  if (raw === undefined || raw === null) {
    return [];
  }
  if (!Array.isArray(raw)) {
    throw createHttpError(400, 'attachmentFiles must be an array when provided');
  }
  const out = [];
  for (const entry of raw) {
    if (!entry || typeof entry !== 'object') {
      continue;
    }
    const name = typeof entry.name === 'string' ? entry.name.trim() : '';
    const contentBase64 = sanitizeAttachmentContentBase64(entry.contentBase64);
    if (!name || !contentBase64) {
      continue;
    }
    const contentType =
      typeof entry.contentType === 'string' && entry.contentType.trim()
        ? entry.contentType.trim()
        : 'application/octet-stream';
    out.push({ name, contentType, contentBase64 });
  }
  if (raw.length > 0 && out.length === 0) {
    throw createHttpError(400, 'attachmentFiles entries must include non-empty name and contentBase64');
  }
  return out;
}

function parseRequestBody(rawBody) {
  if (!rawBody) {
    return {};
  }

  if (typeof rawBody === 'string') {
    try {
      return JSON.parse(rawBody);
    } catch (error) {
      throw createHttpError(400, 'Request body is not valid JSON', error);
    }
  }

  if (typeof rawBody === 'object') {
    return rawBody;
  }

  return {};
}

function enforceMethod(method, allowedMethods) {
  if (!allowedMethods.includes(method)) {
    throw createHttpError(405, `Method ${method} is not allowed`);
  }
}

function getQueryValue(value) {
  if (Array.isArray(value)) {
    return value[0] || '';
  }
  return value || '';
}

function setJsonHeaders(res) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
}

/**
 * Cabeceras sin datos sensibles: visibles en producción (Vercel) desde DevTools → Red → ishikawa.
 * Si SharePoint-Items es 0 pero hay ítems en la lista, revisar permisos OData, variables de entorno y logs de función.
 */
function setDictionaryDiagnosticHeaders(res, sharePointRecords, dictionary) {
  const fco = dictionary.fieldChoiceOptions;
  res.setHeader('X-Ishikawa-SharePoint-Items', String(sharePointRecords.length));
  res.setHeader('X-Ishikawa-Dictionary-Brands', String(dictionary.brands?.length ?? 0));
  res.setHeader(
    'X-Ishikawa-Fco-Section-Count',
    String(Array.isArray(fco?.section) ? fco.section.length : 0)
  );
}

function setRecordsListDiagnosticHeaders(res, allRecords, returnedRecords) {
  res.setHeader('X-Ishikawa-SharePoint-Items', String(allRecords.length));
  res.setHeader('X-Ishikawa-Records-Returned', String(returnedRecords.length));
}

function sendJson(res, statusCode, payload) {
  res.status(statusCode).json(payload);
}

function normalizeStatusCode(statusCode) {
  const parsedStatusCode = Number(statusCode);
  if (Number.isInteger(parsedStatusCode) && parsedStatusCode >= 100 && parsedStatusCode <= 599) {
    return parsedStatusCode;
  }
  return 500;
}
