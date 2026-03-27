import {
  buildDictionaryFromRecords,
  buildRecordPayload,
  createHttpError,
  createListItem,
  enrichDictionaryWithSharePointFieldChoices,
  fetchAllListItems,
  filterRecords,
  getSharePointConfig,
  mapListItemToMachineRecord,
  mergeFieldChoiceOptionsFromRecordsAndDictionary,
  resolveFieldMapWithListSchema,
} from './_sharepoint.js';

const ALLOWED_RESOURCES = Object.freeze(['records', 'dictionary']);
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
      sendJson(res, 200, dictionary);
      return;
    }

    if (requestedResource === 'records' && req.method === 'GET') {
      const records = await safeLoadMappedRecords(sharePointConfigResolved);
      const filters = extractRecordFilters(req.query);
      const filteredRecords = filterRecords(records, filters);
      sendJson(res, 200, { records: filteredRecords });
      return;
    }

    if (requestedResource === 'records' && req.method === 'POST') {
      const requestBody = parseRequestBody(req.body);
      const incomingRecord = requestBody.record;
      if (!incomingRecord || typeof incomingRecord !== 'object') {
        throw createHttpError(400, 'Request body must include a "record" object');
      }

      const payload = buildRecordPayload(incomingRecord, sharePointConfigResolved.fieldMap);
      const createdItem = await createListItem(sharePointConfigResolved, payload);
      const createdRecord = mapListItemToMachineRecord(
        createdItem,
        sharePointConfigResolved.fieldMap,
        sharePointConfigResolved.siteOrigin
      );
      sendJson(res, 201, { record: createdRecord });
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
  } catch {
    /* SharePoint list items failed; empty array avoids 502 and preserves API contract. */
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
