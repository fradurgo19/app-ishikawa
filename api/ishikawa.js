import {
  buildDictionaryFromRecords,
  buildRecordPayload,
  createHttpError,
  createListItem,
  fetchAllListItems,
  filterRecords,
  getSharePointConfig,
  mapListItemToMachineRecord,
} from './_sharepoint.js';

const ALLOWED_RESOURCES = Object.freeze(['records', 'dictionary']);
const ALLOWED_RECORD_FILTERS = Object.freeze([
  'brandId',
  'modelId',
  'sectionId',
  'problem',
  'activityTypeId',
  'activityId',
  'resource',
  'createdBy',
]);

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

    if (requestedResource === 'dictionary') {
      enforceMethod(req.method, ['GET']);
      const records = await loadMappedRecords(sharePointConfig);
      const dictionary = buildDictionaryFromRecords(records);
      sendJson(res, 200, dictionary);
      return;
    }

    if (requestedResource === 'records' && req.method === 'GET') {
      const records = await loadMappedRecords(sharePointConfig);
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

      const payload = buildRecordPayload(incomingRecord, sharePointConfig.fieldMap);
      const createdItem = await createListItem(sharePointConfig, payload);
      const createdRecord = mapListItemToMachineRecord(
        createdItem,
        sharePointConfig.fieldMap,
        sharePointConfig.siteOrigin
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

async function loadMappedRecords(sharePointConfig) {
  const listItems = await fetchAllListItems(sharePointConfig);
  return listItems
    .map((item) =>
      mapListItemToMachineRecord(item, sharePointConfig.fieldMap, sharePointConfig.siteOrigin)
    )
    .filter(isRecordUsable);
}

function isRecordUsable(record) {
  return (
    Boolean(record.brandId) &&
    Boolean(record.modelId) &&
    Boolean(record.sectionId) &&
    Boolean(record.problem) &&
    Boolean(record.activityTypeId) &&
    Boolean(record.activityId)
  );
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
