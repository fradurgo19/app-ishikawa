import type { Attachment } from '../types';

/**
 * Adjuntos nativos de lista vía SharePoint REST (mismo patrón que la app Equipo: token delegado del sitio,
 * no el cuerpo JSON del ítem). POST .../AttachmentFiles/add con application/octet-stream.
 */

function escapeODataStringLiteral(value: string): string {
  return value.replaceAll("'", "''");
}

async function parseJsonSafe(response: Response): Promise<Record<string, unknown>> {
  try {
    const text = await response.text();
    if (!text) {
      return {};
    }
    return JSON.parse(text) as Record<string, unknown>;
  } catch {
    return {};
  }
}

async function getSharePointFormDigest(siteUrl: string, accessToken: string): Promise<string> {
  const base = siteUrl.replace(/\/$/, '');
  const response = await fetch(`${base}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
    },
    body: '{}',
  });
  const payload = await parseJsonSafe(response);
  if (!response.ok) {
    throw new Error(
      `SharePoint contextinfo: HTTP ${response.status} ${JSON.stringify(payload).slice(0, 200)}`
    );
  }
  const d = payload.d as Record<string, unknown> | undefined;
  const digest =
    (typeof payload.FormDigestValue === 'string' && payload.FormDigestValue) ||
    (d && typeof d.FormDigestValue === 'string' && d.FormDigestValue) ||
    '';
  if (!digest) {
    throw new Error('SharePoint: respuesta sin FormDigestValue');
  }
  return digest;
}

/** Guía de producto (SharePoint admite más; Vercel/API tienen otros límites si se usa fallback). */
const MAX_ATTACHMENT_BYTES_CLIENT = 10 * 1024 * 1024;

/**
 * Sube un archivo a la columna nativa Attachments del ítem de lista.
 * Requiere token MSAL con ámbito `https://{host}.sharepoint.com/.default`.
 */
export async function uploadListItemAttachmentRest(options: {
  siteUrl: string;
  listTitle: string;
  itemId: string;
  file: File;
  accessToken: string;
}): Promise<void> {
  const { siteUrl, listTitle, itemId, file, accessToken } = options;
  if (file.size > MAX_ATTACHMENT_BYTES_CLIENT) {
    throw new Error(
      `El archivo "${file.name}" supera el máximo de ${Math.floor(MAX_ATTACHMENT_BYTES_CLIENT / (1024 * 1024))} MB.`
    );
  }

  const base = siteUrl.replace(/\/$/, '');
  const id = Number.parseInt(itemId, 10);
  if (!Number.isInteger(id) || id < 1) {
    throw new Error('Id de ítem de lista no válido para adjuntos');
  }

  const safeTitle = escapeODataStringLiteral(listTitle.trim());
  const safeName = escapeODataStringLiteral((file.name || 'adjunto').trim() || 'adjunto');
  const addUrl = `${base}/_api/web/lists/GetByTitle('${safeTitle}')/items(${id})/AttachmentFiles/add(FileName='${safeName}')`;

  let digest: string | undefined;
  try {
    digest = await getSharePointFormDigest(base, accessToken);
  } catch {
    digest = undefined;
  }

  const body = await file.arrayBuffer();
  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/octet-stream',
  };
  if (digest) {
    headers['X-RequestDigest'] = digest;
  }

  const response = await fetch(addUrl, {
    method: 'POST',
    headers,
    body,
  });

  if (!response.ok) {
    const errBody = await parseJsonSafe(response);
    throw new Error(
      `No se pudo subir "${file.name}": HTTP ${response.status} ${JSON.stringify(errBody).slice(0, 280)}`
    );
  }
}

/**
 * Elimina un archivo de la columna nativa Attachments (DELETE SharePoint REST).
 * 404 se ignora (el adjunto ya no existe).
 */
export async function deleteListItemAttachmentRest(options: {
  siteUrl: string;
  listTitle: string;
  itemId: string;
  fileName: string;
  accessToken: string;
}): Promise<void> {
  const { siteUrl, listTitle, itemId, fileName, accessToken } = options;
  const base = siteUrl.replace(/\/$/, '');
  const id = Number.parseInt(itemId, 10);
  if (!Number.isInteger(id) || id < 1) {
    throw new Error('Id de ítem de lista no válido para eliminar adjunto');
  }
  const name = (fileName || '').trim();
  if (!name) {
    return;
  }
  const safeTitle = escapeODataStringLiteral(listTitle.trim());
  const safeName = escapeODataStringLiteral(name);
  const deleteUrl = `${base}/_api/web/lists/GetByTitle('${safeTitle}')/items(${id})/AttachmentFiles('${safeName}')`;

  let digest: string | undefined;
  try {
    digest = await getSharePointFormDigest(base, accessToken);
  } catch {
    digest = undefined;
  }

  const headers: Record<string, string> = {
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
    const errBody = await parseJsonSafe(response);
    throw new Error(
      `No se pudo eliminar el adjunto "${name}": HTTP ${response.status} ${JSON.stringify(errBody).slice(0, 280)}`
    );
  }
}

interface RestAttachmentFileRow {
  FileName?: string;
  ServerRelativeUrl?: string;
  Length?: number;
}

function toAbsoluteSharePointFileUrl(serverRelativeUrl: string, siteUrl: string): string {
  const relative = (serverRelativeUrl ?? '').trim();
  if (!relative) {
    return '';
  }
  if (relative.startsWith('http://') || relative.startsWith('https://')) {
    return relative;
  }
  try {
    const origin = new URL(siteUrl.replace(/\/$/, '')).origin;
    const path = relative.startsWith('/') ? relative : `/${relative}`;
    return `${origin}${path}`;
  } catch {
    return relative;
  }
}

function extractRestAttachmentFileRows(payload: Record<string, unknown>): RestAttachmentFileRow[] {
  const d = payload.d as Record<string, unknown> | undefined;
  const fromD = d?.results;
  if (Array.isArray(fromD)) {
    return fromD as RestAttachmentFileRow[];
  }
  const value = payload.value;
  if (Array.isArray(value)) {
    return value as RestAttachmentFileRow[];
  }
  return [];
}

/**
 * Lista archivos de la columna nativa Attachments del ítem (GET SharePoint REST).
 * Devuelve URLs absolutas listas para abrir en el navegador (sesión Microsoft).
 */
export async function fetchListItemAttachmentFilesRest(options: {
  siteUrl: string;
  listTitle: string;
  itemId: string;
  accessToken: string;
}): Promise<Attachment[]> {
  const { siteUrl, listTitle, itemId, accessToken } = options;
  const base = siteUrl.replace(/\/$/, '');
  const id = Number.parseInt(itemId, 10);
  if (!Number.isInteger(id) || id < 1) {
    return [];
  }
  const safeTitle = escapeODataStringLiteral(listTitle.trim());
  const url = `${base}/_api/web/lists/GetByTitle('${safeTitle}')/items(${id})/AttachmentFiles`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=verbose',
    },
  });

  if (!response.ok) {
    return [];
  }

  const payload = await parseJsonSafe(response);
  const rows = extractRestAttachmentFileRows(payload);

  return rows.map((row, index) => {
    const fileName = (row.FileName ?? 'Adjunto').trim() || 'Adjunto';
    const serverRelative = (row.ServerRelativeUrl ?? '').trim();
    return {
      id: `attachment-${itemId}-${index}-${fileName}`,
      name: fileName,
      url: toAbsoluteSharePointFileUrl(serverRelative, base),
      type: 'application/octet-stream',
      size: Number(row.Length) || 0,
    };
  });
}
