/**
 * Adjuntos nativos de lista vía SharePoint REST (mismo patrón que la app Equipo: token delegado del sitio,
 * no el cuerpo JSON del ítem). POST .../AttachmentFiles/add con application/octet-stream.
 */

function escapeODataStringLiteral(value: string): string {
  return value.replace(/'/g, "''");
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
