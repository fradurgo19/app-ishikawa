/**
 * Expone solo metadatos no secretos (URL de sitio + título de lista) para que el SPA
 * resuelva Microsoft Graph con permisos delegados sin duplicar VITE_SHAREPOINT_* en el build.
 * No incluye tenant, client id ni secret.
 */
export default function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  if (req.method !== 'GET') {
    res.status(405).json({ message: 'Method not allowed' });
    return;
  }

  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');

  const rawSite = typeof process.env.SHAREPOINT_SITE_URL === 'string' ? process.env.SHAREPOINT_SITE_URL : '';
  const siteUrl = rawSite.replace(/\/$/, '');
  const listTitle =
    typeof process.env.SHAREPOINT_LIST_TITLE === 'string' ? process.env.SHAREPOINT_LIST_TITLE.trim() : '';

  res.status(200).json({ siteUrl, listTitle });
}
