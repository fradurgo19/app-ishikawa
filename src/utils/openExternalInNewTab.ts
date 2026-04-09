/**
 * Texto que puede abrirse en el navegador como URL (http(s) o mailto).
 */
export function looksLikeNavigableUrl(text: string): boolean {
  const t = text.trim();
  return /^https?:\/\//i.test(t) || /^mailto:/i.test(t);
}

/**
 * Abre la URL en una nueva pestaña para visualización (sin forzar descarga desde el cliente).
 * @returns true si se abrió una ventana (no bloqueada por el navegador).
 */
export function openUrlInNewBrowserTab(url: string): boolean {
  const trimmed = url.trim();
  if (!trimmed) {
    return false;
  }
  const handle = window.open(trimmed, '_blank', 'noopener,noreferrer');
  return handle !== null;
}
