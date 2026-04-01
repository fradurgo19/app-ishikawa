import type { Activity } from '../types';

/**
 * Si activityId coincide con un id del catálogo, muestra el nombre; si no, el texto guardado en la lista
 * (p. ej. actividad como texto multilínea en SharePoint).
 */
export function resolveActivityDisplayLabel(activityId: string, activities: Activity[]): string {
  const trimmed = typeof activityId === 'string' ? activityId.trim() : '';
  if (!trimmed) {
    return 'Desconocido';
  }
  const fromCatalog = activities.find((a) => a.id === activityId);
  return fromCatalog?.name ?? activityId;
}
