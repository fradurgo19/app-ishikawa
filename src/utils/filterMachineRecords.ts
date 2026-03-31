import type { MachineRecord } from '../types';

const EXACT_MATCH_KEYS = new Set<keyof MachineRecord>([
  'brandId',
  'modelId',
  'sectionId',
  'tipoEquipoId',
  'activityTypeId',
  'activityId',
  'createdBy',
]);

export function filterMachineRecords(
  records: MachineRecord[],
  filters: Partial<MachineRecord>
): MachineRecord[] {
  const entries = Object.entries(filters).filter(
    ([, value]) => typeof value === 'string' && value.trim().length > 0
  ) as Array<[keyof MachineRecord, string]>;

  if (entries.length === 0) {
    return records;
  }

  return records.filter((record) =>
    entries.every(([key, filterValue]) => {
      const recordValue = String(record[key] ?? '').trim();
      const fv = filterValue.trim();
      if (!recordValue || !fv) {
        return false;
      }
      if (EXACT_MATCH_KEYS.has(key)) {
        return recordValue.toLowerCase() === fv.toLowerCase();
      }
      return recordValue.toLowerCase().includes(fv.toLowerCase());
    })
  );
}
