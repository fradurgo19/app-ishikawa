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

function machineRecordFieldAsFilterString(record: MachineRecord, key: keyof MachineRecord): string {
  const value = record[key];
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'bigint') {
    return String(value);
  }
  return '';
}

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
      const recordValue = machineRecordFieldAsFilterString(record, key);
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
