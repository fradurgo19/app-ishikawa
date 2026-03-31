import type { Activity, ActivityType, Brand, KPIData, MachineRecord, Model, Section } from '../types';

export interface FieldChoiceOptionsShape {
  section: string[];
  activityType: string[];
  activity: string[];
  tipoEquipo?: string[];
  brand?: string[];
  model?: string[];
}

export interface DictionaryFromRecordsShape {
  tiposEquipo: Brand[];
  brands: Brand[];
  models: Model[];
  sections: Section[];
  activityTypes: ActivityType[];
  activities: Activity[];
  kpis: KPIData;
}

function text(value: unknown): string {
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

export function buildDictionaryFromRecords(records: MachineRecord[]): DictionaryFromRecordsShape {
  const uniqueTiposEquipo = new Set<string>();
  const uniqueBrands = new Set<string>();
  const uniqueModels = new Map<string, Model>();
  const uniqueSections = new Map<string, Section>();
  const uniqueActivityTypes = new Set<string>();
  const uniqueActivities = new Map<string, Activity>();

  for (const record of records) {
    if (record.tipoEquipoId) {
      uniqueTiposEquipo.add(record.tipoEquipoId);
    }
    if (record.brandId) {
      uniqueBrands.add(record.brandId);
    }
    if (record.brandId && record.modelId) {
      const modelKey = `${record.brandId}::${record.modelId}`;
      if (!uniqueModels.has(modelKey)) {
        uniqueModels.set(modelKey, {
          id: record.modelId,
          name: record.modelId,
          brandId: record.brandId,
        });
      }
    }
    if (record.brandId && record.modelId && record.sectionId) {
      const sectionKey = `${record.brandId}::${record.modelId}::${record.sectionId}`;
      if (!uniqueSections.has(sectionKey)) {
        uniqueSections.set(sectionKey, {
          id: record.sectionId,
          name: record.sectionId,
          brandId: record.brandId,
          modelId: record.modelId,
        });
      }
    }
    if (record.activityTypeId) {
      uniqueActivityTypes.add(record.activityTypeId);
    }
    if (record.activityTypeId && record.activityId) {
      const activityKey = `${record.activityTypeId}::${record.activityId}`;
      if (!uniqueActivities.has(activityKey)) {
        uniqueActivities.set(activityKey, {
          id: record.activityId,
          name: record.activityId,
          activityTypeId: record.activityTypeId,
        });
      }
    }
  }

  const sectionIdsSeen = new Set(
    Array.from(uniqueSections.values()).map((s) => text(s.id).toLowerCase())
  );
  for (const record of records) {
    const sid = text(record.sectionId);
    if (!sid) {
      continue;
    }
    const lower = sid.toLowerCase();
    if (sectionIdsSeen.has(lower)) {
      continue;
    }
    sectionIdsSeen.add(lower);
    uniqueSections.set(`flat:${lower}`, {
      id: sid,
      name: sid,
      brandId: text(record.brandId),
      modelId: text(record.modelId),
    });
  }

  const brands = Array.from(uniqueBrands)
    .map((value) => ({ id: value, name: value }))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));

  const models = Array.from(uniqueModels.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const sections = Array.from(uniqueSections.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const activityTypes = Array.from(uniqueActivityTypes)
    .map((value) => ({ id: value, name: value }))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));

  const activities = Array.from(uniqueActivities.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const tiposEquipo = Array.from(uniqueTiposEquipo)
    .map((value) => ({ id: value, name: value }))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));

  return {
    tiposEquipo,
    brands,
    models,
    sections,
    activityTypes,
    activities,
    kpis: {
      totalTiposEquipo: tiposEquipo.length,
      totalMarcas: brands.length,
      totalModelos: models.length,
      totalSecciones: sections.length,
      totalRegistros: records.length,
    },
  };
}

export function mergeFieldChoiceOptionsFromRecordsAndDictionary(
  dictionary: DictionaryFromRecordsShape,
  records: MachineRecord[],
  fieldChoiceOptions: FieldChoiceOptionsShape
): FieldChoiceOptionsShape {
  const sectionSet = new Set(
    (fieldChoiceOptions.section || []).map((s) => text(s)).filter(Boolean)
  );
  const activityTypeSet = new Set(
    (fieldChoiceOptions.activityType || []).map((s) => text(s)).filter(Boolean)
  );
  const activitySet = new Set(
    (fieldChoiceOptions.activity || []).map((s) => text(s)).filter(Boolean)
  );
  const tipoEquipoSet = new Set(
    (fieldChoiceOptions.tipoEquipo || []).map((s) => text(s)).filter(Boolean)
  );
  const brandSet = new Set(
    (fieldChoiceOptions.brand || []).map((s) => text(s)).filter(Boolean)
  );
  const modelSet = new Set(
    (fieldChoiceOptions.model || []).map((s) => text(s)).filter(Boolean)
  );

  for (const t of dictionary.tiposEquipo || []) {
    const id = text(t.id);
    if (id) {
      tipoEquipoSet.add(id);
    }
  }
  for (const b of dictionary.brands) {
    const id = text(b.id);
    if (id) {
      brandSet.add(id);
    }
  }
  for (const m of dictionary.models) {
    const id = text(m.id);
    if (id) {
      modelSet.add(id);
    }
  }
  for (const s of dictionary.sections) {
    const id = text(s.id);
    if (id) {
      sectionSet.add(id);
    }
  }
  for (const t of dictionary.activityTypes) {
    const id = text(t.id);
    if (id) {
      activityTypeSet.add(id);
    }
  }
  for (const a of dictionary.activities) {
    const id = text(a.id);
    if (id) {
      activitySet.add(id);
    }
  }

  for (const record of records || []) {
    const sid = text(record.sectionId);
    if (sid) {
      sectionSet.add(sid);
    }
    const tid = text(record.activityTypeId);
    if (tid) {
      activityTypeSet.add(tid);
    }
    const aid = text(record.activityId);
    if (aid) {
      activitySet.add(aid);
    }
    const te = text(record.tipoEquipoId);
    if (te) {
      tipoEquipoSet.add(te);
    }
    const bid = text(record.brandId);
    if (bid) {
      brandSet.add(bid);
    }
    const mid = text(record.modelId);
    if (mid) {
      modelSet.add(mid);
    }
  }

  return {
    section: Array.from(sectionSet).sort((a, b) => a.localeCompare(b, 'es')),
    activityType: Array.from(activityTypeSet).sort((a, b) => a.localeCompare(b, 'es')),
    activity: Array.from(activitySet).sort((a, b) => a.localeCompare(b, 'es')),
    tipoEquipo: Array.from(tipoEquipoSet).sort((a, b) => a.localeCompare(b, 'es')),
    brand: Array.from(brandSet).sort((a, b) => a.localeCompare(b, 'es')),
    model: Array.from(modelSet).sort((a, b) => a.localeCompare(b, 'es')),
  };
}
