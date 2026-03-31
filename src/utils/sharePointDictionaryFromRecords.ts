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

interface RecordUniques {
  uniqueTiposEquipo: Set<string>;
  uniqueBrands: Set<string>;
  uniqueModels: Map<string, Model>;
  uniqueSections: Map<string, Section>;
  uniqueActivityTypes: Set<string>;
  uniqueActivities: Map<string, Activity>;
}

function addTipoEquipoAndBrand(record: MachineRecord, u: RecordUniques): void {
  if (record.tipoEquipoId) {
    u.uniqueTiposEquipo.add(record.tipoEquipoId);
  }
  if (record.brandId) {
    u.uniqueBrands.add(record.brandId);
  }
}

function addModelFromRecord(record: MachineRecord, uniqueModels: Map<string, Model>): void {
  if (!record.brandId || !record.modelId) {
    return;
  }
  const modelKey = `${record.brandId}::${record.modelId}`;
  if (uniqueModels.has(modelKey)) {
    return;
  }
  uniqueModels.set(modelKey, {
    id: record.modelId,
    name: record.modelId,
    brandId: record.brandId,
  });
}

function addSectionFromRecord(record: MachineRecord, uniqueSections: Map<string, Section>): void {
  if (!record.brandId || !record.modelId || !record.sectionId) {
    return;
  }
  const sectionKey = `${record.brandId}::${record.modelId}::${record.sectionId}`;
  if (uniqueSections.has(sectionKey)) {
    return;
  }
  uniqueSections.set(sectionKey, {
    id: record.sectionId,
    name: record.sectionId,
    brandId: record.brandId,
    modelId: record.modelId,
  });
}

function addActivityTypeAndActivity(record: MachineRecord, u: RecordUniques): void {
  if (record.activityTypeId) {
    u.uniqueActivityTypes.add(record.activityTypeId);
  }
  if (!record.activityTypeId || !record.activityId) {
    return;
  }
  const activityKey = `${record.activityTypeId}::${record.activityId}`;
  if (u.uniqueActivities.has(activityKey)) {
    return;
  }
  u.uniqueActivities.set(activityKey, {
    id: record.activityId,
    name: record.activityId,
    activityTypeId: record.activityTypeId,
  });
}

function ingestRecordUniques(record: MachineRecord, u: RecordUniques): void {
  addTipoEquipoAndBrand(record, u);
  addModelFromRecord(record, u.uniqueModels);
  addSectionFromRecord(record, u.uniqueSections);
  addActivityTypeAndActivity(record, u);
}

function mergeFlatSectionsFromRecords(
  records: MachineRecord[],
  uniqueSections: Map<string, Section>,
  sectionIdsSeen: Set<string>
): void {
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
}

function sortedIdNameList(ids: Set<string>): Brand[] {
  return Array.from(ids)
    .map((value) => ({ id: value, name: value }))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));
}

export function buildDictionaryFromRecords(records: MachineRecord[]): DictionaryFromRecordsShape {
  const u: RecordUniques = {
    uniqueTiposEquipo: new Set<string>(),
    uniqueBrands: new Set<string>(),
    uniqueModels: new Map<string, Model>(),
    uniqueSections: new Map<string, Section>(),
    uniqueActivityTypes: new Set<string>(),
    uniqueActivities: new Map<string, Activity>(),
  };

  for (const record of records) {
    ingestRecordUniques(record, u);
  }

  const sectionIdsSeen = new Set(
    Array.from(u.uniqueSections.values()).map((s) => text(s.id).toLowerCase())
  );
  mergeFlatSectionsFromRecords(records, u.uniqueSections, sectionIdsSeen);

  const brands = sortedIdNameList(u.uniqueBrands);

  const models = Array.from(u.uniqueModels.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const sections = Array.from(u.uniqueSections.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const activityTypes = sortedIdNameList(u.uniqueActivityTypes);

  const activities = Array.from(u.uniqueActivities.values()).sort((a, b) =>
    a.name.localeCompare(b.name, 'es')
  );

  const tiposEquipo = sortedIdNameList(u.uniqueTiposEquipo);

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

interface MergedChoiceSets {
  section: Set<string>;
  activityType: Set<string>;
  activity: Set<string>;
  tipoEquipo: Set<string>;
  brand: Set<string>;
  model: Set<string>;
}

function stringSetFromOptionList(values: string[] | undefined): Set<string> {
  return new Set((values || []).map((s) => text(s)).filter(Boolean));
}

function createMergedChoiceSetsFromFieldOptions(
  fieldChoiceOptions: FieldChoiceOptionsShape
): MergedChoiceSets {
  return {
    section: stringSetFromOptionList(fieldChoiceOptions.section),
    activityType: stringSetFromOptionList(fieldChoiceOptions.activityType),
    activity: stringSetFromOptionList(fieldChoiceOptions.activity),
    tipoEquipo: stringSetFromOptionList(fieldChoiceOptions.tipoEquipo),
    brand: stringSetFromOptionList(fieldChoiceOptions.brand),
    model: stringSetFromOptionList(fieldChoiceOptions.model),
  };
}

function addEntityIdsToSet(entities: Array<{ id: string }>, target: Set<string>): void {
  for (const entity of entities) {
    const id = text(entity.id);
    if (id) {
      target.add(id);
    }
  }
}

function mergeDictionaryIdsIntoChoiceSets(
  dictionary: DictionaryFromRecordsShape,
  sets: MergedChoiceSets
): void {
  addEntityIdsToSet(dictionary.tiposEquipo, sets.tipoEquipo);
  addEntityIdsToSet(dictionary.brands, sets.brand);
  addEntityIdsToSet(dictionary.models, sets.model);
  addEntityIdsToSet(dictionary.sections, sets.section);
  addEntityIdsToSet(dictionary.activityTypes, sets.activityType);
  addEntityIdsToSet(dictionary.activities, sets.activity);
}

function addTrimmedTextToSet(raw: unknown, target: Set<string>): void {
  const v = text(raw);
  if (v) {
    target.add(v);
  }
}

function mergeRecordIdsIntoChoiceSets(records: MachineRecord[], sets: MergedChoiceSets): void {
  for (const record of records || []) {
    addTrimmedTextToSet(record.sectionId, sets.section);
    addTrimmedTextToSet(record.activityTypeId, sets.activityType);
    addTrimmedTextToSet(record.activityId, sets.activity);
    addTrimmedTextToSet(record.tipoEquipoId, sets.tipoEquipo);
    addTrimmedTextToSet(record.brandId, sets.brand);
    addTrimmedTextToSet(record.modelId, sets.model);
  }
}

function sortLocaleEs(values: Set<string>): string[] {
  return Array.from(values).sort((a, b) => a.localeCompare(b, 'es'));
}

function mergedChoiceSetsToShape(sets: MergedChoiceSets): FieldChoiceOptionsShape {
  return {
    section: sortLocaleEs(sets.section),
    activityType: sortLocaleEs(sets.activityType),
    activity: sortLocaleEs(sets.activity),
    tipoEquipo: sortLocaleEs(sets.tipoEquipo),
    brand: sortLocaleEs(sets.brand),
    model: sortLocaleEs(sets.model),
  };
}

export function mergeFieldChoiceOptionsFromRecordsAndDictionary(
  dictionary: DictionaryFromRecordsShape,
  records: MachineRecord[],
  fieldChoiceOptions: FieldChoiceOptionsShape
): FieldChoiceOptionsShape {
  const sets = createMergedChoiceSetsFromFieldOptions(fieldChoiceOptions);
  mergeDictionaryIdsIntoChoiceSets(dictionary, sets);
  mergeRecordIdsIntoChoiceSets(records, sets);
  return mergedChoiceSetsToShape(sets);
}
