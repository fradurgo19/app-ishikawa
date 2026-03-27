import React, { useCallback, useEffect, useState } from 'react';
import {
  Activity,
  ActivityType,
  Brand,
  FishboneNode,
  MachineRecord,
  Model,
  Section,
} from '../types';
import { sharePointService } from '../services/sharePointService';
import { normalizeLabel } from '../data/equipmentMatrix';
import { ChevronRight, ChevronDown, Clock, Paperclip, PenTool as Tool } from 'lucide-react';

interface FishboneDiagramProps {
  selectedTipoEquipo?: string;
  selectedBrand?: string;
  selectedModel?: string;
  selectedProblem?: string;
}

export const FishboneDiagram: React.FC<FishboneDiagramProps> = ({
  selectedTipoEquipo,
  selectedBrand,
  selectedModel,
  selectedProblem,
}) => {
  const [fishboneData, setFishboneData] = useState<FishboneNode[]>([]);
  const [loading, setLoading] = useState(true);

  const loadFishboneData = useCallback(async () => {
    setLoading(true);
    try {
      await sharePointService.refreshDictionary?.();
      const [brands, models, sections, activityTypes, activities, records] = await Promise.all([
        sharePointService.getBrands(),
        sharePointService.getModels(),
        sharePointService.getSections(),
        sharePointService.getActivityTypes(),
        sharePointService.getActivities(),
        sharePointService.getRecords(),
      ]);

      const fishboneNodes = buildFishboneNodes(
        { brands, models, sections, activityTypes, activities, records },
        { selectedTipoEquipo, selectedBrand, selectedModel, selectedProblem }
      );

      setFishboneData(fishboneNodes);
    } catch (error) {
      console.error('Error cargando datos del diagrama:', error);
    } finally {
      setLoading(false);
    }
  }, [selectedTipoEquipo, selectedBrand, selectedModel, selectedProblem]);

  useEffect(() => {
    void loadFishboneData();
  }, [loadFishboneData]);

  const toggleNode = (nodeId: string) => {
    const toggleNodeRecursive = (nodes: FishboneNode[]): FishboneNode[] => {
      return nodes.map(node => {
        if (node.id === nodeId) {
          return { ...node, expanded: !node.expanded };
        }
        return { ...node, children: toggleNodeRecursive(node.children) };
      });
    };

    setFishboneData((previousNodes) => toggleNodeRecursive(previousNodes));
  };

  const getNodeIcon = (type: FishboneNode['type']) => {
    switch (type) {
      case 'tiempo':
        return Clock;
      case 'adjunto':
        return Paperclip;
      case 'recurso':
      case 'actividad':
      case 'tipoActividad':
        return Tool;
      default:
        return null;
    }
  };

  const getNodeColor = (type: FishboneNode['type']) => {
    const colors = {
      marca: 'bg-red-100 text-red-800 border-red-200',
      modelo: 'bg-blue-100 text-blue-800 border-blue-200',
      seccion: 'bg-green-100 text-green-800 border-green-200',
      problema: 'bg-yellow-100 text-yellow-800 border-yellow-200',
      tipoActividad: 'bg-purple-100 text-purple-800 border-purple-200',
      actividad: 'bg-indigo-100 text-indigo-800 border-indigo-200',
      recurso: 'bg-teal-100 text-teal-800 border-teal-200',
      tiempo: 'bg-orange-100 text-orange-800 border-orange-200',
      adjunto: 'bg-pink-100 text-pink-800 border-pink-200',
    };
    return colors[type];
  };

  const renderNode = (node: FishboneNode, level = 0) => {
    const hasChildren = node.children.length > 0;
    const Icon = getNodeIcon(node.type);
    
    return (
      <div key={node.id} style={{ marginLeft: `${level}rem` }}>
        <div className="flex items-center mb-2">
          <button
            onClick={() => hasChildren && toggleNode(node.id)}
            className={`flex items-center gap-2 p-2 rounded-lg border-2 transition-all duration-200 ${getNodeColor(node.type)} ${
              hasChildren ? 'hover:shadow-md cursor-pointer' : ''
            }`}
            disabled={!hasChildren}
          >
            {hasChildren && (
              node.expanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />
            )}
            {Icon && <Icon size={16} />}
            <span className="font-medium">{node.label}</span>
          </button>
        </div>
        
        {node.expanded && node.children.length > 0 && (
          <div className="ml-6 border-l-2 border-gray-200 pl-4">
            {node.children.map(child => renderNode(child, level + 1))}
          </div>
        )}
      </div>
    );
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-2 border-red-600 border-t-transparent"></div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-lg p-6 shadow-md">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Diagrama Ishikawa</h2>
      <div className="space-y-4">
        {fishboneData.length > 0 ? (
          fishboneData.map(node => renderNode(node))
        ) : (
          <p className="text-gray-500 text-center py-8">
            No hay datos disponibles para los criterios seleccionados
          </p>
        )}
      </div>
    </div>
  );
};

interface FishboneDataBundle {
  brands: Brand[];
  models: Model[];
  sections: Section[];
  activityTypes: ActivityType[];
  activities: Activity[];
  records: MachineRecord[];
}

interface FishboneFilters {
  selectedTipoEquipo?: string;
  selectedBrand?: string;
  selectedModel?: string;
  selectedProblem?: string;
}

function matchesBrandFilter(filters: FishboneFilters, brand: Brand): boolean {
  if (!filters.selectedBrand) {
    return true;
  }
  return (
    brand.id === filters.selectedBrand ||
    normalizeLabel(brand.name) === normalizeLabel(filters.selectedBrand)
  );
}

function matchesModelFilter(filters: FishboneFilters, model: Model): boolean {
  if (!filters.selectedModel) {
    return true;
  }
  return (
    model.id === filters.selectedModel ||
    normalizeLabel(model.name) === normalizeLabel(filters.selectedModel)
  );
}

function mergeBrandsFromRecords(brands: Brand[], records: MachineRecord[]): Brand[] {
  const map = new Map<string, Brand>();
  for (const b of brands) {
    map.set(b.id.toLowerCase(), b);
  }
  for (const r of records) {
    const id = r.brandId?.trim();
    if (!id) {
      continue;
    }
    const key = id.toLowerCase();
    if (!map.has(key)) {
      map.set(key, { id, name: id });
    }
  }
  return Array.from(map.values());
}

function mergeModelsFromRecords(models: Model[], records: MachineRecord[]): Model[] {
  const map = new Map<string, Model>();
  for (const m of models) {
    map.set(`${m.brandId}::${m.id}`.toLowerCase(), m);
  }
  for (const r of records) {
    const bid = r.brandId?.trim();
    const mid = r.modelId?.trim();
    if (!bid || !mid) {
      continue;
    }
    const key = `${bid}::${mid}`.toLowerCase();
    if (!map.has(key)) {
      map.set(key, { id: mid, name: mid, brandId: bid });
    }
  }
  return Array.from(map.values());
}

/**
 * Secciones bajo marca/modelo: diccionario + valores presentes en registros (SharePoint),
 * para no perder filas cuando el diccionario no tiene el triple exacto.
 */
function sectionsForBrandAndModel(
  brand: Brand,
  model: Model,
  sections: Section[],
  records: MachineRecord[]
): Section[] {
  const fromDict = sections.filter(
    (s) => s.brandId === brand.id && s.modelId === model.id
  );
  const byId = new Map<string, Section>();
  for (const s of fromDict) {
    byId.set(s.id.toLowerCase(), s);
  }
  for (const r of records) {
    if (r.brandId !== brand.id || r.modelId !== model.id) {
      continue;
    }
    const sid = r.sectionId?.trim();
    if (!sid) {
      continue;
    }
    const key = sid.toLowerCase();
    if (!byId.has(key)) {
      byId.set(key, { id: sid, name: sid, brandId: brand.id, modelId: model.id });
    }
  }
  return Array.from(byId.values()).sort((a, b) => a.name.localeCompare(b.name, 'es'));
}

function buildFishboneNodes(
  dataBundle: FishboneDataBundle,
  filters: FishboneFilters
): FishboneNode[] {
  const brands = mergeBrandsFromRecords(dataBundle.brands, dataBundle.records);
  const models = mergeModelsFromRecords(dataBundle.models, dataBundle.records);
  const filteredBrands = brands.filter((brand) => matchesBrandFilter(filters, brand));

  return filteredBrands.map((brand) =>
    buildBrandNode(brand, models, dataBundle.sections, dataBundle, filters)
  );
}

function buildBrandNode(
  brand: Brand,
  models: Model[],
  sections: Section[],
  dataBundle: FishboneDataBundle,
  filters: FishboneFilters
): FishboneNode {
  const brandModels = models.filter(
    (model) => model.brandId === brand.id && matchesModelFilter(filters, model)
  );

  return {
    id: brand.id,
    type: 'marca',
    label: brand.name,
    expanded: false,
    children: brandModels.map((model) =>
      buildModelNode(brand, model, sections, dataBundle.records, dataBundle.activityTypes, dataBundle.activities, filters)
    ),
  };
}

function buildModelNode(
  brand: Brand,
  model: Model,
  sections: Section[],
  records: MachineRecord[],
  activityTypes: ActivityType[],
  activities: Activity[],
  filters: FishboneFilters
): FishboneNode {
  const modelSections = sectionsForBrandAndModel(brand, model, sections, records);

  return {
    id: `${brand.id}-${model.id}`,
    type: 'modelo',
    label: model.name,
    expanded: false,
    children: modelSections.map((section) =>
      buildSectionNode(brand, model, section, records, activityTypes, activities, filters)
    ),
  };
}

function buildSectionNode(
  brand: Brand,
  model: Model,
  section: Section,
  records: MachineRecord[],
  activityTypes: ActivityType[],
  activities: Activity[],
  filters: FishboneFilters
): FishboneNode {
  const sectionRecords = records.filter((record) =>
    shouldIncludeRecord(record, brand.id, model.id, section.id, filters)
  );

  return {
    id: `${brand.id}-${model.id}-${section.id}`,
    type: 'seccion',
    label: section.name,
    expanded: false,
    children: sectionRecords.map((record) =>
      buildProblemNode(record, activityTypes, activities)
    ),
  };
}

function shouldIncludeRecord(
  record: MachineRecord,
  brandId: string,
  modelId: string,
  sectionId: string,
  filters: FishboneFilters
): boolean {
  if (record.brandId !== brandId || record.modelId !== modelId || record.sectionId !== sectionId) {
    return false;
  }

  if (filters.selectedTipoEquipo) {
    if (normalizeLabel(record.tipoEquipoId) !== normalizeLabel(filters.selectedTipoEquipo)) {
      return false;
    }
  }

  if (!filters.selectedProblem) {
    return true;
  }

  return record.problem.toLowerCase().includes(filters.selectedProblem.toLowerCase());
}

function buildProblemNode(
  record: MachineRecord,
  activityTypes: ActivityType[],
  activities: Activity[]
): FishboneNode {
  return {
    id: `problem-${record.id}`,
    type: 'problema',
    label: record.problem,
    expanded: false,
    data: record,
    children: [buildActivityTypeNode(record, activityTypes, activities)],
  };
}

function buildActivityTypeNode(
  record: MachineRecord,
  activityTypes: ActivityType[],
  activities: Activity[]
): FishboneNode {
  const activityTypeLabel =
    activityTypes.find((activityType) => activityType.id === record.activityTypeId)?.name ||
    'Desconocido';

  return {
    id: `activity-type-${record.id}`,
    type: 'tipoActividad',
    label: activityTypeLabel,
    expanded: false,
    children: [buildActivityNode(record, activities)],
  };
}

function buildActivityNode(record: MachineRecord, activities: Activity[]): FishboneNode {
  const activityLabel =
    activities.find((activity) => activity.id === record.activityId)?.name || 'Desconocido';

  return {
    id: `activity-${record.id}`,
    type: 'actividad',
    label: activityLabel,
    expanded: false,
    children: buildDetailNodes(record),
  };
}

function buildDetailNodes(record: MachineRecord): FishboneNode[] {
  const detailNodes: FishboneNode[] = [
    {
      id: `resource-${record.id}`,
      type: 'recurso',
      label: record.resource,
      expanded: false,
      children: [],
    },
    {
      id: `time-${record.id}`,
      type: 'tiempo',
      label: `${record.time} minutos`,
      expanded: false,
      children: [],
    },
  ];

  if (record.attachment) {
    detailNodes.push({
      id: `attachment-${record.id}`,
      type: 'adjunto',
      label: record.attachment.name,
      expanded: false,
      data: record.attachment,
      children: [],
    });
  }

  return detailNodes;
}