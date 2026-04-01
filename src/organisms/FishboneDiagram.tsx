import React, { useCallback, useEffect, useState } from 'react';
import type { LucideIcon } from 'lucide-react';
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

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-2 border-red-600 border-t-transparent"></div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-lg p-6 shadow-md">
      <h2 className="text-2xl font-bold text-gray-900 mb-4">Diagrama Ishikawa</h2>
      <p className="text-sm text-gray-600 mb-6">
        Espina horizontal: las causas se abren hacia arriba y hacia abajo, en alternancia.
      </p>
      <div className="overflow-x-auto overflow-y-visible pb-8 pt-4">
        {fishboneData.length > 0 ? (
          <div className="flex flex-row items-center justify-start min-w-min gap-0 px-2">
            {fishboneData.map((node, index) => (
              <React.Fragment key={node.id}>
                {index > 0 && (
                  <div
                    className="h-1 w-10 sm:w-14 shrink-0 bg-gray-400 rounded-full self-center"
                    aria-hidden
                  />
                )}
                <div className="shrink-0 flex flex-col items-center">
                  <FishboneBranch
                    node={node}
                    onToggle={toggleNode}
                    getNodeColor={getNodeColor}
                    getNodeIcon={getNodeIcon}
                  />
                </div>
              </React.Fragment>
            ))}
            <div
              className="h-1 w-10 sm:w-14 shrink-0 self-center bg-gray-400 rounded-full"
              aria-hidden
            />
            <div
              className="flex shrink-0 items-center gap-2 rounded-lg border-2 border-red-300 bg-red-50 px-4 py-3 text-sm font-semibold text-red-900"
              title="Efecto / foco del análisis"
            >
              <span className="hidden sm:inline">Efecto</span>
              <span className="max-w-[140px] truncate sm:max-w-[200px]">
                {selectedProblem || 'Análisis'}
              </span>
            </div>
          </div>
        ) : (
          <p className="text-gray-500 text-center py-8">
            No hay datos disponibles para los criterios seleccionados
          </p>
        )}
      </div>
    </div>
  );
};

function splitChildrenIntoUpperAndLowerRibs(children: FishboneNode[]): {
  upper: FishboneNode[];
  lower: FishboneNode[];
} {
  const upper: FishboneNode[] = [];
  const lower: FishboneNode[] = [];
  children.forEach((child, index) => {
    if (index % 2 === 0) {
      upper.push(child);
    } else {
      lower.push(child);
    }
  });
  return { upper, lower };
}

interface FishboneBranchProps {
  node: FishboneNode;
  onToggle: (nodeId: string) => void;
  getNodeColor: (type: FishboneNode['type']) => string;
  getNodeIcon: (type: FishboneNode['type']) => LucideIcon | null;
}

function FishboneRibConnector(): React.ReactElement {
  return <div className="h-10 w-px shrink-0 bg-gray-400" aria-hidden />;
}

interface FishboneRibColumnProps {
  child: FishboneNode;
  placement: 'upper' | 'lower';
  onToggle: (nodeId: string) => void;
  getNodeColor: (type: FishboneNode['type']) => string;
  getNodeIcon: (type: FishboneNode['type']) => LucideIcon | null;
}

function FishboneRibColumn({
  child,
  placement,
  onToggle,
  getNodeColor,
  getNodeIcon,
}: FishboneRibColumnProps) {
  const branch = (
    <div className="max-w-[220px]">
      <FishboneBranch
        node={child}
        onToggle={onToggle}
        getNodeColor={getNodeColor}
        getNodeIcon={getNodeIcon}
      />
    </div>
  );
  const connector = <FishboneRibConnector />;

  return (
    <div className="flex flex-col items-center">
      {placement === 'upper' ? (
        <>
          {branch}
          {connector}
        </>
      ) : (
        <>
          {connector}
          {branch}
        </>
      )}
    </div>
  );
}

function FishboneBranch({ node, onToggle, getNodeColor, getNodeIcon }: FishboneBranchProps) {
  const hasChildren = node.children.length > 0;
  const { upper, lower } = splitChildrenIntoUpperAndLowerRibs(node.children);
  const Icon = getNodeIcon(node.type);

  return (
    <div className="flex flex-col items-center">
      {node.expanded && upper.length > 0 && (
        <div className="mb-0 flex flex-row flex-wrap items-end justify-center gap-x-8 gap-y-4">
          {upper.map((child) => (
            <FishboneRibColumn
              key={child.id}
              child={child}
              placement="upper"
              onToggle={onToggle}
              getNodeColor={getNodeColor}
              getNodeIcon={getNodeIcon}
            />
          ))}
        </div>
      )}

      <div className="relative z-10 flex flex-row items-center">
        <FishboneNodeButton
          node={node}
          hasChildren={hasChildren}
          Icon={Icon}
          className={getNodeColor(node.type)}
          onToggle={() => hasChildren && onToggle(node.id)}
        />
      </div>

      {node.expanded && lower.length > 0 && (
        <div className="mt-0 flex flex-row flex-wrap items-start justify-center gap-x-8 gap-y-4">
          {lower.map((child) => (
            <FishboneRibColumn
              key={child.id}
              child={child}
              placement="lower"
              onToggle={onToggle}
              getNodeColor={getNodeColor}
              getNodeIcon={getNodeIcon}
            />
          ))}
        </div>
      )}
    </div>
  );
}

interface FishboneNodeButtonProps {
  node: FishboneNode;
  hasChildren: boolean;
  Icon: LucideIcon | null;
  className: string;
  onToggle: () => void;
}

function FishboneNodeButton({
  node,
  hasChildren,
  Icon,
  className,
  onToggle,
}: FishboneNodeButtonProps) {
  return (
    <button
      type="button"
      onClick={onToggle}
      aria-expanded={hasChildren ? node.expanded : undefined}
      className={`flex max-w-[240px] items-center gap-2 rounded-lg border-2 px-3 py-2 text-left text-sm transition-all duration-200 ${className} ${
        hasChildren ? 'cursor-pointer hover:shadow-md' : 'cursor-default opacity-95'
      }`}
      disabled={!hasChildren}
    >
      {hasChildren &&
        (node.expanded ? <ChevronDown size={16} className="shrink-0" /> : <ChevronRight size={16} className="shrink-0" />)}
      {Icon && <Icon size={16} className="shrink-0" />}
      <span className="font-medium break-words">{node.label}</span>
    </button>
  );
}

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