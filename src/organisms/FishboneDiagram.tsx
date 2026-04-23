import React, { useCallback, useEffect, useRef, useState } from 'react';
import { useLocation, useNavigate } from 'react-router-dom';
import type { LucideIcon } from 'lucide-react';
import {
  Activity,
  ActivityType,
  Brand,
  FishboneAttachmentLeafDetail,
  FishboneDiagramDetailPayload,
  FishboneNode,
  FishboneResourceLeafDetail,
  MachineRecord,
  Model,
  Section,
} from '../types';
import { sharePointService } from '../services/sharePointService';
import { normalizeLabel } from '../data/equipmentMatrix';
import { resolveActivityDisplayLabel } from '../utils/resolveActivityDisplayLabel';
import { looksLikeNavigableUrl, openUrlInNewBrowserTab } from '../utils/openExternalInNewTab';
import { ChevronRight, ChevronDown, Clock, Paperclip, PenTool as Tool } from 'lucide-react';

interface FishboneDiagramProps {
  selectedTipoEquipo?: string;
  selectedBrand?: string;
  selectedModel?: string;
  selectedProblem?: string;
}

function isFishboneResourceLeafDetail(data: unknown): data is FishboneResourceLeafDetail {
  if (!data || typeof data !== 'object') {
    return false;
  }
  const o = data as Record<string, unknown>;
  return typeof o.recordId === 'string' && typeof o.resourceText === 'string' && Array.isArray(o.allAttachments);
}

function isFishboneAttachmentLeafDetail(data: unknown): data is FishboneAttachmentLeafDetail {
  if (!data || typeof data !== 'object') {
    return false;
  }
  const o = data as Record<string, unknown>;
  const att = o.attachment;
  if (!att || typeof att !== 'object') {
    return false;
  }
  const a = att as Record<string, unknown>;
  return typeof o.recordId === 'string' && Array.isArray(o.allAttachments) && typeof a.id === 'string';
}

function fishboneNodeOpensDetailView(node: FishboneNode): boolean {
  if (node.type === 'recurso' && isFishboneResourceLeafDetail(node.data)) {
    return true;
  }
  return node.type === 'adjunto' && isFishboneAttachmentLeafDetail(node.data);
}

/** Intenta abrir recurso/adjunto en nueva pestaña; si hay URL y el navegador no bloquea, no hace falta la pantalla de detalle. */
function findFishboneNodeById(nodes: FishboneNode[], id: string): FishboneNode | undefined {
  for (const n of nodes) {
    if (n.id === id) {
      return n;
    }
    const found = findFishboneNodeById(n.children, id);
    if (found) {
      return found;
    }
  }
  return undefined;
}

function mapToggleFishboneExpanded(nodes: FishboneNode[], nodeId: string): FishboneNode[] {
  return nodes.map((node) => {
    if (node.id === nodeId) {
      return { ...node, expanded: !node.expanded };
    }
    return { ...node, children: mapToggleFishboneExpanded(node.children, nodeId) };
  });
}

function queryFishboneNodeButton(root: HTMLElement, nodeId: string): HTMLElement | null {
  const list = root.querySelectorAll<HTMLElement>('[data-fishbone-node]');
  for (const el of list) {
    if (el.dataset.fishboneNode === nodeId) {
      return el;
    }
  }
  return null;
}

function tryOpenDiagramLeafInNewTab(node: FishboneNode): boolean {
  if (node.type === 'recurso' && isFishboneResourceLeafDetail(node.data)) {
    const text = node.data.resourceText.trim();
    if (looksLikeNavigableUrl(text)) {
      return openUrlInNewBrowserTab(text);
    }
    return false;
  }
  if (node.type === 'adjunto' && isFishboneAttachmentLeafDetail(node.data)) {
    const url = node.data.attachment.url?.trim();
    if (url) {
      return openUrlInNewBrowserTab(url);
    }
    return false;
  }
  return false;
}

export const FishboneDiagram: React.FC<FishboneDiagramProps> = ({
  selectedTipoEquipo,
  selectedBrand,
  selectedModel,
  selectedProblem,
}) => {
  const [fishboneData, setFishboneData] = useState<FishboneNode[]>([]);
  const [loading, setLoading] = useState(true);
  const diagramScrollRef = useRef<HTMLDivElement>(null);
  const navigate = useNavigate();
  const location = useLocation();

  const scrollFishboneNodeIntoView = useCallback((nodeId: string) => {
    const root = diagramScrollRef.current;
    if (!root) {
      return;
    }
    const target = queryFishboneNodeButton(root, nodeId);
    target?.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
  }, []);

  const openLeafDetail = useCallback(
    (node: FishboneNode) => {
      if (tryOpenDiagramLeafInNewTab(node)) {
        return;
      }

      const returnTo = `${location.pathname}${location.search}`;
      const baseState =
        location.state !== null && typeof location.state === 'object'
          ? { ...location.state, fromDataTable: true as const }
          : { fromDataTable: true as const };

      if (node.type === 'recurso' && isFishboneResourceLeafDetail(node.data)) {
        const diagramDetail: FishboneDiagramDetailPayload = {
          kind: 'resource',
          recordId: node.data.recordId,
          resourceText: node.data.resourceText,
          allAttachments: node.data.allAttachments,
        };
        navigate('/fishbone/detail', { state: { ...baseState, returnTo, diagramDetail } });
        return;
      }

      if (node.type === 'adjunto' && isFishboneAttachmentLeafDetail(node.data)) {
        const diagramDetail: FishboneDiagramDetailPayload = {
          kind: 'attachments',
          recordId: node.data.recordId,
          focusAttachmentId: node.data.attachment.id,
          allAttachments: node.data.allAttachments,
        };
        navigate('/fishbone/detail', { state: { ...baseState, returnTo, diagramDetail } });
      }
    },
    [navigate, location.pathname, location.search, location.state]
  );

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

  const toggleNode = useCallback(
    (nodeId: string) => {
      let willExpand = false;
      setFishboneData((previousNodes) => {
        const current = findFishboneNodeById(previousNodes, nodeId);
        willExpand = Boolean(current && !current.expanded && current.children.length > 0);
        return mapToggleFishboneExpanded(previousNodes, nodeId);
      });
      if (willExpand) {
        queueMicrotask(() => {
          requestAnimationFrame(() => {
            scrollFishboneNodeIntoView(nodeId);
          });
        });
      }
    },
    [scrollFishboneNodeIntoView]
  );

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
    <div className="min-h-0 bg-white rounded-lg p-6 shadow-md">
      <h2 className="text-2xl font-bold text-gray-900 mb-4">Diagrama Ishikawa</h2>
      <p className="text-sm text-gray-600 mb-4">
        Vista vertical: el efecto va arriba y la espina desciende por marcas y modelos. Las causas se alternan a
        izquierda y derecha en cada nivel (igual que antes arriba/abajo, ahora en los laterales).
      </p>
      <div
        ref={diagramScrollRef}
        className="min-h-0 max-h-[min(calc(100dvh-15rem),1100px)] overflow-x-auto overflow-y-auto overscroll-x-contain overscroll-y-contain scroll-pl-4 scroll-pr-4 pb-10 pl-3 pr-3 pt-2 sm:scroll-pl-6 sm:scroll-pr-6 sm:pl-5 sm:pr-5 [-webkit-overflow-scrolling:touch]"
      >
        {fishboneData.length > 0 ? (
          <div className="mx-auto flex w-max min-w-full flex-col items-stretch pb-4">
            <div className="flex shrink-0 flex-col items-center px-1 pb-4">
              <div
                className="flex max-w-[min(100%,320px)] flex-col items-center gap-1.5 rounded-lg border-2 border-red-300 bg-red-50 px-3 py-2 text-center text-sm font-semibold text-red-900 sm:flex-row sm:text-left"
                title="Efecto / foco del análisis"
              >
                <span className="hidden sm:inline">Efecto</span>
                <span className="max-w-[min(90vw,320px)] whitespace-pre-wrap break-words sm:max-w-[280px]">
                  {selectedProblem || 'Análisis'}
                </span>
              </div>
            </div>
            {fishboneData.map((node, index) => (
              <div
                key={node.id}
                className={`flex w-full shrink-0 flex-col items-center border-t border-gray-300 ${index === 0 ? 'pt-4' : 'pt-5'}`}
              >
                <FishboneBranch
                  node={node}
                  depth={0}
                  onOpenLeafDetail={openLeafDetail}
                  onToggle={toggleNode}
                  getNodeColor={getNodeColor}
                  getNodeIcon={getNodeIcon}
                />
              </div>
            ))}
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

/**
 * Sin herencia: alterna causas a izquierda y derecha de la espina vertical.
 * Con herencia: los hijos repiten el mismo lateral (izq./der.) que el padre.
 */
function splitChildrenIntoUpperAndLowerRibs(
  children: FishboneNode[],
  inheritedRib?: 'upper' | 'lower'
): {
  upper: FishboneNode[];
  lower: FishboneNode[];
} {
  if (inheritedRib === 'upper') {
    return { upper: [...children], lower: [] };
  }
  if (inheritedRib === 'lower') {
    return { upper: [], lower: [...children] };
  }

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
  /** 0 = marca, 1 = modelo, ≥2 = sección y niveles inferiores (espaciado más amplio entre tarjetas). */
  depth?: number;
  /** Desde qué lado de la espina cuelga esta rama; sus hijos repiten el mismo lado. */
  inheritedRib?: 'upper' | 'lower';
  onOpenLeafDetail?: (node: FishboneNode) => void;
  onToggle: (nodeId: string) => void;
  getNodeColor: (type: FishboneNode['type']) => string;
  getNodeIcon: (type: FishboneNode['type']) => LucideIcon | null;
}

/** Tramo horizontal hacia la espina central (diagrama vertical). */
function FishboneRibHorizontalConnector(): React.ReactElement {
  return <div className="h-px w-4 shrink-0 bg-gray-400 sm:w-6 md:w-8" aria-hidden />;
}

interface FishboneRibColumnProps {
  child: FishboneNode;
  childDepth: number;
  placement: 'upper' | 'lower';
  onOpenLeafDetail?: (node: FishboneNode) => void;
  onToggle: (nodeId: string) => void;
  getNodeColor: (type: FishboneNode['type']) => string;
  getNodeIcon: (type: FishboneNode['type']) => LucideIcon | null;
}

function FishboneRibColumn({
  child,
  childDepth,
  placement,
  onOpenLeafDetail,
  onToggle,
  getNodeColor,
  getNodeIcon,
}: Readonly<FishboneRibColumnProps>) {
  const branch = (
    <div className="min-w-0 w-full max-w-[min(100%,280px)] md:max-w-none">
      <FishboneBranch
        node={child}
        depth={childDepth}
        inheritedRib={placement}
        onOpenLeafDetail={onOpenLeafDetail}
        onToggle={onToggle}
        getNodeColor={getNodeColor}
        getNodeIcon={getNodeIcon}
      />
    </div>
  );
  const connector = <FishboneRibHorizontalConnector />;
  const ribRowPad = childDepth >= 2 ? 'py-5' : 'py-1';
  const ribGap = childDepth >= 2 ? 'gap-5 sm:gap-5' : 'gap-1.5 sm:gap-2';

  if (placement === 'upper') {
    return (
      <div
        className={`flex w-full min-w-0 max-w-full flex-row flex-wrap items-center justify-start md:justify-end sm:flex-nowrap ${ribRowPad} ${ribGap}`}
      >
        {branch}
        {connector}
      </div>
    );
  }

  return (
    <div
      className={`flex w-full min-w-0 max-w-full flex-row flex-wrap items-center justify-start sm:flex-nowrap ${ribRowPad} ${ribGap}`}
    >
      {connector}
      {branch}
    </div>
  );
}

function FishboneBranch({
  node,
  depth = 0,
  inheritedRib,
  onOpenLeafDetail,
  onToggle,
  getNodeColor,
  getNodeIcon,
}: Readonly<FishboneBranchProps>) {
  const hasChildren = node.children.length > 0;
  const { upper, lower } = splitChildrenIntoUpperAndLowerRibs(node.children, inheritedRib);
  const Icon = getNodeIcon(node.type);
  const compactDepth = depth < 2;
  const nextDepth = depth + 1;

  const branchLayout = compactDepth
    ? 'gap-1.5 md:gap-x-2 md:gap-y-1.5 lg:gap-x-3'
    : 'gap-7 md:gap-x-8 md:gap-y-5 lg:gap-x-10';
  const ribStackGap = compactDepth ? 'gap-2' : 'gap-8 md:gap-12';

  if (node.type === 'actividad' && node.expanded && node.children.length > 0) {
    return (
      <div className="flex w-full min-w-0 flex-col items-center gap-5">
        <div className="flex w-full justify-center px-1">
          <FishboneNodeButton
            node={node}
            hasChildren={hasChildren}
            Icon={Icon}
            className={getNodeColor(node.type)}
            onOpenLeafDetail={onOpenLeafDetail}
            onToggle={() => hasChildren && onToggle(node.id)}
          />
        </div>
        <div className="flex w-full min-w-0 max-w-2xl flex-col items-center gap-6 border-t border-gray-200 pt-6">
          {node.children.map((child) => (
            <div key={child.id} className="flex w-full justify-center px-1">
              <FishboneBranch
                node={child}
                depth={nextDepth}
                inheritedRib={undefined}
                onOpenLeafDetail={onOpenLeafDetail}
                onToggle={onToggle}
                getNodeColor={getNodeColor}
                getNodeIcon={getNodeIcon}
              />
            </div>
          ))}
        </div>
      </div>
    );
  }

  return (
    <div
      className={`flex w-full min-w-0 max-w-full flex-col items-stretch md:flex-row md:items-start md:justify-center ${branchLayout}`}
    >
      {node.expanded && upper.length > 0 && (
        <div
          className={`order-2 flex w-full min-w-0 max-w-full shrink flex-col md:order-1 md:max-w-[48%] md:items-end ${ribStackGap}`}
        >
          {upper.map((child) => (
            <FishboneRibColumn
              key={child.id}
              child={child}
              childDepth={nextDepth}
              placement="upper"
              onOpenLeafDetail={onOpenLeafDetail}
              onToggle={onToggle}
              getNodeColor={getNodeColor}
              getNodeIcon={getNodeIcon}
            />
          ))}
        </div>
      )}

      <div className="order-1 flex shrink-0 flex-row items-start justify-center md:order-2 md:pt-0.5">
        <FishboneNodeButton
          node={node}
          hasChildren={hasChildren}
          Icon={Icon}
          className={getNodeColor(node.type)}
          onOpenLeafDetail={onOpenLeafDetail}
          onToggle={() => hasChildren && onToggle(node.id)}
        />
      </div>

      {node.expanded && lower.length > 0 && (
        <div
          className={`order-3 flex w-full min-w-0 max-w-full shrink flex-col md:max-w-[48%] md:items-start ${ribStackGap}`}
        >
          {lower.map((child) => (
            <FishboneRibColumn
              key={child.id}
              child={child}
              childDepth={nextDepth}
              placement="lower"
              onOpenLeafDetail={onOpenLeafDetail}
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
  onOpenLeafDetail?: (node: FishboneNode) => void;
  onToggle: () => void;
}

function FishboneNodeButton({
  node,
  hasChildren,
  Icon,
  className,
  onOpenLeafDetail,
  onToggle,
}: Readonly<FishboneNodeButtonProps>) {
  const opensDetail = Boolean(onOpenLeafDetail) && fishboneNodeOpensDetailView(node);
  const interactive = hasChildren || opensDetail;

  const handleClick = () => {
    if (opensDetail && onOpenLeafDetail) {
      onOpenLeafDetail(node);
      return;
    }
    onToggle();
  };

  return (
    <button
      type="button"
      data-fishbone-node={node.id}
      onClick={handleClick}
      aria-expanded={hasChildren ? node.expanded : undefined}
      className={`flex max-w-[240px] items-center gap-2 rounded-lg border-2 px-3 py-2 text-left text-sm transition-all duration-200 ${className} ${
        interactive ? 'cursor-pointer hover:shadow-md' : 'cursor-default opacity-95'
      }`}
      disabled={!interactive}
    >
      {hasChildren &&
        (node.expanded ? <ChevronDown size={16} className="shrink-0" /> : <ChevronRight size={16} className="shrink-0" />)}
      {Icon && <Icon size={16} className="shrink-0" />}
      <span className="font-medium break-words whitespace-pre-wrap">{node.label}</span>
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

/**
 * Clave estable para agrupar el mismo texto de problema (varios registros → un solo nodo amarillo).
 */
function problemGroupingKey(problem: string): string {
  return normalizeLabel(problem).toLowerCase();
}

/**
 * Hash corto para id de nodo estable sin depender de la longitud del texto del problema.
 */
function shortStringHash(input: string): string {
  let h = 5381;
  for (const ch of input) {
    const cp = ch.codePointAt(0);
    if (cp === undefined) {
      continue;
    }
    h = Math.trunc(Math.imul(h, 33) + cp);
  }
  return (h >>> 0).toString(36);
}

/**
 * Agrupa registros de la misma sección con el mismo problema (texto normalizado), conservando el orden de aparición.
 */
function groupSectionRecordsByProblemText(records: MachineRecord[]): MachineRecord[][] {
  const orderKeys: string[] = [];
  const groups = new Map<string, MachineRecord[]>();
  for (const record of records) {
    const key = problemGroupingKey(record.problem);
    const existing = groups.get(key);
    if (existing) {
      existing.push(record);
    } else {
      groups.set(key, [record]);
      orderKeys.push(key);
    }
  }
  return orderKeys.map((key) => groups.get(key)!);
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
  const problemGroups = groupSectionRecordsByProblemText(sectionRecords);

  return {
    id: `${brand.id}-${model.id}-${section.id}`,
    type: 'seccion',
    label: section.name,
    expanded: false,
    children: problemGroups.map((group) =>
      buildProblemNodeFromRecordGroup(brand.id, model.id, section.id, group, activityTypes, activities)
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

function buildProblemNodeFromRecordGroup(
  brandId: string,
  modelId: string,
  sectionId: string,
  group: MachineRecord[],
  activityTypes: ActivityType[],
  activities: Activity[]
): FishboneNode {
  const representative = group[0];
  const key = problemGroupingKey(representative.problem);
  const idSuffix = shortStringHash(`${brandId}|${modelId}|${sectionId}|${key}`);

  return {
    id: `problem-${brandId}-${modelId}-${sectionId}-${idSuffix}`,
    type: 'problema',
    label: representative.problem,
    expanded: false,
    data: representative,
    children: group.map((record) => buildActivityTypeNode(record, activityTypes, activities)),
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
  const activityLabel = resolveActivityDisplayLabel(record.activityId, activities);

  return {
    id: `activity-${record.id}`,
    type: 'actividad',
    label: activityLabel,
    expanded: false,
    children: buildDetailNodes(record),
  };
}

function listRecordAttachments(record: MachineRecord) {
  if (record.attachments?.length) {
    return record.attachments;
  }
  return record.attachment ? [record.attachment] : [];
}

function buildDetailNodes(record: MachineRecord): FishboneNode[] {
  const allAttachments = listRecordAttachments(record);
  const detailNodes: FishboneNode[] = [
    {
      id: `resource-${record.id}`,
      type: 'recurso',
      label: record.resource,
      expanded: false,
      data: {
        recordId: record.id,
        resourceText: record.resource,
        allAttachments,
      } satisfies FishboneResourceLeafDetail,
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

  for (const att of allAttachments) {
    detailNodes.push({
      id: `attachment-${record.id}-${att.id}`,
      type: 'adjunto',
      label: att.name,
      expanded: false,
      data: {
        recordId: record.id,
        attachment: att,
        allAttachments,
      } satisfies FishboneAttachmentLeafDetail,
      children: [],
    });
  }

  return detailNodes;
}