import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { SelectorCard } from '../molecules/SelectorCard';
import { KPICard } from '../molecules/KPICard';
import { Button } from '../atoms/Button';
import { Select } from '../atoms/Select';
import { Input } from '../atoms/Input';
import { sharePointService } from '../services/sharePointService';
import { Section, KPIData } from '../types';
import {
  getDistinctTiposEquipo,
  getMarcasForTipoEquipo,
  getModelosForTipoYMarca,
  isValidEquipmentCombination,
} from '../data/equipmentMatrix';
import { Factory, Settings, AlertTriangle, BarChart, Plus, Eye, Truck } from 'lucide-react';

export const SelectorPage: React.FC = () => {
  const [selectedCard, setSelectedCard] = useState<'brand' | 'model' | 'problem' | null>(null);
  const [sections, setSections] = useState<Section[]>([]);
  const [kpis, setKpis] = useState<KPIData | null>(null);

  const [selectedTipoEquipo, setSelectedTipoEquipo] = useState('');
  const [selectedBrand, setSelectedBrand] = useState('');
  const [selectedModel, setSelectedModel] = useState('');
  const [problemSearch, setProblemSearch] = useState('');

  const navigate = useNavigate();

  const tiposOptions = useMemo(
    () => getDistinctTiposEquipo().map((t) => ({ value: t, label: t })),
    []
  );

  const marcasOptions = useMemo(() => {
    if (!selectedTipoEquipo) {
      return [];
    }
    return getMarcasForTipoEquipo(selectedTipoEquipo).map((m) => ({ value: m, label: m }));
  }, [selectedTipoEquipo]);

  const modelosOptions = useMemo(() => {
    if (!selectedTipoEquipo || !selectedBrand) {
      return [];
    }
    return getModelosForTipoYMarca(selectedTipoEquipo, selectedBrand).map((m) => ({
      value: m,
      label: m,
    }));
  }, [selectedTipoEquipo, selectedBrand]);

  const loadInitialData = useCallback(async () => {
    try {
      const kpisData = await sharePointService.getKPIs();
      setKpis(kpisData);
    } catch (error) {
      console.error('Error cargando datos iniciales:', error);
    }
  }, []);

  const loadSections = useCallback(async (brandId: string, modelId?: string) => {
    try {
      const sectionsData = await sharePointService.getSections(brandId, modelId);
      setSections(sectionsData);
    } catch (error) {
      console.error('Error cargando secciones:', error);
    }
  }, []);

  useEffect(() => {
    void loadInitialData();
  }, [loadInitialData]);

  useEffect(() => {
    if (selectedBrand) {
      void loadSections(selectedBrand, selectedModel || undefined);
    } else {
      setSections([]);
    }
  }, [selectedBrand, selectedModel, loadSections]);

  const handleViewFishbone = () => {
    const params = new URLSearchParams();
    if (selectedTipoEquipo) params.set('tipoEquipo', selectedTipoEquipo);
    if (selectedBrand) params.set('brand', selectedBrand);
    if (selectedModel) params.set('model', selectedModel);
    if (problemSearch) params.set('problem', problemSearch);

    navigate(`/fishbone?${params.toString()}`);
  };

  const matrixComboValid =
    Boolean(selectedTipoEquipo && selectedBrand && selectedModel) &&
    isValidEquipmentCombination(selectedTipoEquipo, selectedBrand, selectedModel);

  const canViewFishbone = Boolean(problemSearch || matrixComboValid);

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 sm:px-6 lg:px-8 py-8">
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-gray-900">Selector de Datos de Maquinaria</h1>
          <p className="text-gray-600 mt-2">Selecciona tipo de equipo, marca y modelo según la matriz corporativa</p>
        </div>

        {kpis && (
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4 md:gap-6 mb-8">
            <KPICard
              title="Tipos de equipo"
              value={kpis.totalTiposEquipo}
              icon={Truck}
              color="primary"
            />
            <KPICard
              title="Total Marcas"
              value={kpis.totalMarcas}
              icon={Factory}
              color="primary"
            />
            <KPICard
              title="Total Modelos"
              value={kpis.totalModelos}
              icon={Settings}
              color="secondary"
            />
            <KPICard
              title="Total Secciones"
              value={kpis.totalSecciones}
              icon={BarChart}
              color="success"
            />
            <KPICard
              title="Total Registros"
              value={kpis.totalRegistros}
              icon={AlertTriangle}
              color="warning"
            />
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <SelectorCard
            title="Marca"
            description="Según tipo de equipo seleccionado"
            icon={Factory}
            onClick={() => setSelectedCard('brand')}
            selected={selectedCard === 'brand'}
          />
          <SelectorCard
            title="Modelo"
            description="Combinación válida marca / modelo"
            icon={Settings}
            onClick={() => setSelectedCard('model')}
            selected={selectedCard === 'model'}
          />
          <SelectorCard
            title="Problema"
            description="Busca por descripción del problema"
            icon={AlertTriangle}
            onClick={() => setSelectedCard('problem')}
            selected={selectedCard === 'problem'}
          />
        </div>

        <div className="bg-white rounded-lg p-6 shadow-md mb-8">
          <h2 className="text-xl font-semibold text-gray-900 mb-6">Criterios de Filtro</h2>

          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
            <Select
              label="Tipo de equipo"
              options={tiposOptions}
              value={selectedTipoEquipo}
              onChange={(e) => {
                setSelectedTipoEquipo(e.target.value);
                setSelectedBrand('');
                setSelectedModel('');
              }}
              placeholder="Selecciona tipo de equipo"
            />

            <Select
              label="Marca"
              options={marcasOptions}
              value={selectedBrand}
              onChange={(e) => {
                setSelectedBrand(e.target.value);
                setSelectedModel('');
              }}
              placeholder="Selecciona marca"
              disabled={!selectedTipoEquipo}
            />

            <Select
              label="Modelo"
              options={modelosOptions}
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              placeholder="Selecciona modelo"
              disabled={!selectedTipoEquipo || !selectedBrand}
            />

            <Input
              label="Búsqueda de Problemas"
              type="text"
              value={problemSearch}
              onChange={(e) => setProblemSearch(e.target.value)}
              placeholder="Buscar problemas..."
            />
          </div>

          {sections.length > 0 && (
            <div className="mt-6">
              <h3 className="text-sm font-medium text-gray-700 mb-2">
                Secciones Disponibles ({sections.length})
              </h3>
              <div className="flex flex-wrap gap-2">
                {sections.map((section) => (
                  <span
                    key={section.id}
                    className="px-3 py-1 bg-gray-100 text-gray-700 rounded-full text-sm"
                  >
                    {section.name}
                  </span>
                ))}
              </div>
            </div>
          )}

          <div className="flex gap-4 mt-6 pt-6 border-t">
            <Button onClick={handleViewFishbone} disabled={!canViewFishbone} icon={Eye}>
              Ver Diagrama Ishikawa
            </Button>

            <Button variant="outline" onClick={() => navigate('/new-record')} icon={Plus}>
              Crear Nuevo Registro
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};
