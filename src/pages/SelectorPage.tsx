import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
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
} from '../data/equipmentMatrix';
import { Factory, Settings, AlertTriangle, BarChart, Plus, BarChart3, Truck } from 'lucide-react';

export const SelectorPage: React.FC = () => {
  const [selectedCard, setSelectedCard] = useState<'brand' | 'model' | 'problem' | null>(null);
  const [sections, setSections] = useState<Section[]>([]);
  const [kpis, setKpis] = useState<KPIData | null>(null);

  const [selectedTipoEquipo, setSelectedTipoEquipo] = useState('');
  const [selectedBrand, setSelectedBrand] = useState('');
  const [selectedModel, setSelectedModel] = useState('');
  const [problemSearch, setProblemSearch] = useState('');

  const criteriosRef = useRef<HTMLElement>(null);
  const tipoSelectRef = useRef<HTMLSelectElement>(null);
  const brandSelectRef = useRef<HTMLSelectElement>(null);
  const modelSelectRef = useRef<HTMLSelectElement>(null);
  const problemInputRef = useRef<HTMLInputElement>(null);

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

  const scrollCriteriosAndRun = useCallback((action: () => void) => {
    queueMicrotask(() => {
      criteriosRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      action();
    });
  }, []);

  const handleBrandCardClick = useCallback(() => {
    setSelectedCard('brand');
    scrollCriteriosAndRun(() => {
      if (selectedTipoEquipo) {
        brandSelectRef.current?.focus();
      } else {
        tipoSelectRef.current?.focus();
      }
    });
  }, [scrollCriteriosAndRun, selectedTipoEquipo]);

  const handleModelCardClick = useCallback(() => {
    setSelectedCard('model');
    scrollCriteriosAndRun(() => {
      if (selectedTipoEquipo && selectedBrand) {
        modelSelectRef.current?.focus();
      } else if (selectedTipoEquipo) {
        brandSelectRef.current?.focus();
      } else {
        tipoSelectRef.current?.focus();
      }
    });
  }, [scrollCriteriosAndRun, selectedBrand, selectedTipoEquipo]);

  const handleProblemCardClick = useCallback(() => {
    setSelectedCard('problem');
    scrollCriteriosAndRun(() => {
      problemInputRef.current?.focus();
    });
  }, [scrollCriteriosAndRun]);

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
            onClick={handleBrandCardClick}
            selected={selectedCard === 'brand'}
          />
          <SelectorCard
            title="Modelo"
            description="Combinación válida marca / modelo"
            icon={Settings}
            onClick={handleModelCardClick}
            selected={selectedCard === 'model'}
          />
          <SelectorCard
            title="Problema"
            description="Busca por descripción del problema"
            icon={AlertTriangle}
            onClick={handleProblemCardClick}
            selected={selectedCard === 'problem'}
          />
        </div>

        <section ref={criteriosRef} className="bg-white rounded-lg p-6 shadow-md mb-8" aria-label="Criterios de filtro">
          <h2 className="text-xl font-semibold text-gray-900 mb-6">Criterios de Filtro</h2>

          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
            <Select
              ref={tipoSelectRef}
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
              ref={brandSelectRef}
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
              ref={modelSelectRef}
              label="Modelo"
              options={modelosOptions}
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              placeholder="Selecciona modelo"
              disabled={!selectedTipoEquipo || !selectedBrand}
            />

            <Input
              ref={problemInputRef}
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

          <div className="mt-6 space-y-3 border-t pt-6">
            <div className="flex flex-wrap gap-4">
              <Button onClick={() => navigate('/data-table')} icon={BarChart3}>
                Ir a tabla de datos
              </Button>
              <Button variant="outline" onClick={() => navigate('/new-record')} icon={Plus}>
                Crear Nuevo Registro
              </Button>
            </div>
            <p className="text-xs text-gray-500 max-w-xl">
              El diagrama Ishikawa solo se abre desde la tabla de datos (enlace «Diagrama» en cada registro).
            </p>
          </div>
        </section>
      </div>
    </div>
  );
};
