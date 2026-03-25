import React, { useCallback, useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { SelectorCard } from '../molecules/SelectorCard';
import { KPICard } from '../molecules/KPICard';
import { Button } from '../atoms/Button';
import { Select } from '../atoms/Select';
import { Input } from '../atoms/Input';
import { sharePointService } from '../services/sharePointService';
import { Brand, Model, Section, KPIData } from '../types';
import { Factory, Settings, AlertTriangle, BarChart, Plus, Eye } from 'lucide-react';

export const SelectorPage: React.FC = () => {
  const [selectedCard, setSelectedCard] = useState<'brand' | 'model' | 'problem' | null>(null);
  const [brands, setBrands] = useState<Brand[]>([]);
  const [models, setModels] = useState<Model[]>([]);
  const [sections, setSections] = useState<Section[]>([]);
  const [kpis, setKpis] = useState<KPIData | null>(null);
  
  const [selectedBrand, setSelectedBrand] = useState('');
  const [selectedModel, setSelectedModel] = useState('');
  const [problemSearch, setProblemSearch] = useState('');
  
  const navigate = useNavigate();

  const loadInitialData = useCallback(async () => {
    try {
      const [brandsData, kpisData] = await Promise.all([
        sharePointService.getBrands(),
        sharePointService.getKPIs(),
      ]);
      setBrands(brandsData);
      setKpis(kpisData);
    } catch (error) {
      console.error('Error cargando datos iniciales:', error);
    }
  }, []);

  const loadModels = useCallback(async (brandId: string) => {
    try {
      const modelsData = await sharePointService.getModels(brandId);
      setModels(modelsData);
    } catch (error) {
      console.error('Error cargando modelos:', error);
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
      void loadModels(selectedBrand);
      void loadSections(selectedBrand);
    }
  }, [selectedBrand, loadModels, loadSections]);

  useEffect(() => {
    if (selectedBrand && selectedModel) {
      void loadSections(selectedBrand, selectedModel);
    }
  }, [selectedBrand, selectedModel, loadSections]);

  const handleViewFishbone = () => {
    const params = new URLSearchParams();
    if (selectedBrand) params.set('brand', selectedBrand);
    if (selectedModel) params.set('model', selectedModel);
    if (problemSearch) params.set('problem', problemSearch);
    
    navigate(`/fishbone?${params.toString()}`);
  };

  const canViewFishbone = Boolean(selectedBrand || selectedModel || problemSearch);

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-7xl mx-auto px-4 py-8">
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-gray-900">Selector de Datos de Maquinaria</h1>
          <p className="text-gray-600 mt-2">Selecciona criterios para analizar con el diagrama Ishikawa</p>
        </div>

        {/* Tarjetas KPI */}
        {kpis && (
          <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
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

        {/* Tarjetas de Selección */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <SelectorCard
            title="Marca"
            description="Selecciona la marca de maquinaria"
            icon={Factory}
            onClick={() => setSelectedCard('brand')}
            selected={selectedCard === 'brand'}
          />
          <SelectorCard
            title="Modelo"
            description="Elige un modelo específico"
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

        {/* Formularios de Selección */}
        <div className="bg-white rounded-lg p-6 shadow-md mb-8">
          <h2 className="text-xl font-semibold text-gray-900 mb-6">Criterios de Filtro</h2>
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <Select
              label="Marca"
              options={brands.map(b => ({ value: b.id, label: b.name }))}
              value={selectedBrand}
              onChange={(e) => {
                setSelectedBrand(e.target.value);
                setSelectedModel(''); // Reiniciar modelo cuando cambia la marca
              }}
              placeholder="Selecciona una marca"
            />
            
            <Select
              label="Modelo"
              options={models.map(m => ({ value: m.id, label: m.name }))}
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              placeholder="Selecciona un modelo"
              disabled={!selectedBrand}
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
                {sections.map(section => (
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
            <Button
              onClick={handleViewFishbone}
              disabled={!canViewFishbone}
              icon={Eye}
            >
              Ver Diagrama Ishikawa
            </Button>
            
            <Button
              variant="outline"
              onClick={() => navigate('/new-record')}
              icon={Plus}
            >
              Crear Nuevo Registro
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};