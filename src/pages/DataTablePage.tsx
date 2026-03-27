import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Select } from '../atoms/Select';
import { sharePointService } from '../services/sharePointService';
import { MachineRecord, Brand, Model, ActivityType, Activity } from '../types';
import {
  getDistinctTiposEquipo,
  getMarcasForTipoEquipo,
  getModelosForTipoYMarca,
} from '../data/equipmentMatrix';
import { ArrowLeft, Download, Filter } from 'lucide-react';

interface DataTableFilters {
  tipoEquipoId: string;
  brandId: string;
  modelId: string;
  problem: string;
  activityTypeId: string;
  activityId: string;
  resource: string;
}

const INITIAL_FILTERS: DataTableFilters = {
  tipoEquipoId: '',
  brandId: '',
  modelId: '',
  problem: '',
  activityTypeId: '',
  activityId: '',
  resource: '',
};

export const DataTablePage: React.FC = () => {
  const navigate = useNavigate();
  const [records, setRecords] = useState<MachineRecord[]>([]);
  const [filteredRecords, setFilteredRecords] = useState<MachineRecord[]>([]);
  const [brands, setBrands] = useState<Brand[]>([]);
  const [models, setModels] = useState<Model[]>([]);
  const [activityTypes, setActivityTypes] = useState<ActivityType[]>([]);
  const [activities, setActivities] = useState<Activity[]>([]);
  const [loading, setLoading] = useState(true);

  const [filters, setFilters] = useState<DataTableFilters>(INITIAL_FILTERS);

  const tiposFilterOptions = useMemo(
    () => getDistinctTiposEquipo().map((t) => ({ value: t, label: t })),
    []
  );

  const marcasFilterOptions = useMemo(() => {
    if (!filters.tipoEquipoId) {
      return [];
    }
    return getMarcasForTipoEquipo(filters.tipoEquipoId).map((m) => ({ value: m, label: m }));
  }, [filters.tipoEquipoId]);

  const modelosFilterOptions = useMemo(() => {
    if (!filters.tipoEquipoId || !filters.brandId) {
      return [];
    }
    return getModelosForTipoYMarca(filters.tipoEquipoId, filters.brandId).map((m) => ({
      value: m,
      label: m,
    }));
  }, [filters.tipoEquipoId, filters.brandId]);

  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      await sharePointService.refreshDictionary?.();
      const [recordsData, brandsData, modelsData, activityTypesData, activitiesData] = await Promise.all([
        sharePointService.getRecords(),
        sharePointService.getBrands(),
        sharePointService.getModels(),
        sharePointService.getActivityTypes(),
        sharePointService.getActivities(),
      ]);

      setRecords(recordsData);
      setBrands(brandsData);
      setModels(modelsData);
      setActivityTypes(activityTypesData);
      setActivities(activitiesData);
    } catch (error) {
      console.error('Error cargando datos:', error);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadData();
  }, [loadData]);

  useEffect(() => {
    setFilteredRecords(applyRecordFilters(records, filters));
  }, [records, filters]);

  const getBrandName = (brandId: string) => {
    return brands.find((b) => b.id === brandId)?.name || brandId || 'Desconocido';
  };

  const getModelName = (modelId: string) => {
    return models.find((m) => m.id === modelId)?.name || modelId || 'Desconocido';
  };

  const getActivityTypeName = (activityTypeId: string) => {
    return activityTypes.find((at) => at.id === activityTypeId)?.name || 'Desconocido';
  };

  const getActivityName = (activityId: string) => {
    return activities.find((a) => a.id === activityId)?.name || 'Desconocido';
  };

  const handleFilterChange = (key: keyof DataTableFilters, value: string) => {
    setFilters((prev) => {
      const next = { ...prev, [key]: value };
      if (key === 'tipoEquipoId') {
        next.brandId = '';
        next.modelId = '';
      }
      if (key === 'brandId') {
        next.modelId = '';
      }
      if (key === 'activityTypeId') {
        next.activityId = '';
      }
      return next;
    });
  };

  const activityOptionsForFilter = useMemo(() => {
    if (filters.activityTypeId) {
      return activities.filter((a) => a.activityTypeId === filters.activityTypeId);
    }
    return activities;
  }, [activities, filters.activityTypeId]);

  const clearFilters = () => {
    setFilters({ ...INITIAL_FILTERS });
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="animate-spin rounded-full h-8 w-8 border-2 border-red-600 border-t-transparent"></div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 sm:px-6 lg:px-8 py-8">
        <div className="flex items-center gap-4 mb-8">
          <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/selector')}>
            Volver al Selector
          </Button>

          <div>
            <h1 className="text-3xl font-bold text-gray-900">Registros de Datos</h1>
            <p className="text-gray-600 mt-2">Ver y filtrar todos los registros de mantenimiento</p>
          </div>
        </div>

        <div className="bg-white rounded-lg p-6 shadow-md mb-8">
          <div className="flex items-center gap-4 mb-4">
            <Filter className="text-gray-500" size={20} />
            <h2 className="text-lg font-semibold text-gray-900">Filtros</h2>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4 mb-4">
            <Select
              label="Tipo de equipo"
              options={tiposFilterOptions}
              value={filters.tipoEquipoId}
              onChange={(e) => handleFilterChange('tipoEquipoId', e.target.value)}
              placeholder="Todos los tipos"
            />

            <Select
              label="Marca"
              options={marcasFilterOptions}
              value={filters.brandId}
              onChange={(e) => handleFilterChange('brandId', e.target.value)}
              placeholder="Todas las marcas"
              disabled={!filters.tipoEquipoId}
            />

            <Select
              label="Modelo"
              options={modelosFilterOptions}
              value={filters.modelId}
              onChange={(e) => handleFilterChange('modelId', e.target.value)}
              placeholder="Todos los modelos"
              disabled={!filters.tipoEquipoId || !filters.brandId}
            />

            <Select
              label="Tipo de Actividad"
              options={activityTypes.map((at) => ({ value: at.id, label: at.name }))}
              value={filters.activityTypeId}
              onChange={(e) => handleFilterChange('activityTypeId', e.target.value)}
              placeholder="Todos los tipos de actividad"
            />

            <Select
              label="Actividad"
              options={activityOptionsForFilter.map((a) => ({ value: a.id, label: a.name }))}
              value={filters.activityId}
              onChange={(e) => handleFilterChange('activityId', e.target.value)}
              placeholder="Todas las actividades"
            />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <Input
              label="Búsqueda de Problemas"
              type="text"
              value={filters.problem}
              onChange={(e) => handleFilterChange('problem', e.target.value)}
              placeholder="Buscar problemas..."
            />
            <Input
              label="Recurso"
              type="text"
              value={filters.resource}
              onChange={(e) => handleFilterChange('resource', e.target.value)}
              placeholder="Filtrar por recurso..."
            />
          </div>

          <div className="flex gap-4">
            <Button variant="outline" onClick={clearFilters}>
              Limpiar Filtros
            </Button>
            <span className="text-sm text-gray-600 flex items-center">
              Mostrando {filteredRecords.length} de {records.length} registros
            </span>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow-md overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Tipo equipo
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Marca / Modelo
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Problema
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Actividad
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Recurso
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Tiempo
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Acciones
                  </th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {filteredRecords.map((record) => (
                  <tr key={record.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className="text-sm text-gray-900">{record.tipoEquipoId || '—'}</span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="text-sm font-medium text-gray-900">
                        {getBrandName(record.brandId)}
                      </div>
                      <div className="text-sm text-gray-500">{getModelName(record.modelId)}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm text-gray-900 max-w-xs truncate">{record.problem}</div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="text-sm text-gray-900">
                        {getActivityTypeName(record.activityTypeId)}
                      </div>
                      <div className="text-sm text-gray-500">
                        {getActivityName(record.activityId)}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm text-gray-900 max-w-xs truncate">{record.resource}</div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className="px-2 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full">
                        {record.time}m
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      {record.attachment && (
                        <Button
                          size="sm"
                          variant="ghost"
                          icon={Download}
                          onClick={() => alert('La funcionalidad de descarga se implementaría aquí')}
                        >
                          Descargar
                        </Button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {filteredRecords.length === 0 && (
            <div className="text-center py-12">
              <p className="text-gray-500">No se encontraron registros que coincidan con los filtros actuales</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const EXACT_FILTER_KEYS = new Set<keyof DataTableFilters>([
  'brandId',
  'modelId',
  'activityTypeId',
  'activityId',
]);

function applyRecordFilters(records: MachineRecord[], filters: DataTableFilters): MachineRecord[] {
  let result = records;

  Object.entries(filters).forEach(([key, value]) => {
    if (!value) {
      return;
    }

    const filterKey = key as keyof DataTableFilters;

    if (filterKey === 'tipoEquipoId') {
      result = result.filter((record) => record.tipoEquipoId.toLowerCase() === value.toLowerCase());
      return;
    }

    if (EXACT_FILTER_KEYS.has(filterKey)) {
      const recordKey = filterKey as keyof MachineRecord;
      result = result.filter((record) => {
        const recordValue = record[recordKey];
        return typeof recordValue === 'string' && recordValue.toLowerCase() === value.toLowerCase();
      });
      return;
    }

    if (filterKey === 'problem' || filterKey === 'resource') {
      const recordKey = filterKey as keyof MachineRecord;
      result = result.filter((record) => {
        const recordValue = record[recordKey];
        return (
          typeof recordValue === 'string' &&
          recordValue.toLowerCase().includes(value.toLowerCase())
        );
      });
    }
  });

  return result;
}
