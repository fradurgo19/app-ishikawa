import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Select } from '../atoms/Select';
import { Card } from '../atoms/Card';
import { sharePointService } from '../services/sharePointService';
import { Section, ActivityType, Activity } from '../types';
import {
  getDistinctTiposEquipo,
  getMarcasForTipoEquipo,
  getModelosForTipoYMarca,
  isValidEquipmentCombination,
} from '../data/equipmentMatrix';
import { ArrowLeft, Save } from 'lucide-react';

interface FormData {
  tipoEquipoId: string;
  brandId: string;
  modelId: string;
  sectionId: string;
  problem: string;
  activityTypeId: string;
  activityId: string;
  resource: string;
  attachment: string;
  time: number;
}

const DEFAULT_CREATED_BY_USER_ID = '1';

export const NewRecordPage: React.FC = () => {
  const navigate = useNavigate();
  const { register, handleSubmit, watch, setValue, formState: { errors } } = useForm<FormData>({
    defaultValues: {
      tipoEquipoId: '',
      brandId: '',
      modelId: '',
      sectionId: '',
      problem: '',
      activityTypeId: '',
      activityId: '',
      resource: '',
      attachment: '',
      time: 0,
    },
  });

  const [sections, setSections] = useState<Section[]>([]);
  const [activityTypes, setActivityTypes] = useState<ActivityType[]>([]);
  const [activities, setActivities] = useState<Activity[]>([]);
  const [loading, setLoading] = useState(false);

  const watchTipoEquipo = watch('tipoEquipoId');
  const watchBrand = watch('brandId');
  const watchModel = watch('modelId');
  const watchActivityType = watch('activityTypeId');

  const tiposOptions = useMemo(
    () => getDistinctTiposEquipo().map((t) => ({ value: t, label: t })),
    []
  );

  const marcasOptions = useMemo(() => {
    if (!watchTipoEquipo) {
      return [];
    }
    return getMarcasForTipoEquipo(watchTipoEquipo).map((m) => ({ value: m, label: m }));
  }, [watchTipoEquipo]);

  const modelosOptions = useMemo(() => {
    if (!watchTipoEquipo || !watchBrand) {
      return [];
    }
    return getModelosForTipoYMarca(watchTipoEquipo, watchBrand).map((m) => ({
      value: m,
      label: m,
    }));
  }, [watchTipoEquipo, watchBrand]);

  const loadInitialData = useCallback(async () => {
    try {
      const activityTypesData = await sharePointService.getActivityTypes();
      setActivityTypes(activityTypesData);
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

  const loadActivities = useCallback(async (activityTypeId: string) => {
    try {
      const activitiesData = await sharePointService.getActivities(activityTypeId);
      setActivities(activitiesData);
    } catch (error) {
      console.error('Error cargando actividades:', error);
    }
  }, []);

  useEffect(() => {
    void loadInitialData();
  }, [loadInitialData]);

  useEffect(() => {
    if (!watchTipoEquipo) {
      return;
    }

    setValue('brandId', '');
    setValue('modelId', '');
    setValue('sectionId', '');
  }, [watchTipoEquipo, setValue]);

  useEffect(() => {
    if (!watchBrand) {
      return;
    }

    setValue('modelId', '');
    setValue('sectionId', '');
  }, [watchBrand, setValue]);

  useEffect(() => {
    if (!watchBrand) {
      setSections([]);
      return;
    }

    void loadSections(watchBrand, watchModel || undefined);
    setValue('sectionId', '');
  }, [watchBrand, watchModel, loadSections, setValue]);

  useEffect(() => {
    if (!watchActivityType) {
      return;
    }

    void loadActivities(watchActivityType);
    setValue('activityId', '');
  }, [watchActivityType, loadActivities, setValue]);

  const onSubmit = useCallback(
    async (data: FormData) => {
      if (
        !isValidEquipmentCombination(data.tipoEquipoId, data.brandId, data.modelId)
      ) {
        alert('La combinación tipo de equipo / marca / modelo no es válida según la matriz corporativa.');
        return;
      }

      setLoading(true);
      try {
        await sharePointService.createRecord({
          ...data,
          createdBy: DEFAULT_CREATED_BY_USER_ID,
        });

        alert('¡Registro creado exitosamente!');
        navigate('/selector');
      } catch (error) {
        console.error('Error creando registro:', error);
        alert('Error creando registro. Por favor intenta de nuevo.');
      } finally {
        setLoading(false);
      }
    },
    [navigate]
  );

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 sm:px-6 lg:px-8 py-8">
        <div className="flex items-center gap-4 mb-8">
          <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/selector')}>
            Volver al Selector
          </Button>

          <div>
            <h1 className="text-3xl font-bold text-gray-900">Crear Nuevo Registro</h1>
            <p className="text-gray-600 mt-2">Tipo de equipo, marca y modelo según matriz Partequipos</p>
          </div>
        </div>

        <Card className="p-8">
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <Select
                label="Tipo de equipo *"
                options={tiposOptions}
                placeholder="Selecciona tipo de equipo"
                {...register('tipoEquipoId', { required: 'El tipo de equipo es requerido' })}
                error={errors.tipoEquipoId?.message}
              />

              <Select
                label="Marca *"
                options={marcasOptions}
                placeholder="Selecciona marca"
                disabled={!watchTipoEquipo}
                {...register('brandId', { required: 'La marca es requerida' })}
                error={errors.brandId?.message}
              />

              <Select
                label="Modelo *"
                options={modelosOptions}
                placeholder="Selecciona modelo"
                disabled={!watchTipoEquipo || !watchBrand}
                {...register('modelId', { required: 'El modelo es requerido' })}
                error={errors.modelId?.message}
              />
            </div>

            <Select
              label="Sección *"
              options={sections.map((s) => ({ value: s.id, label: s.name }))}
              placeholder="Selecciona una sección"
              disabled={!watchBrand}
              {...register('sectionId', { required: 'La sección es requerida' })}
              error={errors.sectionId?.message}
            />

            <Input
              label="Descripción del Problema *"
              type="text"
              placeholder="Describe el problema o incidencia"
              {...register('problem', { required: 'La descripción del problema es requerida' })}
              error={errors.problem?.message}
            />

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <Select
                label="Tipo de Actividad *"
                options={activityTypes.map((at) => ({ value: at.id, label: at.name }))}
                placeholder="Selecciona tipo de actividad"
                {...register('activityTypeId', { required: 'El tipo de actividad es requerido' })}
                error={errors.activityTypeId?.message}
              />

              <Select
                label="Actividad *"
                options={activities.map((a) => ({ value: a.id, label: a.name }))}
                placeholder="Selecciona actividad"
                disabled={!watchActivityType}
                {...register('activityId', { required: 'La actividad es requerida' })}
                error={errors.activityId?.message}
              />
            </div>

            <Input
              label="Recurso"
              type="text"
              placeholder="Manual de referencia, documentación o enlace de recurso"
              {...register('resource')}
            />

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <Input
                label="Adjunto"
                type="text"
                placeholder="Nombre del archivo adjunto o URL"
                {...register('attachment')}
              />

              <Input
                label="Tiempo (minutos) *"
                type="number"
                min="0"
                placeholder="Tiempo empleado en minutos"
                {...register('time', {
                  required: 'El tiempo es requerido',
                  min: { value: 0, message: 'El tiempo debe ser positivo' },
                })}
                error={errors.time?.message}
              />
            </div>

            <div className="flex gap-4 pt-6 border-t">
              <Button type="submit" loading={loading} icon={Save}>
                Crear Registro
              </Button>

              <Button type="button" variant="outline" onClick={() => navigate('/selector')}>
                Cancelar
              </Button>
            </div>
          </form>
        </Card>
      </div>
    </div>
  );
};
