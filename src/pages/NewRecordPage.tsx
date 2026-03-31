import React, { useCallback, useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Select } from '../atoms/Select';
import { Card } from '../atoms/Card';
import { sharePointService } from '../services/sharePointService';
import { Section, ActivityType, Activity } from '../types';
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
  const { register, handleSubmit, formState: { errors } } = useForm<FormData>({
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
  const [tiposOptions, setTiposOptions] = useState<{ value: string; label: string }[]>([]);
  const [marcasOptions, setMarcasOptions] = useState<{ value: string; label: string }[]>([]);
  const [modelosOptions, setModelosOptions] = useState<{ value: string; label: string }[]>([]);
  const [loading, setLoading] = useState(false);
  /** Evita dejar los selects en disabled cuando la API devuelve listas vacías o falla (antes: disabled si length===0). */
  const [selectOptionsLoading, setSelectOptionsLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const toSelectOptions = (labels: string[]) =>
    labels.map((t) => ({ value: t, label: t }));

  const loadInitialData = useCallback(async () => {
    setSelectOptionsLoading(true);
    setLoadError(null);
    try {
      await sharePointService.refreshDictionary?.();
      const [equipment, activityTypesData, sectionsData, activitiesData] = await Promise.all([
        sharePointService.getNewRecordEquipmentSelectOptions(),
        sharePointService.getActivityTypes(),
        sharePointService.getSectionOptionsForNewRecord(),
        sharePointService.getActivityOptionsForNewRecord(),
      ]);
      setTiposOptions(toSelectOptions(equipment.tipos));
      setMarcasOptions(toSelectOptions(equipment.marcas));
      setModelosOptions(toSelectOptions(equipment.modelos));
      setActivityTypes(activityTypesData);
      setSections(sectionsData);
      setActivities(activitiesData);
    } catch (error) {
      console.error('Error cargando datos iniciales:', error);
      const message =
        error instanceof Error ? error.message : 'No se pudo cargar el diccionario desde /api/ishikawa';
      setLoadError(message);
    } finally {
      setSelectOptionsLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadInitialData();
  }, [loadInitialData]);

  const onSubmit = useCallback(
    async (data: FormData) => {
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
            <p className="text-gray-600 mt-2">
              Los desplegables cargan las opciones de las columnas tipo selección de la lista; el resto son texto
              libre. Las selecciones no se filtran entre sí.
            </p>
          </div>
        </div>

        <Card className="p-8">
          {loadError && (
            <div
              className="mb-6 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900"
              role="alert"
            >
              <p className="font-medium">No se pudieron cargar las opciones desde SharePoint (API)</p>
              <p className="mt-1 text-amber-800">{loadError}</p>
              <p className="mt-2 text-amber-800">
                En local: ejecuta <code className="rounded bg-amber-100 px-1">npx vercel dev</code> (con variables
                SHAREPOINT_* en .env) y deja Vite con proxy a ese puerto, o define{' '}
                <code className="rounded bg-amber-100 px-1">VITE_API_BASE_URL</code> apuntando al origen donde corre
                /api.
              </p>
              <Button type="button" variant="outline" className="mt-3" onClick={() => void loadInitialData()}>
                Reintentar
              </Button>
            </div>
          )}
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <Select
                label="Tipo de equipo *"
                options={tiposOptions}
                placeholder="Selecciona tipo de equipo"
                disabled={selectOptionsLoading}
                {...register('tipoEquipoId', { required: 'El tipo de equipo es requerido' })}
                error={errors.tipoEquipoId?.message}
              />

              <Select
                label="Marca *"
                options={marcasOptions}
                placeholder="Selecciona marca"
                disabled={selectOptionsLoading}
                {...register('brandId', { required: 'La marca es requerida' })}
                error={errors.brandId?.message}
              />

              <Select
                label="Modelo *"
                options={modelosOptions}
                placeholder="Selecciona modelo"
                disabled={selectOptionsLoading}
                {...register('modelId', { required: 'El modelo es requerido' })}
                error={errors.modelId?.message}
              />
            </div>

            <Select
              label="Sección *"
              options={sections.map((s) => ({ value: s.id, label: s.name }))}
              placeholder="Selecciona una sección"
              disabled={selectOptionsLoading}
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
                key={`activity-types-${activityTypes.map((t) => t.id).join('|')}`}
                label="Tipo de Actividad *"
                options={activityTypes.map((at) => ({ value: at.id, label: at.name }))}
                placeholder="Selecciona tipo de actividad"
                disabled={selectOptionsLoading}
                {...register('activityTypeId', { required: 'El tipo de actividad es requerido' })}
                error={errors.activityTypeId?.message}
              />

              <Select
                label="Actividad *"
                options={activities.map((a) => ({ value: a.id, label: a.name }))}
                placeholder="Selecciona actividad"
                disabled={selectOptionsLoading}
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
