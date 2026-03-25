import React, { useCallback, useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Select } from '../atoms/Select';
import { Card } from '../atoms/Card';
import { sharePointService } from '../services/sharePointService';
import { Brand, Model, Section, ActivityType, Activity } from '../types';
import { ArrowLeft, Save } from 'lucide-react';

interface FormData {
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
  const { register, handleSubmit, watch, setValue, formState: { errors } } = useForm<FormData>();
  
  const [brands, setBrands] = useState<Brand[]>([]);
  const [models, setModels] = useState<Model[]>([]);
  const [sections, setSections] = useState<Section[]>([]);
  const [activityTypes, setActivityTypes] = useState<ActivityType[]>([]);
  const [activities, setActivities] = useState<Activity[]>([]);
  const [loading, setLoading] = useState(false);

  const watchBrand = watch('brandId');
  const watchModel = watch('modelId');
  const watchActivityType = watch('activityTypeId');

  const loadInitialData = useCallback(async () => {
    try {
      const [brandsData, activityTypesData] = await Promise.all([
        sharePointService.getBrands(),
        sharePointService.getActivityTypes(),
      ]);
      setBrands(brandsData);
      setActivityTypes(activityTypesData);
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
    if (!watchBrand) {
      return;
    }

    void loadModels(watchBrand);
    void loadSections(watchBrand);
    setValue('modelId', '');
    setValue('sectionId', '');
  }, [watchBrand, loadModels, loadSections, setValue]);

  useEffect(() => {
    if (!watchBrand || !watchModel) {
      return;
    }

    void loadSections(watchBrand, watchModel);
    setValue('sectionId', '');
  }, [watchBrand, watchModel, loadSections, setValue]);

  useEffect(() => {
    if (!watchActivityType) {
      return;
    }

    void loadActivities(watchActivityType);
    setValue('activityId', '');
  }, [watchActivityType, loadActivities, setValue]);

  const onSubmit = useCallback(async (data: FormData) => {
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
  }, [navigate]);

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 sm:px-6 lg:px-8 py-8">
        <div className="flex items-center gap-4 mb-8">
          <Button
            variant="ghost"
            icon={ArrowLeft}
            onClick={() => navigate('/selector')}
          >
            Volver al Selector
          </Button>
          
          <div>
            <h1 className="text-3xl font-bold text-gray-900">Crear Nuevo Registro</h1>
            <p className="text-gray-600 mt-2">Agregar un nuevo registro de mantenimiento de maquinaria</p>
          </div>
        </div>

        <Card className="p-8">
          <form onSubmit={handleSubmit(onSubmit)} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <Select
                label="Marca *"
                options={brands.map(b => ({ value: b.id, label: b.name }))}
                placeholder="Selecciona una marca"
                {...register('brandId', { required: 'La marca es requerida' })}
                error={errors.brandId?.message}
              />
              
              <Select
                label="Modelo *"
                options={models.map(m => ({ value: m.id, label: m.name }))}
                placeholder="Selecciona un modelo"
                disabled={!watchBrand}
                {...register('modelId', { required: 'El modelo es requerido' })}
                error={errors.modelId?.message}
              />
            </div>

            <Select
              label="Sección *"
              options={sections.map(s => ({ value: s.id, label: s.name }))}
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
                options={activityTypes.map(at => ({ value: at.id, label: at.name }))}
                placeholder="Selecciona tipo de actividad"
                {...register('activityTypeId', { required: 'El tipo de actividad es requerido' })}
                error={errors.activityTypeId?.message}
              />
              
              <Select
                label="Actividad *"
                options={activities.map(a => ({ value: a.id, label: a.name }))}
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
                  min: { value: 0, message: 'El tiempo debe ser positivo' }
                })}
                error={errors.time?.message}
              />
            </div>

            <div className="flex gap-4 pt-6 border-t">
              <Button
                type="submit"
                loading={loading}
                icon={Save}
              >
                Crear Registro
              </Button>
              
              <Button
                type="button"
                variant="outline"
                onClick={() => navigate('/selector')}
              >
                Cancelar
              </Button>
            </div>
          </form>
        </Card>
      </div>
    </div>
  );
};