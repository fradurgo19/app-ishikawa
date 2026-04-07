import React, { useCallback, useEffect, useState } from 'react';
import { useForm } from 'react-hook-form';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Textarea } from '../atoms/Textarea';
import { Select } from '../atoms/Select';
import { sharePointService } from '../services/sharePointService';
import { MachineRecord, Section, ActivityType } from '../types';
import { X } from 'lucide-react';

interface EditFormData {
  tipoEquipoId: string;
  brandId: string;
  modelId: string;
  sectionId: string;
  problem: string;
  activityTypeId: string;
  activityId: string;
  resource: string;
  time: number;
}

export interface EditRecordModalProps {
  isOpen: boolean;
  record: MachineRecord | null;
  onClose: () => void;
  onSaved: (record: MachineRecord) => void;
}

const toSelectOptions = (labels: string[]) => labels.map((t) => ({ value: t, label: t }));

export const EditRecordModal: React.FC<EditRecordModalProps> = ({
  isOpen,
  record,
  onClose,
  onSaved,
}) => {
  const { register, handleSubmit, reset, watch, formState: { errors } } = useForm<EditFormData>({
    defaultValues: {
      tipoEquipoId: '',
      brandId: '',
      modelId: '',
      sectionId: '',
      problem: '',
      activityTypeId: '',
      activityId: '',
      resource: '',
      time: 0,
    },
  });

  const [sections, setSections] = useState<Section[]>([]);
  const [activityTypes, setActivityTypes] = useState<ActivityType[]>([]);
  const [tiposOptions, setTiposOptions] = useState<{ value: string; label: string }[]>([]);
  const [marcasOptions, setMarcasOptions] = useState<{ value: string; label: string }[]>([]);
  const [modelosOptions, setModelosOptions] = useState<{ value: string; label: string }[]>([]);
  const [optionsLoading, setOptionsLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);

  const brandWatch = watch('brandId');
  const modelWatch = watch('modelId');

  const loadOptions = useCallback(async () => {
    setOptionsLoading(true);
    setLoadError(null);
    try {
      await sharePointService.refreshDictionary?.();
      const [equipment, activityTypesData] = await Promise.all([
        sharePointService.getNewRecordEquipmentSelectOptions(),
        sharePointService.getActivityTypes(),
      ]);
      setTiposOptions(toSelectOptions(equipment.tipos));
      setMarcasOptions(toSelectOptions(equipment.marcas));
      setModelosOptions(toSelectOptions(equipment.modelos));
      setActivityTypes(activityTypesData);
    } catch (e) {
      const message = e instanceof Error ? e.message : 'Error al cargar opciones';
      setLoadError(message);
    } finally {
      setOptionsLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!isOpen || !record) {
      return;
    }
    void loadOptions();
    reset({
      tipoEquipoId: record.tipoEquipoId ?? '',
      brandId: record.brandId ?? '',
      modelId: record.modelId ?? '',
      sectionId: record.sectionId ?? '',
      problem: record.problem ?? '',
      activityTypeId: record.activityTypeId ?? '',
      activityId: record.activityId ?? '',
      resource: record.resource ?? '',
      time: Number.isFinite(record.time) ? record.time : 0,
    });
  }, [isOpen, record, loadOptions, reset]);

  useEffect(() => {
    if (!isOpen || !brandWatch?.trim() || !modelWatch?.trim()) {
      setSections([]);
      return;
    }
    let cancelled = false;
    void sharePointService.getSections(brandWatch, modelWatch).then((data) => {
      if (!cancelled) {
        setSections(data);
      }
    });
    return () => {
      cancelled = true;
    };
  }, [isOpen, brandWatch, modelWatch]);

  const onSubmit = useCallback(
    async (data: EditFormData) => {
      if (!record) {
        return;
      }
      setSaving(true);
      try {
        const updated = await sharePointService.updateRecord({
          ...record,
          tipoEquipoId: data.tipoEquipoId,
          brandId: data.brandId,
          modelId: data.modelId,
          sectionId: data.sectionId,
          problem: data.problem,
          activityTypeId: data.activityTypeId,
          activityId: data.activityId,
          resource: data.resource,
          time: Number(data.time),
        });
        onSaved(updated);
        onClose();
      } catch (e) {
        console.error('Error actualizando registro:', e);
        alert(e instanceof Error ? e.message : 'No se pudo guardar el registro.');
      } finally {
        setSaving(false);
      }
    },
    [record, onSaved, onClose]
  );

  if (!isOpen || !record) {
    return null;
  }

  return (
    <dialog
      open
      className="fixed inset-0 z-50 m-0 flex h-auto max-h-none w-full max-w-none items-center justify-center border-0 bg-black/50 p-4 backdrop:bg-black/50"
      aria-labelledby="edit-record-modal-title"
      onCancel={(e) => {
        e.preventDefault();
        onClose();
      }}
    >
      <div className="relative w-full max-w-3xl max-h-[90vh] overflow-y-auto rounded-lg bg-white shadow-xl">
        <div className="sticky top-0 z-10 flex items-center justify-between border-b border-gray-200 bg-white px-6 py-4">
          <h2 id="edit-record-modal-title" className="text-lg font-semibold text-gray-900">
            Editar registro
          </h2>
          <button
            type="button"
            onClick={onClose}
            className="rounded p-2 text-gray-500 hover:bg-gray-100"
            aria-label="Cerrar"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        <form onSubmit={handleSubmit(onSubmit)} className="space-y-6 p-6">
          {loadError && (
            <div className="rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900" role="alert">
              {loadError}
            </div>
          )}

          <p className="text-xs text-gray-500">
            Id. ítem: <span className="font-mono">{record.id}</span>
            {record.createdBy ? (
              <>
                {' '}
                · Creado por: {record.createdBy}
              </>
            ) : null}
          </p>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <Select
              label="Tipo de equipo *"
              options={tiposOptions}
              placeholder="Selecciona tipo"
              disabled={optionsLoading}
              {...register('tipoEquipoId', { required: 'Requerido' })}
              error={errors.tipoEquipoId?.message}
            />
            <Select
              label="Marca *"
              options={marcasOptions}
              placeholder="Selecciona marca"
              disabled={optionsLoading}
              {...register('brandId', { required: 'Requerido' })}
              error={errors.brandId?.message}
            />
            <Select
              label="Modelo *"
              options={modelosOptions}
              placeholder="Selecciona modelo"
              disabled={optionsLoading}
              {...register('modelId', { required: 'Requerido' })}
              error={errors.modelId?.message}
            />
          </div>

          <Select
            label="Sección *"
            options={sections.map((s) => ({ value: s.id, label: s.name }))}
            placeholder="Selecciona sección"
            disabled={optionsLoading || !brandWatch || !modelWatch}
            {...register('sectionId', { required: 'Requerido' })}
            error={errors.sectionId?.message}
          />

          <Textarea
            label="Descripción del problema *"
            rows={4}
            {...register('problem', { required: 'Requerido' })}
            error={errors.problem?.message}
          />

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <Select
              key={`edit-activity-types-${activityTypes.map((t) => t.id).join('|')}`}
              label="Tipo de actividad *"
              options={activityTypes.map((at) => ({ value: at.id, label: at.name }))}
              placeholder="Selecciona tipo"
              disabled={optionsLoading}
              {...register('activityTypeId', { required: 'Requerido' })}
              error={errors.activityTypeId?.message}
            />
            <Textarea
              label="Actividad *"
              rows={4}
              placeholder="Describe la actividad realizada"
              {...register('activityId', { required: 'Requerido' })}
              error={errors.activityId?.message}
            />
          </div>

          <Input
            label="Recurso"
            type="text"
            {...register('resource')}
          />

          <Input
            label="Tiempo (minutos) *"
            type="number"
            min={0}
            {...register('time', {
              required: 'Requerido',
              valueAsNumber: true,
              min: { value: 0, message: 'Debe ser ≥ 0' },
            })}
            error={errors.time?.message}
          />

          <p className="text-xs text-gray-500">
            Los adjuntos existentes no se modifican desde aquí; usa Nuevo registro o SharePoint para añadir archivos.
          </p>

          <div className="flex flex-wrap gap-3 border-t border-gray-200 pt-4">
            <Button type="submit" loading={saving}>
              Guardar cambios
            </Button>
            <Button type="button" variant="outline" onClick={onClose} disabled={saving}>
              Cancelar
            </Button>
          </div>
        </form>
      </div>
    </dialog>
  );
};
