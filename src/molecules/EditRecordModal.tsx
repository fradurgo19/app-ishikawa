import React, { useCallback, useEffect, useId, useRef, useState } from 'react';
import { useForm } from 'react-hook-form';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Textarea } from '../atoms/Textarea';
import { Select } from '../atoms/Select';
import { sharePointService } from '../services/sharePointService';
import { Attachment, MachineRecord, Section, ActivityType } from '../types';
import { Download, Trash2, X } from 'lucide-react';

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

function resolveRecordAttachments(source: MachineRecord): Attachment[] {
  if (source.attachments && source.attachments.length > 0) {
    return [...source.attachments];
  }
  if (source.attachment) {
    return [source.attachment];
  }
  return [];
}

function attachmentKey(att: Attachment, index: number): string {
  const name = att.name?.trim() ?? '';
  return `${att.id}-${name}-${index}`;
}

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
  const [existingAttachments, setExistingAttachments] = useState<Attachment[]>([]);
  const [newStagedFiles, setNewStagedFiles] = useState<File[]>([]);

  const initialAttachmentNamesRef = useRef<Set<string>>(new Set());
  const newFilesInputId = useId();
  const newFilesInputRef = useRef<HTMLInputElement>(null);

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
    const initialList = resolveRecordAttachments(record);
    initialAttachmentNamesRef.current = new Set(
      initialList.map((a) => (a.name ?? '').trim()).filter(Boolean)
    );
    setExistingAttachments(initialList);
    setNewStagedFiles([]);
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

  const appendNewFiles = useCallback((picked: File[]) => {
    if (picked.length === 0) {
      return;
    }
    setNewStagedFiles((prev) => [...prev, ...picked]);
  }, []);

  const removeNewFileAt = useCallback((index: number) => {
    setNewStagedFiles((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const removeExistingAttachment = useCallback((index: number) => {
    setExistingAttachments((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const handleNewFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const list = input.files;
    const picked = list && list.length > 0 ? Array.from(list) : [];
    input.value = '';
    appendNewFiles(picked);
  };

  const onSubmit = useCallback(
    async (data: EditFormData) => {
      if (!record) {
        return;
      }
      setSaving(true);
      try {
        const keptNames = new Set(
          existingAttachments.map((a) => (a.name ?? '').trim()).filter(Boolean)
        );
        const removeAttachmentFileNames = [...initialAttachmentNamesRef.current].filter(
          (n) => !keptNames.has(n)
        );

        const hasAttachmentOps =
          removeAttachmentFileNames.length > 0 || newStagedFiles.length > 0;

        const baseUpdate: MachineRecord = {
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
          ...(existingAttachments.length > 0
            ? {
                attachment: existingAttachments[0],
                attachments: existingAttachments,
              }
            : {
                attachment: undefined,
                attachments: undefined,
              }),
        };

        const updated = await sharePointService.updateRecord(
          baseUpdate,
          hasAttachmentOps
            ? {
                addFiles: newStagedFiles.length > 0 ? newStagedFiles : undefined,
                removeAttachmentFileNames:
                  removeAttachmentFileNames.length > 0 ? removeAttachmentFileNames : undefined,
              }
            : undefined
        );
        onSaved(updated);
        onClose();
      } catch (e) {
        console.error('Error actualizando registro:', e);
        alert(e instanceof Error ? e.message : 'No se pudo guardar el registro.');
      } finally {
        setSaving(false);
      }
    },
    [record, existingAttachments, newStagedFiles, onSaved, onClose]
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

          <div className="space-y-3 rounded-lg border border-gray-200 bg-gray-50 p-4">
            <h3 className="text-sm font-medium text-gray-900">Adjuntos</h3>
            <p className="text-xs text-gray-600">
              Puedes quitar archivos actuales o añadir nuevos. Los cambios se aplican al guardar (SharePoint lista
              nativa).
            </p>

            {existingAttachments.length > 0 ? (
              <ul className="m-0 list-none space-y-2 p-0">
                {existingAttachments.map((att, index) => {
                  const href = att.url?.trim();
                  const label = att.name?.trim() || 'Adjunto';
                  return (
                    <li
                      key={attachmentKey(att, index)}
                      className="flex items-center justify-between gap-2 rounded border border-gray-200 bg-white px-3 py-2 text-sm"
                    >
                      <span className="min-w-0 flex-1 truncate" title={label}>
                        {href ? (
                          <a
                            href={href}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex max-w-full items-center gap-1 text-blue-600 hover:underline"
                          >
                            <Download className="h-3.5 w-3.5 shrink-0" aria-hidden />
                            <span className="truncate">{label}</span>
                          </a>
                        ) : (
                          <span className="text-gray-800">{label}</span>
                        )}
                      </span>
                      <button
                        type="button"
                        className="shrink-0 rounded p-1 text-red-600 hover:bg-red-50"
                        aria-label={`Quitar adjunto ${label}`}
                        onClick={() => removeExistingAttachment(index)}
                      >
                        <Trash2 className="h-4 w-4" aria-hidden />
                      </button>
                    </li>
                  );
                })}
              </ul>
            ) : (
              <p className="text-xs text-gray-500">No hay adjuntos en este registro.</p>
            )}

            <div className="space-y-2">
              <label htmlFor={newFilesInputId} className="block text-sm font-medium text-gray-700">
                Añadir archivos
              </label>
              <input
                id={newFilesInputId}
                ref={newFilesInputRef}
                type="file"
                multiple
                accept="image/*,.pdf,.doc,.docx,.xls,.xlsx"
                className="hidden"
                tabIndex={-1}
                onChange={handleNewFilesChange}
              />
              <button
                type="button"
                className="inline-flex rounded-md border border-gray-300 bg-white px-3 py-2 text-sm font-medium text-gray-700 shadow-sm hover:bg-gray-50"
                onClick={() => newFilesInputRef.current?.click()}
              >
                Elegir archivos
              </button>
              {newStagedFiles.length > 0 ? (
                <ul className="m-0 list-none space-y-1 rounded border border-gray-200 bg-white p-2 text-sm">
                  {newStagedFiles.map((file, index) => (
                    <li key={`${file.name}-${file.size}-${index}`} className="flex items-center justify-between gap-2">
                      <span className="truncate" title={file.name}>
                        {file.name}
                      </span>
                      <button
                        type="button"
                        className="shrink-0 rounded p-1 text-red-600 hover:bg-red-50"
                        aria-label={`Quitar ${file.name} de la lista`}
                        onClick={() => removeNewFileAt(index)}
                      >
                        <Trash2 className="h-4 w-4" aria-hidden />
                      </button>
                    </li>
                  ))}
                </ul>
              ) : null}
            </div>
          </div>

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
