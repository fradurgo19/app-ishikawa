import { Attachment, MachineRecord } from '../types';

/**
 * Payload para crear un registro nuevo copiando los datos de uno existente (misma API que Nuevo registro).
 * No incluye id ni marcas de tiempo; el primer adjunto se reutiliza como metadato si existe.
 */
export function buildDuplicateCreatePayload(
  record: MachineRecord
): Omit<MachineRecord, 'id' | 'createdAt' | 'updatedAt'> & { attachment?: Attachment } {
  const time = Number(record.time);
  const base = {
    tipoEquipoId: record.tipoEquipoId,
    brandId: record.brandId,
    modelId: record.modelId,
    sectionId: record.sectionId,
    problem: record.problem,
    activityTypeId: record.activityTypeId,
    activityId: record.activityId,
    resource: record.resource,
    time: Number.isFinite(time) ? time : 0,
    createdBy: record.createdBy,
  };
  const primary = record.attachments?.[0] ?? record.attachment;
  if (primary) {
    return { ...base, attachment: primary };
  }
  return base;
}
