import type { AttachmentFilePayload } from '../types';

/** Límite por archivo para evitar cuerpos HTTP enormes y errores en Graph (ajustable). */
const MAX_BYTES_PER_FILE = 15 * 1024 * 1024;

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== 'string') {
        reject(new Error('No se pudo leer el archivo'));
        return;
      }
      const comma = result.indexOf(',');
      const base64 = comma >= 0 ? result.slice(comma + 1) : result;
      resolve(base64);
    };
    reader.onerror = () => reject(new Error('Error al leer el archivo'));
    reader.readAsDataURL(file);
  });
}

export async function filesToAttachmentPayloads(files: File[]): Promise<AttachmentFilePayload[]> {
  const payloads: AttachmentFilePayload[] = [];
  for (const file of files) {
    if (file.size > MAX_BYTES_PER_FILE) {
      throw new Error(
        `El archivo "${file.name}" supera el tamaño máximo permitido (${Math.floor(MAX_BYTES_PER_FILE / (1024 * 1024))} MB).`
      );
    }
    const contentBase64 = await readFileAsBase64(file);
    payloads.push({
      name: file.name,
      contentType: file.type || 'application/octet-stream',
      contentBase64,
    });
  }
  return payloads;
}
