/**
 * Matriz oficial Tipo de equipo → Marca → Modelo (valores alineados con SharePoint).
 * Una fila por combinación válida; se usa para filtros en cascada y validación.
 */
const RAW_ROWS = `
BULLDOZER|CASE|1150L
BULLDOZER|CASE|1650M
CARGADOR|LIUGONG|856H
CARGADOR|LIUGONG|856HE
CARGADOR|CASE|821G
COMPACTADOR|DYNAPAC|CC1200
COMPACTADOR|DYNAPAC|CA1300D
COMPACTADOR NEUMATIVO|DYNAPAC|CP1200
EXCAVADORA|CASE|CX220C
EXCAVADORA|CASE|CX240C
EXCAVADORA|CASE|CX350C
EXCAVADORA|LBX|210X3
EXCAVADORA|LBX|130X3
EXCAVADORA|LBX|300X3
EXCAVADORA|LBX|360X3
EXCAVADORA|YANMAR|VIO80-1
EXCAVADORA|CASE|CX220C LC
EXCAVADORA|HITACHI|ZX200-5G
EXCAVADORA|HITACHI|ZX130-5G
EXCAVADORA|HITACHI|ZX210LC-5B
EXCAVADORA|HITACHI|ZX350LC-5B
EXCAVADORA|HITACHI|ZX130-5B
EXCAVADORA|LIUGONG|922F
EXCAVADORA|LIUGONG|933F
EXCAVADORA|LIUGONG|915F
EXCAVADORA|LIUGONG|920F
EXCAVADORA|HITACHI|ZX75US-7
EXCAVADORA|HITACHI|ZX75US-3
EXCAVADORA|HITACHI|ZX135US-3
EXCAVADORA|HITACHI|ZX200-3
EXCAVADORA|HITACHI|ZX70-3
EXCAVADORA|HITACHI|ZX75US
EXCAVADORA|HITACHI|ZX225US-3
EXCAVADORA|HITACHI|ZX120-5B
EXCAVADORA|HITACHI|ZX120-3
EXCAVADORA|HITACHI|ZX135US-5B
EXCAVADORA|HITACHI|ZX130K-3
EXCAVADORA|HITACHI|ZX40U-5B
EXCAVADORA|HITACHI|ZX50U-3
EXCAVADORA|HITACHI|ZX225USR-3
EXCAVADORA|HITACHI|ZX210K-5B
EXCAVADORA|HITACHI|ZX75US-5B
EXCAVADORA|HITACHI|ZX135USK-5B
EXCAVADORA|HITACHI|ZX200-5B
EXCAVADORA|HITACHI|ZX50U-5B
EXCAVADORA|HITACHI|ZX225US-5B
EXCAVADORA|HITACHI|ZX75USK-5B
EXCAVADORA|HITACHI|ZX240LC-5B
EXCAVADORA|HITACHI|ZX200-6
EXCAVADORA|HITACHI|ZX225USR-5B
EXCAVADORA|HITACHI|ZX135US-6
EXCAVADORA|HITACHI|ZX210LCH-5B
EXCAVADORA|HITACHI|ZX350LC-6N
EXCAVADORA|HITACHI|ZX200LC-6
EXCAVADORA|HITACHI|ZX-200-6
EXCAVADORA|HITACHI|ZX120-6
EXCAVADORA|HITACHI|ZX225USRLCK-6
EXCAVADORA|HITACHI|ZX210K-6
EXCAVADORA|HITACHI|ZX130K-6
EXCAVADORA|HITACHI|ZX330LC-5B
EXCAVADORA|HITACHI|ZX240-6
EXCAVADORA|HITACHI|ZX17U-5A
EXCAVADORA|HITACHI|ZX350H-5B
EXCAVADORA|HITACHI|ZX350-5B
EXCAVADORA|HITACHI|ZX210LCK-6
EXCAVADORA|HITACHI|ZX135US-6N
EXCAVADORA|HITACHI|ZX135US
EXCAVADORA|HITACHI|ZX210 LC
EXCAVADORA|HITACHI|ZX300 LC-6
EXCAVADORA|HITACHI|ZX345US LC-6N
EXCAVADORA|HITACHI|ZX350LC-6
EXCAVADORA|HITACHI|ZX17U-2
EXCAVADORA|HITACHI|ZX225USR-6
EXCAVADORA|HITACHI|ZX350K-5B
EXCAVADORA|HITACHI|ZX225USRLC-5B
EXCAVADORA|HITACHI|ZX135USK-6
EXCAVADORA|HITACHI|ZX225USRK-6
EXCAVADORA|HITACHI|ZX250K-6
EXCAVADORA|HITACHI|ZX225US-6
EXCAVADORA|HITACHI|ZX330-6
EXCAVADORA|HITACHI|ZX200X-5B
EXCAVADORA|LIUGONG|915FW
EXCAVADORA|YANMAR|VIO80-7
FRESADORA|BOMAG|BM 1000/20
FRESADORA|BOMAG|BM 500/15-2
MINICARGADOR|CASE|SR175B
MINICARGADOR|CASE|SR200B
MINICARGADOR|CASE|SR220B
MINICARGADOR|CASE|SR250B
MINICARGADOR|CASE|SR210B
MINICARGADOR|CASE|SR240B
MINICARGADOR|CASE|SR270B
MINIEXCAVADORA|YANMAR|VIO50-6B
MINIEXCAVADORA|YANMAR|VIO17-1B
MINIEXCAVADORA|YANMAR|VIO35-6B
MINIEXCAVADORA|YANMAR|VIO80-1
MINIEXCAVADORA|YANMAR|VIO35-7
MINIEXCAVADORA|YANMAR|VIO50-6
MINIEXCAVADORA|YANMAR|VIO17
MINIPAVIMENTADORA|DYNAPAC|F80W
MOTONIVELADORA|CASE|845B
MOTONIVELADORA|CASE|845C
MOTONIVELADORA|CASE|865C
MOTONIVELADORA|LIUGONG|4165D
PAVIMENTADORA|DYNAPAC|F1800C
RETROCARGADOR|CASE|575SV
RETROCARGADOR|CASE|580SN
RETROCARGADOR|CASE|580N
RETROCARGADOR|CASE|580SV
RETROCARGADOR|CASE|851FX
RODILLO COMBI|DYNAPAC|CC1400CVI
RODILLO TANDEM|DYNAPAC|CC1300VI
RODILLO TANDEM|DYNAPAC|CC1400VI
RODILLO TANDEM|DYNAPAC|CC2200VI
RODILLO TANDEM|DYNAPAC|CC1200VI
VIBROCOMPACTADOR|CASE|1107EX
VIBROCOMPACTADOR|DYNAPAC|CA1500D
VIBROCOMPACTADOR|DYNAPAC|CA25D
VIBROCOMPACTADOR|DYNAPAC|CA15D
`.trim();

export interface EquipmentMatrixRow {
  tipoEquipo: string;
  marca: string;
  modelo: string;
}

/** Normaliza etiquetas de la matriz y valores de SharePoint para comparación estable. */
export function normalizeLabel(value: string): string {
  return value.trim().replace(/\s+/g, ' ');
}

function parseRows(): EquipmentMatrixRow[] {
  const seen = new Set<string>();
  const rows: EquipmentMatrixRow[] = [];

  for (const line of RAW_ROWS.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed) {
      continue;
    }

    const parts = trimmed.split('|').map((part) => normalizeLabel(part));
    if (parts.length < 3) {
      continue;
    }

    const [tipoEquipo, marca, ...modeloParts] = parts;
    const modelo = normalizeLabel(modeloParts.join('|'));
    const key = `${tipoEquipo}|${marca}|${modelo}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    rows.push({ tipoEquipo, marca, modelo });
  }

  return rows;
}

export const EQUIPMENT_MATRIX_ROWS: EquipmentMatrixRow[] = parseRows();

/** Modelos definidos en la matriz (p. ej. Hitachi) para enriquecer desplegables de nuevo registro aunque no estén aún en SharePoint. */
export function getDistinctModeloIdsFromMatrix(): string[] {
  const set = new Set<string>();
  for (const row of EQUIPMENT_MATRIX_ROWS) {
    const id = row.modelo.trim();
    if (id) {
      set.add(id);
    }
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

export function getDistinctTiposEquipo(): string[] {
  const set = new Set(EQUIPMENT_MATRIX_ROWS.map((row) => row.tipoEquipo));
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

export function getMarcasForTipoEquipo(tipoEquipo: string): string[] {
  const normalized = normalizeLabel(tipoEquipo);
  const set = new Set(
    EQUIPMENT_MATRIX_ROWS.filter((row) => row.tipoEquipo === normalized).map((row) => row.marca)
  );
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

export function getModelosForTipoYMarca(tipoEquipo: string, marca: string): string[] {
  const t = normalizeLabel(tipoEquipo);
  const m = normalizeLabel(marca);
  return EQUIPMENT_MATRIX_ROWS.filter((row) => row.tipoEquipo === t && row.marca === m)
    .map((row) => row.modelo)
    .sort((a, b) => a.localeCompare(b, 'es'));
}

export function isValidEquipmentCombination(
  tipoEquipo: string,
  marca: string,
  modelo: string
): boolean {
  const t = normalizeLabel(tipoEquipo);
  const m = normalizeLabel(marca);
  const mo = normalizeLabel(modelo);
  return EQUIPMENT_MATRIX_ROWS.some(
    (row) => row.tipoEquipo === t && row.marca === m && row.modelo === mo
  );
}
