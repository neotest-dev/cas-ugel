import * as XLSX from "xlsx";

export interface Worker {
  n: string;
  dni: string;
  apPaterno: string;
  apMaterno: string;
  nombres: string;
  fechaNac: string;
  cargo: string;
  codEssalud: string;
  cuentaBanco: string;
  leyendaRD: string;
  sistemaPensionario: string;
  cussp: string;
  fechaAfiliacion: string;
  aporteObligatorio: string;
  comision: string;
  primaSeguro: string;
  montoMensual: string;
  onp: string;
  prima: string;
  integra: string;
  profuturo: string;
  habitat: string;
  totalDscto: string;
  totalLiquido: string;
}

const MESES = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];

export function parsePeriodFromFilename(filename: string): { mes: string; anio: string } {
  // Ejemplo: "03 2026-PLLA CAS SEDE UGEL04-TSE_FINAL.xlsx"
  const m = filename.match(/(\d{1,2})[\s_-]+(\d{4})/);
  if (m) {
    const mesIdx = parseInt(m[1], 10) - 1;
    const anio = m[2];
    if (mesIdx >= 0 && mesIdx < 12) {
      return { mes: MESES[mesIdx], anio };
    }
  }
  const now = new Date();
  return { mes: MESES[now.getMonth()], anio: String(now.getFullYear()) };
}

function normHeader(s: any): string {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .replace(/[\n\r]/g, " ")
    .trim()
    .toUpperCase();
}

function fmtCell(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    // Si parece fecha serial (entre 1900-01-01 y 2100)
    return String(v);
  }
  if (v instanceof Date) {
    const d = v;
    return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
  }
  return String(v).trim();
}

function fmtNumber(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  if (isNaN(n)) return String(v);
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Mapeo: claves -> posibles nombres de columna (normalizados)
const COLUMN_MAP: Record<keyof Worker, string[]> = {
  n: ["N°", "Nº", "N", "NRO", "NUMERO"],
  dni: ["DNI"],
  apPaterno: ["APELLIDO PATERNO", "AP PATERNO", "AP. PATERNO"],
  apMaterno: ["APELLIDO MATERNO", "AP MATERNO", "AP. MATERNO"],
  nombres: ["NOMBRES", "NOMBRE"],
  fechaNac: ["FECHA DE NACIMIENTO", "FECHA NACIMIENTO", "F. NACIMIENTO", "F NACIMIENTO"],
  cargo: ["CARGO"],
  codEssalud: ["CODIGO ESSALUD", "CÓDIGO ESSALUD", "COD ESSALUD", "ESSALUD"],
  cuentaBanco: ["N° CUENTA BANCO LA NACION", "Nº CUENTA BANCO LA NACION", "CUENTA BANCO LA NACION", "CUENTA BANCO", "N° CUENTA"],
  leyendaRD: ["LEYENDA - RD", "LEYENDA RD", "LEYENDA"],
  sistemaPensionario: ["SISTEMA PENSIONARIO", "SIST PENSIONARIO", "REG PENSIONARIO"],
  cussp: ["CUSSP", "CUSPP"],
  fechaAfiliacion: ["FECHA DE AFILIACION", "FECHA AFILIACION", "F. AFILIACION"],
  aporteObligatorio: ["APORTE OBLIGATORIO"],
  comision: ["COMISION", "COMISIÓN"],
  primaSeguro: ["PRIMA SEGURO", "PRIMA DE SEGURO"],
  montoMensual: ["MONTO MENSUAL"],
  onp: ["ONP 13% DL. 19990", "ONP 13%", "ONP"],
  prima: ["PRIMA"],
  integra: ["INTEGRA"],
  profuturo: ["PROFUTURO"],
  habitat: ["HABITAT", "HÁBITAT"],
  totalDscto: ["TOTAL DSCTO", "TOTAL DESCUENTO", "TOTAL DCTO"],
  totalLiquido: ["TOTAL LIQUIDO", "TOTAL LÍQUIDO", "LIQUIDO A PAGAR"],
};

const REQUIRED: (keyof Worker)[] = [
  "dni", "apPaterno", "apMaterno", "nombres", "cargo", "montoMensual"
];

export interface ParseResult {
  workers: Worker[];
  period: { mes: string; anio: string };
}

export async function parseExcelFile(file: File): Promise<ParseResult> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === "CAS-SEDE");
  if (!sheetName) {
    throw new Error('No se encontró la hoja "CAS-SEDE" en el archivo.');
  }
  const sheet = wb.Sheets[sheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });

  // Encabezados combinados en filas 4,5,6 (índices 3,4,5). Concatenamos.
  const h1 = rows[3] || [];
  const h2 = rows[4] || [];
  const h3 = rows[5] || [];
  const maxCols = Math.max(h1.length, h2.length, h3.length);
  const headers: string[] = [];
  for (let c = 0; c < maxCols; c++) {
    const combined = [h1[c], h2[c], h3[c]]
      .map(v => normHeader(v))
      .filter(Boolean)
      .join(" ");
    headers.push(combined);
  }

  // Mapeo columna -> índice
  const colIndex: Partial<Record<keyof Worker, number>> = {};
  const missing: string[] = [];
  (Object.keys(COLUMN_MAP) as (keyof Worker)[]).forEach(key => {
    const candidates = COLUMN_MAP[key];
    let found = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (!h) continue;
      for (const cand of candidates) {
        if (h.includes(cand)) { found = i; break; }
      }
      if (found >= 0) break;
    }
    if (found >= 0) colIndex[key] = found;
    else if (REQUIRED.includes(key)) missing.push(candidates[0]);
  });

  if (missing.length) {
    throw new Error(`Faltan columnas obligatorias: ${missing.join(", ")}`);
  }

  // Procesar desde fila 7 (índice 6)
  const workers: Worker[] = [];
  for (let r = 6; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const dniIdx = colIndex.dni!;
    const dni = fmtCell(row[dniIdx]);
    const ap = fmtCell(row[colIndex.apPaterno!]);
    if (!dni && !ap) continue;
    const get = (k: keyof Worker) => {
      const i = colIndex[k];
      return i === undefined ? "" : fmtCell(row[i]);
    };
    const getNum = (k: keyof Worker) => {
      const i = colIndex[k];
      return i === undefined ? "" : fmtNumber(row[i]);
    };
    workers.push({
      n: get("n") || String(workers.length + 1),
      dni,
      apPaterno: ap,
      apMaterno: get("apMaterno"),
      nombres: get("nombres"),
      fechaNac: get("fechaNac"),
      cargo: get("cargo"),
      codEssalud: get("codEssalud"),
      cuentaBanco: get("cuentaBanco"),
      leyendaRD: get("leyendaRD"),
      sistemaPensionario: get("sistemaPensionario"),
      cussp: get("cussp"),
      fechaAfiliacion: get("fechaAfiliacion"),
      aporteObligatorio: getNum("aporteObligatorio"),
      comision: getNum("comision"),
      primaSeguro: getNum("primaSeguro"),
      montoMensual: getNum("montoMensual"),
      onp: getNum("onp"),
      prima: getNum("prima"),
      integra: getNum("integra"),
      profuturo: getNum("profuturo"),
      habitat: getNum("habitat"),
      totalDscto: getNum("totalDscto"),
      totalLiquido: getNum("totalLiquido"),
    });
  }

  const period = parsePeriodFromFilename(file.name);
  return { workers, period };
}

function pad(label: string, value: string, width = 60): string {
  return `${label}${value}`;
}

export function buildBoletaText(w: Worker, mes: string, anio: string): string {
  const isONP = w.sistemaPensionario.toUpperCase().includes("ONP");
  const apellidos = `${w.apPaterno} ${w.apMaterno}`.trim();

  const lines: string[] = [];
  lines.push(`BOLETA N° ${w.n}`);
  lines.push("");
  lines.push("DIRECCION REGIONAL LA LIBERTAD");
  lines.push("*B9 UGEL 04 SUR ESTE");
  lines.push("RUC - 20539889622");
  lines.push("");
  lines.push(`${mes} - ${anio}`);
  lines.push("");
  lines.push(`Apellidos                    : ${apellidos}`);
  lines.push(`Nombres                      : ${w.nombres}`);
  lines.push(`Fecha de Nacimiento          : ${w.fechaNac}`);
  lines.push(`Documento de Identidad       : (Lib.Electoral o D.N.) ${w.dni}`);
  lines.push(`Establecimiento              : UGEL Nº 04 TRUJILLO SUR ESTE`);
  lines.push(`Cargo                        : ${w.cargo}`);
  lines.push(`Tipo de Servidor             : ADMINISTRATIVO CONTRATADO`);
  lines.push(`Regimen Laboral              : D.LEG.Nº 1057 - CAS`);
  lines.push(`Niv.Mag./Grupo Ocup./Horas   : 0/0/40 Horas`);
  lines.push(`Tiempo de Servicio (AA-MM-DD): --        ESSALUD : ${w.codEssalud}`);
  lines.push(`Fecha de Registro            : Ingr.:            Termino:`);
  lines.push(`Cta. TeleAhorro o Nro.Cheque : CTA- ${w.cuentaBanco}`);
  lines.push(`Leyenda Permanente           : ${w.leyendaRD}`);
  lines.push(`Leyenda Mensual              :`);
  lines.push("");
  lines.push("------------------------------------------------------------");
  lines.push("PENSIONES");
  lines.push("------------------------------------------------------------");

  let descuentoLine = "";
  if (isONP) {
    lines.push(`Reg.Pensionario              : ONP /W`);
    descuentoLine = `-ONP                          S/.        ${w.onp}`;
  } else {
    lines.push(`Reg.Pensionario              : ${w.sistemaPensionario} / ${w.cussp}`);
    lines.push(`CFija                        : ${w.aporteObligatorio}`);
    lines.push(`FAfiliacion                  : ${w.fechaAfiliacion}`);
    lines.push(`CVariable                    : ${w.comision}`);
    lines.push(`Seguro                       : ${w.primaSeguro}`);
    const afpMonto = [w.prima, w.integra, w.profuturo, w.habitat].find(v => v && v !== "0.00") || "";
    descuentoLine = `-AFP                          S/.        ${afpMonto}`;
  }

  lines.push("");
  lines.push("------------------------------------------------------------");
  lines.push("INGRESOS");
  lines.push("------------------------------------------------------------");
  lines.push(`+Honorario                    S/.        ${w.montoMensual}`);
  lines.push("");
  lines.push("------------------------------------------------------------");
  lines.push("DESCUENTOS");
  lines.push("------------------------------------------------------------");
  lines.push(descuentoLine);
  lines.push("");
  lines.push("------------------------------------------------------------");
  lines.push("TOTALES");
  lines.push("------------------------------------------------------------");
  lines.push(`T-HONORARIO                   S/.        ${w.montoMensual}`);
  lines.push(`T-DSCTO                       S/.        ${w.totalDscto}`);
  lines.push(`T-LIQUI                       S/.        ${w.totalLiquido}`);
  lines.push("");
  lines.push(`MImponible                    S/.        ${w.montoMensual}`);
  lines.push("");
  lines.push(`Mensajes :`);

  return lines.join("\n");
}
