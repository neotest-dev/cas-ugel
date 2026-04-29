/// <reference lib="webworker" />
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

function parsePeriodFromFilename(filename: string): { mes: string; anio: string } {
  const m = filename.match(/(\d{1,2})[\s_-]+(\d{4})/);
  if (m) {
    const mesIdx = parseInt(m[1], 10) - 1;
    const anio = m[2];
    if (mesIdx >= 0 && mesIdx < 12) return { mes: MESES[mesIdx], anio };
  }
  const now = new Date();
  return { mes: MESES[now.getMonth()], anio: String(now.getFullYear()) };
}

function normHeader(s: any): string {
  return String(s ?? "").replace(/\s+/g, " ").replace(/[\n\r]/g, " ").trim().toUpperCase();
}

function fmtCell(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}/${v.getFullYear()}`;
  }
  return String(v).trim();
}

function fmtNumber(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  if (isNaN(n)) return String(v);
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

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

const REQUIRED: (keyof Worker)[] = ["dni", "apPaterno", "apMaterno", "nombres", "cargo", "montoMensual"];

self.onmessage = async (e: MessageEvent<{ buffer: ArrayBuffer; filename: string }>) => {
  try {
    const { buffer, filename } = e.data;
    const SHEET = "CAS-SEDE";

    // Lectura mínima: solo la hoja necesaria, sin estilos, sin macros, sin formato HTML
    const wb = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellFormula: false,
      cellHTML: false,
      cellStyles: false,
      cellNF: false,
      bookVBA: false,
      bookFiles: false,
      bookProps: false,
      bookSheets: false,
      sheets: [SHEET],
      // matchear case-insensitive lo hace SheetJS internamente para `sheets` por nombre exacto;
      // si no, fallback más abajo.
    } as any);

    let sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === SHEET);
    if (!sheetName) {
      // Fallback: releer todas las hojas si no encontró por filtro
      const wb2 = XLSX.read(buffer, {
        type: "array", cellDates: true, cellFormula: false,
        cellHTML: false, cellStyles: false, cellNF: false, bookVBA: false,
      } as any);
      sheetName = wb2.SheetNames.find(n => n.trim().toUpperCase() === SHEET);
      if (!sheetName) throw new Error('No se encontró la hoja "CAS-SEDE" en el archivo.');
      (wb as any).Sheets = wb2.Sheets;
      (wb as any).SheetNames = wb2.SheetNames;
    }

    const sheet = wb.Sheets[sheetName];

    // Limitar el rango de lectura: descartar filas previas a la 7 (índice 6)
    // y leer todas las filas válidas
    const ref = sheet["!ref"];
    if (!ref) throw new Error("La hoja CAS-SEDE está vacía.");
    const range = XLSX.utils.decode_range(ref);

    // Necesitamos filas 4,5,6 (encabezados) y 7+ (datos)
    const headerRange = { s: { r: 3, c: range.s.c }, e: { r: 5, c: range.e.c } };
    const dataRange = { s: { r: 6, c: range.s.c }, e: range.e };

    const headerRows: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1, raw: true, defval: "",
      range: XLSX.utils.encode_range(headerRange),
    });
    const dataRows: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1, raw: false, defval: "",
      range: XLSX.utils.encode_range(dataRange),
    });

    const h1 = headerRows[0] || [];
    const h2 = headerRows[1] || [];
    const h3 = headerRows[2] || [];
    const maxCols = Math.max(h1.length, h2.length, h3.length);
    const headers: string[] = new Array(maxCols);
    for (let c = 0; c < maxCols; c++) {
      const a = normHeader(h1[c]);
      const b = normHeader(h2[c]);
      const d = normHeader(h3[c]);
      headers[c] = [a, b, d].filter(Boolean).join(" ");
    }

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

    if (missing.length) throw new Error(`Faltan columnas obligatorias: ${missing.join(", ")}`);

    const workers: Worker[] = [];
    const dniIdx = colIndex.dni!;
    const apIdx = colIndex.apPaterno!;
    const keys = Object.keys(colIndex) as (keyof Worker)[];
    const numericKeys = new Set<keyof Worker>([
      "aporteObligatorio", "comision", "primaSeguro", "montoMensual",
      "onp", "prima", "integra", "profuturo", "habitat", "totalDscto", "totalLiquido"
    ]);

    for (let r = 0; r < dataRows.length; r++) {
      const row = dataRows[r];
      if (!row) continue;
      const dni = fmtCell(row[dniIdx]);
      const ap = fmtCell(row[apIdx]);
      if (!dni && !ap) continue;
      const w: any = { n: "", dni: "", apPaterno: "", apMaterno: "", nombres: "", fechaNac: "", cargo: "",
        codEssalud: "", cuentaBanco: "", leyendaRD: "", sistemaPensionario: "", cussp: "", fechaAfiliacion: "",
        aporteObligatorio: "", comision: "", primaSeguro: "", montoMensual: "", onp: "", prima: "",
        integra: "", profuturo: "", habitat: "", totalDscto: "", totalLiquido: "" };
      for (const k of keys) {
        const i = colIndex[k]!;
        const v = row[i];
        w[k] = numericKeys.has(k) ? fmtNumber(v) : fmtCell(v);
      }
      if (!w.n) w.n = String(workers.length + 1);
      workers.push(w as Worker);
    }

    const period = parsePeriodFromFilename(filename);
    (self as any).postMessage({ ok: true, workers, period });
  } catch (err: any) {
    (self as any).postMessage({ ok: false, error: err?.message || "Error desconocido" });
  }
};
