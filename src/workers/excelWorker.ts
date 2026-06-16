/// <reference lib="webworker" />
import ExcelJS, { CellValue } from "exceljs";
import { inferCategoryIdFromText } from "@/lib/boleta";

type RowValues = CellValue[];

type MsgIn = {
  buffer: ArrayBuffer;
  filename: string;
};

type MsgOut =
  | {
    ok: true;
    workers: Worker[];
    period: { mes: string; anio: string };
    categoryId: string | null;
    debug?: {
      headerRowIdx: number;
      col: Partial<Record<keyof Worker, number>>;
      headers: string[];
    };
  }
  | {
    ok: false;
    error: string;
  };

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
  leyendaMensual: string;
  sistemaPensionario: string;
  cussp: string;
  fechaAfiliacion: string;
  fechaDevengue: string;
  aporteObligatorio: string;
  comision: string;
  primaSeguro: string;
  montoMensual: string;
  descuentoPension: string;
  onp: string;
  prima: string;
  integra: string;
  profuturo: string;
  habitat: string;
  totalDscto: string;
  otrosDsctos: string;
  dsctoEntidades: string;
  dsctoJudicial: string;
  totalLiquido: string;
}

const MESES = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];

function parsePeriodFromFilename(filename: string): { mes: string; anio: string } {
  const m = filename.match(/(\d{1,2})[\s_-]+(\d{4})/);

  if (m) {
    return {
      mes: MESES[parseInt(m[1], 10) - 1] || "",
      anio: m[2]
    };
  }

  return { mes: "", anio: "" };
}

function norm(v: CellValue | undefined): string {
  return txt(v)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/º/g, "°")
    .replace(/[^a-zA-Z0-9°]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function txt(v: CellValue | undefined): string {
  if (v === null || v === undefined) return "";

  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2, "0")}/${String(
      v.getMonth() + 1
    ).padStart(2, "0")}/${v.getFullYear()}`;
  }

  if (typeof v === "object") {
    if ("richText" in v && Array.isArray(v.richText)) {
      return v.richText.map(item => String(item.text ?? "")).join("");
    }
    if ("text" in v) return String(v.text).trim();
    if ("result" in v) {
      if (v.result === null || v.result === undefined) return "";
      if (typeof v.result === "object") {
        return txt(v.result as any);
      }
      return String(v.result).trim();
    }
    if ("formula" in v) {
      return "";
    }
    return "";
  }

  return String(v).trim();
}

function money(v: CellValue | undefined): string {
  if (v === null || v === undefined || v === "") {
    return "0.00";
  }

  if (typeof v === "object") {
    if ("result" in v) {
      if (v.result === null || v.result === undefined) return "0.00";
      if (typeof v.result === "object") {
        return money(v.result as any);
      }
      const n = Number(v.result);
      return Number.isNaN(n) ? "0.00" : n.toFixed(2);
    }

    if ("text" in v) {
      const n = Number(v.text);
      return Number.isNaN(n) ? "0.00" : n.toFixed(2);
    }

    if ("richText" in v && Array.isArray(v.richText)) {
      const textVal = v.richText.map(item => String(item.text ?? "")).join("");
      const cleanVal = textVal.replace(/[^\d.-]/g, "");
      const n = Number(cleanVal);
      return Number.isNaN(n) ? "0.00" : n.toFixed(2);
    }
    return "0.00";
  }

  const cleanVal = String(v).replace(/[^\d.-]/g, "");
  const n = Number(cleanVal);

  return Number.isNaN(n) ? "0.00" : n.toFixed(2);
}

function colLetterToNumber(col: string): number {
  let num = 0;

  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }

  return num;
}

type ColumnSpec = {
  fallback?: string[];
  headers?: string[];
};

function isDataRow(row: RowValues | undefined): boolean {
  if (!row) return false;
  for (let i = 1; i < row.length; i++) {
    const val = String(row[i] ?? "").trim();
    if (/^\d{8}$/.test(val)) {
      return true;
    }
  }
  return false;
}

const COLUMN_MAP: Record<keyof Worker, ColumnSpec> = {
  n: { fallback: ["A"], headers: ["N°", "Nº", "N.", "NUMERO", "NUM"] },
  dni: { fallback: ["B"], headers: ["DNI", "N° DNI", "Nº DNI", "DOCUMENTO DE IDENTIDAD", "DOC. IDENTIDAD", "DOCUMENTO"] },
  apPaterno: { fallback: ["C"], headers: ["APELLIDO PATERNO", "AP. PATERNO", "APE. PATERNO", "PATERNO"] },
  apMaterno: { fallback: ["D"], headers: ["APELLIDO MATERNO", "AP. MATERNO", "APE. MATERNO", "MATERNO"] },
  nombres: { fallback: ["E"], headers: ["NOMBRES", "NOMBRE"] },
  fechaNac: { fallback: ["BF"], headers: ["FECHA NACIMIENTO", "FECHA DE NACIMIENTO", "FEC. NACIMIENTO", "FEC. NAC.", "F. NAC."] },
  cargo: { fallback: ["G", "F"], headers: ["CARGO", "PUESTO", "CARGO/PUESTO"] },
  codEssalud: { fallback: ["BQ"], headers: ["CODIGO ESSALUD", "COD. ESSALUD", "COD.ESSALUD", "ESSALUD AUTOGENERADO", "AUTOGENERADO", "COD. AUTOGENERADO", "CODIGO AUTOGENERADO"] },
  cuentaBanco: { fallback: ["AD"], headers: ["Nº CUENTA BANCO LA NACION", "N° CUENTA BANCO LA NACION", "NRO CUENTA BANCO LA NACION", "CUENTA BANCO LA NACION", "NRO. CTA BANCO", "NRO. CUENTA", "CTA. BANCO", "CTA BANCO", "CUENTA BANCO", "BANCO LA NACION", "BANCO DE LA NACION"] },
  leyendaRD: { fallback: ["BU"], headers: ["LEYENDA - RD"] },
  leyendaMensual: { fallback: ["BT"], headers: ["LEYENDA MENSUAL"] },
  sistemaPensionario: { fallback: ["BG"], headers: ["SISTEMA PENSIONARIO", "SIST. PENSION", "SIST. PENS.", "REG. PENSION", "REGIMEN PENSIONARIO", "AFP/ONP"] },
  cussp: { fallback: ["BO"], headers: ["CUSSP", "CODIGO CUSSP", "N° CUSSP", "Nº CUSSP"] },
  fechaAfiliacion: { fallback: ["BJ"], headers: ["FECHA DE INGRESO DE REGISTRO", "FECHA INGRESO DE REGISTRO", "FECHA DE INGRESO", "FECHA INGRESO", "FECHA AFILIACION", "F. AFILIACION", "FEC. AFIL.", "FECHA DE AFILIACION"] },
  fechaDevengue: { fallback: ["BK"], headers: ["FECHA DE TERMINO DE REGISTRO", "FECHA TERMINO DE REGISTRO", "FECHA DE TERMINO", "FECHA TERMINO", "FECHA DEVENGUE", "F. DEVENGUE", "FEC. DEV.", "FECHA DE DEVENGUE"] },
  aporteObligatorio: { fallback: ["BL"], headers: ["APORTE OBLIGATORIO", "APORTE OBLIG", "APORTE OB."] },
  comision: { fallback: ["BM"], headers: ["COMISION", "COM. VARIABLE", "COMIS."] },
  primaSeguro: { fallback: ["BN"], headers: ["PRIMA SEGURO", "PRIMA SEG", "PRIMA SEG.", "SEG.", "SEGURO"] },
  montoMensual: { fallback: ["P", "H"], headers: ["PAGO TOTAL MENSUAL"] },
  descuentoPension: { fallback: ["BH", "S"], headers: ["MONTO SISTEMA PENSION", "ONP", "PRIMA", "INTEGRA", "PROFUTURO", "HABITAT", "DESCUENTO PENSION", "TOT. DSCTO. PENS"] },
  onp: { fallback: ["S"], headers: ["ONP", "ONP 13%", "DECRETO LEY 19990", "D.L. 19990", "19990"] },
  prima: { fallback: ["T"], headers: ["PRIMA"] },
  integra: { fallback: ["U"], headers: ["INTEGRA"] },
  profuturo: { fallback: ["V"], headers: ["PROFUTURO"] },
  habitat: { fallback: ["W"], headers: ["HABITAT", "HABITAD"] },
  totalDscto: { fallback: ["AA"], headers: ["TOTAL DSCTO"] },
  otrosDsctos: { headers: ["OTROS DSCTOS", "OTROS DESCUENTOS"] },
  dsctoEntidades: { headers: ["DESCUENTO ENTIDADES", "DSCTO ENTIDADES", "ENTIDADES"] },
  dsctoJudicial: { headers: ["DSCTO JUDICIAL", "DESCUENTO JUDICIAL", "JUDICIAL"] },
  totalLiquido: { fallback: ["AB", "AC"], headers: ["TOTAL LIQUIDO"] }
};

function findHeaderColumn(headers: string[], names: string[]): number | undefined {
  for (let i = 1; i < headers.length; i++) {
    const header = headers[i] ?? "";
    for (const name of names) {
      const normalizedName = norm(name);
      if (normalizedName && header.includes(normalizedName)) {
        return i;
      }
    }
  }
  return undefined;
}

function resolveColumn(spec: ColumnSpec, headers: string[]): number | undefined {
  const byHeader = spec.headers ? findHeaderColumn(headers, spec.headers) : undefined;
  if (byHeader) return byHeader;

  if (!spec.fallback) return undefined;

  for (const letter of spec.fallback) {
    return colLetterToNumber(letter);
  }
}

self.onmessage = async (e: MessageEvent<MsgIn>) => {
  try {
    const { buffer, filename } = e.data;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const ws = workbook.getWorksheet("CAS-SEDE");

    if (!ws) {
      const fail: MsgOut = { ok: false, error: "No existe hoja CAS-SEDE" };
      self.postMessage(fail);
      return;
    }

    const rows: CellValue[][] = [];
    const totalRows = Math.max(300, ws.rowCount || 0);
    const limitRows = Math.min(1000, totalRows);

    for (let i = 1; i <= limitRows; i++) {
      const row = ws.getRow(i);
      rows[i] = [];
      for (let c = 1; c <= 120; c++) {
        rows[i][c] = row.getCell(c).value;
      }
    }

    const titleText = [rows[1], rows[2], rows[3]]
      .flatMap((row) => (row ?? []).map((value) => txt(value)))
      .filter(Boolean)
      .join(" ");

    // Buscar fila de encabezados dinámicamente
    let headerRowIdx = 5; // default fallback
    let maxMatches = 0;
    const headerKeywords = ["DNI", "PATERNO", "MATERNO", "NOMBRES", "CARGO"];

    for (let r = 1; r <= 30; r++) {
      const row = rows[r];
      if (!row) continue;
      let matches = 0;
      for (let c = 1; c < row.length; c++) {
        const val = norm(row[c]);
        if (headerKeywords.some(kw => val.includes(kw))) {
          matches++;
        }
      }
      if (matches > maxMatches && matches >= 3) {
        maxMatches = matches;
        headerRowIdx = r;
      }
    }

    // Recopilar filas de cabecera: unificar filas 4, 5 y 6 del excel cargado
    const headerRows: RowValues[] = [];
    for (const r of [4, 5, 6]) {
      const row = rows[r];
      if (row) {
        headerRows.push(row);
      }
    }

    const maxCols = 120;
    const headers: string[] = [];

    for (let c = 1; c <= maxCols; c++) {
      headers[c] = headerRows
        .map(row => norm(row[c]))
        .filter(Boolean)
        .join(" ");
    }

    const col: Partial<Record<keyof Worker, number>> = {};

    (Object.keys(COLUMN_MAP) as (keyof Worker)[]).forEach((key) => {
      col[key] = resolveColumn(COLUMN_MAP[key], headers);
    });

    const workers: Worker[] = [];

    const limit = Math.min(300, rows.length - 1);
    for (let r = headerRowIdx + 1; r <= limit; r++) {
      const row = rows[r];
      if (!row) continue;

      const dni = txt(row[col.dni ?? 0]).replace(/\D/g, "");

      if (dni.length !== 8) continue;

      // Obtener valores de celdas de pensión específicas
      const cellONP = money(row[col.onp ?? 0]);
      const cellPrima = money(row[col.prima ?? 0]);
      const cellIntegra = money(row[col.integra ?? 0]);
      const cellProfuturo = money(row[col.profuturo ?? 0]);
      const cellHabitat = money(row[col.habitat ?? 0]);

      let sysPension = txt(row[col.sistemaPensionario ?? 0]).toUpperCase().trim();

      // Si la columna principal está vacía, o dice simplemente "AFP" o "ONP"
      if (!sysPension || sysPension === "" || sysPension === "AFP" || sysPension === "ONP") {
        if (Number(cellONP) > 0) sysPension = "ONP";
        else if (Number(cellPrima) > 0) sysPension = "PRIMA";
        else if (Number(cellIntegra) > 0) sysPension = "INTEGRA";
        else if (Number(cellProfuturo) > 0) sysPension = "PROFUTURO";
        else if (Number(cellHabitat) > 0) sysPension = "HABITAT";
      }

      // Obtener o calcular el aporte obligatorio (CFija)
      let aporteObligVal = money(row[col.aporteObligatorio ?? 0]);
      if (Number(aporteObligVal) === 0) {
        if (sysPension === "ONP") aporteObligVal = cellONP;
        else if (sysPension === "PRIMA") aporteObligVal = cellPrima;
        else if (sysPension === "INTEGRA") aporteObligVal = cellIntegra;
        else if (sysPension === "PROFUTURO") aporteObligVal = cellProfuturo;
        else if (sysPension === "HABITAT") aporteObligVal = cellHabitat;
      }

      // Si descuentoPension sale 0.00, calcularlo según el tipo de pensión
      let descPension = money(row[col.descuentoPension ?? 0]);
      if (Number(descPension) === 0) {
        if (sysPension === "ONP") {
          descPension = cellONP;
        } else {
          // Es una AFP (PRIMA, INTEGRA, PROFUTURO, HABITAT)
          const aOblig = Number(aporteObligVal);
          const comisionVal = Number(money(row[col.comision ?? 0]));
          const seguroVal = Number(money(row[col.primaSeguro ?? 0]));
          descPension = (aOblig + comisionVal + seguroVal).toFixed(2);
        }
      }

      workers.push({
        n: txt(row[col.n ?? 0]) || String(workers.length + 1),
        dni,
        apPaterno: txt(row[col.apPaterno ?? 0]),
        apMaterno: txt(row[col.apMaterno ?? 0]),
        nombres: txt(row[col.nombres ?? 0]),
        fechaNac: txt(row[col.fechaNac ?? 0]),
        cargo: txt(row[col.cargo ?? 0]),
        codEssalud: txt(row[col.codEssalud ?? 0]),
        cuentaBanco: txt(row[col.cuentaBanco ?? 0]),
        leyendaRD: txt(row[col.leyendaRD ?? 0]),
        leyendaMensual: txt(row[col.leyendaMensual ?? 0]),
        sistemaPensionario: sysPension,
        cussp: txt(row[col.cussp ?? 0]),
        fechaAfiliacion: txt(row[col.fechaAfiliacion ?? 0]),
        fechaDevengue: txt(row[col.fechaDevengue ?? 0]),
        aporteObligatorio: aporteObligVal,
        comision: money(row[col.comision ?? 0]),
        primaSeguro: money(row[col.primaSeguro ?? 0]),
        montoMensual: money(row[col.montoMensual ?? 0]),
        descuentoPension: descPension,
        onp: cellONP,
        prima: cellPrima,
        integra: cellIntegra,
        profuturo: cellProfuturo,
        habitat: cellHabitat,
        totalDscto: money(row[col.totalDscto ?? 0]),
        otrosDsctos: money(row[col.otrosDsctos ?? 0]),
        dsctoEntidades: money(row[col.dsctoEntidades ?? 0]),
        dsctoJudicial: money(row[col.dsctoJudicial ?? 0]),
        totalLiquido: money(row[col.totalLiquido ?? 0]),
      });
    }

    const ok: MsgOut = {
      ok: true,
      workers,
      period: parsePeriodFromFilename(filename),
      categoryId: inferCategoryIdFromText(titleText),
      debug: {
        headerRowIdx,
        col,
        headers: headers.slice(0, 100)
      }
    };

    self.postMessage(ok);

  } catch (error: unknown) {
    const fail: MsgOut = {
      ok: false,
      error: error instanceof Error ? error.message : "Error procesando archivo"
    };

    self.postMessage(fail);
  }
};
