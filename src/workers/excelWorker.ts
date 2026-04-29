/// <reference lib="webworker" />
import ExcelJS, { CellValue } from "exceljs";

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
  sistemaPensionario: string;
  cussp: string;
  fechaAfiliacion: string;
  fechaDevengue: string;
  aporteObligatorio: string;
  comision: string;
  primaSeguro: string;
  montoMensual: string;
  descuentoPension: string;
  totalDscto: string;
  totalLiquido: string;
}

const MESES = [
  "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
  "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
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
  return String(v ?? "")
      .replace(/\n/g, " ")
      .replace(/\r/g, " ")
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
    if ("text" in v) return String(v.text).trim();

    if ("result" in v) return String(v.result ?? "").trim();
  }

  return String(v).trim();
}

function money(v: CellValue | undefined): string {
  if (v === null || v === undefined || v === "") return "";

  // Si es fórmula de Excel
  if (typeof v === "object" && v !== null) {
    if ("result" in v) {
      const n = Number(v.result);

      if (Number.isNaN(n)) return "";

      return n.toFixed(2);
    }

    if ("text" in v) {
      const n = Number(v.text);

      if (Number.isNaN(n)) return "";

      return n.toFixed(2);
    }
  }

  const n = Number(v);

  if (Number.isNaN(n)) return "";

  return n.toFixed(2);
}

function colLetterToNumber(col: string): number {
  let num = 0;

  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }

  return num;
}

const COLUMN_MAP: Record<keyof Worker, string[]> = {
  n:["__COL_A__"],
  dni:["__COL_B__"],
  apPaterno:["__COL_C__"],
  apMaterno:["__COL_D__"],
  nombres:["__COL_E__"],

  fechaNac:["__COL_BF__"],
  cargo:["__COL_G__"],

  codEssalud:["__COL_BQ__"],
  cuentaBanco:["__COL_AD__"],
  leyendaRD:["__COL_BU__"],

  sistemaPensionario:["__COL_BG__"],
  cussp:["__COL_BO__"],

  fechaAfiliacion:["__COL_BJ__"],
  fechaDevengue:["__COL_BK__"],
  aporteObligatorio:["__COL_BL__"],
  comision:["__COL_BM__"],
  primaSeguro:["__COL_BN__"],

  // ESTOS AÚN POR HEADER (como pediste)
  montoMensual:["__COL_P__"],
  descuentoPension:["__COL_BH__"],
  totalDscto:["__COL_AA__"],
  totalLiquido:["__COL_AB__"]
};

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

    const rows: RowValues[] = [];

    for (let i = 1; i <= 100; i++) {
      const row = ws.getRow(i);
      rows[i] = row.values as RowValues;
    }

    const h1 = rows[4] ?? [];
    const h2 = rows[5] ?? [];
    const h3 = rows[6] ?? [];

    const maxCols = Math.max(h1.length, h2.length, h3.length);

    const headers: string[] = [];

    for (let c = 1; c <= maxCols; c++) {
      headers[c] = [h1[c], h2[c], h3[c]]
          .map(v => norm(v))
          .filter(Boolean)
          .join(" ");
    }

    const col: Partial<Record<keyof Worker, number>> = {};

    (Object.keys(COLUMN_MAP) as (keyof Worker)[]).forEach((key) => {
      const names = COLUMN_MAP[key];

      if (names[0].startsWith("__COL_")) {
        const letter = names[0]
            .replace("__COL_", "")
            .replace("__", "");

        col[key] = colLetterToNumber(letter);
        return;
      }

      for (let i = 1; i < headers.length; i++) {
        const h = headers[i] ?? "";

        if (names.some(name => h.includes(name))) {
          col[key] = i;
          break;
        }
      }
    });

    const workers: Worker[] = [];

    for (let r = 7; r <= 100; r++) {
      const row = rows[r];
      if (!row) continue;

      const dni = txt(row[col.dni ?? 0]).replace(/\D/g, "");

      if (dni.length !== 8) continue;

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
        sistemaPensionario: txt(row[col.sistemaPensionario ?? 0]),
        cussp: txt(row[col.cussp ?? 0]),
        fechaAfiliacion: txt(row[col.fechaAfiliacion ?? 0]),
        fechaDevengue: txt(row[col.fechaDevengue ?? 0]),
        aporteObligatorio: money(row[col.aporteObligatorio ?? 0]),
        comision: money(row[col.comision ?? 0]),
        primaSeguro: money(row[col.primaSeguro ?? 0]),
        montoMensual: money(row[col.montoMensual ?? 0]),
        descuentoPension: money(row[col.descuentoPension ?? 0]),
        totalDscto: money(row[col.totalDscto ?? 0]),
        totalLiquido: money(row[col.totalLiquido ?? 0]),
      });
    }

    const ok: MsgOut = {
      ok: true,
      workers,
      period: parsePeriodFromFilename(filename)
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