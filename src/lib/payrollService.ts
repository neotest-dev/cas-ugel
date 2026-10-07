import { supabase } from "./supabase";
import { Worker } from "./boleta";

export interface CargaPlanillaItem {
  id: string;
  anio: string;
  mes: string;
  categoria_id: string;
  nombre_archivo: string;
  total_trabajadores: number;
  created_at: string;
  categoria_label?: string;
}

export interface BoletaHistoricaItem {
  boleta_id: number;
  n: string;
  dni: string;
  ap_paterno: string;
  ap_materno: string;
  nombres: string;
  fecha_nac: string;
  cargo: string;
  cod_essalud: string;
  cuenta_banco: string;
  leyenda_rd: string;
  leyenda_mensual: string;
  sistema_pensionario: string;
  cussp: string;
  fecha_afiliacion: string;
  fecha_devengue: string;
  monto_mensual: string;
  descuento_pension: string;
  onp: string;
  prima: string;
  integra: string;
  profuturo: string;
  habitat: string;
  aporte_obligatorio: string;
  comision: string;
  prima_seguro: string;
  total_dscto: string;
  otros_dsctos: string;
  dscto_entidades: string;
  dscto_judicial: string;
  total_liquido: string;
  mes: string;
  anio: string;
  categoria_label: string;
  boleta_texto_personalizado?: string | null;
}

export function boletaItemToWorker(item: BoletaHistoricaItem): Worker {
  const parseVal = (v: any) => parseFloat(String(v ?? "0").replace(/,/g, "").trim()) || 0;
  const monto = parseVal(item.monto_mensual);
  const dscto = parseVal(item.total_dscto);
  let liqui = parseVal(item.total_liquido);
  if ((liqui <= 1 && monto > 1) || (monto > 0 && Math.abs(liqui - (monto - dscto)) > 5)) {
    liqui = Math.max(0, monto - dscto);
  }

  return {
    n: item.n || "1",
    dni: item.dni,
    apPaterno: item.ap_paterno,
    apMaterno: item.ap_materno,
    nombres: item.nombres,
    fechaNac: item.fecha_nac || "",
    cargo: item.cargo || "",
    codEssalud: item.cod_essalud || "",
    cuentaBanco: item.cuenta_banco || "",
    leyendaRD: item.leyenda_rd || "",
    leyendaMensual: item.leyenda_mensual || "",
    sistemaPensionario: item.sistema_pensionario || "",
    cussp: item.cussp || "",
    fechaAfiliacion: item.fecha_afiliacion || "",
    fechaDevengue: item.fecha_devengue || "",
    aporteObligatorio: String(item.aporte_obligatorio ?? "0.00"),
    comision: String(item.comision ?? "0.00"),
    primaSeguro: String(item.prima_seguro ?? "0.00"),
    montoMensual: String(item.monto_mensual ?? "0.00"),
    descuentoPension: String(item.descuento_pension ?? "0.00"),
    onp: String(item.onp ?? "0.00"),
    prima: String(item.prima ?? "0.00"),
    integra: String(item.integra ?? "0.00"),
    profuturo: String(item.profuturo ?? "0.00"),
    habitat: String(item.habitat ?? "0.00"),
    totalDscto: String(item.total_dscto ?? "0.00"),
    otrosDsctos: String(item.otros_dsctos ?? "0.00"),
    dsctoEntidades: String(item.dscto_entidades ?? "0.00"),
    dsctoJudicial: String(item.dscto_judicial ?? "0.00"),
    totalLiquido: liqui > 0 ? liqui.toFixed(2) : String(item.total_liquido ?? "0.00"),
  };
}

function chunkArray<T>(array: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

export const CATEGORIA_LABELS_MAP: Record<string, string> = {
  sede: "CAS SEDE",
  jec: "CAS JEC",
  orquestando: "CAS ORQUESTANDO",
  seho: "CAS HOSPITALARIOS",
  hospitalarios: "CAS HOSPITALARIOS",
  ebe: "CAS MEDICA-CEBE",
  "medica-cebe": "CAS MEDICA-CEBE",
  winanq: "CAS WIÑANQ",
  convivencia: "CAS CONVIVENCIA",
  mantenimiento: "CAS MANTENIMIENTO",
};

export function resolveCategoriaLabel(
  categoriaId?: string | null,
  fallbackLabel?: string | null
): string {
  if (!categoriaId) return fallbackLabel || "CAS";
  const normalizedId = categoriaId.toLowerCase().trim();
  if (CATEGORIA_LABELS_MAP[normalizedId]) {
    return CATEGORIA_LABELS_MAP[normalizedId];
  }
  return fallbackLabel || categoriaId.toUpperCase();
}

export async function syncCategoriasInDb(): Promise<void> {
  try {
    await supabase.from("categorias").upsert([
      { id: "sede", label: "CAS SEDE" },
      { id: "jec", label: "CAS JEC" },
      { id: "orquestando", label: "CAS ORQUESTANDO" },
      { id: "seho", label: "CAS HOSPITALARIOS" },
      { id: "ebe", label: "CAS MEDICA-CEBE" },
      { id: "winanq", label: "CAS WIÑANQ" },
      { id: "convivencia", label: "CAS CONVIVENCIA" },
      { id: "mantenimiento", label: "CAS MANTENIMIENTO" },
    ]);
  } catch (err) {
    console.warn("Sincronización de categorías omitida:", err);
  }
}

export interface ExistingFileInfo {
  id: string;
  nombre_archivo: string;
  total_trabajadores: number;
  created_at: string;
}

export interface DuplicatePlanillaInfo {
  period: string;
  categoriaLabel: string;
  count: number;
  totalTrabajadores: number;
  existingFiles: ExistingFileInfo[];
}

export interface SavePlanillaParams {
  filename: string;
  categoriaId: string;
  period: { mes: string; anio: string };
  workers: Worker[];
  overwrite?: boolean;
  saveMode?: "check" | "append" | "replace";
  onProgress?: (step: string, percent: number) => void;
}

export interface SavePlanillaResult {
  ok: boolean;
  error?: string;
  isDuplicate?: boolean;
  duplicateInfo?: DuplicatePlanillaInfo;
  totalSaved?: number;
}

export async function getExistingPlanillas(
  anio: string,
  mes: string,
  categoriaId: string
): Promise<CargaPlanillaItem[]> {
  const { data, error } = await supabase
    .from("cargas_planilla")
    .select("id, anio, mes, categoria_id, nombre_archivo, total_trabajadores, created_at")
    .eq("anio", anio)
    .eq("mes", mes)
    .eq("categoria_id", categoriaId)
    .order("created_at", { ascending: false });

  if (error) {
    console.error("Error al verificar planillas existentes:", error);
    return [];
  }

  return (data || []).map((item) => ({
    id: String(item.id),
    anio: String(item.anio),
    mes: String(item.mes),
    categoria_id: String(item.categoria_id),
    nombre_archivo: String(item.nombre_archivo),
    total_trabajadores: Number(item.total_trabajadores) || 0,
    created_at: String(item.created_at),
    categoria_label: resolveCategoriaLabel(item.categoria_id),
  }));
}

export async function checkPlanillaExists(
  anio: string,
  mes: string,
  categoriaId: string
): Promise<CargaPlanillaItem | null> {
  const list = await getExistingPlanillas(anio, mes, categoriaId);
  return list.length > 0 ? list[0] : null;
}

export async function savePlanillaToDatabase(
  params: SavePlanillaParams
): Promise<SavePlanillaResult> {
  const {
    filename,
    categoriaId,
    period,
    workers,
    overwrite,
    saveMode = "check",
    onProgress,
  } = params;

  try {
    // Sincronizar nombres oficiales en la tabla de categorías
    syncCategoriasInDb().catch(() => {});

    onProgress?.("Verificando registros existentes...", 10);
    const existingList = await getExistingPlanillas(period.anio, period.mes, categoriaId);

    const isReplace = saveMode === "replace" || overwrite === true;
    const isAppend = saveMode === "append";

    if (existingList.length > 0 && !isReplace && !isAppend) {
      const totalPrevWorkers = existingList.reduce(
        (acc, cur) => acc + (cur.total_trabajadores || 0),
        0
      );
      const catLabel = resolveCategoriaLabel(categoriaId);

      return {
        ok: false,
        isDuplicate: true,
        error: "DUPLICATE_PERIOD",
        duplicateInfo: {
          period: `${period.mes} ${period.anio}`,
          categoriaLabel: catLabel,
          count: existingList.length,
          totalTrabajadores: totalPrevWorkers,
          existingFiles: existingList.map((e) => ({
            id: e.id,
            nombre_archivo: e.nombre_archivo,
            total_trabajadores: e.total_trabajadores,
            created_at: e.created_at,
          })),
        },
      };
    }

    if (existingList.length > 0 && isReplace) {
      onProgress?.("Eliminando versión previa...", 20);
      const matchSameName = existingList.find(
        (e) => e.nombre_archivo.trim().toLowerCase() === filename.trim().toLowerCase()
      );
      const targetToDelete = matchSameName || existingList[0];

      const { error: delError } = await supabase
        .from("cargas_planilla")
        .delete()
        .eq("id", targetToDelete.id);

      if (delError) {
        throw new Error(`Error al reemplazar la planilla previa: ${delError.message}`);
      }
    }

    onProgress?.("Registrando cabecera de planilla...", 30);
    const { data: carga, error: cargaError } = await supabase
      .from("cargas_planilla")
      .insert({
        anio: period.anio,
        mes: period.mes,
        categoria_id: categoriaId,
        nombre_archivo: filename,
        total_trabajadores: workers.length,
      })
      .select("id")
      .single();

    if (cargaError || !carga) {
      throw new Error(`Error al crear el registro de carga: ${cargaError?.message}`);
    }

    const cargaId = carga.id;

    // 1. Guardar/actualizar trabajadores (maestro)
    onProgress?.("Actualizando padrón de trabajadores...", 45);
    const trabajadoresPayload = workers.map((w) => {
      const cleanDigits = (w.cuentaBanco || "").replace(/\D/g, "");
      const ultimos4 = cleanDigits.length >= 4 ? cleanDigits.slice(-4) : cleanDigits || null;

      return {
        dni: w.dni,
        ap_paterno: w.apPaterno,
        ap_materno: w.apMaterno,
        nombres: w.nombres,
        fecha_nac: w.fechaNac || null,
        cod_essalud: w.codEssalud || null,
        cuenta_banco_ultimos4: ultimos4,
        updated_at: new Date().toISOString(),
      };
    });

    // Chunking de 100 en 100 para evitar límites de payload
    const workerChunks = chunkArray(trabajadoresPayload, 100);
    for (let i = 0; i < workerChunks.length; i++) {
      const { error: upsertErr } = await supabase
        .from("trabajadores")
        .upsert(workerChunks[i], {
          onConflict: "dni",
          ignoreDuplicates: false,
        });

      if (upsertErr) {
        console.error("Error al actualizar trabajadores chunk:", upsertErr);
      }
    }

    // 2. Guardar boletas detalle
    onProgress?.("Guardando boletas detalladas...", 65);
    const boletasPayload = workers.map((w) => ({
      carga_id: cargaId,
      dni: w.dni,
      n: w.n,
      cargo: w.cargo,
      cuenta_banco: w.cuentaBanco,
      leyenda_rd: w.leyendaRD,
      leyenda_mensual: w.leyendaMensual,
      sistema_pensionario: w.sistemaPensionario,
      cussp: w.cussp,
      fecha_afiliacion: w.fechaAfiliacion,
      fecha_devengue: w.fechaDevengue,
      monto_mensual: Number(w.montoMensual) || 0,
      descuento_pension: Number(w.descuentoPension) || 0,
      onp: Number(w.onp) || 0,
      prima: Number(w.prima) || 0,
      integra: Number(w.integra) || 0,
      profuturo: Number(w.profuturo) || 0,
      habitat: Number(w.habitat) || 0,
      aporte_obligatorio: Number(w.aporteObligatorio) || 0,
      comision: Number(w.comision) || 0,
      prima_seguro: Number(w.primaSeguro) || 0,
      total_dscto: Number(w.totalDscto) || 0,
      otros_dsctos: Number(w.otrosDsctos) || 0,
      dscto_entidades: Number(w.dsctoEntidades) || 0,
      dscto_judicial: Number(w.dsctoJudicial) || 0,
      total_liquido: Number(w.totalLiquido) || 0,
    }));

    const boletaChunks = chunkArray(boletasPayload, 100);
    for (let i = 0; i < boletaChunks.length; i++) {
      const progress = 65 + Math.round(((i + 1) / boletaChunks.length) * 30);
      onProgress?.(`Guardando boletas (${i * 100 + boletaChunks[i].length} / ${workers.length})...`, progress);

      const { error: boletaErr } = await supabase
        .from("boletas_detalle")
        .insert(boletaChunks[i]);

      if (boletaErr) {
        throw new Error(`Error al insertar lote de boletas: ${boletaErr.message}`);
      }
    }

    onProgress?.("Planilla guardada exitosamente.", 100);
    return { ok: true, totalSaved: workers.length };
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Error inesperado al guardar";
    console.error("Error en savePlanillaToDatabase:", err);
    return { ok: false, error: msg };
  }
}

export async function fetchCargasPlanilla(): Promise<CargaPlanillaItem[]> {
  const { data, error } = await supabase
    .from("cargas_planilla")
    .select(`
      id,
      anio,
      mes,
      categoria_id,
      nombre_archivo,
      total_trabajadores,
      created_at,
      categorias ( label )
    `)
    .order("created_at", { ascending: false });

  syncCategoriasInDb().catch(() => {});

  if (error) {
    console.error("Error al obtener historial de cargas:", error);
    return [];
  }

  return (data || []).map((item) => ({
    id: item.id,
    anio: item.anio,
    mes: item.mes,
    categoria_id: item.categoria_id,
    nombre_archivo: item.nombre_archivo,
    total_trabajadores: item.total_trabajadores,
    created_at: item.created_at,
    categoria_label: resolveCategoriaLabel(item.categoria_id, item.categorias?.label),
  }));
}

export async function deleteCargaPlanilla(id: string): Promise<{ ok: boolean; error?: string }> {
  const { error } = await supabase
    .from("cargas_planilla")
    .delete()
    .eq("id", id);

  if (error) {
    return { ok: false, error: error.message };
  }

  return { ok: true };
}

export async function deleteCargasPlanilla(ids: string[]): Promise<{ ok: boolean; error?: string }> {
  if (!ids.length) return { ok: true };
  const { error } = await supabase
    .from("cargas_planilla")
    .delete()
    .in("id", ids);

  if (error) {
    return { ok: false, error: error.message };
  }

  return { ok: true };
}

export interface AdminBoletaSearchFilters {
  searchTerm?: string;
  categoriaId?: string;
  desde?: string;
  hasta?: string;
}

const MESES_NUMERO: Record<string, string> = {
  ENERO: "01", FEBRERO: "02", MARZO: "03", ABRIL: "04", MAYO: "05", JUNIO: "06",
  JULIO: "07", AGOSTO: "08", SEPTIEMBRE: "09", SETIEMBRE: "09", OCTUBRE: "10",
  NOVIEMBRE: "11", DICIEMBRE: "12",
};

const BOLETA_SEARCH_SELECT = `
       id,
      n,
      dni,
      cargo,
      cuenta_banco,
      leyenda_rd,
      leyenda_mensual,
      sistema_pensionario,
      cussp,
      fecha_afiliacion,
      fecha_devengue,
      monto_mensual,
      descuento_pension,
      onp,
      prima,
      integra,
      profuturo,
      habitat,
      aporte_obligatorio,
      comision,
      prima_seguro,
      total_dscto,
      otros_dsctos,
      dscto_entidades,
      dscto_judicial,
      total_liquido,
       boleta_texto_personalizado,
       cargas_planilla (
         id,
         mes,
         anio,
         categoria_id,
         categorias ( label )
      ),
      trabajadores (
        dni,
        ap_paterno,
        ap_materno,
        nombres,
        fecha_nac,
        cod_essalud
      )
     `;

interface AdminBoletaDbRow {
  id: number;
  n: string;
  dni: string;
  cargo: string;
  cuenta_banco: string;
  leyenda_rd: string;
  leyenda_mensual: string;
  sistema_pensionario: string;
  cussp: string;
  fecha_afiliacion: string;
  fecha_devengue: string;
  monto_mensual: string | number | null;
  descuento_pension: string | number | null;
  onp: string | number | null;
  prima: string | number | null;
  integra: string | number | null;
  profuturo: string | number | null;
  habitat: string | number | null;
  aporte_obligatorio: string | number | null;
  comision: string | number | null;
  prima_seguro: string | number | null;
  total_dscto: string | number | null;
  otros_dsctos: string | number | null;
  dscto_entidades: string | number | null;
  dscto_judicial: string | number | null;
  total_liquido: string | number | null;
  boleta_texto_personalizado?: string | null;
  cargas_planilla?: {
    mes?: string | null;
    anio?: string | null;
    categoria_id?: string | null;
    categorias?: { label?: string | null } | null;
  } | null;
  trabajadores?: {
    ap_paterno?: string | null;
    ap_materno?: string | null;
    nombres?: string | null;
    fecha_nac?: string | null;
    cod_essalud?: string | null;
  } | null;
}

/**
 * Finds worker DNIs matching every word, returning a relevance score per DNI.
 * Preferred path: SQL function `buscar_trabajadores` (accent-insensitive and ranked).
 * Fallback when the function is not installed: plain ILIKE without ranking.
 */
async function findWorkersByTokens(tokens: string[]): Promise<Map<string, number> | null> {
  const pageSize = 1000;
  const scores = new Map<string, number>();

  let rpcFailed = false;
  for (let from = 0; ; from += pageSize) {
    const { data, error } = await supabase
      .rpc("buscar_trabajadores", { p_term: tokens.join(" ") })
      .range(from, from + pageSize - 1);
    if (error) {
      rpcFailed = true;
      break;
    }
    for (const row of (data ?? []) as Array<{ dni: string; score: number }>) {
      scores.set(row.dni, row.score);
    }
    if (!data || data.length < pageSize) break;
  }
  if (!rpcFailed) return scores;
  scores.clear();

  for (let from = 0; ; from += pageSize) {
    let query = supabase.from("trabajadores").select("dni");
    for (const token of tokens) {
      query = query.or(
        `dni.ilike.%${token}%,ap_paterno.ilike.%${token}%,ap_materno.ilike.%${token}%,nombres.ilike.%${token}%`
      );
    }
    const { data, error } = await query.range(from, from + pageSize - 1);
    if (error) {
      console.error("Error al buscar trabajadores:", error);
      return null;
    }
    for (const row of data ?? []) scores.set(row.dni, 0);
    if (!data || data.length < pageSize) break;
  }
  return scores;
}

export async function searchAdminBoletas(
  filters: AdminBoletaSearchFilters
): Promise<BoletaHistoricaItem[]> {
  const term = filters.searchTerm?.trim() ?? "";
  if (!term && !filters.categoriaId && !filters.desde && !filters.hasta) return [];

  // Each word must match at least one of: DNI, paternal surname, maternal surname, names.
  // Words may appear in any order (e.g. "SARITA FLORIAN" finds "SARITA JUDITH DIAZ FLORIAN").
  const tokens = term
    .split(/\s+/)
    .map((token) => token.replace(/[,()%*\\"'_]/g, ""))
    .filter(Boolean);
  if (term && !tokens.length) return [];

  let dnis: string[] | null = null;
  let scoreByDni = new Map<string, number>();
  if (tokens.length) {
    const matches = await findWorkersByTokens(tokens);
    if (!matches || !matches.size) return [];
    scoreByDni = matches;
    dnis = [...matches.keys()];
  }

  let cargaIds: string[] | null = null;
  if (filters.categoriaId || filters.desde || filters.hasta) {
    let cargasQuery = supabase.from("cargas_planilla").select("id, mes, anio, categoria_id");
    if (filters.categoriaId) cargasQuery = cargasQuery.eq("categoria_id", filters.categoriaId);
    const yearFrom = filters.desde?.slice(0, 4);
    const yearTo = filters.hasta?.slice(0, 4);
    if (yearFrom) cargasQuery = cargasQuery.gte("anio", yearFrom);
    if (yearTo) cargasQuery = cargasQuery.lte("anio", yearTo);

    const { data: cargas, error: cargasError } = await cargasQuery;
    if (cargasError) {
      console.error("Error al filtrar periodos de planilla:", cargasError);
      return [];
    }

    cargaIds = (cargas ?? [])
      .filter((carga) => {
        const month = MESES_NUMERO[String(carga.mes ?? "").trim().toUpperCase()];
        if (!month) return false;
        const period = `${carga.anio}-${month}`;
        const from = filters.desde?.slice(0, 7);
        const to = filters.hasta?.slice(0, 7);
        return (!from || period >= from) && (!to || period <= to);
      })
      .map((carga) => String(carga.id));
    if (!cargaIds.length) return [];
  }

  const dnisBatches = dnis ? chunkArray(dnis, 100) : [null];
  const cargaBatches = cargaIds ? chunkArray(cargaIds, 100) : [null];
  const rows: AdminBoletaDbRow[] = [];
  const pageSize = 1000;

  for (const dniBatch of dnisBatches) {
    for (const cargaBatch of cargaBatches) {
      for (let from = 0; ; from += pageSize) {
        let query = supabase
          .from("boletas_detalle")
          .select(BOLETA_SEARCH_SELECT)
          .order("created_at", { ascending: false })
          .range(from, from + pageSize - 1);
        if (dniBatch) query = query.in("dni", dniBatch);
        if (cargaBatch) query = query.in("carga_id", cargaBatch);

        let result = await query;
        if (result.error && (result.error.message?.includes("boleta_texto_personalizado") || result.error.code === "42703")) {
          // Reaplicar los filtros en el reintento cuando el esquema aún no tenga la columna opcional.
          let fallback = supabase
            .from("boletas_detalle")
            .select(BOLETA_SEARCH_SELECT.replace("boleta_texto_personalizado,", ""))
            .order("created_at", { ascending: false })
            .range(from, from + pageSize - 1);
          if (dniBatch) fallback = fallback.in("dni", dniBatch);
          if (cargaBatch) fallback = fallback.in("carga_id", cargaBatch);
          result = await fallback;
        }
        if (result.error) {
          console.error("Error al buscar boletas:", result.error);
          return [];
        }
        rows.push(...((result.data ?? []) as unknown as AdminBoletaDbRow[]));
        if (!result.data || result.data.length < pageSize) break;
      }
    }
  }

  const uniqueRows = [...new Map(rows.map((row) => [row.id, row])).values()];
  return uniqueRows.map((item) => ({
    boleta_id: item.id,
    n: item.n,
    dni: item.dni,
    ap_paterno: item.trabajadores?.ap_paterno ?? "",
    ap_materno: item.trabajadores?.ap_materno ?? "",
    nombres: item.trabajadores?.nombres ?? "",
    fecha_nac: item.trabajadores?.fecha_nac ?? "",
    cargo: item.cargo,
    cod_essalud: item.trabajadores?.cod_essalud ?? "",
    cuenta_banco: item.cuenta_banco,
    leyenda_rd: item.leyenda_rd,
    leyenda_mensual: item.leyenda_mensual,
    sistema_pensionario: item.sistema_pensionario,
    cussp: item.cussp,
    fecha_afiliacion: item.fecha_afiliacion,
    fecha_devengue: item.fecha_devengue,
    monto_mensual: String(item.monto_mensual ?? "0.00"),
    descuento_pension: String(item.descuento_pension ?? "0.00"),
    onp: String(item.onp ?? "0.00"),
    prima: String(item.prima ?? "0.00"),
    integra: String(item.integra ?? "0.00"),
    profuturo: String(item.profuturo ?? "0.00"),
    habitat: String(item.habitat ?? "0.00"),
    aporte_obligatorio: String(item.aporte_obligatorio ?? "0.00"),
    comision: String(item.comision ?? "0.00"),
    prima_seguro: String(item.prima_seguro ?? "0.00"),
    total_dscto: String(item.total_dscto ?? "0.00"),
    otros_dsctos: String(item.otros_dsctos ?? "0.00"),
    dscto_entidades: String(item.dscto_entidades ?? "0.00"),
    dscto_judicial: String(item.dscto_judicial ?? "0.00"),
    total_liquido: String(item.total_liquido ?? "0.00"),
    mes: item.cargas_planilla?.mes ?? "",
    anio: item.cargas_planilla?.anio ?? "",
    categoria_label: resolveCategoriaLabel(
      item.cargas_planilla?.categoria_id,
      item.cargas_planilla?.categorias?.label
    ),
    boleta_texto_personalizado: item.boleta_texto_personalizado || null,
  })).sort((a, b) => {
    const scoreDiff = (scoreByDni.get(b.dni) ?? 0) - (scoreByDni.get(a.dni) ?? 0);
    if (scoreDiff) return scoreDiff;
    const periodA = `${a.anio}-${MESES_NUMERO[a.mes.toUpperCase()] ?? "00"}`;
    const periodB = `${b.anio}-${MESES_NUMERO[b.mes.toUpperCase()] ?? "00"}`;
    const periodDiff = periodB.localeCompare(periodA);
    if (periodDiff) return periodDiff;
    const nameA = `${a.ap_paterno} ${a.ap_materno} ${a.nombres}`.trim();
    const nameB = `${b.ap_paterno} ${b.ap_materno} ${b.nombres}`.trim();
    return nameA.localeCompare(nameB, "es", { sensitivity: "base" });
  });
}

/**
 * Saves the first known version of a boleta so it can be restored later.
 * Uses ignoreDuplicates, so only the very first snapshot is kept.
 */
export async function saveBoletaOriginal(item: BoletaHistoricaItem): Promise<void> {
  const { error } = await supabase
    .from("boletas_originales")
    .upsert({ boleta_id: item.boleta_id, datos: item }, { onConflict: "boleta_id", ignoreDuplicates: true });
  if (error) console.warn("No se pudo guardar el respaldo original de la boleta:", error.message);
}

export async function fetchBoletaOriginal(
  boletaId: number
): Promise<{ ok: boolean; item?: BoletaHistoricaItem | null; error?: string }> {
  const { data, error } = await supabase
    .from("boletas_originales")
    .select("datos")
    .eq("boleta_id", boletaId)
    .maybeSingle();
  if (error) return { ok: false, error: error.message };
  return { ok: true, item: (data?.datos as BoletaHistoricaItem | undefined) ?? null };
}

export async function guardarBoletaTexto(
  boletaId: number,
  texto: string | null
): Promise<{ ok: boolean; error?: string }> {
  try {
    // 1. Intentar primero mediante función RPC segura
    const { error: rpcError } = await supabase.rpc("guardar_edicion_boleta", {
      p_boleta_id: boletaId,
      p_texto: texto,
    });

    if (!rpcError) {
      return { ok: true };
    }

    // 2. Si el RPC no existe, intentar UPDATE directo en la tabla
    const { error: updateError } = await supabase
      .from("boletas_detalle")
      .update({ boleta_texto_personalizado: texto })
      .eq("id", boletaId);

    if (updateError) {
      if (
        updateError.message?.includes("boleta_texto_personalizado") ||
        updateError.code === "42703"
      ) {
        return {
          ok: false,
          error:
            "Falta crear la columna en Supabase. Ejecuta el script SQL en el Editor SQL de Supabase.",
        };
      }
      return { ok: false, error: updateError.message };
    }

    return { ok: true };
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Error inesperado al guardar";
    return { ok: false, error: msg };
  }
}

export interface BoletaFormData {
  boleta_id: number;
  dni: string;
  ap_paterno: string;
  ap_materno: string;
  nombres: string;
  fecha_nac: string;
  cargo: string;
  cod_essalud: string;
  cuenta_banco: string;
  leyenda_rd: string;
  leyenda_mensual: string;
  sistema_pensionario: string;
  cussp: string;
  fecha_afiliacion: string;
  fecha_devengue: string;
  monto_mensual: string;
  onp: string;
  prima: string;
  integra: string;
  profuturo: string;
  habitat: string;
  total_dscto: string;
  otros_dsctos: string;
  dscto_entidades: string;
  dscto_judicial: string;
  total_liquido: string;
}

export async function actualizarDatosBoletaYTrabajador(
  data: BoletaFormData
): Promise<{ ok: boolean; error?: string }> {
  try {
    // 1. Intentar mediante función RPC segura en PostgreSQL
    const { error: rpcError } = await supabase.rpc("actualizar_boleta_y_trabajador", {
      p_boleta_id: data.boleta_id,
      p_dni: data.dni.trim(),
      p_ap_paterno: data.ap_paterno.trim().toUpperCase(),
      p_ap_materno: data.ap_materno.trim().toUpperCase(),
      p_nombres: data.nombres.trim().toUpperCase(),
      p_fecha_nac: data.fecha_nac?.trim() || null,
      p_cod_essalud: data.cod_essalud?.trim() || null,
      p_cargo: data.cargo.trim().toUpperCase(),
      p_cuenta_banco: data.cuenta_banco.trim(),
      p_leyenda_rd: data.leyenda_rd.trim(),
      p_leyenda_mensual: data.leyenda_mensual.trim(),
      p_sistema_pensionario: data.sistema_pensionario.trim().toUpperCase(),
      p_cussp: data.cussp.trim(),
      p_fecha_afiliacion: data.fecha_afiliacion.trim(),
      p_fecha_devengue: data.fecha_devengue.trim(),
      p_monto_mensual: Number(data.monto_mensual) || 0,
      p_onp: Number(data.onp) || 0,
      p_prima: Number(data.prima) || 0,
      p_integra: Number(data.integra) || 0,
      p_profuturo: Number(data.profuturo) || 0,
      p_habitat: Number(data.habitat) || 0,
      p_total_dscto: Number(data.total_dscto) || 0,
      p_otros_dsctos: Number(data.otros_dsctos) || 0,
      p_dscto_entidades: Number(data.dscto_entidades) || 0,
      p_dscto_judicial: Number(data.dscto_judicial) || 0,
      p_total_liquido: Number(data.total_liquido) || 0,
    });

    if (!rpcError) {
      return { ok: true };
    }

    // 2. Si la función RPC aún no está creada, ejecutar UPDATE directo
    const cleanDigits = (data.cuenta_banco || "").replace(/\D/g, "");
    const ultimos4 = cleanDigits.length >= 4 ? cleanDigits.slice(-4) : cleanDigits || null;

    const { error: trabErr } = await supabase
      .from("trabajadores")
      .update({
        ap_paterno: data.ap_paterno.trim().toUpperCase(),
        ap_materno: data.ap_materno.trim().toUpperCase(),
        nombres: data.nombres.trim().toUpperCase(),
        fecha_nac: data.fecha_nac?.trim() || null,
        cod_essalud: data.cod_essalud?.trim() || null,
        cuenta_banco_ultimos4: ultimos4,
        updated_at: new Date().toISOString(),
      })
      .eq("dni", data.dni.trim());

    if (trabErr) {
      console.warn("Advertencia al actualizar tabla trabajadores:", trabErr);
    }

    const { error: bolErr } = await supabase
      .from("boletas_detalle")
      .update({
        cargo: data.cargo.trim().toUpperCase(),
        cuenta_banco: data.cuenta_banco.trim(),
        leyenda_rd: data.leyenda_rd.trim(),
        leyenda_mensual: data.leyenda_mensual.trim(),
        sistema_pensionario: data.sistema_pensionario.trim().toUpperCase(),
        cussp: data.cussp.trim(),
        fecha_afiliacion: data.fecha_afiliacion.trim(),
        fecha_devengue: data.fecha_devengue.trim(),
        monto_mensual: Number(data.monto_mensual) || 0,
        onp: Number(data.onp) || 0,
        prima: Number(data.prima) || 0,
        integra: Number(data.integra) || 0,
        profuturo: Number(data.profuturo) || 0,
        habitat: Number(data.habitat) || 0,
        total_dscto: Number(data.total_dscto) || 0,
        otros_dsctos: Number(data.otros_dsctos) || 0,
        dscto_entidades: Number(data.dscto_entidades) || 0,
        dscto_judicial: Number(data.dscto_judicial) || 0,
        total_liquido: Number(data.total_liquido) || 0,
        boleta_texto_personalizado: null,
      })
      .eq("id", data.boleta_id);

    if (bolErr) {
      return { ok: false, error: bolErr.message };
    }

    return { ok: true };
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Error inesperado al guardar datos";
    return { ok: false, error: msg };
  }
}
