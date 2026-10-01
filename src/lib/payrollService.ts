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
    totalLiquido: String(item.total_liquido ?? "0.00"),
  };
}

function chunkArray<T>(array: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

export async function checkPlanillaExists(anio: string, mes: string, categoriaId: string): Promise<CargaPlanillaItem | null> {
  const { data, error } = await supabase
    .from("cargas_planilla")
    .select("id, anio, mes, categoria_id, nombre_archivo, total_trabajadores, created_at")
    .eq("anio", anio)
    .eq("mes", mes)
    .eq("categoria_id", categoriaId)
    .maybeSingle();

  if (error) {
    console.error("Error al verificar planilla existente:", error);
    return null;
  }

  return data as CargaPlanillaItem | null;
}

export async function savePlanillaToDatabase(params: {
  filename: string;
  categoriaId: string;
  period: { mes: string; anio: string };
  workers: Worker[];
  overwrite?: boolean;
  onProgress?: (step: string, percent: number) => void;
}): Promise<{ ok: boolean; error?: string; totalSaved?: number }> {
  const { filename, categoriaId, period, workers, overwrite, onProgress } = params;

  try {
    onProgress?.("Verificando registros existentes...", 10);
    const existing = await checkPlanillaExists(period.anio, period.mes, categoriaId);

    if (existing && !overwrite) {
      return {
        ok: false,
        error: `Ya existe una planilla registrada para ${categoriaId.toUpperCase()} - ${period.mes} ${period.anio} (${existing.total_trabajadores} trabajadores). ¿Deseas reemplazarla?`,
      };
    }

    if (existing && overwrite) {
      onProgress?.("Eliminando versión anterior...", 20);
      const { error: delError } = await supabase
        .from("cargas_planilla")
        .delete()
        .eq("id", existing.id);

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

  if (error) {
    console.error("Error al obtener historial de cargas:", error);
    return [];
  }

  return (data || []).map((item: any) => ({
    id: item.id,
    anio: item.anio,
    mes: item.mes,
    categoria_id: item.categoria_id,
    nombre_archivo: item.nombre_archivo,
    total_trabajadores: item.total_trabajadores,
    created_at: item.created_at,
    categoria_label: item.categorias?.label ?? item.categoria_id.toUpperCase(),
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

export async function searchAdminBoletas(searchTerm: string): Promise<BoletaHistoricaItem[]> {
  const term = searchTerm.trim();
  if (!term) return [];

  // Búsqueda en trabajadores por DNI, Apellidos o Nombres
  const { data: matchingWorkers, error: workerErr } = await supabase
    .from("trabajadores")
    .select("dni")
    .or(`dni.ilike.%${term}%,ap_paterno.ilike.%${term}%,ap_materno.ilike.%${term}%,nombres.ilike.%${term}%`)
    .limit(30);

  if (workerErr || !matchingWorkers || matchingWorkers.length === 0) {
    return [];
  }

  const dnis = matchingWorkers.map((w) => w.dni);

  // Intentar consultar con boleta_texto_personalizado
  let res = await supabase
    .from("boletas_detalle")
    .select(`
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
        mes,
        anio,
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
    `)
    .in("dni", dnis)
    .order("created_at", { ascending: false })
    .limit(50);

  // Si la columna boleta_texto_personalizado aún no existe en BD, reintentar sin ella
  if (res.error && (res.error.message?.includes("boleta_texto_personalizado") || res.error.code === "42703")) {
    res = await supabase
      .from("boletas_detalle")
      .select(`
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
        cargas_planilla (
          mes,
          anio,
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
      `)
      .in("dni", dnis)
      .order("created_at", { ascending: false })
      .limit(50);
  }

  const { data, error } = res;

  if (error || !data) {
    console.error("Error al buscar boletas:", error);
    return [];
  }

  return data.map((item: any) => ({
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
    categoria_label: item.cargas_planilla?.categorias?.label ?? "CAS",
    boleta_texto_personalizado: item.boleta_texto_personalizado || null,
  }));
}

export async function consultarBoletasTrabajador(
  dni: string,
  claveOCuenta: string
): Promise<{ ok: boolean; boletas?: BoletaHistoricaItem[]; error?: string }> {
  try {
    const { data, error } = await supabase.rpc("consultar_boletas_trabajador", {
      p_dni: dni.trim(),
      p_clave_o_cuenta: claveOCuenta.trim(),
    });

    if (error) {
      return { ok: false, error: error.message };
    }

    if (!data || data.length === 0) {
      return {
        ok: false,
        error: "No se encontraron boletas con esos datos. Verifica tu DNI y que la clave o los últimos 4 dígitos de tu cuenta bancaria sean correctos.",
      };
    }

    const boletas: BoletaHistoricaItem[] = data.map((item: any) => ({
      boleta_id: item.boleta_id,
      n: item.n,
      dni: item.doc_identidad,
      ap_paterno: item.ap_paterno,
      ap_materno: item.ap_materno,
      nombres: item.nombres,
      fecha_nac: item.fecha_nac,
      cargo: item.cargo,
      cod_essalud: item.cod_essalud,
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
      mes: item.mes,
      anio: item.anio,
      categoria_label: item.categoria_label,
      boleta_texto_personalizado: item.boleta_texto_personalizado || null,
    }));

    return { ok: true, boletas };
  } catch (err) {
    return { ok: false, error: "Error de conexión al consultar las boletas." };
  }
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

export async function cambiarClaveTrabajador(
  dni: string,
  ultimos4Cuenta: string,
  nuevaClave: string
): Promise<{ ok: boolean; error?: string }> {
  try {
    const { data, error } = await supabase.rpc("crear_o_cambiar_clave_trabajador", {
      p_dni: dni.trim(),
      p_ultimos4_cuenta: ultimos4Cuenta.trim(),
      p_nueva_clave: nuevaClave.trim(),
    });

    if (error) {
      return { ok: false, error: error.message };
    }

    if (!data) {
      return {
        ok: false,
        error: "Los últimos 4 dígitos de la cuenta bancaria no coinciden con los registrados en la UGEL para este DNI.",
      };
    }

    return { ok: true };
  } catch (err) {
    return { ok: false, error: "Error de red al actualizar la clave." };
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
