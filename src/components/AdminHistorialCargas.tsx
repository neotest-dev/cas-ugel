import { useState, useEffect, useMemo, useRef } from "react";
import {
  fetchCargasPlanilla,
  CargaPlanillaItem,
  resolveCategoriaLabel,
  deleteCargaPlanilla,
} from "@/lib/payrollService";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import {
  History,
  RefreshCw,
  FileSpreadsheet,
  Users,
  Calendar,
  Loader2,
  ChevronDown,
  ChevronRight,
  Trash2,
  ExternalLink,
  Search,
  Eye,
  ChevronsDown,
  ChevronsUp,
} from "lucide-react";
import { toast } from "@/hooks/use-toast";

const MESES_ORDEN = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
];

const MESES_CORTOS: Record<string, string> = {
  ENERO: "ENE", FEBRERO: "FEB", MARZO: "MAR", ABRIL: "ABR",
  MAYO: "MAY", JUNIO: "JUN", JULIO: "JUL", AGOSTO: "AGO",
  SEPTIEMBRE: "SET", SETIEMBRE: "SET", OCTUBRE: "OCT",
  NOVIEMBRE: "NOV", DICIEMBRE: "DIC",
};

function normalizeMes(mes: string): string {
  const m = String(mes || "").trim().toUpperCase();
  if (m === "SETIEMBRE") return "SEPTIEMBRE";
  return m;
}

interface MonthData {
  mes: string;
  cargas: CargaPlanillaItem[];
  totalBoletas: number;
}

export interface HistorialNavigateParams {
  categoriaId: string;
  mes: string;
  anio: string;
}

interface AdminHistorialCargasProps {
  onNavigateToConsulta?: (params: HistorialNavigateParams) => void;
}

const CATEGORY_COLORS: Record<string, string> = {
  sede: "bg-blue-50 text-blue-800 border-blue-200 hover:bg-blue-100",
  jec: "bg-violet-50 text-violet-800 border-violet-200 hover:bg-violet-100",
  orquestando: "bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100",
  seho: "bg-rose-50 text-rose-800 border-rose-200 hover:bg-rose-100",
  ebe: "bg-teal-50 text-teal-800 border-teal-200 hover:bg-teal-100",
  winanq: "bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100",
  convivencia: "bg-orange-50 text-orange-800 border-orange-200 hover:bg-orange-100",
  mantenimiento: "bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200",
};

/**
 * Genera la URL con parámetros para abrir en una nueva pestaña (target="_blank")
 */
function buildConsultaUrl(params: HistorialNavigateParams): string {
  const searchParams = new URLSearchParams();
  searchParams.set("tab", "admin");
  searchParams.set("subtab", "ventanilla");
  if (params.categoriaId) searchParams.set("categoria", params.categoriaId);
  if (params.mes) searchParams.set("mes", params.mes);
  if (params.anio) searchParams.set("anio", params.anio);
  return `/?${searchParams.toString()}`;
}

export function AdminHistorialCargas({ onNavigateToConsulta }: AdminHistorialCargasProps) {
  const [cargas, setCargas] = useState<CargaPlanillaItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [selectedYear, setSelectedYear] = useState("");
  // Estado de meses expandidos (permite abrir y cerrar individualmente)
  const [expandedMonths, setExpandedMonths] = useState<Record<string, boolean>>({});
  const initialOpenDoneRef = useRef(false);

  const [deleteTarget, setDeleteTarget] = useState<CargaPlanillaItem | null>(null);
  const [deleting, setDeleting] = useState(false);

  const loadData = async () => {
    setLoading(true);
    try {
      const data = await fetchCargasPlanilla();
      setCargas(data);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  // Años disponibles (incluyendo año actual por defecto)
  const availableYears = useMemo(() => {
    const yearsSet = new Set<string>();
    for (const c of cargas) {
      if (c.anio) yearsSet.add(String(c.anio).trim());
    }
    const currentYear = String(new Date().getFullYear());
    yearsSet.add(currentYear);
    return [...yearsSet].sort((a, b) => b.localeCompare(a));
  }, [cargas]);

  // Auto-seleccionar año más reciente
  useEffect(() => {
    if (!selectedYear && availableYears.length > 0) {
      setSelectedYear(availableYears[0]);
    }
  }, [availableYears, selectedYear]);

  // Agrupación por mes del año seleccionado
  const monthsData = useMemo((): MonthData[] => {
    if (!selectedYear) return [];

    const yearCargas = cargas.filter((c) => String(c.anio).trim() === String(selectedYear).trim());
    const mesMap = new Map<string, CargaPlanillaItem[]>();

    for (const carga of yearCargas) {
      const mesNorm = normalizeMes(carga.mes);
      if (!mesMap.has(mesNorm)) mesMap.set(mesNorm, []);
      mesMap.get(mesNorm)!.push(carga);
    }

    const knownMonths = MESES_ORDEN.filter((m) => mesMap.has(m));
    const extraMonths = [...mesMap.keys()].filter((m) => !MESES_ORDEN.includes(m));
    const allPresentMonths = [...knownMonths, ...extraMonths];

    return allPresentMonths
      .map((mes) => {
        const cargasMes = mesMap.get(mes)!;
        return {
          mes,
          cargas: cargasMes.sort(
            (a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime()
          ),
          totalBoletas: cargasMes.reduce((sum, c) => sum + (c.total_trabajadores || 0), 0),
        };
      })
      .reverse(); // Más reciente primero
  }, [cargas, selectedYear]);

  // En la carga inicial de datos, abrir solo el mes más reciente una sola vez
  useEffect(() => {
    if (!initialOpenDoneRef.current && monthsData.length > 0) {
      setExpandedMonths({ [monthsData[0].mes]: true });
      initialOpenDoneRef.current = true;
    }
  }, [monthsData]);

  // Alternar apertura/cierre de un mes específico (dropdown)
  const toggleMonth = (mes: string) => {
    setExpandedMonths((prev) => ({
      ...prev,
      [mes]: !prev[mes],
    }));
  };

  const expandAll = () => {
    const all: Record<string, boolean> = {};
    for (const m of monthsData) all[m.mes] = true;
    setExpandedMonths(all);
  };

  const collapseAll = () => {
    setExpandedMonths({});
  };

  // Estadísticas del año seleccionado
  const yearStats = useMemo(() => {
    const yearCargas = cargas.filter((c) => String(c.anio).trim() === String(selectedYear).trim());
    return {
      totalCargas: yearCargas.length,
      totalBoletas: yearCargas.reduce((sum, c) => sum + (c.total_trabajadores || 0), 0),
      mesesConDatos: new Set(yearCargas.map((c) => normalizeMes(c.mes))).size,
    };
  }, [cargas, selectedYear]);

  // Mini calendario de 12 meses
  const calendarMonthsStatus = useMemo(() => {
    const yearCargas = cargas.filter((c) => String(c.anio).trim() === String(selectedYear).trim());
    const map = new Map<string, { totalBoletas: number; totalCargas: number }>();
    for (const c of yearCargas) {
      const mesNorm = normalizeMes(c.mes);
      const curr = map.get(mesNorm) || { totalBoletas: 0, totalCargas: 0 };
      curr.totalBoletas += c.total_trabajadores || 0;
      curr.totalCargas += 1;
      map.set(mesNorm, curr);
    }
    return MESES_ORDEN.map((mes) => ({
      mes,
      corto: MESES_CORTOS[mes] || mes.slice(0, 3),
      hasData: map.has(mes),
      totalBoletas: map.get(mes)?.totalBoletas || 0,
      totalCargas: map.get(mes)?.totalCargas || 0,
    }));
  }, [cargas, selectedYear]);

  // Categorías por mes
  const getCategoriesForMonth = (cargas: CargaPlanillaItem[]) => {
    const catMap = new Map<string, { id: string; label: string; totalBoletas: number; count: number }>();
    for (const c of cargas) {
      const existing = catMap.get(c.categoria_id);
      if (existing) {
        existing.totalBoletas += c.total_trabajadores || 0;
        existing.count += 1;
      } else {
        catMap.set(c.categoria_id, {
          id: c.categoria_id,
          label: resolveCategoriaLabel(c.categoria_id, c.categoria_label),
          totalBoletas: c.total_trabajadores || 0,
          count: 1,
        });
      }
    }
    return [...catMap.values()];
  };

  const handleDelete = async () => {
    if (!deleteTarget) return;
    setDeleting(true);
    try {
      const res = await deleteCargaPlanilla(deleteTarget.id);
      if (res.ok) {
        toast({
          title: "Planilla eliminada",
          description: `Se eliminó "${deleteTarget.nombre_archivo}" correctamente.`,
        });
        await loadData();
      } else {
        toast({
          title: "Error al eliminar",
          description: res.error || "No se pudo eliminar la carga de planilla.",
          variant: "destructive",
        });
      }
    } finally {
      setDeleting(false);
      setDeleteTarget(null);
    }
  };

  return (
    <div className="space-y-3">
      <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        {/* 1. Encabezado Principal */}
        <div className="bg-[#0b223d] text-white px-4 py-2.5 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <History className="h-4 w-4 text-amber-400" />
            <h3 className="text-xs font-bold uppercase tracking-wider">
              Historial de Cargas de Planilla
            </h3>
          </div>

          <div className="flex items-center gap-2">
            {availableYears.length > 1 && (
              <Select value={selectedYear} onValueChange={setSelectedYear}>
                <SelectTrigger className="h-7 w-28 bg-[#153457] border-[#31577e] text-white text-xs font-bold rounded">
                  <Calendar className="h-3 w-3 mr-1 text-amber-300" />
                  <SelectValue placeholder="Año" />
                </SelectTrigger>
                <SelectContent>
                  {availableYears.map((year) => (
                    <SelectItem key={year} value={year} className="text-xs font-semibold">
                      {year}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}

            <Button
              variant="outline"
              size="sm"
              onClick={loadData}
              disabled={loading}
              className="h-7 text-xs bg-white text-slate-800 border-slate-300 hover:bg-slate-100 rounded px-2.5 font-semibold"
            >
              <RefreshCw className={`h-3 w-3 mr-1 ${loading ? "animate-spin" : ""}`} />
              Actualizar
            </Button>
          </div>
        </div>

        {/* 2. Barra de Selección Rápida de Año */}
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2">
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Calendar className="h-3.5 w-3.5 text-blue-600" />
              Año:
            </span>
            <div className="flex flex-wrap gap-1">
              {availableYears.map((year) => {
                const count = cargas.filter((c) => String(c.anio).trim() === year).length;
                const isSelected = selectedYear === year;
                return (
                  <button
                    key={year}
                    type="button"
                    onClick={() => setSelectedYear(year)}
                    className={`inline-flex items-center gap-1.5 px-3 py-1 rounded text-xs font-bold transition-all ${
                      isSelected
                        ? "bg-[#0b223d] text-white shadow-sm ring-2 ring-blue-500/20"
                        : "bg-white text-slate-700 border border-slate-300 hover:bg-slate-100"
                    }`}
                  >
                    <span>{year}</span>
                    {count > 0 && (
                      <span
                        className={`text-[10px] px-1.5 py-0.2 rounded-full font-semibold ${
                          isSelected ? "bg-white/20 text-white" : "bg-slate-100 text-slate-600"
                        }`}
                      >
                        {count}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          <span className="text-[11px] text-slate-500 hidden sm:inline">
            Haga clic en una categoría o en <strong>Consultar</strong> para abrir la búsqueda en una <strong>nueva pestaña</strong>
          </span>
        </div>

        <div className="p-4 space-y-4">
          {/* Cargando */}
          {loading && cargas.length === 0 ? (
            <div className="py-12 text-center">
              <Loader2 className="mx-auto h-6 w-6 animate-spin text-[#0d6efd] mb-2" />
              <p className="text-xs text-slate-500">Cargando historial de planillas...</p>
            </div>
          ) : cargas.length === 0 ? (
            <div className="py-12 text-center">
              <FileSpreadsheet className="mx-auto h-10 w-10 text-slate-300 mb-2" />
              <h4 className="text-xs font-bold text-slate-800">No existen planillas publicadas</h4>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Vaya a &quot;Importar Planilla Excel&quot; para publicar el primer archivo de boletas.
              </p>
            </div>
          ) : (
            <>
              {/* 3. Tarjetas de Resumen del Año */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="bg-gradient-to-br from-blue-50 to-blue-100/50 border border-blue-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-blue-900">{yearStats.mesesConDatos} / 12</div>
                  <div className="text-[11px] text-blue-700 font-semibold">Meses con planilla en {selectedYear}</div>
                </div>
                <div className="bg-gradient-to-br from-emerald-50 to-emerald-100/50 border border-emerald-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-emerald-900">{yearStats.totalCargas}</div>
                  <div className="text-[11px] text-emerald-700 font-semibold">Archivos Excel cargados</div>
                </div>
                <div className="bg-gradient-to-br from-amber-50 to-amber-100/50 border border-amber-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-amber-900">{yearStats.totalBoletas.toLocaleString()}</div>
                  <div className="text-[11px] text-amber-700 font-semibold">Boletas generadas en el año</div>
                </div>
              </div>

              {/* 4. Mini Calendario Resumen de los 12 Meses */}
              <div className="bg-slate-50 border border-slate-200 rounded p-3">
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider block mb-2">
                  Vista rápida anual {selectedYear} (haga clic para desplegar):
                </span>
                <div className="grid grid-cols-4 sm:grid-cols-6 lg:grid-cols-12 gap-1.5">
                  {calendarMonthsStatus.map((m) => {
                    const isOpen = !!expandedMonths[m.mes];
                    return (
                      <button
                        key={m.mes}
                        type="button"
                        disabled={!m.hasData}
                        onClick={() => {
                          if (m.hasData) {
                            setExpandedMonths((prev) => ({ ...prev, [m.mes]: !prev[m.mes] }));
                          }
                        }}
                        className={`p-1.5 rounded text-center transition border ${
                          m.hasData
                            ? isOpen
                              ? "bg-[#0b223d] text-white border-[#0b223d] shadow-sm font-bold"
                              : "bg-white text-emerald-800 border-emerald-300 hover:bg-emerald-50 font-bold"
                            : "bg-slate-100 text-slate-400 border-slate-200 opacity-60 cursor-not-allowed text-[11px]"
                        }`}
                        title={
                          m.hasData
                            ? `${m.mes}: ${m.totalBoletas} boletas en ${m.totalCargas} archivo(s) - Haga clic para alternar detalle`
                            : `${m.mes}: Sin planilla cargada`
                        }
                      >
                        <div className="text-xs leading-none">{m.corto}</div>
                        <div className="text-[9px] mt-1 font-mono leading-none">
                          {m.hasData ? (
                            <span className={isOpen ? "text-amber-300" : "text-emerald-700"}>
                              {m.totalBoletas}
                            </span>
                          ) : (
                            "0"
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* 5. Lista de Meses con Planillas */}
              {monthsData.length === 0 ? (
                <div className="py-8 text-center bg-slate-50 border border-slate-200 rounded">
                  <Calendar className="mx-auto h-8 w-8 text-slate-300 mb-2" />
                  <p className="text-xs text-slate-600 font-semibold">
                    No hay planillas registradas para el año {selectedYear}.
                  </p>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    Seleccione otro año en la barra superior o cargue un nuevo archivo Excel.
                  </p>
                </div>
              ) : (
                <div className="space-y-3">
                  {/* Barra de control para desplegar / colapsar todos */}
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                      Meses con planillas publicadas ({monthsData.length}):
                    </span>
                    <div className="flex items-center gap-1.5">
                      <Button
                        type="button"
                        variant="ghost"
                        size="sm"
                        onClick={expandAll}
                        className="h-7 text-[11px] text-slate-600 hover:text-slate-900 px-2 font-semibold"
                      >
                        <ChevronsDown className="h-3.5 w-3.5 mr-1" />
                        Expandir todos
                      </Button>
                      <Button
                        type="button"
                        variant="ghost"
                        size="sm"
                        onClick={collapseAll}
                        className="h-7 text-[11px] text-slate-600 hover:text-slate-900 px-2 font-semibold"
                      >
                        <ChevronsUp className="h-3.5 w-3.5 mr-1" />
                        Cerrar todos
                      </Button>
                    </div>
                  </div>

                  {monthsData.map((monthData) => {
                    const isExpanded = !!expandedMonths[monthData.mes];
                    const categories = getCategoriesForMonth(monthData.cargas);

                    return (
                      <div
                        key={monthData.mes}
                        className={`border rounded overflow-hidden transition-all ${
                          isExpanded
                            ? "border-blue-400 shadow-sm"
                            : "border-slate-200 hover:border-slate-300"
                        }`}
                      >
                        {/* Cabecera del Mes (Haga clic para abrir/cerrar el dropdown) */}
                        <div
                          onClick={() => toggleMonth(monthData.mes)}
                          className={`w-full flex flex-col sm:flex-row sm:items-center justify-between px-4 py-3 cursor-pointer transition select-none gap-2 ${
                            isExpanded ? "bg-blue-50/40" : "bg-slate-50 hover:bg-slate-100"
                          }`}
                        >
                          {/* Identificador del Mes */}
                          <div className="flex items-center gap-3">
                            <div className="h-10 w-12 rounded bg-[#0b223d] text-white flex flex-col items-center justify-center shrink-0 shadow-sm">
                              <span className="text-[9px] font-bold tracking-wider leading-none text-amber-300">
                                {MESES_CORTOS[monthData.mes] || monthData.mes.slice(0, 3)}
                              </span>
                              <span className="text-sm font-bold leading-none mt-0.5">
                                {selectedYear.slice(2)}
                              </span>
                            </div>
                            <div>
                              <div className="flex items-center gap-2">
                                <h4 className="text-sm font-bold text-slate-900 leading-tight">
                                  {monthData.mes} {selectedYear}
                                </h4>
                                <span className="bg-slate-200 text-slate-700 text-[10px] font-bold px-1.5 py-0.2 rounded">
                                  {monthData.cargas.length} archivo{monthData.cargas.length !== 1 ? "s" : ""}
                                </span>
                              </div>
                              <p className="text-[11px] text-slate-500 mt-0.5">
                                Total: <strong>{monthData.totalBoletas.toLocaleString()}</strong> boletas de pago
                              </p>
                            </div>
                          </div>

                          {/* Chips de Categorías y Acciones (abren en nueva pestaña target="_blank") */}
                          <div className="flex items-center flex-wrap gap-1.5">
                            {/* Chips de Categorías clickeables directamente con target="_blank" */}
                            {categories.map((cat) => {
                              const url = buildConsultaUrl({
                                categoriaId: cat.id,
                                mes: monthData.mes,
                                anio: selectedYear,
                              });
                              return (
                                <a
                                  key={cat.id}
                                  href={url}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  onClick={(e) => e.stopPropagation()}
                                  className={`inline-flex items-center gap-1 text-[11px] font-bold px-2.5 py-1 rounded border transition shadow-2xs hover:scale-105 active:scale-95 ${
                                    CATEGORY_COLORS[cat.id] ||
                                    "bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100"
                                  }`}
                                  title={`Abrir en nueva pestaña: ${cat.label} de ${monthData.mes} ${selectedYear}`}
                                >
                                  <span>{cat.label.replace("CAS ", "")}</span>
                                  <span className="bg-white/70 px-1 py-0.2 rounded text-[10px] font-bold ml-0.5">
                                    {cat.totalBoletas}
                                  </span>
                                  <ExternalLink className="h-2.5 w-2.5 opacity-60 ml-0.5" />
                                </a>
                              );
                            })}

                            {/* Botón Ver Todo el Mes (target="_blank") */}
                            <a
                              href={buildConsultaUrl({
                                categoriaId: "",
                                mes: monthData.mes,
                                anio: selectedYear,
                              })}
                              target="_blank"
                              rel="noopener noreferrer"
                              onClick={(e) => e.stopPropagation()}
                              className="inline-flex items-center gap-1 text-[11px] font-bold px-2.5 py-1 rounded border border-blue-300 bg-white text-blue-700 hover:bg-blue-50 shadow-2xs transition"
                              title={`Abrir en nueva pestaña: TODAS las boletas de ${monthData.mes} ${selectedYear}`}
                            >
                              <Eye className="h-3 w-3 text-blue-600" />
                              <span>Ver todo</span>
                              <ExternalLink className="h-2.5 w-2.5 text-blue-500 opacity-70" />
                            </a>

                            {/* Indicador y botón para plegar/desplegar el dropdown del mes */}
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                toggleMonth(monthData.mes);
                              }}
                              className="inline-flex items-center gap-1 px-2 py-1 rounded text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-200/60 transition"
                              title={isExpanded ? "Cerrar detalle" : "Ver detalle"}
                            >
                              <span className="text-[11px] hidden sm:inline">
                                {isExpanded ? "Cerrar" : "Detalle"}
                              </span>
                              {isExpanded ? (
                                <ChevronDown className="h-4 w-4 text-slate-700" />
                              ) : (
                                <ChevronRight className="h-4 w-4 text-slate-500" />
                              )}
                            </button>
                          </div>
                        </div>

                        {/* Detalle Expandido (Dropdown) */}
                        {isExpanded && (
                          <div className="border-t border-slate-200 bg-white p-4 space-y-3">
                            {/* Barra de Acceso Rápido a Consultar (target="_blank") */}
                            <div className="bg-blue-50/70 border border-blue-200/80 rounded p-3">
                              <span className="text-[11px] font-bold text-blue-900 uppercase tracking-wider block mb-2">
                                🔍 Abrir consulta en nueva pestaña ({monthData.mes} {selectedYear}):
                              </span>
                              <div className="flex flex-wrap gap-2">
                                <a
                                  href={buildConsultaUrl({
                                    categoriaId: "",
                                    mes: monthData.mes,
                                    anio: selectedYear,
                                  })}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="inline-flex items-center gap-1.5 h-8 px-3 text-xs bg-[#0b223d] hover:bg-[#153457] text-white font-bold rounded shadow-sm transition"
                                >
                                  <Users className="h-3.5 w-3.5 text-amber-300" />
                                  <span>Ver todas las categorías ({monthData.totalBoletas})</span>
                                  <ExternalLink className="h-3 w-3 text-slate-300" />
                                </a>

                                {categories.map((cat) => (
                                  <a
                                    key={cat.id}
                                    href={buildConsultaUrl({
                                      categoriaId: cat.id,
                                      mes: monthData.mes,
                                      anio: selectedYear,
                                    })}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    className={`inline-flex items-center gap-1.5 h-8 px-3 text-xs font-bold rounded border transition ${
                                      CATEGORY_COLORS[cat.id] || "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                                    }`}
                                  >
                                    <span>{cat.label}</span>
                                    <span className="bg-white/80 px-1.5 py-0.2 rounded text-[10px] font-bold">
                                      {cat.totalBoletas}
                                    </span>
                                    <ExternalLink className="h-3 w-3 opacity-60" />
                                  </a>
                                ))}
                              </div>
                            </div>

                            {/* Lista de Archivos Excel Cargados */}
                            <div>
                              <div className="flex items-center justify-between mb-2">
                                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                                  Archivos Excel cargados ({monthData.cargas.length}):
                                </span>
                                <button
                                  type="button"
                                  onClick={() => toggleMonth(monthData.mes)}
                                  className="text-[11px] text-slate-500 hover:text-slate-800 font-semibold underline flex items-center gap-1"
                                >
                                  <ChevronDown className="h-3 w-3 rotate-180" />
                                  Cerrar detalle de {monthData.mes}
                                </button>
                              </div>

                              <div className="divide-y divide-slate-100 border border-slate-200 rounded overflow-hidden">
                                {monthData.cargas.map((carga) => (
                                  <div
                                    key={carga.id}
                                    className="flex flex-col sm:flex-row sm:items-center justify-between p-3 bg-white hover:bg-slate-50/80 transition gap-2 group"
                                  >
                                    {/* Info del Archivo */}
                                    <div className="flex items-start sm:items-center gap-2.5 min-w-0">
                                      <FileSpreadsheet className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5 sm:mt-0" />
                                      <div className="min-w-0">
                                        <p
                                          className="text-xs font-semibold text-slate-900 truncate font-mono"
                                          title={carga.nombre_archivo}
                                        >
                                          {carga.nombre_archivo}
                                        </p>
                                        <p className="text-[11px] text-slate-500">
                                          Publicado el:{" "}
                                          {new Date(carga.created_at).toLocaleString("es-PE", {
                                            day: "2-digit",
                                            month: "short",
                                            year: "numeric",
                                            hour: "2-digit",
                                            minute: "2-digit",
                                          })}
                                        </p>
                                      </div>
                                    </div>

                                    {/* Categoría, Boletas y Botones de Acción */}
                                    <div className="flex items-center justify-between sm:justify-end gap-2 shrink-0">
                                      <span
                                        className={`text-[10px] font-bold px-2 py-0.5 rounded border ${
                                          CATEGORY_COLORS[carga.categoria_id] ||
                                          "bg-slate-50 text-slate-700 border-slate-200"
                                        }`}
                                      >
                                        {resolveCategoriaLabel(carga.categoria_id, carga.categoria_label)}
                                      </span>

                                      <span className="inline-flex items-center gap-1 text-xs text-slate-700 font-bold px-2 py-0.5 bg-slate-100 rounded">
                                        <Users className="h-3 w-3 text-slate-500" />
                                        {carga.total_trabajadores} boletas
                                      </span>

                                      {/* Botón Consultar (target="_blank") */}
                                      <a
                                        href={buildConsultaUrl({
                                          categoriaId: carga.categoria_id,
                                          mes: monthData.mes,
                                          anio: selectedYear,
                                        })}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        className="inline-flex items-center gap-1 text-xs font-bold text-blue-700 hover:text-blue-900 bg-blue-50 hover:bg-blue-100 px-2.5 py-1 rounded border border-blue-200 transition"
                                        title={`Abrir en nueva pestaña las boletas de ${carga.nombre_archivo}`}
                                      >
                                        <Search className="h-3 w-3" />
                                        <span>Consultar</span>
                                        <ExternalLink className="h-2.5 w-2.5 opacity-60" />
                                      </a>

                                      {/* Botón Eliminar */}
                                      <button
                                        type="button"
                                        onClick={() => setDeleteTarget(carga)}
                                        className="p-1 text-red-500 hover:text-red-700 hover:bg-red-50 rounded transition"
                                        title="Eliminar este archivo de planilla"
                                      >
                                        <Trash2 className="h-3.5 w-3.5" />
                                      </button>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          )}
        </div>
      </div>

      {/* Modal de confirmación para eliminar */}
      <AlertDialog open={!!deleteTarget} onOpenChange={(open) => !open && setDeleteTarget(null)}>
        <AlertDialogContent className="bg-white rounded border border-slate-300">
          <AlertDialogHeader>
            <AlertDialogTitle className="text-sm font-bold text-slate-900">
              ¿Eliminar esta planilla?
            </AlertDialogTitle>
            <AlertDialogDescription className="text-xs text-slate-600 space-y-1">
              <p>Esta acción eliminará permanentemente la carga y sus boletas asociadas:</p>
              <p className="font-mono font-semibold text-slate-800 bg-slate-100 p-2 rounded break-all">
                {deleteTarget?.nombre_archivo}
              </p>
              <p className="text-slate-500">
                Periodo: <strong>{deleteTarget?.mes} {deleteTarget?.anio}</strong> ·{" "}
                <strong>{deleteTarget?.total_trabajadores}</strong> trabajadores
              </p>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel className="rounded text-xs h-8" disabled={deleting}>
              Cancelar
            </AlertDialogCancel>
            <AlertDialogAction
              onClick={handleDelete}
              disabled={deleting}
              className="rounded bg-red-600 hover:bg-red-700 text-white text-xs h-8 font-bold"
            >
              {deleting ? (
                <>
                  <Loader2 className="mr-1 h-3 w-3 animate-spin" /> Eliminando...
                </>
              ) : (
                "Sí, eliminar"
              )}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
