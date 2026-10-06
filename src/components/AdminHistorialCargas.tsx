import { useState, useEffect, useMemo } from "react";
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
  ChevronRight,
  FileText,
  Trash2,
  ExternalLink,
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

export function AdminHistorialCargas({ onNavigateToConsulta }: AdminHistorialCargasProps) {
  const [cargas, setCargas] = useState<CargaPlanillaItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [selectedYear, setSelectedYear] = useState("");
  const [expandedMonth, setExpandedMonth] = useState<string | null>(null);
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

  // Años disponibles
  const availableYears = useMemo(() => {
    const years = [...new Set(cargas.map((c) => c.anio))].sort((a, b) => b.localeCompare(a));
    return years;
  }, [cargas]);

  // Auto-seleccionar el año más reciente
  useEffect(() => {
    if (!selectedYear && availableYears.length > 0) {
      setSelectedYear(availableYears[0]);
    }
  }, [availableYears, selectedYear]);

  // Datos del año seleccionado agrupados por mes
  const monthsData = useMemo((): MonthData[] => {
    if (!selectedYear) return [];

    const yearCargas = cargas.filter((c) => c.anio === selectedYear);
    const mesMap = new Map<string, CargaPlanillaItem[]>();

    for (const carga of yearCargas) {
      const mesNorm = carga.mes.trim().toUpperCase();
      if (!mesMap.has(mesNorm)) mesMap.set(mesNorm, []);
      mesMap.get(mesNorm)!.push(carga);
    }

    return MESES_ORDEN
      .filter((m) => mesMap.has(m))
      .map((mes) => {
        const cargasMes = mesMap.get(mes)!;
        return {
          mes,
          cargas: cargasMes.sort((a, b) =>
            new Date(b.created_at).getTime() - new Date(a.created_at).getTime()
          ),
          totalBoletas: cargasMes.reduce((sum, c) => sum + c.total_trabajadores, 0),
        };
      })
      .reverse(); // Más reciente primero
  }, [cargas, selectedYear]);

  // Estadísticas del año
  const yearStats = useMemo(() => {
    const yearCargas = cargas.filter((c) => c.anio === selectedYear);
    return {
      totalCargas: yearCargas.length,
      totalBoletas: yearCargas.reduce((sum, c) => sum + c.total_trabajadores, 0),
      mesesConDatos: new Set(yearCargas.map((c) => c.mes.trim().toUpperCase())).size,
    };
  }, [cargas, selectedYear]);

  // Categorías únicas por mes (para badges del grid)
  const getCategoriesForMonth = (cargas: CargaPlanillaItem[]) => {
    const catMap = new Map<string, { id: string; label: string; totalBoletas: number; count: number }>();
    for (const c of cargas) {
      const existing = catMap.get(c.categoria_id);
      if (existing) {
        existing.totalBoletas += c.total_trabajadores;
        existing.count += 1;
      } else {
        catMap.set(c.categoria_id, {
          id: c.categoria_id,
          label: resolveCategoriaLabel(c.categoria_id, c.categoria_label),
          totalBoletas: c.total_trabajadores,
          count: 1,
        });
      }
    }
    return [...catMap.values()];
  };

  const handleCategoryClick = (categoriaId: string, mes: string) => {
    if (onNavigateToConsulta) {
      onNavigateToConsulta({ categoriaId, mes, anio: selectedYear });
    }
  };

  const handleDelete = async () => {
    if (!deleteTarget) return;
    setDeleting(true);
    try {
      const res = await deleteCargaPlanilla(deleteTarget.id);
      if (res.ok) {
        toast({ title: "Planilla eliminada", description: `Se eliminó "${deleteTarget.nombre_archivo}" correctamente.` });
        await loadData();
      } else {
        toast({ title: "Error", description: res.error || "No se pudo eliminar.", variant: "destructive" });
      }
    } finally {
      setDeleting(false);
      setDeleteTarget(null);
    }
  };

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

  return (
    <div className="space-y-3">
      <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        {/* Encabezado */}
        <div className="bg-[#0b223d] text-white px-4 py-2.5 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <History className="h-4 w-4 text-amber-400" />
            <h3 className="text-xs font-bold uppercase tracking-wider">
              Historial de Planillas
            </h3>
          </div>

          <div className="flex items-center gap-2">
            {availableYears.length > 0 && (
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

        <div className="p-4">
          {/* Cargando */}
          {loading && cargas.length === 0 ? (
            <div className="py-12 text-center">
              <Loader2 className="mx-auto h-6 w-6 animate-spin text-[#0d6efd] mb-2" />
              <p className="text-xs text-slate-500">Cargando registros...</p>
            </div>
          ) : cargas.length === 0 ? (
            <div className="py-12 text-center">
              <FileSpreadsheet className="mx-auto h-10 w-10 text-slate-300 mb-2" />
              <h4 className="text-xs font-bold text-slate-800">
                No existen planillas publicadas
              </h4>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Vaya a &quot;Importar Planilla Excel&quot; para publicar el primer archivo.
              </p>
            </div>
          ) : (
            <>
              {/* Resumen del año */}
              <div className="grid grid-cols-3 gap-3 mb-4">
                <div className="bg-blue-50 border border-blue-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-blue-800">{yearStats.mesesConDatos}</div>
                  <div className="text-[11px] text-blue-600 font-semibold">Meses con planilla</div>
                </div>
                <div className="bg-emerald-50 border border-emerald-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-emerald-800">{yearStats.totalCargas}</div>
                  <div className="text-[11px] text-emerald-600 font-semibold">Archivos cargados</div>
                </div>
                <div className="bg-amber-50 border border-amber-200 rounded p-3 text-center">
                  <div className="text-2xl font-bold text-amber-800">{yearStats.totalBoletas.toLocaleString()}</div>
                  <div className="text-[11px] text-amber-600 font-semibold">Total de boletas</div>
                </div>
              </div>

              {/* Grilla de meses */}
              {monthsData.length === 0 ? (
                <div className="py-8 text-center">
                  <Calendar className="mx-auto h-8 w-8 text-slate-300 mb-2" />
                  <p className="text-xs text-slate-500">No hay planillas registradas para {selectedYear}.</p>
                </div>
              ) : (
                <div className="space-y-2">
                  {monthsData.map((monthData) => {
                    const isExpanded = expandedMonth === monthData.mes;
                    const categories = getCategoriesForMonth(monthData.cargas);

                    return (
                      <div
                        key={monthData.mes}
                        className="border border-slate-200 rounded overflow-hidden transition-shadow hover:shadow-sm"
                      >
                        {/* Encabezado del mes */}
                        <button
                          type="button"
                          onClick={() => setExpandedMonth(isExpanded ? null : monthData.mes)}
                          className="w-full flex items-center justify-between px-4 py-3 bg-slate-50 hover:bg-slate-100 transition text-left"
                        >
                          <div className="flex items-center gap-3">
                            <div className="h-9 w-12 rounded bg-[#0b223d] text-white flex flex-col items-center justify-center shrink-0">
                              <span className="text-[9px] font-bold leading-none opacity-70">
                                {MESES_CORTOS[monthData.mes] || monthData.mes.slice(0, 3)}
                              </span>
                              <span className="text-sm font-bold leading-none">
                                {selectedYear.slice(2)}
                              </span>
                            </div>
                            <div>
                              <h4 className="text-sm font-bold text-slate-900">
                                {monthData.mes}
                              </h4>
                              <p className="text-[11px] text-slate-500">
                                {monthData.cargas.length} archivo{monthData.cargas.length !== 1 ? "s" : ""} ·{" "}
                                <strong>{monthData.totalBoletas}</strong> boletas
                              </p>
                            </div>
                          </div>

                          <div className="flex items-center gap-2">
                            {/* Chips de categorías (vista compacta) */}
                            <div className="hidden sm:flex flex-wrap gap-1">
                              {categories.map((cat) => (
                                <span
                                  key={cat.id}
                                  className={`inline-flex items-center gap-1 text-[10px] font-semibold px-2 py-0.5 rounded border ${
                                    CATEGORY_COLORS[cat.id] || "bg-slate-50 text-slate-700 border-slate-200"
                                  }`}
                                >
                                  {cat.label.replace("CAS ", "")}
                                  <span className="opacity-70">({cat.totalBoletas})</span>
                                </span>
                              ))}
                            </div>
                            <ChevronRight
                              className={`h-4 w-4 text-slate-400 transition-transform duration-200 ${
                                isExpanded ? "rotate-90" : ""
                              }`}
                            />
                          </div>
                        </button>

                        {/* Detalle expandido */}
                        {isExpanded && (
                          <div className="border-t border-slate-200 bg-white p-3 space-y-2">
                            {/* Categorías clickeables */}
                            <div className="mb-3">
                              <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider block mb-2">
                                Filtrar por categoría en Consultar:
                              </span>
                              <div className="flex flex-wrap gap-1.5">
                                {categories.map((cat) => (
                                  <button
                                    key={cat.id}
                                    type="button"
                                    onClick={() => handleCategoryClick(cat.id, monthData.mes)}
                                    className={`inline-flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded border transition cursor-pointer ${
                                      CATEGORY_COLORS[cat.id] || "bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100"
                                    }`}
                                    title={`Ver boletas de ${cat.label} - ${monthData.mes} ${selectedYear}`}
                                  >
                                    <Users className="h-3 w-3" />
                                    {cat.label}
                                    <span className="bg-white/60 px-1.5 py-0.5 rounded text-[10px] font-bold">
                                      {cat.totalBoletas}
                                    </span>
                                    <ExternalLink className="h-2.5 w-2.5 opacity-50" />
                                  </button>
                                ))}
                              </div>
                            </div>

                            {/* Lista de archivos */}
                            <div className="space-y-1">
                              <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider block mb-1">
                                Archivos cargados:
                              </span>
                              {monthData.cargas.map((carga) => (
                                <div
                                  key={carga.id}
                                  className="flex items-center justify-between bg-slate-50 border border-slate-100 rounded px-3 py-2 gap-2 group"
                                >
                                  <div className="flex items-center gap-2 min-w-0">
                                    <FileText className="h-3.5 w-3.5 text-slate-400 shrink-0" />
                                    <div className="min-w-0">
                                      <p className="text-[11px] font-mono text-slate-700 truncate" title={carga.nombre_archivo}>
                                        {carga.nombre_archivo}
                                      </p>
                                      <p className="text-[10px] text-slate-400">
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
                                  <div className="flex items-center gap-2 shrink-0">
                                    <span
                                      className={`text-[10px] font-semibold px-2 py-0.5 rounded border ${
                                        CATEGORY_COLORS[carga.categoria_id] || "bg-slate-50 text-slate-700 border-slate-200"
                                      }`}
                                    >
                                      {resolveCategoriaLabel(carga.categoria_id, carga.categoria_label)}
                                    </span>
                                    <span className="inline-flex items-center gap-1 text-[11px] text-slate-600 font-semibold">
                                      <Users className="h-3 w-3 text-slate-400" />
                                      {carga.total_trabajadores}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => setDeleteTarget(carga)}
                                      className="opacity-0 group-hover:opacity-100 p-1 text-red-400 hover:text-red-600 hover:bg-red-50 rounded transition"
                                      title="Eliminar esta carga"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </div>
                                </div>
                              ))}
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

      {/* Dialog de confirmación para eliminar */}
      <AlertDialog open={!!deleteTarget} onOpenChange={(open) => !open && setDeleteTarget(null)}>
        <AlertDialogContent className="bg-white rounded border border-slate-300">
          <AlertDialogHeader>
            <AlertDialogTitle className="text-sm font-bold text-slate-900">
              ¿Eliminar esta planilla?
            </AlertDialogTitle>
            <AlertDialogDescription className="text-xs text-slate-600 space-y-1">
              <p>Se eliminará permanentemente:</p>
              <p className="font-mono font-semibold text-slate-800">
                {deleteTarget?.nombre_archivo}
              </p>
              <p>
                ({deleteTarget?.total_trabajadores} trabajadores · {deleteTarget?.mes} {deleteTarget?.anio})
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
