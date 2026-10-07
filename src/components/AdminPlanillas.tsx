import { useCallback, useEffect, useMemo, useState } from "react";
import {
  fetchCargasPlanilla,
  CargaPlanillaItem,
  resolveCategoriaLabel,
  deleteCargaPlanilla,
  deleteCargasPlanilla,
} from "@/lib/payrollService";
import { AdminVentanilla, VentanillaInitialFilters } from "@/components/AdminVentanilla";
import { Button } from "@/components/ui/button";
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
import { toast } from "@/hooks/use-toast";
import {
  Calendar,
  ChevronRight,
  FileSpreadsheet,
  Home,
  Loader2,
  RefreshCw,
  Search,
  Trash2,
} from "lucide-react";

const MESES_ORDEN = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
];

const CATEGORY_COLORS: Record<string, string> = {
  sede: "border-blue-200 bg-blue-50 text-blue-900 hover:border-blue-400",
  jec: "border-violet-200 bg-violet-50 text-violet-900 hover:border-violet-400",
  orquestando: "border-amber-200 bg-amber-50 text-amber-900 hover:border-amber-400",
  seho: "border-rose-200 bg-rose-50 text-rose-900 hover:border-rose-400",
  ebe: "border-teal-200 bg-teal-50 text-teal-900 hover:border-teal-400",
  winanq: "border-emerald-200 bg-emerald-50 text-emerald-900 hover:border-emerald-400",
  convivencia: "border-orange-200 bg-orange-50 text-orange-900 hover:border-orange-400",
  mantenimiento: "border-slate-300 bg-slate-100 text-slate-800 hover:border-slate-500",
};
const DEFAULT_CATEGORY_COLOR = "border-slate-200 bg-white text-slate-800 hover:border-slate-400";

const normalizeMes = (mes: string) => {
  const m = String(mes || "").trim().toUpperCase();
  return m === "SETIEMBRE" ? "SEPTIEMBRE" : m;
};

const toTitleCase = (value: string) =>
  value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();

const sumBoletas = (items: CargaPlanillaItem[]) =>
  items.reduce((total, item) => total + (item.total_trabajadores || 0), 0);

const cardBase =
  "rounded-lg border bg-white p-4 text-left shadow-sm transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500";

export function AdminPlanillas() {
  const [cargas, setCargas] = useState<CargaPlanillaItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [anio, setAnio] = useState("");
  const [mes, setMes] = useState("");
  const [categoriaId, setCategoriaId] = useState("");
  const [searching, setSearching] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<CargaPlanillaItem | null>(null);
  const [deleting, setDeleting] = useState(false);
  const [deleteMonthTarget, setDeleteMonthTarget] = useState<{
    mes: string;
    anio: string;
    count: number;
    ids: string[];
  } | null>(null);
  const [deletingMonth, setDeletingMonth] = useState(false);

  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      setCargas(await fetchCargasPlanilla());
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const years = useMemo(() => {
    const set = new Set(cargas.map((c) => String(c.anio).trim()).filter(Boolean));
    return [...set].sort((a, b) => b.localeCompare(a));
  }, [cargas]);

  const yearCargas = useMemo(
    () => cargas.filter((c) => String(c.anio).trim() === anio),
    [cargas, anio]
  );

  const monthCargas = useMemo(
    () => yearCargas.filter((c) => normalizeMes(c.mes) === mes),
    [yearCargas, mes]
  );

  const categories = useMemo(() => {
    const map = new Map<string, { id: string; label: string; total: number }>();
    for (const c of monthCargas) {
      const current = map.get(c.categoria_id);
      if (current) {
        current.total += c.total_trabajadores || 0;
      } else {
        map.set(c.categoria_id, {
          id: c.categoria_id,
          label: resolveCategoriaLabel(c.categoria_id, c.categoria_label),
          total: c.total_trabajadores || 0,
        });
      }
    }
    return [...map.values()].sort((a, b) => b.total - a.total);
  }, [monthCargas]);

  const categoryLabel =
    categories.find((c) => c.id === categoriaId)?.label ?? categoriaId.toUpperCase();

  const ventanillaFilters = useMemo<VentanillaInitialFilters>(
    () => ({ categoriaId, mes, anio }),
    [categoriaId, mes, anio]
  );

  const goHome = () => {
    setSearching(false);
    setAnio("");
    setMes("");
    setCategoriaId("");
  };
  const goYear = () => {
    setMes("");
    setCategoriaId("");
  };
  const goMonth = () => setCategoriaId("");

  const handleDelete = async () => {
    if (!deleteTarget) return;
    setDeleting(true);
    try {
      const res = await deleteCargaPlanilla(deleteTarget.id);
      if (res.ok) {
        toast({ title: "Planilla eliminada", description: deleteTarget.nombre_archivo });
        await loadData();
      } else {
        toast({
          title: "Error al eliminar",
          description: res.error || "No se pudo eliminar la planilla.",
          variant: "destructive",
        });
      }
    } finally {
      setDeleting(false);
      setDeleteTarget(null);
    }
  };

  const handleDeleteMonth = async () => {
    if (!deleteMonthTarget) return;
    setDeletingMonth(true);
    try {
      const res = await deleteCargasPlanilla(deleteMonthTarget.ids);
      if (res.ok) {
        toast({
          title: "Planillas del mes eliminadas",
          description: `Se eliminaron correctamente los ${deleteMonthTarget.count} archivos de ${toTitleCase(deleteMonthTarget.mes)} ${deleteMonthTarget.anio}.`,
        });
        await loadData();
        goYear();
      } else {
        toast({
          title: "Error al eliminar planillas",
          description: res.error || "No se pudieron eliminar las planillas del mes.",
          variant: "destructive",
        });
      }
    } finally {
      setDeletingMonth(false);
      setDeleteMonthTarget(null);
    }
  };

  const crumbClass = "rounded px-1.5 py-0.5 font-semibold text-blue-700 hover:bg-blue-50 hover:underline";

  const breadcrumb = (
    <nav aria-label="Ruta de navegación" className="flex flex-wrap items-center gap-1 text-sm">
      <button type="button" onClick={goHome} className={`${crumbClass} inline-flex items-center gap-1`}>
        <Home className="h-3.5 w-3.5" /> Inicio
      </button>
      {searching && (
        <>
          <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
          <span className="px-1.5 font-bold text-slate-900">Buscar trabajador</span>
        </>
      )}
      {anio && (
        <>
          <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
          {mes ? (
            <button type="button" onClick={goYear} className={crumbClass}>{anio}</button>
          ) : (
            <span className="px-1.5 font-bold text-slate-900">{anio}</span>
          )}
        </>
      )}
      {mes && (
        <>
          <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
          {categoriaId ? (
            <button type="button" onClick={goMonth} className={crumbClass}>{toTitleCase(mes)}</button>
          ) : (
            <span className="px-1.5 font-bold text-slate-900">{toTitleCase(mes)}</span>
          )}
        </>
      )}
      {categoriaId && (
        <>
          <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
          <span className="px-1.5 font-bold text-slate-900">{categoryLabel}</span>
        </>
      )}
    </nav>
  );

  const renderContent = () => {
    if (searching) {
      return <AdminVentanilla key="search" />;
    }

    if (categoriaId) {
      return (
        <AdminVentanilla key={`${anio}-${mes}-${categoriaId}`} initialFilters={ventanillaFilters} hideSearch />
      );
    }

    if (loading && !cargas.length) {
      return (
        <div className="flex items-center justify-center gap-2 py-16 text-sm text-slate-500">
          <Loader2 className="h-4 w-4 animate-spin" /> Cargando planillas...
        </div>
      );
    }

    if (!anio) {
      return years.length ? (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          {years.map((year) => {
            const items = cargas.filter((c) => String(c.anio).trim() === year);
            return (
              <button
                key={year}
                type="button"
                onClick={() => setAnio(year)}
                className={`${cardBase} border-slate-200 hover:border-blue-400`}
              >
                <Calendar className="mb-2 h-5 w-5 text-blue-600" />
                <p className="text-2xl font-extrabold text-[#0b223d]">{year}</p>
                <p className="text-xs text-slate-500">{sumBoletas(items)} boletas</p>
              </button>
            );
          })}
        </div>
      ) : (
        <p className="py-16 text-center text-sm text-slate-500">
          Aún no hay planillas. Impórtalas desde la pestaña "Importar Excel".
        </p>
      );
    }

    if (!mes) {
      return (
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
          {MESES_ORDEN.map((name) => {
            const items = yearCargas.filter((c) => normalizeMes(c.mes) === name);
            const hasData = items.length > 0;
            return (
              <button
                key={name}
                type="button"
                disabled={!hasData}
                onClick={() => setMes(name)}
                className={`${cardBase} ${
                  hasData
                    ? "border-slate-200 hover:border-blue-400"
                    : "cursor-not-allowed border-slate-100 bg-slate-50 opacity-50 shadow-none"
                }`}
              >
                <p className="text-base font-bold text-[#0b223d]">{toTitleCase(name)}</p>
                <p className="text-xs text-slate-500">
                  {hasData ? `${sumBoletas(items)} boletas` : "Sin planilla"}
                </p>
              </button>
            );
          })}
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {categories.map((category) => (
            <button
              key={category.id}
              type="button"
              onClick={() => setCategoriaId(category.id)}
              className={`${cardBase} ${CATEGORY_COLORS[category.id] ?? DEFAULT_CATEGORY_COLOR}`}
            >
              <p className="text-sm font-bold uppercase">CAS {category.label}</p>
              <p className="text-xs opacity-80">{category.total} boletas</p>
            </button>
          ))}
        </div>

        <details className="rounded-lg border border-slate-200 bg-white text-sm shadow-sm" open>
          <summary className="cursor-pointer select-none px-4 py-3 font-semibold text-slate-700 flex items-center justify-between">
            <span>Archivos Excel cargados ({monthCargas.length})</span>
            {monthCargas.length > 0 && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setDeleteMonthTarget({
                    mes,
                    anio,
                    count: monthCargas.length,
                    ids: monthCargas.map((c) => String(c.id)),
                  });
                }}
                className="inline-flex items-center gap-1.5 rounded border border-red-200 bg-red-50 px-2.5 py-1 text-xs font-bold text-red-700 hover:bg-red-100 transition shadow-sm"
              >
                <Trash2 className="h-3.5 w-3.5 text-red-600" /> Borrar todo el mes ({monthCargas.length})
              </button>
            )}
          </summary>
          <ul className="divide-y divide-slate-100 border-t border-slate-100">
            {monthCargas.map((carga) => (
              <li key={carga.id} className="flex items-center gap-3 px-4 py-2.5">
                <FileSpreadsheet className="h-4 w-4 shrink-0 text-emerald-600" />
                <span className="min-w-0 flex-1 truncate font-mono text-xs" title={carga.nombre_archivo}>
                  {carga.nombre_archivo}
                </span>
                <span className="shrink-0 text-xs text-slate-500">{carga.total_trabajadores} boletas</span>
                <button
                  type="button"
                  onClick={() => setDeleteTarget(carga)}
                  aria-label={`Eliminar ${carga.nombre_archivo}`}
                  className="shrink-0 rounded p-1.5 text-red-600 hover:bg-red-50"
                >
                  <Trash2 className="h-4 w-4" />
                </button>
              </li>
            ))}
          </ul>
        </details>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-col gap-3 rounded-lg border border-slate-200 bg-white p-3 shadow-sm sm:flex-row sm:items-center sm:justify-between">
        {breadcrumb}
        <div className="flex flex-wrap gap-2">
          {mes && !categoriaId && monthCargas.length > 0 && (
            <Button
              size="sm"
              variant="destructive"
              onClick={() => setDeleteMonthTarget({
                mes,
                anio,
                count: monthCargas.length,
                ids: monthCargas.map((c) => String(c.id)),
              })}
              className="h-9 text-xs font-bold"
            >
              <Trash2 className="mr-1.5 h-3.5 w-3.5" /> Borrar todo ({toTitleCase(mes)})
            </Button>
          )}
          {!searching && !categoriaId && (
            <Button size="sm" onClick={() => setSearching(true)} className="h-9 bg-[#0d6efd] text-xs font-bold hover:bg-[#0b5ed7]">
              <Search className="mr-1.5 h-3.5 w-3.5" /> Buscar trabajador
            </Button>
          )}
          <Button size="sm" variant="outline" onClick={loadData} disabled={loading} className="h-9 text-xs">
            <RefreshCw className={`mr-1.5 h-3.5 w-3.5 ${loading ? "animate-spin" : ""}`} /> Actualizar
          </Button>
        </div>
      </div>

      {renderContent()}

      <AlertDialog open={!!deleteTarget} onOpenChange={(open) => !open && setDeleteTarget(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>¿Eliminar esta planilla?</AlertDialogTitle>
            <AlertDialogDescription>
              Se borrarán el archivo <strong>{deleteTarget?.nombre_archivo}</strong> y sus boletas. Esta acción no se puede deshacer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={deleting}>Cancelar</AlertDialogCancel>
            <AlertDialogAction onClick={handleDelete} disabled={deleting} className="bg-red-600 hover:bg-red-700">
              {deleting ? "Eliminando..." : "Eliminar"}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Diálogo para eliminar todas las planillas del mes */}
      <AlertDialog open={!!deleteMonthTarget} onOpenChange={(open) => !open && setDeleteMonthTarget(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle className="text-red-700">
              ¿Eliminar todas las planillas de {toTitleCase(deleteMonthTarget?.mes || "")} {deleteMonthTarget?.anio}?
            </AlertDialogTitle>
            <AlertDialogDescription className="space-y-2 text-slate-600">
              <p>
                Se eliminarán de forma permanente los <strong>{deleteMonthTarget?.count} archivos Excel</strong> cargados en este mes y todas sus boletas correspondientes en la base de datos.
              </p>
              <p className="text-xs font-semibold text-amber-700 bg-amber-50 p-2 rounded border border-amber-200">
                ⚠️ Esta acción no se puede deshacer.
              </p>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={deletingMonth}>Cancelar</AlertDialogCancel>
            <AlertDialogAction
              onClick={handleDeleteMonth}
              disabled={deletingMonth}
              className="bg-red-600 hover:bg-red-700 font-bold"
            >
              {deletingMonth ? "Eliminando..." : `Sí, borrar ${deleteMonthTarget?.count} planillas`}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
