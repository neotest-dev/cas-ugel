import { useState, useMemo, useEffect } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Skeleton } from "@/components/ui/skeleton";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { toast } from "@/hooks/use-toast";
import {
  searchAdminBoletas,
  AdminBoletaSearchFilters,
  actualizarDatosBoletaYTrabajador,
  BoletaFormData,
  BoletaHistoricaItem,
  boletaItemToWorker,
} from "@/lib/payrollService";
import { buildBoletaText, CATEGORIAS_PLANILLA, Worker } from "@/lib/boleta";
import { exportBoletaToPDF, exportBoletasToPDF } from "@/lib/pdfExport";
import { PrintBoletaPortal } from "./PrintBoletaPortal";
import { BoletaFormEditor } from "./BoletaFormEditor";
import {
  Search,
  Printer,
  Download,
  Calendar,
  X,
  Loader2,
  Users,
  Pencil,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  Save,
  Sliders,
} from "lucide-react";

const MESES_FECHA: Record<string, string> = {
  ENERO: "01", FEBRERO: "02", MARZO: "03", ABRIL: "04", MAYO: "05", JUNIO: "06",
  JULIO: "07", AGOSTO: "08", SEPTIEMBRE: "09", SETIEMBRE: "09", OCTUBRE: "10",
  NOVIEMBRE: "11", DICIEMBRE: "12",
};

const MONTH_OPTIONS: Array<[string, string]> = [
  ["01", "Enero"], ["02", "Febrero"], ["03", "Marzo"], ["04", "Abril"], ["05", "Mayo"], ["06", "Junio"],
  ["07", "Julio"], ["08", "Agosto"], ["09", "Septiembre"], ["10", "Octubre"], ["11", "Noviembre"], ["12", "Diciembre"],
];

const YEAR_OPTIONS = Array.from({ length: 6 }, (_, index) => String(new Date().getFullYear() - index));

export interface VentanillaInitialFilters {
  categoriaId?: string;
  mes?: string;
  anio?: string;
}

interface AdminVentanillaProps {
  initialFilters?: VentanillaInitialFilters | null;
  hideSearch?: boolean;
}

export function AdminVentanilla({ initialFilters, hideSearch = false }: AdminVentanillaProps = {}) {
  const [query, setQuery] = useState("");
  const [categoryId, setCategoryId] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [results, setResults] = useState<BoletaHistoricaItem[]>([]);
  const [selectedIds, setSelectedIds] = useState<number[]>([]);
  const [printTexts, setPrintTexts] = useState<string[]>([]);
  const [selectedBoleta, setSelectedBoleta] = useState<BoletaHistoricaItem | null>(null);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);

  // Modo de visualización: "form" (cajas de texto por campo) o "preview" (formato Courier A4)
  const [subTab, setSubTab] = useState<"form" | "preview">("preview");
  const [formData, setFormData] = useState<BoletaFormData | null>(null);
  const [zoom, setZoom] = useState(100);
  const [savingDb, setSavingDb] = useState(false);
  const [editedBoletaText, setEditedBoletaText] = useState("");
  const [activeHistorialFilterLabel, setActiveHistorialFilterLabel] = useState<string | null>(null);
  const [filterYear, setFilterYear] = useState("");
  const [filterMonth, setFilterMonth] = useState("");

  const clearDisplayedResults = () => {
    setResults([]);
    setSelectedBoleta(null);
    setSelectedIds([]);
    setHasSearched(false);
  };

  // Year/month selects -> date range used by the search service
  const applyPeriod = (year: string, month: string) => {
    setFilterYear(year);
    setFilterMonth(month);
    if (!year) {
      setDateFrom("");
      setDateTo("");
    } else if (month) {
      const lastDay = new Date(Number(year), Number(month), 0).getDate();
      setDateFrom(`${year}-${month}-01`);
      setDateTo(`${year}-${month}-${String(lastDay).padStart(2, "0")}`);
    } else {
      setDateFrom(`${year}-01-01`);
      setDateTo(`${year}-12-31`);
    }
    clearDisplayedResults();
  };

  const resetFilters = () => {
    setCategoryId("");
    applyPeriod("", "");
  };

  const handleSearch = async (e?: React.FormEvent) => {
    e?.preventDefault();
    const term = query.trim();
    if (!term && !categoryId && !dateFrom && !dateTo) return;
    if (dateFrom && dateTo && dateFrom > dateTo) {
      toast({ title: "Rango de fechas inválido", description: "La fecha inicial debe ser anterior a la fecha final.", variant: "destructive" });
      return;
    }

    setLoading(true);
    setHasSearched(true);
    setSelectedIds([]);
    try {
      const filters: AdminBoletaSearchFilters = {
        searchTerm: term,
        categoriaId: categoryId,
        desde: dateFrom,
        hasta: dateTo,
      };
      const data = await searchAdminBoletas(filters);
      setResults(data);
      if (data.length > 0) {
        setSelectedBoleta(data[0]);
      } else {
        setSelectedBoleta(null);
      }
    } finally {
      setLoading(false);
    }
  };

  // Aplicar filtros iniciales desde el historial
  useEffect(() => {
    if (initialFilters && (initialFilters.categoriaId || initialFilters.mes)) {
      const catId = initialFilters.categoriaId || "";
      const mes = (initialFilters.mes || "").toUpperCase();
      const anio = initialFilters.anio || "";

      setCategoryId(catId);

      const catObj = CATEGORIAS_PLANILLA.find((c) => c.id === catId);
      const catText = catObj ? catObj.label : (catId ? catId.toUpperCase() : "Todas las categorías");
      const periodText = mes && anio ? `${mes} ${anio}` : anio || "";
      setActiveHistorialFilterLabel(`${catText} · ${periodText}`);

      // Calcular rango de fechas para el mes
      if (mes && anio && MESES_FECHA[mes]) {
        const mesNum = MESES_FECHA[mes];
        const from = `${anio}-${mesNum}-01`;
        // Último día del mes
        const lastDay = new Date(Number(anio), Number(mesNum), 0).getDate();
        const to = `${anio}-${mesNum}-${String(lastDay).padStart(2, "0")}`;
        setDateFrom(from);
        setDateTo(to);
      }

      // Auto-ejecutar la búsqueda tras un pequeño delay para que los estados se actualicen
      const timer = setTimeout(() => {
        const mesNum = mes && MESES_FECHA[mes] ? MESES_FECHA[mes] : "";
        const from = mesNum ? `${anio}-${mesNum}-01` : "";
        const lastDay = mesNum ? new Date(Number(anio), Number(mesNum), 0).getDate() : 0;
        const to = mesNum ? `${anio}-${mesNum}-${String(lastDay).padStart(2, "0")}` : "";

        setLoading(true);
        setHasSearched(true);
        setSelectedIds([]);

        searchAdminBoletas({
          searchTerm: "",
          categoriaId: catId,
          desde: from,
          hasta: to,
        }).then((data) => {
          setResults(data);
          if (data.length > 0) {
            setSelectedBoleta(data[0]);
          } else {
            setSelectedBoleta(null);
          }
        }).finally(() => {
          setLoading(false);
        });
      }, 100);

      return () => clearTimeout(timer);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialFilters]);

  // Inicializar formData al seleccionar una boleta
  useEffect(() => {
    if (selectedBoleta) {
      setFormData({
        boleta_id: selectedBoleta.boleta_id,
        dni: selectedBoleta.dni,
        ap_paterno: selectedBoleta.ap_paterno,
        ap_materno: selectedBoleta.ap_materno,
        nombres: selectedBoleta.nombres,
        fecha_nac: selectedBoleta.fecha_nac || "",
        cargo: selectedBoleta.cargo || "",
        cod_essalud: selectedBoleta.cod_essalud || "",
        cuenta_banco: selectedBoleta.cuenta_banco || "",
        leyenda_rd: selectedBoleta.leyenda_rd || "",
        leyenda_mensual: selectedBoleta.leyenda_mensual || "",
        sistema_pensionario: selectedBoleta.sistema_pensionario || "",
        cussp: selectedBoleta.cussp || "",
        fecha_afiliacion: selectedBoleta.fecha_afiliacion || "",
        fecha_devengue: selectedBoleta.fecha_devengue || "",
        monto_mensual: selectedBoleta.monto_mensual || "0.00",
        onp: selectedBoleta.onp || "0.00",
        prima: selectedBoleta.prima || "0.00",
        integra: selectedBoleta.integra || "0.00",
        profuturo: selectedBoleta.profuturo || "0.00",
        habitat: selectedBoleta.habitat || "0.00",
        total_dscto: selectedBoleta.total_dscto || "0.00",
        otros_dsctos: selectedBoleta.otros_dsctos || "0.00",
        dscto_entidades: selectedBoleta.dscto_entidades || "0.00",
        dscto_judicial: selectedBoleta.dscto_judicial || "0.00",
        total_liquido: selectedBoleta.total_liquido || "0.00",
      });
    } else {
      setFormData(null);
    }
  }, [selectedBoleta]);

  // Trabajador activo generado a partir de formData en tiempo real
  const activeWorker = useMemo<Worker | null>(() => {
    if (!formData || !selectedBoleta) return null;
    return {
      n: selectedBoleta.n || "1",
      dni: formData.dni,
      apPaterno: formData.ap_paterno,
      apMaterno: formData.ap_materno,
      nombres: formData.nombres,
      fechaNac: formData.fecha_nac,
      cargo: formData.cargo,
      codEssalud: formData.cod_essalud,
      cuentaBanco: formData.cuenta_banco,
      leyendaRD: formData.leyenda_rd,
      leyendaMensual: formData.leyenda_mensual,
      sistemaPensionario: formData.sistema_pensionario,
      cussp: formData.cussp,
      fechaAfiliacion: formData.fecha_afiliacion,
      fechaDevengue: formData.fecha_devengue,
      montoMensual: formData.monto_mensual,
      descuentoPension: formData.onp,
      onp: formData.onp,
      prima: formData.prima,
      integra: formData.integra,
      profuturo: formData.profuturo,
      habitat: formData.habitat,
      aporteObligatorio: "0.00",
      comision: "0.00",
      primaSeguro: "0.00",
      totalDscto: formData.total_dscto,
      otrosDsctos: formData.otros_dsctos || "0.00",
      dsctoEntidades: formData.dscto_entidades || "0.00",
      dsctoJudicial: formData.dscto_judicial || "0.00",
      totalLiquido: formData.total_liquido,
    };
  }, [formData, selectedBoleta]);

  const liveBoletaText = useMemo(() => {
    if (!activeWorker || !selectedBoleta) return "";
    return buildBoletaText(activeWorker, selectedBoleta.mes, selectedBoleta.anio);
  }, [activeWorker, selectedBoleta]);

  useEffect(() => {
    setEditedBoletaText(liveBoletaText);
  }, [liveBoletaText]);

  const handleFieldChange = (field: keyof BoletaFormData, value: string) => {
    if (!formData) return;
    setFormData((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleAutoCalculate = () => {
    if (!formData) return;
    const onp = parseFloat(formData.onp) || 0;
    const prima = parseFloat(formData.prima) || 0;
    const integra = parseFloat(formData.integra) || 0;
    const profuturo = parseFloat(formData.profuturo) || 0;
    const habitat = parseFloat(formData.habitat) || 0;
    const otrosDsctos = parseFloat(formData.otros_dsctos) || 0;
    const dsctoEntidades = parseFloat(formData.dscto_entidades) || 0;
    const dsctoJudicial = parseFloat(formData.dscto_judicial) || 0;
    const totalDscto = (
      onp + prima + integra + profuturo + habitat + otrosDsctos + dsctoEntidades + dsctoJudicial
    ).toFixed(2);
    const monto = parseFloat(formData.monto_mensual) || 0;
    const totalLiquido = Math.max(0, monto - parseFloat(totalDscto)).toFixed(2);

    setFormData((prev) =>
      prev
        ? {
            ...prev,
            total_dscto: totalDscto,
            total_liquido: totalLiquido,
          }
        : prev
    );

    toast({
      title: "Cálculo automático aplicado",
      description: `Total Descuentos: S/. ${totalDscto} · Total Líquido: S/. ${totalLiquido}`,
    });
  };

  const handlePrint = () => {
    window.print();
  };

  const selectedBoletas = useMemo(
    () => results.filter((item) => selectedIds.includes(item.boleta_id)),
    [results, selectedIds]
  );

  const getPrintableText = (item: BoletaHistoricaItem) =>
    item.boleta_texto_personalizado || buildBoletaText(boletaItemToWorker(item), item.mes, item.anio);

  const handlePrintSelected = () => {
    const texts = selectedBoletas.map(getPrintableText);
    if (!texts.length) return;
    setPrintTexts(texts);
    window.addEventListener("afterprint", () => setPrintTexts([]), { once: true });
    window.setTimeout(() => window.print(), 100);
  };

  const handleDownloadSelected = () => {
    exportBoletasToPDF(
      selectedBoletas.map((item) => ({ worker: boletaItemToWorker(item), text: getPrintableText(item) })),
      `Boletas_CAS_${selectedBoletas.length}`
    );
  };

  const toggleBoleta = (boletaId: number) => {
    setSelectedIds((current) => current.includes(boletaId)
      ? current.filter((id) => id !== boletaId)
      : [...current, boletaId]);
  };

  const handlePDF = () => {
    const textToExport = editedBoletaText || liveBoletaText;
    if (!activeWorker || !textToExport) return;
    exportBoletaToPDF(
      activeWorker,
      textToExport,
      `Ventanilla_${selectedBoleta?.dni}_${selectedBoleta?.mes}_${selectedBoleta?.anio}`
    );
  };

  const handleResetOriginalFields = () => {
    if (!selectedBoleta) return;
    setFormData({
      boleta_id: selectedBoleta.boleta_id,
      dni: selectedBoleta.dni,
      ap_paterno: selectedBoleta.ap_paterno,
      ap_materno: selectedBoleta.ap_materno,
      nombres: selectedBoleta.nombres,
      fecha_nac: selectedBoleta.fecha_nac || "",
      cargo: selectedBoleta.cargo || "",
      cod_essalud: selectedBoleta.cod_essalud || "",
      cuenta_banco: selectedBoleta.cuenta_banco || "",
      leyenda_rd: selectedBoleta.leyenda_rd || "",
      leyenda_mensual: selectedBoleta.leyenda_mensual || "",
      sistema_pensionario: selectedBoleta.sistema_pensionario || "",
      cussp: selectedBoleta.cussp || "",
      fecha_afiliacion: selectedBoleta.fecha_afiliacion || "",
      fecha_devengue: selectedBoleta.fecha_devengue || "",
      monto_mensual: selectedBoleta.monto_mensual || "0.00",
      onp: selectedBoleta.onp || "0.00",
      prima: selectedBoleta.prima || "0.00",
      integra: selectedBoleta.integra || "0.00",
      profuturo: selectedBoleta.profuturo || "0.00",
      habitat: selectedBoleta.habitat || "0.00",
      total_dscto: selectedBoleta.total_dscto || "0.00",
      otros_dsctos: selectedBoleta.otros_dsctos || "0.00",
      dscto_entidades: selectedBoleta.dscto_entidades || "0.00",
      dscto_judicial: selectedBoleta.dscto_judicial || "0.00",
      total_liquido: selectedBoleta.total_liquido || "0.00",
    });
    // The preview textarea is independently editable, so restoring fields alone
    // may not change liveBoletaText (for example, after editing only the preview).
    // Reset the visible document directly from the original saved boleta too.
    setEditedBoletaText(
      buildBoletaText(
        boletaItemToWorker(selectedBoleta),
        selectedBoleta.mes,
        selectedBoleta.anio
      )
    );
    toast({
      title: "Datos restablecidos",
      description: "Se han recuperado los datos cargados originalmente desde la planilla.",
    });
  };

  const handleSaveToDatabase = async () => {
    if (!selectedBoleta || !formData) return;
    setSavingDb(true);
    try {
      const res = await actualizarDatosBoletaYTrabajador(formData);

      if (!res.ok) {
        toast({
          title: "Error al guardar en Base de Datos",
          description: res.error || "No se pudo actualizar la información en Supabase.",
          variant: "destructive",
        });
        return;
      }

      // Actualizar registro en memoria y en la lista de resultados
      const updatedBoleta: BoletaHistoricaItem = {
        ...selectedBoleta,
        ap_paterno: formData.ap_paterno,
        ap_materno: formData.ap_materno,
        nombres: formData.nombres,
        fecha_nac: formData.fecha_nac,
        cargo: formData.cargo,
        cod_essalud: formData.cod_essalud,
        cuenta_banco: formData.cuenta_banco,
        leyenda_rd: formData.leyenda_rd,
        leyenda_mensual: formData.leyenda_mensual,
        sistema_pensionario: formData.sistema_pensionario,
        cussp: formData.cussp,
        fecha_afiliacion: formData.fecha_afiliacion,
        fecha_devengue: formData.fecha_devengue,
        monto_mensual: formData.monto_mensual,
        onp: formData.onp,
        prima: formData.prima,
        integra: formData.integra,
        profuturo: formData.profuturo,
        habitat: formData.habitat,
        total_dscto: formData.total_dscto,
        otros_dsctos: formData.otros_dsctos,
        dscto_entidades: formData.dscto_entidades,
        dscto_judicial: formData.dscto_judicial,
        total_liquido: formData.total_liquido,
        boleta_texto_personalizado: null,
      };

      setSelectedBoleta(updatedBoleta);
      setResults((prev) =>
        prev.map((b) => (b.boleta_id === selectedBoleta.boleta_id ? updatedBoleta : b))
      );

      toast({
        title: "¡Guardado exitoso en Base de Datos!",
        description: `Los datos de ${formData.nombres} ${formData.ap_paterno} (DNI ${formData.dni}) se actualizaron correctamente en las tablas de Supabase.`,
      });
    } finally {
      setSavingDb(false);
    }
  };


  return (
    <div className="space-y-4">
      {/* Buscador: texto libre + filtros opcionales (CAS, año, mes) */}
      {!hideSearch && (
        <div className="rounded-lg border border-slate-200 bg-white shadow-sm">
          <div className="flex items-center justify-between gap-2 border-b border-slate-100 px-4 py-3">
            <div>
              <h3 className="text-sm font-bold text-[#0b223d]">Buscar trabajador</h3>
              <p className="text-xs text-slate-500">Escribe nombres, apellidos o DNI, en cualquier orden.</p>
            </div>
          </div>

          <form onSubmit={handleSearch} className="space-y-4 p-4">
            <div className="flex flex-col gap-2 sm:flex-row">
              <div className="relative flex-1">
                <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                <Input
                  type="text"
                  autoFocus
                  placeholder="Ej: Sarita Florian, Judith Díaz o 41234567"
                  value={query}
                  onChange={(e) => {
                    setQuery(e.target.value);
                    clearDisplayedResults();
                  }}
                  className="h-11 bg-white pl-9 pr-10 text-base sm:text-sm"
                  aria-label="Buscar por nombres, apellidos o DNI"
                />
                {query && (
                  <button
                    type="button"
                    onClick={() => { setQuery(""); clearDisplayedResults(); }}
                    aria-label="Borrar búsqueda"
                    className="absolute right-1.5 top-1/2 flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                  >
                    <X className="h-4 w-4" />
                  </button>
                )}
              </div>
              <Button
                type="submit"
                disabled={loading || (!query.trim() && !categoryId && !dateFrom && !dateTo)}
                className="h-11 w-full bg-[#0d6efd] px-6 text-sm font-bold text-white hover:bg-[#0b5ed7] sm:w-auto"
              >
                {loading ? (
                  <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Buscando...</>
                ) : (
                  <><Search className="mr-2 h-4 w-4" /> Buscar</>
                )}
              </Button>
            </div>

            <div className="rounded-md bg-slate-50 p-3">
              <div className="mb-2 flex items-center justify-between">
                <p className="flex items-center gap-1.5 text-xs font-semibold text-slate-600">
                  <Sliders className="h-3.5 w-3.5 text-blue-600" /> Filtrar por (opcional)
                </p>
                {(categoryId || filterYear) && (
                  <button type="button" onClick={resetFilters} className="text-xs font-semibold text-blue-700 hover:underline">
                    Quitar filtros
                  </button>
                )}
              </div>
              <div className="grid gap-3 sm:grid-cols-3">
                <div className="space-y-1">
                  <Label className="text-xs text-slate-600">Tipo de CAS</Label>
                  <Select value={categoryId || "all"} onValueChange={(value) => { setCategoryId(value === "all" ? "" : value); clearDisplayedResults(); }}>
                    <SelectTrigger className="h-10 bg-white text-sm"><SelectValue placeholder="Todos" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Todos los CAS</SelectItem>
                      {CATEGORIAS_PLANILLA.map((category) => (
                        <SelectItem key={category.id} value={category.id}>{category.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1">
                  <Label className="text-xs text-slate-600">Año</Label>
                  <Select value={filterYear || "all"} onValueChange={(value) => applyPeriod(value === "all" ? "" : value, value === "all" ? "" : filterMonth)}>
                    <SelectTrigger className="h-10 bg-white text-sm"><SelectValue placeholder="Todos" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Todos los años</SelectItem>
                      {YEAR_OPTIONS.map((year) => (
                        <SelectItem key={year} value={year}>{year}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1">
                  <Label className="text-xs text-slate-600">Mes</Label>
                  <Select disabled={!filterYear} value={filterMonth || "all"} onValueChange={(value) => applyPeriod(filterYear, value === "all" ? "" : value)}>
                    <SelectTrigger className="h-10 bg-white text-sm"><SelectValue placeholder={filterYear ? "Todo el año" : "Elige un año"} /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Todo el año</SelectItem>
                      {MONTH_OPTIONS.map(([value, label]) => (
                        <SelectItem key={value} value={value}>{label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
            </div>
          </form>
        </div>
      )}

      {/* Skeletons mientras carga */}
      {loading && (
        <div className="grid gap-4 lg:grid-cols-[320px_1fr]">
          {/* Lista skeleton */}
          <div className="space-y-2">
            <Skeleton className="h-9 w-full rounded" />
            <div className="divide-y divide-slate-100 rounded border border-slate-200 bg-white">
              {Array.from({ length: 6 }).map((_, i) => (
                <div key={i} className="flex flex-col gap-2 p-3">
                  <Skeleton className="h-3.5 w-3/4 rounded" />
                  <Skeleton className="h-3 w-1/3 rounded" />
                  <div className="mt-1 flex items-center justify-between">
                    <Skeleton className="h-5 w-20 rounded-full" />
                    <Skeleton className="h-4 w-14 rounded" />
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Boleta skeleton */}
          <div className="rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="flex items-center justify-between border-b border-slate-200 p-4">
              <div className="space-y-2">
                <Skeleton className="h-6 w-28 rounded-md" />
                <Skeleton className="h-5 w-48 rounded" />
                <Skeleton className="h-3.5 w-32 rounded" />
              </div>
              <Skeleton className="h-9 w-44 rounded-lg" />
            </div>
            <div className="space-y-3 p-6">
              {Array.from({ length: 20 }).map((_, i) => (
                <Skeleton
                  key={i}
                  className="h-3 rounded"
                  style={{ width: `${55 + Math.sin(i * 1.7) * 35}%` }}
                />
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Mensaje de Sin Resultados */}
      {hasSearched && results.length === 0 && !loading && (
        <div className="text-center py-8 bg-white border border-slate-300 rounded p-6">
          <Users className="mx-auto h-8 w-8 text-slate-400 mb-2" />
          <h4 className="text-sm font-bold text-slate-800">
            No se encontraron registros coincidentes
          </h4>
          <p className="text-xs text-slate-500 mt-1">
            Revisa la ortografía o prueba con menos palabras o sin filtros.
          </p>
        </div>
      )}

      {/* Resultados y Previsualización */}
      {results.length > 0 && !loading && (
        <div className="space-y-3">
          <div className="flex flex-col gap-2 rounded border border-slate-300 bg-white p-3 shadow-sm sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-3 text-xs">
              <span className="font-bold text-slate-800">{results.length} boletas encontradas</span>
              <label className="inline-flex cursor-pointer items-center gap-1.5 text-slate-600">
                <input type="checkbox" checked={selectedIds.length === results.length} onChange={(event) => setSelectedIds(event.target.checked ? results.map((item) => item.boleta_id) : [])} className="h-4 w-4 rounded border-slate-300 accent-blue-700" />
                Seleccionar todas
              </label>
              <span className="text-slate-500">{selectedIds.length} seleccionadas</span>
            </div>
            <div className="flex flex-wrap gap-2">
              <Button size="sm" variant="outline" onClick={handlePrintSelected} disabled={!selectedIds.length} className="h-8 text-xs">
                <Printer className="mr-1.5 h-3.5 w-3.5" /> Imprimir seleccionadas
              </Button>
              <Button size="sm" onClick={handleDownloadSelected} disabled={!selectedIds.length} className="h-8 bg-[#dc3545] text-xs text-white hover:bg-[#bb2d3b]">
                <Download className="mr-1.5 h-3.5 w-3.5" /> PDF combinado
              </Button>
            </div>
          </div>
          <div className="grid gap-4 lg:grid-cols-[320px_1fr]">
          {/* Lista de Resultados (List-Group) */}
          <div className="space-y-2">
            <div className="bg-slate-200 text-slate-800 px-3 py-2 rounded-t text-xs font-bold border border-slate-300 flex items-center justify-between">
              <span>Resultados ({results.length})</span>
              <span className="text-[11px] text-slate-600 font-normal">Haga clic para ver</span>
            </div>

            <div className="bg-white border border-slate-300 rounded-b divide-y divide-slate-200 max-h-[600px] overflow-y-auto">
              {results.map((item) => {
                const isSelected = item.boleta_id === selectedBoleta?.boleta_id;
                return (
                  <div
                    key={item.boleta_id}
                    className={`flex items-start gap-2 p-3 text-xs transition ${
                      isSelected
                        ? "bg-[#e7f1ff] text-[#084298] font-bold border-l-4 border-l-[#0d6efd]"
                        : "hover:bg-slate-50 text-slate-800"
                    }`}
                  >
                    <input type="checkbox" aria-label={`Seleccionar boleta de ${item.ap_paterno} ${item.ap_materno}, ${item.mes} ${item.anio}`} checked={selectedIds.includes(item.boleta_id)} onChange={() => toggleBoleta(item.boleta_id)} className="mt-1 h-4 w-4 shrink-0 rounded border-slate-300 accent-blue-700" />
                    <button type="button" onClick={() => setSelectedBoleta(item)} className="min-w-0 flex-1 text-left">
                    <div className="flex items-start justify-between">
                      <div className="min-w-0">
                        <p className="font-bold text-slate-900 truncate">
                          {item.ap_paterno} {item.ap_materno}, {item.nombres}
                        </p>
                        <p className="text-[11px] text-slate-500 font-mono">
                          DNI: {item.dni}
                        </p>
                      </div>
                      <span className="bg-slate-100 border border-slate-200 px-1.5 py-0.5 rounded text-[10px] font-semibold uppercase text-slate-600">
                        {item.categoria_label}
                      </span>
                    </div>

                    <div className="mt-2 flex items-center justify-between border-t border-slate-100 pt-1.5 text-[11px]">
                      <span className="rounded bg-[#0b223d] px-2 py-0.5 text-[11px] font-extrabold uppercase tracking-wide text-white">
                        {item.mes} {item.anio}
                      </span>
                      <span className="font-mono font-bold text-emerald-700">
                        S/. {item.total_liquido}
                      </span>
                    </div>
                    </button>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Boleta seleccionada: Imprimir (por defecto) o Editar datos */}
          {selectedBoleta && (
            <div className="min-w-0 overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
              <div className="flex flex-col gap-3 border-b border-slate-200 p-4 sm:flex-row sm:items-center sm:justify-between">
                <div className="min-w-0">
                  <div className="mb-1.5 inline-flex items-center gap-1.5 rounded-md bg-[#0b223d] px-2.5 py-1 text-xs font-extrabold uppercase tracking-wide text-white">
                    <Calendar className="h-3.5 w-3.5 text-amber-400" />
                    {selectedBoleta.mes} {selectedBoleta.anio}
                  </div>
                  <h4 className="truncate text-base font-bold text-slate-900">
                    {selectedBoleta.nombres} {selectedBoleta.ap_paterno} {selectedBoleta.ap_materno}
                  </h4>
                  <p className="text-xs text-slate-500">
                    DNI {selectedBoleta.dni} · {selectedBoleta.categoria_label}
                  </p>
                </div>

                <div role="tablist" aria-label="Modo de la boleta" className="inline-flex shrink-0 rounded-lg bg-slate-100 p-1">
                  {([
                    { id: "preview", label: "Imprimir", icon: Printer },
                    { id: "form", label: "Editar datos", icon: Pencil },
                  ] as const).map(({ id, label, icon: Icon }) => (
                    <button
                      key={id}
                      type="button"
                      role="tab"
                      aria-selected={subTab === id}
                      onClick={() => setSubTab(id)}
                      className={`inline-flex items-center gap-1.5 rounded-md px-3.5 py-1.5 text-xs font-bold transition ${
                        subTab === id ? "bg-white text-[#0b223d] shadow-sm" : "text-slate-500 hover:text-slate-800"
                      }`}
                    >
                      <Icon className="h-3.5 w-3.5" />
                      {label}
                    </button>
                  ))}
                </div>
              </div>

              {/* MODO IMPRIMIR */}
              {subTab === "preview" && (
                <div>
                  <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 bg-slate-50 px-4 py-2.5">
                    <div className="flex flex-wrap gap-2">
                      <Button size="sm" onClick={handlePrint} className="h-9 bg-[#0d6efd] text-xs font-bold text-white hover:bg-[#0b5ed7]">
                        <Printer className="mr-1.5 h-3.5 w-3.5" /> Imprimir
                      </Button>
                      <Button size="sm" variant="outline" onClick={handlePDF} className="h-9 text-xs font-bold">
                        <Download className="mr-1.5 h-3.5 w-3.5" /> Descargar PDF
                      </Button>
                    </div>
                    <div className="flex items-center gap-1">
                      <Button type="button" variant="outline" size="sm" onClick={() => setZoom((z) => Math.max(70, z - 10))} className="h-8 w-8 p-0" aria-label="Alejar">
                        <ZoomOut className="h-3.5 w-3.5" />
                      </Button>
                      <button type="button" onClick={() => setZoom(100)} className="w-12 text-center font-mono text-xs font-bold text-slate-600 hover:text-slate-900" title="Restablecer zoom">
                        {zoom}%
                      </button>
                      <Button type="button" variant="outline" size="sm" onClick={() => setZoom((z) => Math.min(160, z + 10))} className="h-8 w-8 p-0" aria-label="Acercar">
                        <ZoomIn className="h-3.5 w-3.5" />
                      </Button>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center justify-between gap-2 px-4 pt-3">
                    <p className="text-xs text-slate-500">Puedes editar el texto libremente antes de imprimir. Los cambios aquí <strong>no</strong> se guardan en la base de datos.</p>
                    {editedBoletaText !== liveBoletaText ? (
                      <button
                        type="button"
                        onClick={() => setEditedBoletaText(liveBoletaText)}
                        className="inline-flex items-center gap-1.5 rounded-md border border-amber-400 bg-amber-50 px-3 py-1.5 text-xs font-bold text-amber-900 shadow-sm hover:bg-amber-100"
                      >
                        <RotateCcw className="h-3.5 w-3.5" /> Restaurar texto
                      </button>
                    ) : (
                      <span className="text-xs text-slate-400 italic">Sin cambios en el texto</span>
                    )}
                  </div>

                  <div className="overflow-x-auto bg-slate-50 p-4">
                    <textarea
                      value={editedBoletaText}
                      onChange={(e) => setEditedBoletaText(e.target.value)}
                      rows={Math.max(45, (editedBoletaText || "").split("\n").length + 2)}
                      style={{
                        fontSize: `${(11 * zoom) / 100}px`,
                        lineHeight: 1.35,
                        width: `${Math.round(80 * (zoom / 100))}ch`,
                        minWidth: "68ch",
                      }}
                      aria-label="Texto de la boleta a imprimir"
                      className="mx-auto block resize-none whitespace-pre rounded border border-slate-300 bg-white p-6 font-mono text-black shadow-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                      spellCheck={false}
                    />
                  </div>
                </div>
              )}

              {/* MODO EDITAR (base de datos) */}
              {subTab === "form" && formData && (
                <div>
                  <p className="px-4 pt-3 text-xs text-slate-500">
                    Los cambios se guardan de forma permanente en la base de datos al presionar <strong>Guardar cambios</strong>.
                  </p>
                  <BoletaFormEditor
                    formData={formData}
                    onChange={handleFieldChange}
                    onAutoCalculate={handleAutoCalculate}
                  />
                  <div className="sticky bottom-0 flex flex-wrap items-center justify-end gap-2 border-t border-slate-200 bg-white/95 px-4 py-3 backdrop-blur">
                    <Button type="button" variant="outline" size="sm" onClick={handleResetOriginalFields} className="h-9 text-xs font-semibold" title="Descartar los cambios que aún no guardaste">
                      <RotateCcw className="mr-1.5 h-3.5 w-3.5" /> Deshacer cambios
                    </Button>
                    <Button size="sm" onClick={handleSaveToDatabase} disabled={savingDb} className="h-9 bg-[#198754] text-xs font-bold text-white hover:bg-[#157347]">
                      {savingDb ? (
                        <><Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> Guardando...</>
                      ) : (
                        <><Save className="mr-1.5 h-3.5 w-3.5" /> Guardar en BD</>
                      )}
                    </Button>
                  </div>
                </div>
              )}
            </div>
          )}
          </div>
        </div>
      )}

      {/* Portal de Impresión Limpia (1 sola página A4) */}
      <PrintBoletaPortal text={printTexts.length ? printTexts : editedBoletaText || liveBoletaText} />
    </div>
  );
}
