import { useState, useMemo, useEffect } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
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
import { TutorialDialog } from "./TutorialDialog";
import {
  Search,
  Printer,
  Download,
  Building,
  Calendar,
  X,
  Loader2,
  User,
  Users,
  FileText,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  Save,
  CheckCircle2,
  Sliders,
} from "lucide-react";

export function AdminVentanilla() {
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
  const [subTab, setSubTab] = useState<"form" | "preview">("form");
  const [formData, setFormData] = useState<BoletaFormData | null>(null);
  const [zoom, setZoom] = useState(100);
  const [savingDb, setSavingDb] = useState(false);
  const [editedBoletaText, setEditedBoletaText] = useState("");

  const clearDisplayedResults = () => {
    setResults([]);
    setSelectedBoleta(null);
    setSelectedIds([]);
    setHasSearched(false);
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
      otrosDsctos: "0.00",
      dsctoEntidades: "0.00",
      dsctoJudicial: "0.00",
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
    const totalDscto = (onp + prima + integra + profuturo + habitat).toFixed(2);
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
      {/* Buscador de Ventanilla (Panel Estilo Bootstrap) */}
      <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        <div className="bg-[#0b223d] text-white px-4 py-2.5 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Users className="h-4 w-4 text-amber-400" />
            <h3 className="text-xs font-bold uppercase tracking-wider">
              Consultar · Búsqueda General de Trabajadores
            </h3>
          </div>
          <span className="text-[11px] text-slate-300">
            Sin clave requerida
          </span>
          <TutorialDialog audience="admin" />
        </div>

        <div className="p-4 bg-slate-50 border-b border-slate-200">
          <form onSubmit={handleSearch} className="space-y-3">
            <div className="flex flex-col sm:flex-row items-center gap-2">
            <div className="relative flex-1 w-full">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
              <Input
                type="text"
                placeholder="Escriba DNI, Apellidos o Nombres del trabajador..."
                value={query}
                onChange={(e) => {
                  setQuery(e.target.value);
                  clearDisplayedResults();
                }}
                className="h-10 rounded border-slate-300 pl-9 pr-8 text-sm bg-white focus:border-blue-600"
              />
              {(query || categoryId || dateFrom || dateTo) && (
                <button
                  type="button"
                  onClick={() => {
                    setQuery("");
                    setCategoryId("");
                    setDateFrom("");
                    setDateTo("");
                    clearDisplayedResults();
                  }}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X className="h-4 w-4" />
                </button>
              )}
            </div>

            <Button
              type="submit"
                disabled={loading || (!query.trim() && !categoryId && !dateFrom && !dateTo)}
              className="h-10 px-5 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] font-bold text-white text-xs shadow-sm shrink-0 w-full sm:w-auto"
            >
              {loading ? (
                <>
                  <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> Buscando...
                </>
              ) : (
                <>
                  <Search className="mr-1.5 h-3.5 w-3.5" /> Buscar en Base de Datos
                </>
              )}
            </Button>
            </div>

            <div className="rounded border border-slate-200 bg-white p-3">
              <div className="mb-2 flex items-center gap-2 text-xs font-bold text-slate-700">
                <Sliders className="h-3.5 w-3.5 text-blue-700" /> Filtros avanzados
              </div>
              <div className="grid gap-3 sm:grid-cols-3">
                <div className="space-y-1">
                  <Label className="text-[11px] font-semibold text-slate-600">Tipo de CAS</Label>
                  <Select value={categoryId || "all"} onValueChange={(value) => { setCategoryId(value === "all" ? "" : value); clearDisplayedResults(); }}>
                    <SelectTrigger className="h-9 text-xs"><SelectValue placeholder="Todos los CAS" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Todos los CAS</SelectItem>
                      {CATEGORIAS_PLANILLA.map((category) => (
                        <SelectItem key={category.id} value={category.id}>{category.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-1">
                  <Label htmlFor="admin-date-from" className="text-[11px] font-semibold text-slate-600">Desde</Label>
                  <Input id="admin-date-from" type="date" value={dateFrom} onChange={(event) => { setDateFrom(event.target.value); clearDisplayedResults(); }} className="h-9 text-xs" />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="admin-date-to" className="text-[11px] font-semibold text-slate-600">Hasta</Label>
                  <Input id="admin-date-to" type="date" value={dateTo} onChange={(event) => { setDateTo(event.target.value); clearDisplayedResults(); }} className="h-9 text-xs" />
                </div>
              </div>
              {(categoryId || dateFrom || dateTo) && (
                <button
                  type="button"
                  onClick={() => { setCategoryId(""); setDateFrom(""); setDateTo(""); clearDisplayedResults(); }}
                  className="mt-2 text-[11px] font-semibold text-blue-700 hover:underline"
                >Limpiar filtros</button>
              )}
            </div>
          </form>

          <p className="mt-1.5 text-[11px] text-slate-500">
            * Busca por persona, tipo de CAS o periodo. Puedes combinar los filtros; el rango incluye los meses seleccionados.
          </p>
        </div>
      </div>

      {/* Mensaje de Sin Resultados */}
      {hasSearched && results.length === 0 && !loading && (
        <div className="text-center py-8 bg-white border border-slate-300 rounded p-6">
          <Users className="mx-auto h-8 w-8 text-slate-400 mb-2" />
          <h4 className="text-sm font-bold text-slate-800">
            No se encontraron registros coincidentes
          </h4>
          <p className="text-xs text-slate-500 mt-1">
            Verifique la búsqueda o pruebe con otro tipo de CAS o rango de fechas.
          </p>
        </div>
      )}

      {/* Resultados y Previsualización */}
      {results.length > 0 && (
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

                    <div className="mt-2 flex items-center justify-between text-[11px] pt-1.5 border-t border-slate-100">
                      <span className="text-slate-600 font-medium">
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

          {/* Panel de Visualización e Impresión */}
          {selectedBoleta && (
            <div className="space-y-3">
              <div className="bg-white border border-slate-300 rounded p-3 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                <div>
                  <h4 className="font-bold text-sm text-slate-900">
                    {selectedBoleta.nombres} {selectedBoleta.ap_paterno} {selectedBoleta.ap_materno}
                  </h4>
                  <p className="text-xs text-slate-600">
                    DNI: <strong>{selectedBoleta.dni}</strong> · Periodo: <strong>{selectedBoleta.mes} {selectedBoleta.anio}</strong> ({selectedBoleta.categoria_label})
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <Button
                    size="sm"
                    onClick={handleSaveToDatabase}
                    disabled={savingDb}
                    className="h-9 rounded bg-[#198754] hover:bg-[#157347] text-white text-xs font-semibold shadow-sm"
                    title="Guardar de forma permanente los cambios en la Base de Datos para esta boleta"
                  >
                    {savingDb ? (
                      <>
                        <Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" /> Guardando...
                      </>
                    ) : (
                      <>
                        <Save className="h-3.5 w-3.5 mr-1" /> Guardar en BD
                      </>
                    )}
                  </Button>
                  <Button
                    size="sm"
                    onClick={handlePrint}
                    className="h-9 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white text-xs font-semibold shadow-sm"
                  >
                    <Printer className="h-3.5 w-3.5 mr-1" /> Imprimir Boleta
                  </Button>
                  <Button
                    size="sm"
                    onClick={handlePDF}
                    className="h-9 rounded bg-[#dc3545] hover:bg-[#bb2d3b] text-white text-xs font-semibold shadow-sm"
                  >
                    <Download className="h-3.5 w-3.5 mr-1" /> Descargar PDF
                  </Button>
                </div>
              </div>

              {/* Barra de Pestañas de Vista: Formulario vs Vista Previa */}
              <div className="bg-white border border-slate-300 rounded overflow-hidden shadow-sm">
                <div className="bg-slate-100 border-b border-slate-300 px-3 pt-2 flex flex-wrap items-center justify-between gap-2">
                  <div className="flex space-x-1">
                    <button
                      type="button"
                      onClick={() => setSubTab("form")}
                      className={`px-3.5 py-1.5 text-xs font-bold rounded-t border-t border-l border-r transition flex items-center gap-1.5 ${
                        subTab === "form"
                          ? "bg-white text-[#0b223d] border-slate-300 -mb-[1px] shadow-sm font-extrabold"
                          : "bg-slate-200/80 text-slate-600 border-transparent hover:text-slate-900"
                      }`}
                    >
                      <Sliders className="h-3.5 w-3.5 text-blue-600" />
                      <span>Formulario de Edición (Por Campos)</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setSubTab("preview")}
                      className={`px-3.5 py-1.5 text-xs font-bold rounded-t border-t border-l border-r transition flex items-center gap-1.5 ${
                        subTab === "preview"
                          ? "bg-white text-[#0b223d] border-slate-300 -mb-[1px] shadow-sm font-extrabold"
                          : "bg-slate-200/80 text-slate-600 border-transparent hover:text-slate-900"
                      }`}
                    >
                      <FileText className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Vista Previa Oficial (Impresión Courier)</span>
                    </button>
                  </div>

                  <div className="flex items-center gap-2 pb-1.5 sm:pb-0">
                    <Button
                      type="button"
                      variant="outline"
                      size="sm"
                      onClick={handleResetOriginalFields}
                      className="h-6 px-2 text-[10px] text-amber-800 border-amber-300 bg-amber-50 hover:bg-amber-100 rounded font-semibold"
                      title="Restablecer los valores originales de la planilla"
                    >
                      <RotateCcw className="h-2.5 w-2.5 mr-1" /> Revertir datos originales
                    </Button>
                  </div>
                </div>

                {/* Vista 1: Formulario Estructurado por Campos */}
                {subTab === "form" && formData && (
                  <div className="p-3 bg-slate-50">
                    <div className="bg-[#f0f7ff] border border-[#d0e3ff] p-2.5 rounded mb-3 text-[11px] text-[#084298] flex items-center justify-between">
                      <span>
                        Modifique los campos correspondientes (como Fecha de Nacimiento, Cargo, Leyenda, etc.) y presione el botón verde <strong>"Guardar en BD"</strong> para grabarlos en la base de datos de Supabase.
                      </span>
                    </div>

                    <BoletaFormEditor
                      formData={formData}
                      onChange={handleFieldChange}
                      onAutoCalculate={handleAutoCalculate}
                    />
                  </div>
                )}

                {/* Vista 2: Visor Oficial Courier con Zoom y Textarea */}
                {subTab === "preview" && (
                  <div>
                    {/* Controles de Zoom */}
                    <div className="bg-slate-50 border-b border-slate-300 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-700 font-bold">
                      <div className="flex items-center gap-2">
                        <span>FORMATO OFICIAL IMPRESIÓN · UGEL Nº 04 TSE</span>
                        <span className="bg-[#e7f1ff] border border-[#b6d4fe] text-[#084298] text-[10px] px-1.5 py-0.5 rounded font-semibold">
                          Hoja A4
                        </span>
                      </div>

                      <div className="flex items-center gap-1 font-sans">
                        <span className="text-[11px] text-slate-500 mr-1 font-normal">Zoom:</span>
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          onClick={() => setZoom((z) => Math.max(70, z - 10))}
                          className="h-6 w-6 p-0 text-xs rounded border-slate-300 bg-white"
                          title="Alejar (Zoom Out)"
                        >
                          <ZoomOut className="h-3 w-3" />
                        </Button>
                        <span className="text-xs font-mono font-bold w-12 text-center text-slate-700">
                          {zoom}%
                        </span>
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          onClick={() => setZoom((z) => Math.min(160, z + 10))}
                          className="h-6 w-6 p-0 text-xs rounded border-slate-300 bg-white"
                          title="Acercar (Zoom In)"
                        >
                          <ZoomIn className="h-3 w-3" />
                        </Button>
                        <Button
                          type="button"
                          variant="ghost"
                          size="sm"
                          onClick={() => setZoom(100)}
                          className="h-6 px-1.5 text-[11px] text-slate-600 hover:text-slate-900 rounded font-normal"
                          title="Restablecer tamaño normal (100%)"
                        >
                          100%
                        </Button>
                      </div>
                    </div>

                    <div className="bg-[#f0f7ff] border-b border-[#d0e3ff] px-4 py-1.5 text-[11px] text-[#084298]">
                      <span>Vista previa exacta en Courier New monoespaciado. Cualquier modificación en el formulario se actualiza aquí automáticamente.</span>
                    </div>

                    <div className="p-4 sm:p-6 overflow-x-auto bg-[#fafafa]">
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
                        className="font-mono text-black whitespace-pre bg-white p-6 rounded border border-slate-300 shadow-sm mx-auto block resize-none focus:outline-none focus:ring-1 focus:ring-blue-500"
                        spellCheck={false}
                      />
                    </div>
                  </div>
                )}
              </div>
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
