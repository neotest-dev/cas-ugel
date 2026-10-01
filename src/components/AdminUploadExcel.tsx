import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  CATEGORIAS_PLANILLA,
  Worker,
  buildBoletaText,
  parsePeriodFromFilename,
} from "@/lib/boleta";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
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
import { Progress } from "@/components/ui/progress";
import {
  ChevronLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  FileUp,
  Loader2,
  Printer,
  Search,
  Upload,
  X,
  Database,
  CheckCircle2,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { toast } from "@/hooks/use-toast";
import { exportBoletaToPDF } from "@/lib/pdfExport";
import { savePlanillaToDatabase } from "@/lib/payrollService";
import { PrintBoletaPortal } from "./PrintBoletaPortal";
import ExcelWorker from "../workers/excelWorker.ts?worker";

const HEAVY_FILE_BYTES = 3 * 1024 * 1024;

interface AdminUploadExcelProps {
  onPlanillaSaved?: () => void;
}

export function AdminUploadExcel({ onPlanillaSaved }: AdminUploadExcelProps) {
  const [workers, setWorkers] = useState<Worker[]>([]);
  const [period, setPeriod] = useState<{ mes: string; anio: string }>({
    mes: "",
    anio: "",
  });
  const [activeIdx, setActiveIdx] = useState(0);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("Procesando planilla...");
  const [search, setSearch] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [selectedCategoryId, setSelectedCategoryId] = useState("");
  const [currentFilename, setCurrentFilename] = useState("");
  const [editedBoletaText, setEditedBoletaText] = useState("");
  const [zoom, setZoom] = useState(100);

  const [savingToDb, setSavingToDb] = useState(false);
  const [saveProgress, setSaveProgress] = useState(0);
  const [saveStepMsg, setSaveStepMsg] = useState("");
  const [showOverwriteModal, setShowOverwriteModal] = useState(false);
  const [pendingOverwriteMsg, setPendingOverwriteMsg] = useState("");

  const fileRef = useRef<HTMLInputElement>(null);
  const workerRef = useRef<globalThis.Worker | null>(null);

  const selectedCategory = useMemo(
    () =>
      CATEGORIAS_PLANILLA.find(
        (category) => category.id === selectedCategoryId,
      ) ?? null,
    [selectedCategoryId],
  );

  const getWorker = useCallback(() => {
    if (!workerRef.current) {
      workerRef.current = new ExcelWorker();
    }
    return workerRef.current;
  }, []);

  useEffect(() => {
    return () => {
      workerRef.current?.terminate();
      workerRef.current = null;
    };
  }, []);

  const clearLoadedState = useCallback(() => {
    setWorkers([]);
    setActiveIdx(0);
    setSearch("");
    setPeriod({ mes: "", anio: "" });
    setEditedBoletaText("");
    setSelectedCategoryId("");
    setCurrentFilename("");
  }, []);

  const handleFile = useCallback(
    async (file: File) => {
      setLoading(true);
      setCurrentFilename(file.name);
      setLoadingMsg(
        file.size > HEAVY_FILE_BYTES
          ? "Archivo pesado detectado, procesando registros..."
          : "Leyendo hoja de planilla CAS...",
      );

      try {
        const buffer = await file.arrayBuffer();
        const worker = getWorker();
        const result = await new Promise<{
          ok: boolean;
          workers?: Worker[];
          period?: { mes: string; anio: string };
          categoryId?: string | null;
          debug?: any;
          error?: string;
        }>((resolve) => {
          const onMessage = (event: MessageEvent) => {
            worker.removeEventListener("message", onMessage);
            resolve(event.data);
          };

          worker.addEventListener("message", onMessage);
          worker.postMessage({ buffer, filename: file.name }, [buffer]);
        });

        if (!result.ok) {
          throw new Error(result.error || "Error desconocido al procesar el archivo Excel.");
        }

        const detectedCategoryId = result.categoryId ?? null;
        if (!detectedCategoryId) {
          toast({
            title: "Categoría no identificada",
            description: "Por favor seleccione la categoría CAS en el menú desplegable.",
            variant: "destructive",
          });
        } else {
          setSelectedCategoryId(detectedCategoryId);
        }

        const nextWorkers = result.workers || [];
        if (!nextWorkers.length) {
          toast({
            title: "Planilla sin trabajadores",
            description: "No se encontraron registros de trabajadores en la hoja procesada.",
            variant: "destructive",
          });
          return;
        }

        setWorkers(nextWorkers);
        setPeriod(result.period || parsePeriodFromFilename(file.name));
        setActiveIdx(0);
        toast({
          title: "Planilla cargada",
          description: `Se detectaron ${nextWorkers.length} trabajadores listos para publicar.`,
        });
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Error desconocido";
        toast({
          title: "Error al procesar archivo",
          description: errorMessage,
          variant: "destructive",
        });
      } finally {
        setLoading(false);
      }
    },
    [getWorker],
  );

  const onDrop = (event: React.DragEvent) => {
    event.preventDefault();
    setDragOver(false);
    if (loading) return;
    const file = event.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  const filtered = useMemo(() => {
    if (!search.trim()) return workers;
    const query = search.toLowerCase();
    return workers.filter((worker) =>
      `${worker.apPaterno} ${worker.apMaterno} ${worker.nombres} ${worker.dni}`
        .toLowerCase()
        .includes(query),
    );
  }, [workers, search]);

  useEffect(() => {
    if (!search.trim() || !filtered.length) return;
    const firstMatchIdx = workers.indexOf(filtered[0]);
    if (firstMatchIdx >= 0 && firstMatchIdx !== activeIdx) {
      setActiveIdx(firstMatchIdx);
    }
  }, [activeIdx, filtered, search, workers]);

  const active = workers[activeIdx];
  const boletaText = useMemo(
    () => (active ? buildBoletaText(active, period.mes, period.anio) : ""),
    [active, period],
  );

  useEffect(() => {
    setEditedBoletaText(boletaText);
  }, [boletaText]);

  const handlePrev = () => setActiveIdx((index) => Math.max(0, index - 1));
  const handleNext = () =>
    setActiveIdx((index) => Math.min(workers.length - 1, index + 1));

  const handlePrint = () => window.print();

  const handlePDF = () => {
    if (!active || !editedBoletaText) return;
    exportBoletaToPDF(active, editedBoletaText);
  };

  const handleSaveToDatabase = async (overwrite = false) => {
    if (!workers.length || !selectedCategoryId || !period.mes || !period.anio) {
      toast({
        title: "Datos incompletos",
        description: "Asegúrese de seleccionar la categoría y el periodo correspondiente.",
        variant: "destructive",
      });
      return;
    }

    setSavingToDb(true);
    setSaveProgress(5);
    setSaveStepMsg("Iniciando guardado...");

    try {
      const res = await savePlanillaToDatabase({
        filename: currentFilename || "planilla.xlsx",
        categoriaId: selectedCategoryId,
        period,
        workers,
        overwrite,
        onProgress: (step, percent) => {
          setSaveStepMsg(step);
          setSaveProgress(percent);
        },
      });

      if (!res.ok) {
        if (res.error && res.error.includes("Ya existe una planilla registrada")) {
          setPendingOverwriteMsg(res.error);
          setShowOverwriteModal(true);
          return;
        }

        toast({
          title: "Error al guardar en Base de Datos",
          description: res.error || "Ocurrió un error inesperado.",
          variant: "destructive",
        });
        return;
      }

      toast({
        title: "Planilla Guardada con Éxito",
        description: `Se publicaron ${res.totalSaved} boletas para ${selectedCategory?.label} (${period.mes} ${period.anio}).`,
      });

      onPlanillaSaved?.();
    } finally {
      setSavingToDb(false);
    }
  };

  // 1. Vista de Subida (Estilo Panel de Carga Bootstrap)
  if (!workers.length) {
    return (
      <div className="mx-auto max-w-2xl py-2">
        <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
          <div className="bg-[#0b223d] text-white px-4 py-3 border-b border-slate-300 flex items-center justify-between">
            <h3 className="text-xs font-bold uppercase tracking-wider flex items-center gap-2">
              <FileSpreadsheet className="h-4 w-4 text-amber-400" />
              Módulo de Importación y Publicación de Planillas
            </h3>
            <span className="text-[11px] text-slate-300">Formatos .xlsx, .xls</span>
          </div>

          <div className="p-6">
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current?.click()}
              className={`border-2 border-dashed p-8 text-center rounded cursor-pointer transition ${
                dragOver
                  ? "border-[#0d6efd] bg-[#f0f7ff]"
                  : "border-slate-300 bg-slate-50 hover:bg-slate-100 hover:border-slate-400"
              }`}
            >
              {loading ? (
                <div className="flex flex-col items-center justify-center py-4 space-y-2">
                  <Loader2 className="h-8 w-8 animate-spin text-[#0d6efd]" />
                  <p className="text-sm font-semibold text-slate-800">{loadingMsg}</p>
                  <p className="text-xs text-slate-500">
                    Extrayendo trabajadores, remuneraciones y aportes pensionarios...
                  </p>
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center space-y-3">
                  <div className="h-12 w-12 rounded bg-white border border-slate-300 flex items-center justify-center text-[#0b223d] shadow-sm">
                    <Upload className="h-6 w-6" />
                  </div>
                  <div>
                    <p className="text-sm font-bold text-slate-800">
                      Seleccione o arrastre el archivo Excel de la planilla
                    </p>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Hojas compatibles: CAS-SEDE, CAS JEC, ORQUESTANDO, etc.
                    </p>
                  </div>
                  <Button
                    type="button"
                    className="h-9 px-4 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white text-xs font-bold shadow-sm"
                  >
                    Examinar archivo en este equipo
                  </Button>
                </div>
              )}

              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.xlsm"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) handleFile(file);
                  e.target.value = "";
                }}
              />
            </div>

            <div className="mt-4 pt-3 border-t border-slate-200 grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs text-slate-600">
              <div className="flex items-center gap-1.5 bg-slate-50 p-2 rounded border border-slate-200">
                <CheckCircle2 className="h-3.5 w-3.5 text-emerald-600 shrink-0" />
                <span>8 Categorías CAS</span>
              </div>
              <div className="flex items-center gap-1.5 bg-slate-50 p-2 rounded border border-slate-200">
                <Database className="h-3.5 w-3.5 text-blue-600 shrink-0" />
                <span>Almacenamiento SQL</span>
              </div>
              <div className="flex items-center gap-1.5 bg-slate-50 p-2 rounded border border-slate-200">
                <FileSpreadsheet className="h-3.5 w-3.5 text-purple-600 shrink-0" />
                <span>Control de Duplicados</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // 2. Vista de Previsualización y Publicación en BD
  return (
    <div className="space-y-4">
      {/* Barra de Control y Botones Bootstrap */}
      <div className="bg-white border border-slate-300 rounded shadow-sm p-4 space-y-3">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 border-b border-slate-200 pb-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="bg-[#e7f1ff] border border-[#b6d4fe] text-[#084298] text-[10px] font-bold px-2 py-0.5 rounded uppercase">
                Planilla Leída
              </span>
              <span className="text-xs text-slate-500 font-mono truncate max-w-[250px]">
                {currentFilename}
              </span>
            </div>
            <h3 className="text-base font-bold text-slate-900 mt-1">
              {selectedCategory?.label ?? "Categoría"} · {period.mes} {period.anio}
            </h3>
            <p className="text-xs text-slate-500">
              Total de registros: <strong>{workers.length} trabajadores</strong> detectados.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <Button
              size="sm"
              onClick={() => handleSaveToDatabase(false)}
              disabled={savingToDb}
              className="h-9 rounded bg-[#198754] hover:bg-[#157347] text-white font-bold text-xs shadow-sm"
            >
              {savingToDb ? (
                <>
                  <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> Guardando en BD...
                </>
              ) : (
                <>
                  <Database className="mr-1.5 h-3.5 w-3.5" /> Publicar en Base de Datos
                </>
              )}
            </Button>

            <Button
              size="sm"
              onClick={handlePrint}
              className="h-9 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white font-semibold text-xs shadow-sm"
            >
              <Printer className="mr-1 h-3.5 w-3.5" /> Imprimir
            </Button>

            <Button
              size="sm"
              onClick={handlePDF}
              className="h-9 rounded bg-[#dc3545] hover:bg-[#bb2d3b] text-white font-semibold text-xs shadow-sm"
            >
              <Download className="mr-1 h-3.5 w-3.5" /> PDF
            </Button>

            <Button
              size="sm"
              variant="outline"
              onClick={clearLoadedState}
              className="h-9 rounded border-slate-300 text-slate-700 hover:bg-slate-100 text-xs font-semibold"
            >
              <FileUp className="mr-1 h-3.5 w-3.5" /> Descartar
            </Button>
          </div>
        </div>

        {/* Progreso de Guardado */}
        {savingToDb && (
          <div className="bg-[#e8f5e9] border border-[#c8e6c9] rounded p-2.5 space-y-1.5">
            <div className="flex items-center justify-between text-xs font-semibold text-[#1b5e20]">
              <span>{saveStepMsg}</span>
              <span>{saveProgress}%</span>
            </div>
            <Progress value={saveProgress} className="h-2 bg-[#c8e6c9]" />
          </div>
        )}

        {/* Formulario de Validación: Categoría, Periodo y Filtro */}
        <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4 pt-1">
          <div className="space-y-1">
            <label className="text-[11px] font-bold text-slate-700">Categoría CAS:</label>
            <Select
              value={selectedCategoryId}
              onValueChange={setSelectedCategoryId}
            >
              <SelectTrigger className="h-9 rounded border-slate-300 bg-white text-xs">
                <SelectValue placeholder="Seleccione categoría" />
              </SelectTrigger>
              <SelectContent>
                {CATEGORIAS_PLANILLA.map((cat) => (
                  <SelectItem key={cat.id} value={cat.id} className="text-xs">
                    {cat.label}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div className="space-y-1">
            <label className="text-[11px] font-bold text-slate-700">Mes:</label>
            <Input
              value={period.mes}
              onChange={(e) => setPeriod({ ...period, mes: e.target.value.toUpperCase() })}
              className="h-9 rounded border-slate-300 bg-white text-xs font-bold"
            />
          </div>

          <div className="space-y-1">
            <label className="text-[11px] font-bold text-slate-700">Año:</label>
            <Input
              value={period.anio}
              onChange={(e) => setPeriod({ ...period, anio: e.target.value })}
              className="h-9 rounded border-slate-300 bg-white text-xs font-bold"
            />
          </div>

          <div className="space-y-1">
            <label className="text-[11px] font-bold text-slate-700">Filtrar trabajador:</label>
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400" />
              <Input
                placeholder="Nombre o DNI..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className="h-9 rounded border-slate-300 pl-8 pr-7 text-xs bg-white"
              />
              {search && (
                <button
                  type="button"
                  onClick={() => setSearch("")}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Navegador y Previsualización */}
      <div className="grid gap-4 lg:grid-cols-[300px_1fr]">
        {/* Selector de Trabajador */}
        <div className="space-y-2">
          <div className="bg-slate-200 text-slate-800 px-3 py-2 rounded-t text-xs font-bold border border-slate-300 flex items-center justify-between">
            <span>Trabajadores ({filtered.length})</span>
            <div className="flex items-center gap-1">
              <Button
                variant="outline"
                size="sm"
                onClick={handlePrev}
                disabled={activeIdx === 0}
                className="h-6 w-6 p-0 rounded bg-white"
              >
                <ChevronLeft className="h-3 w-3" />
              </Button>
              <span className="text-[11px] font-mono font-bold text-slate-700 px-1">
                {activeIdx + 1}/{workers.length}
              </span>
              <Button
                variant="outline"
                size="sm"
                onClick={handleNext}
                disabled={activeIdx === workers.length - 1}
                className="h-6 w-6 p-0 rounded bg-white"
              >
                <ChevronRight className="h-3 w-3" />
              </Button>
            </div>
          </div>

          <div className="bg-white border border-slate-300 rounded-b divide-y divide-slate-200 max-h-[560px] overflow-y-auto">
            {filtered.map((w, idx) => {
              const realIdx = workers.indexOf(w);
              const isSelected = realIdx === activeIdx;
              return (
                <button
                  key={`${w.dni}-${idx}`}
                  type="button"
                  onClick={() => setActiveIdx(realIdx)}
                  className={`w-full text-left p-2.5 text-xs transition block ${
                    isSelected
                      ? "bg-[#e7f1ff] text-[#084298] font-bold border-l-4 border-l-[#0d6efd]"
                      : "hover:bg-slate-50 text-slate-800"
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="truncate max-w-[200px] font-semibold text-slate-900">
                      {w.apPaterno} {w.nombres}
                    </span>
                    <span className="font-mono text-[10px] text-slate-500">#{w.n}</span>
                  </div>
                  <div className="mt-0.5 flex items-center justify-between text-[11px] text-slate-500">
                    <span>DNI {w.dni}</span>
                    <span className="font-mono font-bold text-emerald-700">S/. {w.totalLiquido}</span>
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* Visor Oficial Courier con Controles de Zoom y Edición */}
        <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
          <div className="bg-slate-100 border-b border-slate-300 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-700 font-bold">
            <span>VISTA PREVIA · {active?.apPaterno} {active?.nombres}</span>

            {/* Controles de Zoom */}
            <div className="flex items-center gap-1 font-sans">
              <span className="text-[11px] text-slate-500 mr-1 font-normal">Zoom:</span>
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => setZoom((z) => Math.max(70, z - 10))}
                className="h-6 w-6 p-0 text-xs rounded border-slate-300 bg-white"
                title="Alejar"
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
                title="Acercar"
              >
                <ZoomIn className="h-3 w-3" />
              </Button>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                onClick={() => setZoom(100)}
                className="h-6 px-1.5 text-[11px] text-slate-600 hover:text-slate-900 rounded font-normal"
                title="Restablecer tamaño normal"
              >
                100%
              </Button>
            </div>
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
      </div>

      {/* Modal Confirmación de Sobreescritura */}
      <AlertDialog open={showOverwriteModal} onOpenChange={setShowOverwriteModal}>
        <AlertDialogContent className="bg-white rounded border border-slate-300">
          <AlertDialogHeader>
            <AlertDialogTitle className="text-sm font-bold text-slate-900">
              ¿Desea reemplazar la planilla existente?
            </AlertDialogTitle>
            <AlertDialogDescription className="text-xs text-slate-600">
              {pendingOverwriteMsg}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel className="rounded text-xs h-8">Cancelar</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                setShowOverwriteModal(false);
                handleSaveToDatabase(true);
              }}
              className="rounded bg-[#d63384] hover:bg-[#b02a6b] text-white text-xs h-8 font-bold"
            >
              Sí, reemplazar planilla
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Portal de Impresión Limpia (1 sola página A4) */}
      <PrintBoletaPortal text={editedBoletaText || boletaText} />
    </div>
  );
}
