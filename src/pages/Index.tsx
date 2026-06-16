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
  ChevronLeft,
  ChevronRight,
  CheckCircle2,
  Code2,
  Download,
  FileCheck2,
  FileSpreadsheet,
  FileUp,
  Info,
  Loader2,
  Printer,
  Search,
  Upload,
  X,
} from "lucide-react";
import { toast } from "@/hooks/use-toast";
import jsPDF from "jspdf";
import ExcelWorker from "../workers/excelWorker.ts?worker";

const HEAVY_FILE_BYTES = 3 * 1024 * 1024;

const Index = () => {
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
  const [editedBoletaText, setEditedBoletaText] = useState("");
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
  }, []);

  const handleFile = useCallback(
    async (file: File) => {
      setLoading(true);
      setLoadingMsg(
        file.size > HEAVY_FILE_BYTES
          ? "Archivo pesado detectado, optimizando carga..."
          : "Procesando planilla...",
      );

      try {
        const buffer = await file.arrayBuffer();
        const worker = getWorker();
        const result = await new Promise<{
          ok: boolean;
          workers?: Worker[];
          period?: { mes: string; anio: string };
          categoryId?: string | null;
          debug?: {
            headerRowIdx: number;
            col: Record<string, number>;
            headers: string[];
            sheetNames?: string[];
            titleText?: string;
            categoryId?: string | null;
            headerRowIndexes?: number[];
            candidateRows?: number;
            headerCandidates?: Array<{
              rowIndex: number;
              matches: number;
              dniRows: number;
              score: number;
            }>;
          };
          error?: string;
        }>((resolve) => {
          const onMessage = (event: MessageEvent) => {
            worker.removeEventListener("message", onMessage);
            resolve(event.data);
          };

          worker.addEventListener("message", onMessage);
          worker.postMessage({ buffer, filename: file.name }, [buffer]);
        });

        console.info("[CAS Upload] Resultado del worker", {
          fileName: file.name,
          fileSize: file.size,
          ok: result.ok,
          workersCount: result.workers?.length ?? 0,
          categoryId: result.categoryId ?? null,
          period: result.period,
          debug: result.debug,
          error: result.error,
        });

        if (!result.ok) {
          console.error("[CAS Upload] Error reportado por worker", {
            fileName: file.name,
            error: result.error,
          });
          throw new Error(result.error || "Error desconocido");
        }

        const detectedCategoryId = result.categoryId ?? null;

        if (!detectedCategoryId) {
          console.warn("[CAS Upload] Categoria no identificada", {
            fileName: file.name,
            titleText: result.debug?.titleText,
            sheetNames: result.debug?.sheetNames,
          });

          toast({
            title: "No se pudo identificar la categoría",
            description:
              "No pude reconocer la categoría desde el contenido del Excel. Verifica la planilla antes de continuar.",
            variant: "destructive",
          });
          return;
        }

        const detectedCategory = CATEGORIAS_PLANILLA.find(
          (category) => category.id === detectedCategoryId,
        );

        setSelectedCategoryId(detectedCategoryId);

        const nextWorkers = result.workers || [];

        if (!nextWorkers.length) {
          console.warn("[CAS Upload] Sin trabajadores detectados", {
            fileName: file.name,
            debug: result.debug,
          });

          const debugInfo = result.debug
            ? ` (Hojas: ${result.debug.sheetNames?.join(", ") || "no disponible"}, Fila Cabecera: ${result.debug.headerRowIdx}, Filas usadas: ${result.debug.headerRowIndexes?.join(", ") || "no disponible"}, Col DNI: ${result.debug.col?.dni ?? "no encontrada"}, Filas con DNI: ${result.debug.candidateRows ?? "no disponible"}, Cabeceras 1-10: ${result.debug.headers?.slice(1, 11).join(", ")})`
            : "";
          toast({
            title: "Sin datos",
            description: `No se encontraron trabajadores en la hoja CAS-SEDE.${debugInfo}`,
            variant: "destructive",
          });
          return;
        }

        setWorkers(nextWorkers);
        setPeriod(result.period || parsePeriodFromFilename(file.name));
        setActiveIdx(0);
        toast({
          title: "Archivo cargado",
          description: `${detectedCategory?.label ?? "Categoría Detectada"}: ${nextWorkers.length} trabajadores · ${result.period?.mes} ${result.period?.anio}`,
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

    if (loading) {
      return;
    }

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

  const handlePrev = () => setActiveIdx((index) => Math.max(0, index - 1));
  const handleNext = () =>
    setActiveIdx((index) => Math.min(workers.length - 1, index + 1));

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!workers.length) return;
      if (event.key === "ArrowLeft") handlePrev();
      if (event.key === "ArrowRight") handleNext();
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [workers.length]);

  useEffect(() => {
    setEditedBoletaText(boletaText);
  }, [boletaText]);

  const handlePrint = () => window.print();

  const handlePDF = () => {
    if (!active || !editedBoletaText) return;

    const pdf = new jsPDF({ unit: "mm", format: "a4" });
    pdf.setFont("courier", "normal");
    pdf.setFontSize(9);

    let y = 15;
    for (const line of editedBoletaText.split("\n")) {
      if (y > 285) {
        pdf.addPage();
        y = 15;
      }
      pdf.text(line, 15, y);
      y += 4.2;
    }

    const fileName =
      `Boleta_${active.n}_${active.apPaterno}_${active.nombres}.pdf`.replace(
        /\s+/g,
        "_",
      );
    pdf.save(fileName);
  };

  const reset = () => clearLoadedState();

  if (!workers.length) {
    return (
      <div className="min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top_left,#dbeafe_0,transparent_34%),linear-gradient(135deg,#f8fafc_0%,#eef4ff_48%,#f8fafc_100%)] flex flex-col">
        <header className="motion-rise border-b border-white/70 bg-white/80 shadow-sm shadow-blue-950/5 backdrop-blur-xl">
          <div className="container py-3.5 flex items-center justify-between gap-4">
            <div className="flex items-center gap-3 min-w-0">
              <div className="h-11 w-11 rounded-2xl bg-white shadow-sm border border-blue-100 flex items-center justify-center ring-4 ring-blue-50">
                <img src="/cas.png" alt="Logo CAS" className="h-5 w-5" />
              </div>
              <div className="min-w-0">
                <h1 className="text-base font-bold tracking-tight text-slate-950 sm:text-lg">
                  Generador de Boletas CAS
                </h1>
                <p className="text-xs font-medium text-slate-500 sm:text-sm">
                  UGEL 04 - TRUJILLO SUR ESTE
                </p>
              </div>
            </div>
            <div className="hidden items-center gap-2 sm:flex">
              <span className="inline-flex items-center gap-1.5 rounded-full border border-slate-200 bg-white/80 px-3 py-1.5 text-xs font-semibold text-slate-600 shadow-sm">
                <Code2 className="h-3.5 w-3.5 text-blue-600" />
                neotest-dev
              </span>
              <span className="rounded-full border border-blue-100 bg-blue-50/80 px-3 py-1.5 text-xs font-semibold text-blue-700 shadow-sm">
                v1.3
              </span>
            </div>
          </div>
        </header>

        <main className="flex-1 flex items-start justify-center p-4 pt-8 sm:p-6 sm:pt-10 md:p-8 md:pt-12">
          <div className="w-full max-w-3xl space-y-4 md:space-y-5">
            <div
              onDragOver={(event) => {
                event.preventDefault();
                if (!loading) setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => {
                if (loading) return;
                fileRef.current?.click();
              }}
              className={`motion-soft-pop group relative overflow-hidden rounded-[28px] border p-5 text-center shadow-xl shadow-blue-950/5 transition-all sm:p-6 md:p-8 ${loading ? "cursor-wait" : "cursor-pointer"} ${dragOver ? "border-blue-400 bg-blue-50/90 ring-4 ring-blue-100" : "border-white/80 bg-white/90 hover:-translate-y-0.5 hover:border-blue-200 hover:bg-white hover:shadow-2xl hover:shadow-blue-950/10"
                }`}
            >
              <div className="pointer-events-none absolute inset-x-8 top-0 h-px bg-gradient-to-r from-transparent via-blue-300/70 to-transparent" />
              <div className="pointer-events-none absolute -right-24 -top-24 h-56 w-56 rounded-full bg-blue-100/70 blur-3xl transition group-hover:bg-blue-200/70" />
              <div className="pointer-events-none absolute -bottom-28 -left-20 h-56 w-56 rounded-full bg-cyan-100/60 blur-3xl" />

              {loading ? (
                <div className="relative flex min-h-[210px] flex-col items-center justify-center gap-4 md:min-h-[240px]">
                  <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-600 text-white shadow-lg shadow-blue-600/25">
                    <Loader2 className="h-7 w-7 animate-spin" />
                  </div>
                  <div className="space-y-1">
                    <p className="text-lg font-semibold text-slate-900">
                      Analizando tu planilla
                    </p>
                    <p className="text-sm text-slate-500">{loadingMsg}</p>
                  </div>
                </div>
              ) : (
                <div className="relative mx-auto flex max-w-xl flex-col items-center gap-4 py-4 md:py-6">
                  <div className="inline-flex items-center gap-2 rounded-full border border-blue-100 bg-blue-50/80 px-3 py-1.5 text-xs font-semibold uppercase tracking-[0.18em] text-blue-700">
                    <FileSpreadsheet className="h-3.5 w-3.5" />
                    Carga inteligente CAS
                  </div>

                  <div className="relative">
                    <div className="absolute inset-0 rounded-[30px] bg-blue-500 blur-xl opacity-20 transition group-hover:opacity-30" />
                    <div className="motion-float relative flex h-16 w-16 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-600 to-blue-500 text-white shadow-xl shadow-blue-600/25">
                      <Upload className="h-7 w-7" />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <h2 className="text-xl font-bold tracking-tight text-slate-950 md:text-2xl">
                      Sube tu planilla Excel de CAS
                    </h2>
                    <p className="mx-auto max-w-lg text-sm leading-6 text-slate-600">
                      Arrastra el archivo aquí o selecciona tu planilla para generar boletas con detección automática de categoría.
                    </p>
                  </div>

                  <div className="flex flex-col items-center gap-2 sm:flex-row">
                    <span className="inline-flex h-10 items-center justify-center rounded-xl bg-blue-600 px-4 text-sm font-semibold text-white shadow-lg shadow-blue-600/20 transition group-hover:bg-blue-700">
                      Seleccionar archivo
                    </span>
                    <span className="text-sm font-medium text-slate-500">
                      o suéltalo en esta zona
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center justify-center gap-2 text-xs text-slate-500">
                    <span className="rounded-full bg-slate-100 px-3 py-1 font-medium text-slate-600">
                      .xlsx
                    </span>
                    <span className="rounded-full bg-slate-100 px-3 py-1 font-medium text-slate-600">
                      .xls
                    </span>
                    <span className="rounded-full bg-slate-100 px-3 py-1 font-medium text-slate-600">
                      .xlsm
                    </span>
                  </div>
                </div>
              )}

              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.xlsm"
                className="hidden"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) handleFile(file);
                  event.target.value = "";
                }}
              />
            </div>

            <div className="grid gap-3 md:grid-cols-3">
              <div className="motion-rise motion-delay-1 rounded-2xl border border-white/80 bg-white/85 p-3.5 text-sm shadow-sm shadow-blue-950/5 backdrop-blur">
                <div className="mb-2 flex h-8 w-8 items-center justify-center rounded-xl bg-blue-50 text-blue-700">
                  <CheckCircle2 className="h-4 w-4" />
                </div>
                <span className="text-slate-500">Categoría</span>
                <p className="mt-1 font-semibold text-slate-900">
                  Detección automática
                </p>
              </div>
              <div className="motion-rise motion-delay-2 rounded-2xl border border-white/80 bg-white/85 p-3.5 text-sm shadow-sm shadow-blue-950/5 backdrop-blur">
                <div className="mb-2 flex h-8 w-8 items-center justify-center rounded-xl bg-emerald-50 text-emerald-700">
                  <FileCheck2 className="h-4 w-4" />
                </div>
                <span className="text-slate-500">Validación</span>
                <p className="mt-1 font-semibold text-slate-900">
                  Por contenido del Excel
                </p>
              </div>
              <div className="motion-rise motion-delay-3 rounded-2xl border border-white/80 bg-white/85 p-3.5 text-sm shadow-sm shadow-blue-950/5 backdrop-blur">
                <div className="mb-2 flex h-8 w-8 items-center justify-center rounded-xl bg-slate-100 text-slate-700">
                  <FileSpreadsheet className="h-4 w-4" />
                </div>
                <span className="text-slate-500">
                  Formato recomendado
                </span>
                <p className="mt-1 font-semibold text-slate-900">
                  <code className="font-mono">.xlsx</code>
                </p>
              </div>
            </div>
          </div>
        </main>
      </div>
    );
  }

  return (
    <div className="min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top_left,#dbeafe_0,transparent_32%),linear-gradient(135deg,#f8fafc_0%,#eef4ff_48%,#f8fafc_100%)] flex flex-col">
      <header className="motion-rise no-print sticky top-0 z-20 border-b border-white/70 bg-white/85 shadow-sm shadow-blue-950/5 backdrop-blur-xl">
        <div className="container py-4 space-y-4">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
            <div className="flex items-start gap-3 min-w-0">
              <div className="motion-soft-pop h-12 w-12 rounded-2xl bg-white border border-blue-100 shadow-sm ring-4 ring-blue-50 flex items-center justify-center shrink-0">
                <img src="/cas.png" alt="Logo CAS" className="h-6 w-6" />
              </div>
              <div className="min-w-0">
                <p className="text-[11px] font-bold uppercase tracking-[0.22em] text-blue-700">
                  Gestión de boletas
                </p>
                <h1 className="mt-1 text-lg font-bold tracking-tight text-slate-950">
                  Boletas CAS · UGEL 04 TSE
                </h1>
                <div className="mt-2 flex flex-wrap items-center gap-2 text-xs font-semibold">
                  <span className="rounded-full border border-slate-200 bg-white px-3 py-1 text-slate-700 shadow-sm">
                    {selectedCategory?.label ?? "Categoria no definida"}
                  </span>
                  <span className="rounded-full border border-blue-100 bg-blue-50 px-3 py-1 text-blue-700">
                    {period.mes} {period.anio}
                  </span>
                  <span className="rounded-full border border-emerald-100 bg-emerald-50 px-3 py-1 text-emerald-700">
                    {workers.length} trabajadores
                  </span>
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <span className="hidden items-center gap-1.5 rounded-xl border border-slate-200 bg-white/80 px-3 py-2 text-xs font-semibold text-slate-600 shadow-sm lg:inline-flex">
                <Code2 className="h-3.5 w-3.5 text-blue-600" />
                neotest-dev
              </span>
              <Button
                size="sm"
                variant="default"
                className="h-11 rounded-xl bg-blue-600 px-4 text-white shadow-lg shadow-blue-600/20 transition hover:-translate-y-0.5 hover:bg-blue-700"
                onClick={handlePrint}
              >
                <Printer className="h-4 w-4 mr-1.5" /> Imprimir
              </Button>
              <Button
                size="sm"
                variant="destructive"
                className="h-11 rounded-xl px-4 shadow-lg shadow-red-500/15 transition hover:-translate-y-0.5"
                onClick={handlePDF}
              >
                <Download className="h-4 w-4 mr-1.5" /> PDF
              </Button>
              <Button
                size="sm"
                variant="default"
                className="h-11 rounded-xl bg-emerald-600 px-4 text-white shadow-lg shadow-emerald-600/20 transition hover:-translate-y-0.5 hover:bg-emerald-700"
                onClick={reset}
              >
                <FileUp className="h-4 w-4 mr-1.5" /> Nueva planilla
              </Button>
            </div>
          </div>

          <div className="rounded-2xl border border-blue-100/80 bg-blue-50/60 px-4 py-3 text-sm text-slate-600 shadow-sm">
            <span className="font-semibold text-blue-800">Busca y navega:</span>{" "}
            escribe un nombre o DNI para abrir la boleta al instante, o elige un trabajador desde el selector.
          </div>

          <div className="grid gap-3 xl:grid-cols-[minmax(0,1fr)_auto_auto] xl:items-center">
            <div className="flex flex-col gap-2 w-full min-w-0 sm:flex-row">
              <div className="relative flex-grow">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-blue-600" />
                <Input
                  placeholder="Buscar nombre o DNI..."
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  className="h-11 w-full rounded-xl border-blue-100 bg-white/90 pl-10 pr-10 shadow-sm focus-visible:ring-blue-200"
                />
                {search ? (
                  <button
                    type="button"
                    aria-label="Limpiar búsqueda"
                    className="absolute right-2 top-1/2 inline-flex h-7 w-7 -translate-y-1/2 items-center justify-center rounded-lg text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
                    onClick={() => setSearch("")}
                  >
                    <X className="h-4 w-4" />
                  </button>
                ) : null}
              </div>
              <Select
                value={String(activeIdx)}
                onValueChange={(value) => setActiveIdx(Number(value))}
              >
                <SelectTrigger className="h-11 w-full rounded-xl border-blue-100 bg-white/90 shadow-sm sm:w-[360px]">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent className="max-h-80">
                  {filtered.map((worker) => {
                    const realIdx = workers.indexOf(worker);
                    return (
                      <SelectItem key={realIdx} value={String(realIdx)}>
                        {worker.n}. {worker.apPaterno} {worker.apMaterno},{" "}
                        {worker.nombres}
                      </SelectItem>
                    );
                  })}
                </SelectContent>
              </Select>
            </div>

            <div className="flex items-center gap-2 justify-self-start xl:justify-self-center">
              <Button
                size="icon"
                variant="outline"
                className="h-11 w-11 rounded-xl border-blue-100 bg-white/90 shadow-sm disabled:bg-slate-100"
                onClick={handlePrev}
                disabled={activeIdx === 0}
              >
                <ChevronLeft className="h-4 w-4" />
              </Button>
              <Button
                size="icon"
                variant="outline"
                className="h-11 w-11 rounded-xl border-blue-100 bg-white/90 shadow-sm disabled:bg-slate-100"
                onClick={handleNext}
                disabled={activeIdx === workers.length - 1}
              >
                <ChevronRight className="h-4 w-4" />
              </Button>
            </div>

            <div className="justify-self-start xl:justify-self-end">
              <div className="h-11 min-w-[230px] rounded-xl border border-blue-100 bg-blue-50/70 px-3 flex items-center gap-2 text-sm shadow-sm">
                <span className="text-slate-500">Categoría fija:</span>
                <span className="font-bold text-slate-950">
                  {selectedCategory?.label ?? "Sin categoria"}
                </span>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="relative flex-1 px-4 py-6 md:px-6 md:py-8">
        <div className="pointer-events-none absolute inset-x-0 top-0 h-40 bg-gradient-to-b from-white/50 to-transparent" />
        <div className="relative mx-auto grid w-full max-w-7xl gap-5 xl:grid-cols-[320px_minmax(0,1fr)]">
          <aside className="no-print space-y-4 xl:sticky xl:top-[170px] xl:self-start">
            <div className="motion-rise motion-delay-1 overflow-hidden rounded-3xl border border-white/80 bg-white/90 shadow-xl shadow-blue-950/5 backdrop-blur">
              <div className="border-b border-slate-100 bg-gradient-to-br from-blue-600 to-blue-500 p-5 text-white">
                <p className="text-xs font-semibold uppercase tracking-[0.18em] text-blue-100">
                  Trabajador activo
                </p>
                <h2 className="mt-3 text-xl font-bold leading-tight">
                  {active?.apPaterno} {active?.apMaterno}
                </h2>
                <p className="mt-1 text-sm text-blue-50">{active?.nombres}</p>
              </div>
              <div className="space-y-3 p-5 text-sm">
                <div className="flex items-center justify-between gap-3 rounded-2xl bg-slate-50 px-4 py-3">
                  <span className="text-slate-500">DNI</span>
                  <span className="font-bold text-slate-950">{active?.dni}</span>
                </div>
                <div className="rounded-2xl bg-slate-50 px-4 py-3">
                  <span className="text-slate-500">Cargo</span>
                  <p className="mt-1 font-semibold leading-snug text-slate-950">
                    {active?.cargo || "Sin cargo registrado"}
                  </p>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div className="rounded-2xl bg-emerald-50 px-4 py-3">
                    <span className="text-xs font-medium text-emerald-700">
                      Líquido
                    </span>
                    <p className="mt-1 font-bold text-emerald-900">
                      S/ {active?.totalLiquido || "0.00"}
                    </p>
                  </div>
                  <div className="rounded-2xl bg-blue-50 px-4 py-3">
                    <span className="text-xs font-medium text-blue-700">
                      Pensión
                    </span>
                    <p className="mt-1 truncate font-bold text-blue-900">
                      {active?.sistemaPensionario || "-"}
                    </p>
                  </div>
                </div>
              </div>
            </div>

            <div className="motion-rise motion-delay-2 rounded-3xl border border-white/80 bg-white/80 p-5 shadow-sm shadow-blue-950/5 backdrop-blur">
              <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-500">
                Progreso
              </p>
              <div className="mt-4 flex items-end justify-between">
                <span className="text-3xl font-black tracking-tight text-slate-950">
                  {activeIdx + 1}
                </span>
                <span className="pb-1 text-sm font-semibold text-slate-500">
                  de {workers.length} boletas
                </span>
              </div>
              <div className="mt-4 h-2 overflow-hidden rounded-full bg-slate-100">
                <div
                  className="motion-progress h-full rounded-full bg-gradient-to-r from-blue-600 to-emerald-500"
                  style={{ width: `${((activeIdx + 1) / workers.length) * 100}%` }}
                />
              </div>
            </div>
          </aside>

          <section className="motion-soft-pop motion-delay-2 min-w-0">
            <div
              id="boleta-print"
              className="mx-auto w-full max-w-[860px] border border-slate-200 bg-white p-5 shadow-2xl shadow-slate-950/10 md:p-8 lg:p-10"
            >
              <div
                className="no-print mb-5 flex items-start gap-3 rounded-2xl border border-blue-100 bg-blue-50/80 px-4 py-3 text-sm text-blue-900"
                role="alert"
              >
                <Info className="mt-0.5 h-4 w-4 flex-shrink-0" />
                <span className="sr-only">Info</span>
                <div>
                  <span className="font-bold">Texto editable:</span> puedes corregir errores o añadir datos faltantes antes de imprimir o exportar.
                </div>
              </div>

              <textarea
                className="boleta-mono no-print min-h-[calc(29.7cm-7rem)] w-full resize-none rounded-2xl border border-slate-100 bg-white p-4 shadow-inner focus:outline-none focus:ring-2 focus:ring-blue-100 md:p-6"
                value={editedBoletaText}
                onChange={(event) => setEditedBoletaText(event.target.value)}
                rows={editedBoletaText.split("\n").length}
                readOnly={loading}
              />

              <pre className="boleta-mono print-only print-text">
                {editedBoletaText}
              </pre>
            </div>
          </section>
        </div>
      </main>
    </div>
  );
};

export default Index;
