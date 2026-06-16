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
  Download,
  FileUp,
  Info,
  Loader2,
  Printer,
  Search,
  Upload,
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

        if (!result.ok) {
          throw new Error(result.error || "Error desconocido");
        }

        const detectedCategoryId = result.categoryId ?? null;

        if (!detectedCategoryId) {
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
          const debugInfo = result.debug
            ? ` (Fila Cabecera: ${result.debug.headerRowIdx}, Col DNI: ${result.debug.col?.dni ?? "no encontrada"}, Cabeceras 1-10: ${result.debug.headers?.slice(1, 11).join(", ")})`
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
      <div className="min-h-screen bg-slate-50 flex flex-col">
        <header className="border-b border-border/80 bg-white/90 backdrop-blur">
          <div className="container py-5 flex items-center justify-between gap-4">
            <div className="flex items-center gap-3 min-w-0">
              <div className="h-11 w-11 rounded-2xl bg-white shadow-sm border flex items-center justify-center">
                <img src="/cas.png" alt="Logo CAS" className="h-6 w-6" />
              </div>
              <div className="min-w-0">
                <h1 className="text-lg font-semibold tracking-tight">
                  Generador de Boletas CAS
                </h1>
                <p className="text-sm text-muted-foreground">
                  UGEL 04 TSE · Carga tu archivo CAS y genera boletas de forma automática.
                </p>
              </div>
            </div>
            <span className="text-xs text-muted-foreground whitespace-nowrap">
              v1.1 (neotest-dev)
            </span>
          </div>
        </header>

        <main className="flex-1 flex items-center justify-center p-6 md:p-10">
          <div className="w-full max-w-4xl space-y-6">
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
              className={`rounded-[28px] border-2 border-dashed p-10 md:p-14 text-center transition-all bg-white cursor-pointer ${dragOver ? "border-blue-600 bg-blue-50/30" : "border-slate-200 hover:border-slate-400 hover:bg-slate-50/50"
                }`}
            >
              {loading ? (
                <div className="flex flex-col items-center gap-3">
                  <Loader2 className="h-10 w-10 animate-spin text-blue-600" />
                  <p className="text-sm text-muted-foreground">{loadingMsg}</p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-5">
                  <div className="h-16 w-16 rounded-full bg-blue-600 text-white flex items-center justify-center shadow-md animate-pulse">
                    <Upload className="h-6 w-6" />
                  </div>
                  <div className="space-y-2">
                    <p className="text-lg font-semibold text-slate-800">
                      Sube tu planilla Excel de CAS
                    </p>
                    <p className="text-sm text-muted-foreground">
                      Arrastra aquí o haz clic para seleccionar .xlsx, .xls, .xlsm
                    </p>
                    <p className="text-xs text-muted-foreground/80 mt-2 max-w-md mx-auto">
                      El sistema identificará de manera automática el tipo de CAS (Sede, JEC, Orquestando, SEHO, EBE, Winanq, Convivencia o Mantenimiento) a partir de su contenido.
                    </p>
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
              <div className="rounded-2xl border bg-white px-4 py-3 text-sm">
                <span className="text-muted-foreground">Categoría</span>
                <p className="font-semibold mt-1">Detección automática</p>
              </div>
              <div className="rounded-2xl border bg-white px-4 py-3 text-sm">
                <span className="text-muted-foreground">Validación</span>
                <p className="font-semibold mt-1">Por contenido del Excel</p>
              </div>
              <div className="rounded-2xl border bg-white px-4 py-3 text-sm">
                <span className="text-muted-foreground">
                  Formato recomendado
                </span>
                <p className="font-semibold mt-1">
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
    <div className="min-h-screen bg-slate-100/70 flex flex-col">
      <header className="no-print sticky top-0 z-10 border-b border-border/80 bg-white/95 backdrop-blur">
        <div className="container py-4 space-y-4">
          <div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between">
            <div className="flex items-start gap-3 min-w-0">
              <div className="h-11 w-11 rounded-2xl bg-white border shadow-sm flex items-center justify-center shrink-0">
                <img src="/cas.png" alt="Logo CAS" className="h-6 w-6" />
              </div>
              <div className="min-w-0">
                <h1 className="text-base font-semibold leading-none pt-1">
                  Boletas CAS · UGEL 04 TSE
                </h1>
                <div className="mt-2 flex flex-wrap items-center gap-2 text-xs">
                  <span className="rounded-full bg-slate-100 px-2.5 py-1 font-medium text-slate-700">
                    {selectedCategory?.label ?? "Categoria no definida"}
                  </span>
                  <span className="rounded-full bg-blue-50 px-2.5 py-1 font-medium text-blue-700">
                    {period.mes} {period.anio}
                  </span>
                  <span className="rounded-full bg-emerald-50 px-2.5 py-1 font-medium text-emerald-700">
                    {workers.length} trabajadores
                  </span>
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <Button
                size="sm"
                variant="default"
                className="h-10 bg-blue-600 text-white hover:bg-blue-700"
                onClick={handlePrint}
              >
                <Printer className="h-4 w-4 mr-1.5" /> Imprimir
              </Button>
              <Button
                size="sm"
                variant="destructive"
                className="h-10"
                onClick={handlePDF}
              >
                <Download className="h-4 w-4 mr-1.5" /> PDF
              </Button>
              <Button
                size="sm"
                variant="default"
                className="h-10 bg-emerald-600 text-white hover:bg-emerald-700"
                onClick={reset}
              >
                <FileUp className="h-4 w-4 mr-1.5" /> Nueva planilla
              </Button>
            </div>
          </div>

          <div className="grid gap-3 xl:grid-cols-[minmax(0,1fr)_auto_auto] xl:items-center">
            <div className="flex items-center gap-2 w-full min-w-0">
              <div className="relative flex-grow">
                <Search className="absolute left-2 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-muted-foreground" />
                <Input
                  placeholder="Buscar nombre o DNI..."
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  className="pl-7 h-10 w-full bg-white"
                />
              </div>
              <Select
                value={String(activeIdx)}
                onValueChange={(value) => setActiveIdx(Number(value))}
              >
                <SelectTrigger className="h-10 w-full sm:w-[330px] bg-white">
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
                className="h-10 w-10 bg-white"
                onClick={handlePrev}
                disabled={activeIdx === 0}
              >
                <ChevronLeft className="h-4 w-4" />
              </Button>
              <Button
                size="icon"
                variant="outline"
                className="h-10 w-10 bg-white"
                onClick={handleNext}
                disabled={activeIdx === workers.length - 1}
              >
                <ChevronRight className="h-4 w-4" />
              </Button>
            </div>

            <div className="justify-self-start xl:justify-self-end">
              <div className="h-10 min-w-[220px] rounded-xl border bg-slate-50 px-3 flex items-center gap-2 text-sm">
                <span className="text-muted-foreground">Categoria fija:</span>
                <span className="font-semibold text-foreground">
                  {selectedCategory?.label ?? "Sin categoria"}
                </span>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 py-8 px-4 flex justify-center">
        <div
          id="boleta-print"
          className="bg-white shadow-sm border border-border w-full max-w-[820px] mx-auto p-8 md:p-12"
        >
          <div
            className="no-print flex items-center gap-2 p-2 mb-4 text-sm text-blue-800 rounded-lg bg-blue-50 dark:bg-gray-800 dark:text-blue-400"
            role="alert"
          >
            <Info className="h-4 w-4 flex-shrink-0" />
            <span className="sr-only">Info</span>
            <div>
              <span className="font-medium">Nota:</span> El texto de la boleta
              es editable. Puedes corregir errores o anadir datos faltantes
              directamente aqui.
            </div>
          </div>

          <textarea
            className="boleta-mono no-print w-full h-full min-h-[calc(29.7cm-6rem)] resize-none border-none focus:outline-none"
            value={editedBoletaText}
            onChange={(event) => setEditedBoletaText(event.target.value)}
            rows={editedBoletaText.split("\n").length}
            readOnly={loading}
          />

          <pre className="boleta-mono print-only print-text">
            {editedBoletaText}
          </pre>
        </div>
      </main>
    </div>
  );
};

export default Index;
