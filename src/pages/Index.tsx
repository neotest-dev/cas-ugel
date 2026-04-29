import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Worker, buildBoletaText, parsePeriodFromFilename } from "@/lib/boleta";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { Upload, ChevronLeft, ChevronRight, Printer, Download, FileUp, Loader2, Search } from "lucide-react";
import { toast } from "@/hooks/use-toast";
import jsPDF from "jspdf";

const HEAVY_FILE_BYTES = 3 * 1024 * 1024; // 3 MB

const Index = () => {
  const [workers, setWorkers] = useState<Worker[]>([]);
  const [period, setPeriod] = useState<{ mes: string; anio: string }>({ mes: "", anio: "" });
  const [activeIdx, setActiveIdx] = useState(0);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("Procesando planilla…");
  const [search, setSearch] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const workerRef = useRef<globalThis.Worker | null>(null);

  // Mantener instancia única del Web Worker
  const getWorker = useCallback(() => {
    if (!workerRef.current) {
      workerRef.current = new window.Worker(
        new URL("../workers/excelWorker.ts", import.meta.url),
        { type: "module" }
      );
    }
    return workerRef.current;
  }, []);

  useEffect(() => {
    return () => {
      workerRef.current?.terminate();
      workerRef.current = null;
    };
  }, []);

  const handleFile = useCallback(async (file: File) => {
    setLoading(true);
    const heavy = file.size > HEAVY_FILE_BYTES;
    setLoadingMsg(heavy ? "Archivo pesado detectado, optimizando carga…" : "Procesando planilla…");
    try {
      const buffer = await file.arrayBuffer();
      const w = getWorker();
      const result = await new Promise<{ ok: boolean; workers?: Worker[]; period?: { mes: string; anio: string }; error?: string }>((resolve) => {
        const onMsg = (e: MessageEvent) => {
          w.removeEventListener("message", onMsg);
          resolve(e.data);
        };
        w.addEventListener("message", onMsg);
        w.postMessage({ buffer, filename: file.name }, [buffer]);
      });
      if (!result.ok) throw new Error(result.error || "Error desconocido");
      const ws = result.workers || [];
      if (!ws.length) {
        toast({ title: "Sin datos", description: "No se encontraron trabajadores en la hoja CAS-SEDE.", variant: "destructive" });
        return;
      }
      setWorkers(ws);
      setPeriod(result.period || parsePeriodFromFilename(file.name));
      setActiveIdx(0);
      toast({ title: "Archivo cargado", description: `${ws.length} trabajadores · ${result.period?.mes} ${result.period?.anio}` });
    } catch (e: any) {
      toast({ title: "Error al procesar archivo", description: e?.message || "Error desconocido", variant: "destructive" });
    } finally {
      setLoading(false);
    }
  }, [getWorker]);

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  };

  const filtered = useMemo(() => {
    if (!search.trim()) return workers;
    const q = search.toLowerCase();
    return workers.filter(w =>
      `${w.apPaterno} ${w.apMaterno} ${w.nombres} ${w.dni}`.toLowerCase().includes(q)
    );
  }, [workers, search]);

  const active = workers[activeIdx];
  const boletaText = useMemo(
    () => active ? buildBoletaText(active, period.mes, period.anio) : "",
    [active, period]
  );

  const handlePrev = () => setActiveIdx(i => Math.max(0, i - 1));
  const handleNext = () => setActiveIdx(i => Math.min(workers.length - 1, i + 1));

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!workers.length) return;
      if (e.key === "ArrowLeft") handlePrev();
      if (e.key === "ArrowRight") handleNext();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [workers.length]);

  const handlePrint = () => window.print();

  const handlePDF = () => {
    if (!active) return;
    const pdf = new jsPDF({ unit: "mm", format: "a4" });
    pdf.setFont("courier", "normal");
    pdf.setFontSize(9);
    const lines = boletaText.split("\n");
    let y = 15;
    lines.forEach(line => {
      if (y > 285) { pdf.addPage(); y = 15; }
      pdf.text(line, 15, y);
      y += 4.2;
    });
    const fname = `Boleta_${active.n}_${active.apPaterno}_${active.nombres}.pdf`.replace(/\s+/g, "_");
    pdf.save(fname);
  };

  const reset = () => {
    setWorkers([]);
    setActiveIdx(0);
    setSearch("");
    setPeriod({ mes: "", anio: "" });
  };

  if (!workers.length) {
    return (
      <div className="min-h-screen bg-background flex flex-col">
        <header className="border-b border-border">
          <div className="container py-4 flex items-center justify-between">
            <h1 className="text-lg font-semibold tracking-tight">Generador de Boletas CAS · UGEL 04</h1>
            <span className="text-xs text-muted-foreground">v1.0</span>
          </div>
        </header>
        <main className="flex-1 flex items-center justify-center p-6">
          <div className="w-full max-w-2xl">
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current?.click()}
              className={`cursor-pointer rounded-2xl border-2 border-dashed transition-all p-12 text-center
                ${dragOver ? "border-foreground bg-muted" : "border-border hover:border-foreground/50 hover:bg-muted/40"}`}
            >
              {loading ? (
                <div className="flex flex-col items-center gap-3">
                  <Loader2 className="h-10 w-10 animate-spin" />
                  <p className="text-sm text-muted-foreground">{loadingMsg}</p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-4">
                  <div className="h-14 w-14 rounded-full bg-foreground text-background flex items-center justify-center">
                    <Upload className="h-6 w-6" />
                  </div>
                  <div>
                    <p className="text-base font-medium">Sube archivo de planilla CAS-SEDE</p>
                    <p className="text-sm text-muted-foreground mt-1">
                      Arrastra aquí o haz clic para seleccionar · .xlsx, .xls, .xlsm
                    </p>
                  </div>
                </div>
              )}
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.xlsm"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleFile(f);
                  e.target.value = "";
                }}
              />
            </div>
            <div className="mt-6 text-center text-xs text-muted-foreground space-y-1">
              <p>El sistema leerá la hoja <code className="font-mono">CAS-SEDE</code> desde la fila 7.</p>
              <p>Para mayor velocidad use archivo <code className="font-mono">.xlsx</code></p>
            </div>
          </div>
        </main>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-muted/30 flex flex-col">
      {/* Top bar */}
      <header className="no-print border-b border-border bg-background sticky top-0 z-10">
        <div className="container py-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div className="flex items-center gap-3 min-w-0">
            <h1 className="text-sm font-semibold whitespace-nowrap">Boletas CAS · UGEL 04</h1>
            <span className="text-xs text-muted-foreground hidden md:inline">{period.mes} {period.anio}</span>
            <span className="text-xs text-muted-foreground">{workers.length} trab.</span>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <div className="relative">
              <Search className="absolute left-2 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-muted-foreground" />
              <Input
                placeholder="Buscar nombre o DNI…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className="pl-7 h-9 w-48"
              />
            </div>
            <Select
              value={String(activeIdx)}
              onValueChange={(v) => setActiveIdx(Number(v))}
            >
              <SelectTrigger className="h-9 w-64">
                <SelectValue />
              </SelectTrigger>
              <SelectContent className="max-h-80">
                {filtered.map((w) => {
                  const realIdx = workers.indexOf(w);
                  return (
                    <SelectItem key={realIdx} value={String(realIdx)}>
                      {w.n}. {w.apPaterno} {w.apMaterno}, {w.nombres}
                    </SelectItem>
                  );
                })}
              </SelectContent>
            </Select>
            <Button size="icon" variant="outline" className="h-9 w-9" onClick={handlePrev} disabled={activeIdx === 0}>
              <ChevronLeft className="h-4 w-4" />
            </Button>
            <Button size="icon" variant="outline" className="h-9 w-9" onClick={handleNext} disabled={activeIdx === workers.length - 1}>
              <ChevronRight className="h-4 w-4" />
            </Button>
            <Button size="sm" variant="outline" className="h-9" onClick={handlePrint}>
              <Printer className="h-4 w-4 mr-1.5" /> Imprimir
            </Button>
            <Button size="sm" className="h-9" onClick={handlePDF}>
              <Download className="h-4 w-4 mr-1.5" /> PDF
            </Button>
            <Button size="sm" variant="ghost" className="h-9" onClick={reset}>
              <FileUp className="h-4 w-4 mr-1.5" /> Nuevo
            </Button>
          </div>
        </div>
      </header>

      {/* Boleta */}
      <main className="flex-1 py-8 px-4 flex justify-center">
        <div
          id="boleta-print"
          className="bg-white shadow-sm border border-border w-full max-w-[820px] mx-auto p-8 md:p-12"
          style={{ minHeight: "29.7cm" }}
        >
          <pre className="boleta-mono">{boletaText}</pre>
        </div>
      </main>
    </div>
  );
};

export default Index;
