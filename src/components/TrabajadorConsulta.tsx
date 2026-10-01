import { useState, useMemo } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { toast } from "@/hooks/use-toast";
import {
  consultarBoletasTrabajador,
  BoletaHistoricaItem,
  boletaItemToWorker,
} from "@/lib/payrollService";
import { buildBoletaText } from "@/lib/boleta";
import { exportBoletaToPDF } from "@/lib/pdfExport";
import { CambiarClaveDialog } from "./CambiarClaveDialog";
import { PrintBoletaPortal } from "./PrintBoletaPortal";
import {
  Search,
  Eye,
  EyeOff,
  Printer,
  Download,
  KeyRound,
  FileText,
  Calendar,
  Building,
  UserCheck,
  ArrowLeft,
  Loader2,
  Info,
  ShieldCheck,
  ZoomIn,
  ZoomOut,
} from "lucide-react";

export function TrabajadorConsulta() {
  const [dni, setDni] = useState("");
  const [claveOCuenta, setClaveOCuenta] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [boletas, setBoletas] = useState<BoletaHistoricaItem[]>([]);
  const [selectedBoletaId, setSelectedBoletaId] = useState<number | null>(null);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [zoom, setZoom] = useState(100);

  const handleConsultar = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!dni || dni.length !== 8) {
      toast({
        title: "DNI inválido",
        description: "El DNI debe contener exactamente 8 dígitos numéricos.",
        variant: "destructive",
      });
      return;
    }

    if (!claveOCuenta.trim()) {
      toast({
        title: "Dato de seguridad requerido",
        description:
          "Ingresa tu clave personal o los últimos 4 dígitos de tu cuenta de ahorros del Banco de la Nación.",
        variant: "destructive",
      });
      return;
    }

    setLoading(true);
    try {
      const res = await consultarBoletasTrabajador(dni, claveOCuenta);
      if (!res.ok || !res.boletas || res.boletas.length === 0) {
        toast({
          title: "Consulta no encontrada",
          description:
            res.error ||
            "Verifica tu DNI y que la clave o los últimos 4 dígitos de tu cuenta bancaria sean correctos.",
          variant: "destructive",
        });
        return;
      }

      setBoletas(res.boletas);
      setSelectedBoletaId(res.boletas[0].boleta_id);
      toast({
        title: "Boletas localizadas",
        description: `Se cargaron ${res.boletas.length} boleta(s) para ${res.boletas[0].nombres} ${res.boletas[0].ap_paterno}.`,
      });
    } finally {
      setLoading(false);
    }
  };

  const selectedBoleta = useMemo(
    () => boletas.find((b) => b.boleta_id === selectedBoletaId) || boletas[0] || null,
    [boletas, selectedBoletaId]
  );

  const activeWorker = useMemo(
    () => (selectedBoleta ? boletaItemToWorker(selectedBoleta) : null),
    [selectedBoleta]
  );

  const boletaText = useMemo(() => {
    if (!activeWorker || !selectedBoleta) return "";
    if (selectedBoleta.boleta_texto_personalizado) {
      return selectedBoleta.boleta_texto_personalizado;
    }
    return buildBoletaText(activeWorker, selectedBoleta.mes, selectedBoleta.anio);
  }, [activeWorker, selectedBoleta]);

  const handlePrint = () => {
    window.print();
  };

  const handlePDF = () => {
    if (!activeWorker || !boletaText) return;
    exportBoletaToPDF(activeWorker, boletaText, `Boleta_${selectedBoleta?.mes}_${selectedBoleta?.anio}`);
  };

  const handleResetConsulta = () => {
    setBoletas([]);
    setSelectedBoletaId(null);
  };

  return (
    <div className="space-y-4">
      {/* 1. Formulario de Búsqueda (Estilo Bootstrap Card) */}
      {boletas.length === 0 ? (
        <div className="mx-auto max-w-xl py-3">
          <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
            {/* Encabezado del Panel */}
            <div className="bg-[#0b223d] text-white px-5 py-3.5 border-b border-slate-300 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <FileText className="h-5 w-5 text-amber-400" />
                <h2 className="text-sm font-bold uppercase tracking-wide">
                  Consulta de Boleta de Pago CAS
                </h2>
              </div>
              <span className="text-[11px] bg-[#1a3d66] text-slate-200 px-2 py-0.5 rounded border border-[#234e80]">
                UGEL 04 TSE
              </span>
            </div>

            <div className="p-5 sm:p-6 space-y-4">
              {/* Alerta Informativa (Bootstrap alert-info) */}
              <div className="bg-[#e7f1ff] border border-[#b6d4fe] text-[#084298] p-3 rounded text-xs flex items-start gap-2.5">
                <Info className="h-4 w-4 shrink-0 mt-0.5 text-[#0d6efd]" />
                <div className="space-y-1">
                  <p className="font-semibold leading-snug">
                    Instrucciones para consultar su boleta:
                  </p>
                  <p className="text-[11px] text-[#084298] leading-relaxed">
                    Ingrese su <strong>número de DNI</strong> y su <strong>contraseña personal</strong>. Si es su primera vez en el sistema o no recuerda su clave, ingrese los <strong>últimos 4 dígitos de su cuenta de ahorros del Banco de la Nación</strong> donde percibe sus haberes.
                  </p>
                </div>
              </div>

              <form onSubmit={handleConsultar} className="space-y-4 pt-1">
                {/* Campo DNI */}
                <div className="space-y-1.5">
                  <Label htmlFor="cons-dni" className="text-xs font-bold text-slate-700">
                    Número de Documento de Identidad (DNI):
                  </Label>
                  <div className="relative">
                    <Input
                      id="cons-dni"
                      type="text"
                      inputMode="numeric"
                      maxLength={8}
                      placeholder="Ejemplo: 41234567"
                      value={dni}
                      onChange={(e) => setDni(e.target.value.replace(/\D/g, ""))}
                      className="h-10 rounded border-slate-300 bg-white font-mono text-sm focus:border-blue-600 focus:ring-1 focus:ring-blue-600"
                      required
                    />
                  </div>
                </div>

                {/* Campo Contraseña o Cuenta Bancaria */}
                <div className="space-y-1.5">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="cons-clave" className="text-xs font-bold text-slate-700">
                      Contraseña Personal o Últimos 4 Dígitos de Cuenta:
                    </Label>
                  </div>

                  <div className="relative">
                    <Input
                      id="cons-clave"
                      type={showPassword ? "text" : "password"}
                      placeholder="••••"
                      value={claveOCuenta}
                      onChange={(e) => setClaveOCuenta(e.target.value)}
                      className="h-10 rounded border-slate-300 bg-white pr-10 text-sm tracking-wider focus:border-blue-600 focus:ring-1 focus:ring-blue-600"
                      required
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                      title={showPassword ? "Ocultar" : "Mostrar"}
                    >
                      {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                    </button>
                  </div>
                </div>

                {/* Botón Consultar (Estilo Bootstrap btn-primary) */}
                <Button
                  type="submit"
                  disabled={loading}
                  className="w-full h-10 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white font-bold text-sm shadow-sm transition"
                >
                  {loading ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" /> Verificando en sistema...
                    </>
                  ) : (
                    <>
                      <Search className="mr-2 h-4 w-4" /> Consultar Boleta
                    </>
                  )}
                </Button>

                {/* Enlace para cambiar/crear contraseña */}
                <div className="pt-2 text-center border-t border-slate-200">
                  <button
                    type="button"
                    onClick={() => setDialogOpen(true)}
                    className="text-xs text-[#0d6efd] hover:text-[#0a58ca] font-semibold hover:underline inline-flex items-center gap-1.5"
                  >
                    <KeyRound className="h-3.5 w-3.5" />
                    ¿Desea crear o cambiar su contraseña personal?
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      ) : (
        /* 2. Vista de Resultados (Panel Clásico de Consulta) */
        <div className="space-y-4">
          {/* Ficha Resumen del Trabajador */}
          <div className="bg-white border border-slate-300 rounded p-4 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <div className="flex items-start gap-3">
              <div className="h-9 w-9 rounded bg-[#e7f1ff] text-[#0d6efd] border border-[#b6d4fe] flex items-center justify-center shrink-0">
                <UserCheck className="h-5 w-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-slate-900 text-sm sm:text-base">
                    {selectedBoleta?.ap_paterno} {selectedBoleta?.ap_materno}, {selectedBoleta?.nombres}
                  </span>
                  <span className="bg-slate-100 border border-slate-300 px-2 py-0.5 text-xs font-mono font-bold text-slate-700 rounded">
                    DNI {selectedBoleta?.dni}
                  </span>
                </div>
                <p className="text-xs text-slate-600 mt-0.5">
                  Cargo: <strong>{selectedBoleta?.cargo}</strong> · Régimen: <strong>D.Leg. 1057 (CAS)</strong>
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <Button
                variant="outline"
                size="sm"
                onClick={handleResetConsulta}
                className="h-9 rounded border-slate-300 text-slate-700 hover:bg-slate-100 text-xs font-semibold"
              >
                <ArrowLeft className="h-3.5 w-3.5 mr-1" /> Nueva Consulta
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

          <div className="grid gap-4 lg:grid-cols-[280px_1fr]">
            {/* Lista de Boletas (Estilo Bootstrap List-Group) */}
            <div className="space-y-2">
              <div className="bg-[#0b223d] text-white px-3 py-2 rounded-t text-xs font-bold uppercase tracking-wider flex items-center justify-between">
                <span>Historial de Meses</span>
                <span className="bg-[#153457] px-1.5 py-0.5 rounded text-[10px] font-mono">
                  {boletas.length} reg.
                </span>
              </div>

              <div className="bg-white border border-slate-300 rounded-b divide-y divide-slate-200 overflow-hidden">
                {boletas.map((b) => {
                  const isSelected = b.boleta_id === selectedBoletaId;
                  return (
                    <button
                      key={b.boleta_id}
                      type="button"
                      onClick={() => setSelectedBoletaId(b.boleta_id)}
                      className={`w-full text-left p-3 text-xs transition block ${
                        isSelected
                          ? "bg-[#e7f1ff] text-[#084298] font-bold border-l-4 border-l-[#0d6efd]"
                          : "hover:bg-slate-50 text-slate-800"
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-semibold text-slate-900">
                          {b.mes} {b.anio}
                        </span>
                        <span className="bg-slate-100 border border-slate-200 text-slate-600 px-1.5 py-0.5 rounded text-[10px] uppercase font-semibold">
                          {b.categoria_label}
                        </span>
                      </div>
                      <div className="mt-1 flex items-center justify-between text-[11px]">
                        <span className="text-slate-500">Líquido a percibir:</span>
                        <span className="font-mono font-bold text-emerald-700">
                          S/. {b.total_liquido}
                        </span>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Visor Oficial de Boleta en Papel A4 con Controles de Zoom */}
            <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
              <div className="bg-slate-100 border-b border-slate-300 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-700">
                <span className="font-bold flex items-center gap-1.5">
                  <Building className="h-4 w-4 text-slate-500" />
                  Boleta de Pago Oficial · Periodo: {selectedBoleta?.mes} {selectedBoleta?.anio}
                </span>

                {/* Controles de Zoom */}
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

              <div className="p-4 sm:p-6 overflow-x-auto bg-[#fafafa]">
                <pre
                  style={{
                    fontSize: `${(11 * zoom) / 100}px`,
                    lineHeight: 1.35,
                    width: `${Math.round(80 * (zoom / 100))}ch`,
                    minWidth: "68ch",
                  }}
                  className="font-mono text-black whitespace-pre select-all bg-white p-6 rounded border border-slate-300 shadow-sm mx-auto block"
                >
                  {boletaText}
                </pre>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal para crear o cambiar contraseña */}
      <CambiarClaveDialog
        open={dialogOpen}
        onOpenChange={setDialogOpen}
        defaultDni={dni}
      />

      {/* Portal de Impresión Limpia (1 sola página A4) */}
      <PrintBoletaPortal text={boletaText} />
    </div>
  );
}
