import { useState, useEffect } from "react";
import {
  fetchCargasPlanilla,
  CargaPlanillaItem,
} from "@/lib/payrollService";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  History,
  RefreshCw,
  FileSpreadsheet,
  Users,
  Calendar,
  Loader2,
} from "lucide-react";

export function AdminHistorialCargas() {
  const [cargas, setCargas] = useState<CargaPlanillaItem[]>([]);
  const [loading, setLoading] = useState(false);

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

  return (
    <div className="space-y-3">
      <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        {/* Encabezado Clásico */}
        <div className="bg-[#0b223d] text-white px-4 py-2.5 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <History className="h-4 w-4 text-amber-400" />
            <h3 className="text-xs font-bold uppercase tracking-wider">
              Historial de Planillas Registradas en Base de Datos
            </h3>
          </div>

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

        {/* Tabla Estilo Bootstrap */}
        <div className="p-0">
          {loading && cargas.length === 0 ? (
            <div className="py-12 text-center">
              <Loader2 className="mx-auto h-6 w-6 animate-spin text-[#0d6efd] mb-2" />
              <p className="text-xs text-slate-500">Cargando registros...</p>
            </div>
          ) : cargas.length === 0 ? (
            <div className="py-12 text-center p-4">
              <FileSpreadsheet className="mx-auto h-10 w-10 text-slate-300 mb-2" />
              <h4 className="text-xs font-bold text-slate-800">
                No existen planillas publicadas
              </h4>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Vaya a &quot;Importar Planilla Excel&quot; para publicar el primer archivo.
              </p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <Table className="text-xs">
                <TableHeader>
                  <TableRow className="bg-slate-100 border-b border-slate-300 hover:bg-slate-100">
                    <TableHead className="font-bold text-slate-800">Periodo</TableHead>
                    <TableHead className="font-bold text-slate-800">Categoría CAS</TableHead>
                    <TableHead className="font-bold text-slate-800">Boletas</TableHead>
                    <TableHead className="font-bold text-slate-800">Archivo Original</TableHead>
                    <TableHead className="font-bold text-slate-800">Fecha de Registro</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {cargas.map((item, index) => (
                    <TableRow
                      key={item.id}
                      className={index % 2 === 0 ? "bg-white" : "bg-slate-50/70"}
                    >
                      <TableCell className="font-bold text-slate-900">
                        <span className="inline-flex items-center gap-1.5">
                          <Calendar className="h-3.5 w-3.5 text-[#0d6efd]" />
                          {item.mes} {item.anio}
                        </span>
                      </TableCell>
                      <TableCell>
                        <span className="bg-[#e7f1ff] border border-[#b6d4fe] text-[#084298] px-2 py-0.5 rounded text-[11px] font-semibold">
                          {item.categoria_label}
                        </span>
                      </TableCell>
                      <TableCell>
                        <span className="inline-flex items-center gap-1 text-slate-800 font-mono font-bold">
                          <Users className="h-3 w-3 text-slate-400" />
                          {item.total_trabajadores}
                        </span>
                      </TableCell>
                      <TableCell className="text-slate-600 font-mono text-[11px] max-w-[200px] truncate" title={item.nombre_archivo}>
                        {item.nombre_archivo}
                      </TableCell>
                      <TableCell className="text-slate-600 text-[11px]">
                        {new Date(item.created_at).toLocaleString("es-PE", {
                          day: "2-digit",
                          month: "short",
                          year: "numeric",
                          hour: "2-digit",
                          minute: "2-digit",
                        })}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
