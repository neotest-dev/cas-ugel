import { useState } from "react";
import { BookOpen, Search, Upload, UserRound } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

type TutorialAudience = "trabajador" | "admin";
type AdminSection = "buscar" | "importar";

export function TutorialDialog({ audience }: { audience: TutorialAudience }) {
  const [open, setOpen] = useState(false);
  const [section, setSection] = useState<AdminSection>("buscar");
  const isAdmin = audience === "admin";

  return (
    <>
      <Button
        type="button"
        variant="outline"
        size="sm"
        onClick={() => setOpen(true)}
        className="gap-1.5 border-slate-300 bg-white text-xs font-semibold text-slate-700"
      >
        <BookOpen className="h-3.5 w-3.5 text-blue-700" />
        {isAdmin ? "Tutoriales" : "¿Cómo consultar?"}
      </Button>
      <Dialog open={open} onOpenChange={setOpen}>
        <DialogContent className="max-h-[85vh] overflow-y-auto border-slate-200 bg-white sm:max-w-2xl">
          <DialogHeader className="pr-6 text-left">
            <DialogTitle className="flex items-center gap-2 text-[#0b223d]">
              {isAdmin ? <BookOpen className="h-5 w-5 text-amber-500" /> : <UserRound className="h-5 w-5 text-blue-700" />}
              {isAdmin ? "Guía de Oficina de Planillas" : "Guía del Portal del Trabajador"}
            </DialogTitle>
            <DialogDescription>
              {isAdmin
                ? "Instrucciones para consultar boletas e importar planillas Excel."
                : "Consulta tu historial y descarga o imprime tus boletas de pago."}
            </DialogDescription>
          </DialogHeader>

          {isAdmin ? (
            <>
              <div className="flex flex-wrap gap-2 border-b pb-3">
                <Button type="button" size="sm" variant={section === "buscar" ? "default" : "outline"} onClick={() => setSection("buscar")}>
                  <Search className="mr-1.5 h-3.5 w-3.5" /> Buscar boletas
                </Button>
                <Button type="button" size="sm" variant={section === "importar" ? "default" : "outline"} onClick={() => setSection("importar")}>
                  <Upload className="mr-1.5 h-3.5 w-3.5" /> Importar a la base de datos
                </Button>
              </div>
              {section === "buscar" ? (
                <ol className="list-decimal space-y-3 pl-5 text-sm leading-relaxed text-slate-700">
                  <li>En <strong>Consultar</strong>, escribe DNI, apellidos o nombres. También puedes buscar usando solo los filtros avanzados.</li>
                  <li>Opcionalmente elige un tipo de CAS y las fechas desde/hasta. Puedes combinar los filtros, por ejemplo DNI más un rango de dos años.</li>
                  <li>Presiona <strong>Buscar</strong>. Revisa nombre, DNI, tipo CAS y periodo en los resultados; selecciona una boleta para abrirla.</li>
                  <li>Marca las casillas de las boletas que necesites o usa <strong>Seleccionar todas</strong>. Luego imprime o descarga el PDF combinado.</li>
                </ol>
              ) : (
                <ol className="list-decimal space-y-3 pl-5 text-sm leading-relaxed text-slate-700">
                  <li>Abre <strong>Importar Planilla Excel</strong> y selecciona o arrastra el archivo Excel de la planilla.</li>
                  <li>Verifica la categoría CAS y el mes/año detectados. Corrígelos si no corresponden.</li>
                  <li>Revisa el total de trabajadores y las boletas procesadas en la vista previa antes de publicar.</li>
                  <li>Presiona <strong>Publicar en Base de Datos</strong> y espera a que termine el progreso.</li>
                  <li>Si ya existe una planilla para la misma categoría y periodo, confirma el reemplazo solo después de verificar que elegiste el archivo correcto.</li>
                  <li>Confirma el registro en <strong>Historial de Planillas</strong>.</li>
                </ol>
              )}
            </>
          ) : (
            <ol className="list-decimal space-y-3 pl-5 text-sm leading-relaxed text-slate-700">
              <li>Ingresa tu DNI de ocho dígitos y tu contraseña personal. Si aún no tienes contraseña, puedes usar los últimos cuatro dígitos de la cuenta bancaria registrada.</li>
              <li>Presiona <strong>Consultar Boleta</strong>. El sistema mostrará tu historial de periodos disponibles.</li>
              <li>Selecciona un mes para ver la boleta correspondiente y confirma que el nombre y DNI sean los tuyos.</li>
              <li>Usa <strong>Descargar PDF</strong> para guardar la boleta o <strong>Imprimir Boleta</strong> para imprimirla.</li>
              <li>Si no recuerdas tu contraseña, usa el enlace para crearla o cambiarla. Si no aparece una boleta esperada, verifica el periodo y comunícate con la Oficina de Planillas.</li>
            </ol>
          )}
        </DialogContent>
      </Dialog>
    </>
  );
}
