import { useState, ReactNode } from "react";
import { BookOpen, FolderOpen, Pencil, Search, Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

type SectionId = "importar" | "navegar" | "buscar" | "boleta";

interface Section {
  id: SectionId;
  label: string;
  icon: typeof Upload;
  intro: string;
  steps: ReactNode[];
  tip?: string;
}

const SECTIONS: Section[] = [
  {
    id: "importar",
    label: "Cargar planilla",
    icon: Upload,
    intro: "Sube el Excel del mes para generar las boletas de todos los trabajadores.",
    steps: [
      <>Entra a la pestaña <strong>Importar Excel</strong>.</>,
      <>Selecciona o arrastra el archivo Excel de la planilla (por ejemplo, <em>09 2026-PLLA CAS SEDE.xlsm</em>).</>,
      <>Revisa que el <strong>tipo de CAS</strong>, el <strong>mes</strong> y el <strong>año</strong> detectados sean correctos. Si no lo son, corrígelos.</>,
      <>Verifica en la vista previa el total de trabajadores y de boletas.</>,
      <>Presiona <strong>Publicar en Base de Datos</strong> y espera a que termine la barra de progreso.</>,
      <>Al terminar, la app te lleva a <strong>Planillas</strong> para que confirmes que el mes aparece con sus boletas.</>,
    ],
    tip: "Si ya existe una planilla del mismo tipo de CAS y mes, la app te pedirá confirmar el reemplazo. Asegúrate de haber elegido el archivo correcto.",
  },
  {
    id: "navegar",
    label: "Navegar planillas",
    icon: FolderOpen,
    intro: "Encuentra boletas recorriendo carpetas, sin escribir nada.",
    steps: [
      <>En <strong>Planillas</strong>, elige el <strong>año</strong>.</>,
      <>Elige el <strong>mes</strong>. Los meses sin planilla aparecen apagados.</>,
      <>Elige el <strong>tipo de CAS</strong> (SEDE, JEC, Mantenimiento, etc.). Verás cuántas boletas tiene cada uno.</>,
      <>Se abre la lista de trabajadores de ese CAS. Haz clic en uno para ver su boleta.</>,
      <>Usa la ruta de arriba (<em>Inicio › 2026 › Septiembre › CAS SEDE</em>) para volver a cualquier nivel.</>,
    ],
    tip: "En el nivel de tipos de CAS hay un desplegable \"Archivos Excel cargados\", donde puedes eliminar una planilla subida por error.",
  },
  {
    id: "buscar",
    label: "Buscar trabajador",
    icon: Search,
    intro: "Ubica a una persona rápido por su nombre, apellidos o DNI.",
    steps: [
      <>Presiona <strong>Buscar trabajador</strong> en la parte superior de Planillas.</>,
      <>Escribe el <strong>DNI</strong> o cualquier combinación de <strong>nombres y apellidos</strong>, en cualquier orden. Ejemplo: para <em>SARITA JUDITH DÍAZ FLORIAN</em> sirve escribir <em>Sarita Florian</em> o <em>Judith Florian</em>.</>,
      <>Si lo necesitas, reduce los resultados con <strong>Tipo de CAS</strong>, <strong>Año</strong> y <strong>Mes</strong>. Son opcionales.</>,
      <>Presiona <strong>Buscar</strong> o la tecla Enter.</>,
      <>Haz clic en un resultado de la lista para ver su boleta.</>,
    ],
    tip: "También puedes buscar solo con filtros, sin escribir nombre. Por ejemplo, todo el año 2026 de un CAS.",
  },
  {
    id: "boleta",
    label: "Imprimir y editar",
    icon: Pencil,
    intro: "Imprime la boleta para el trabajador o corrige sus datos en el sistema.",
    steps: [
      <>Abre una boleta. Por defecto estás en <strong>Imprimir</strong>: verás cómo saldrá en papel.</>,
      <>Presiona <strong>Imprimir</strong> o <strong>Descargar PDF</strong>. Para varias boletas, marca las casillas de la lista y usa <strong>Imprimir seleccionadas</strong> o <strong>PDF combinado</strong>.</>,
      <>Para corregir datos (cargo, fechas, montos, etc.), cambia a <strong>Editar datos</strong>.</>,
      <>Modifica los campos. Con <strong>Calcular totales</strong> se recalculan el total de descuentos y el líquido a pagar.</>,
      <>Presiona <strong>Guardar cambios</strong> para grabarlos en la base de datos. <strong>Deshacer cambios</strong> devuelve los datos originales de la planilla.</>,
    ],
    tip: "Los ajustes que hagas al texto en el modo Imprimir solo valen para esa impresión y no se guardan. Para cambios permanentes usa Editar datos.",
  },
];

export function TutorialDialog() {
  const [open, setOpen] = useState(false);
  const [sectionId, setSectionId] = useState<SectionId>("importar");
  const section = SECTIONS.find((item) => item.id === sectionId) ?? SECTIONS[0];

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
        Ayuda
      </Button>

      <Dialog open={open} onOpenChange={setOpen}>
        <DialogContent className="max-h-[90vh] overflow-y-auto border-slate-200 bg-white sm:max-w-2xl">
          <DialogHeader className="pr-6 text-left">
            <DialogTitle className="flex items-center gap-2 text-[#0b223d]">
              <BookOpen className="h-5 w-5 text-amber-500" />
              ¿Cómo uso el sistema?
            </DialogTitle>
            <DialogDescription>Elige una tarea para ver los pasos.</DialogDescription>
          </DialogHeader>

          <div role="tablist" aria-label="Tutoriales" className="grid grid-cols-2 gap-2 sm:grid-cols-4">
            {SECTIONS.map(({ id, label, icon: Icon }) => (
              <button
                key={id}
                type="button"
                role="tab"
                aria-selected={sectionId === id}
                onClick={() => setSectionId(id)}
                className={`flex flex-col items-center gap-1 rounded-lg border px-2 py-3 text-xs font-bold transition ${
                  sectionId === id
                    ? "border-[#0b223d] bg-[#0b223d] text-white shadow-sm"
                    : "border-slate-200 bg-white text-slate-700 hover:border-slate-400"
                }`}
              >
                <Icon className="h-4 w-4" />
                {label}
              </button>
            ))}
          </div>

          <div className="space-y-4">
            <p className="text-sm text-slate-600">{section.intro}</p>

            <ol className="space-y-2.5">
              {section.steps.map((step, index) => (
                <li key={index} className="flex gap-3 text-sm leading-relaxed text-slate-700">
                  <span className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-100 text-xs font-bold text-blue-800">
                    {index + 1}
                  </span>
                  <span>{step}</span>
                </li>
              ))}
            </ol>

            {section.tip && (
              <p className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs leading-relaxed text-amber-900">
                <strong>Importante:</strong> {section.tip}
              </p>
            )}
          </div>
        </DialogContent>
      </Dialog>
    </>
  );
}
