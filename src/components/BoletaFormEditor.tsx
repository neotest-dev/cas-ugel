import { ReactNode } from "react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { BoletaFormData } from "@/lib/payrollService";
import { User, Briefcase, ShieldCheck, Calculator } from "lucide-react";

interface BoletaFormEditorProps {
  formData: BoletaFormData;
  onChange: (field: keyof BoletaFormData, value: string) => void;
  onAutoCalculate: () => void;
}

const inputClass = "h-9 bg-white text-sm";
const moneyClass = `${inputClass} font-mono`;

function Section({ icon, title, action, children }: { icon: ReactNode; title: string; action?: ReactNode; children: ReactNode }) {
  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4">
      <div className="mb-3 flex items-center justify-between gap-2">
        <h5 className="flex items-center gap-2 text-sm font-bold text-[#0b223d]">
          {icon}
          {title}
        </h5>
        {action}
      </div>
      {children}
    </section>
  );
}

function Field({ label, className = "", children }: { label: string; className?: string; children: ReactNode }) {
  return (
    <div className={`space-y-1 ${className}`}>
      <Label className="text-xs font-semibold text-slate-600">{label}</Label>
      {children}
    </div>
  );
}

export function BoletaFormEditor({ formData, onChange, onAutoCalculate }: BoletaFormEditorProps) {
  const text = (field: keyof BoletaFormData, placeholder: string, upper = false, extra = "") => (
    <Input
      value={formData[field] as string}
      onChange={(e) => onChange(field, upper ? e.target.value.toUpperCase() : e.target.value)}
      placeholder={placeholder}
      className={`${inputClass} ${upper ? "uppercase" : ""} ${extra}`}
    />
  );

  const money = (field: keyof BoletaFormData, extra = "") => (
    <Input
      value={formData[field] as string}
      onChange={(e) => onChange(field, e.target.value)}
      inputMode="decimal"
      placeholder="0.00"
      className={`${moneyClass} ${extra}`}
    />
  );

  return (
    <div className="space-y-4 bg-slate-50 p-4">
      <Section icon={<User className="h-4 w-4 text-blue-600" />} title="Datos personales">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <Field label="DNI">
            <Input value={formData.dni} disabled className={`${inputClass} cursor-not-allowed bg-slate-100 font-mono font-bold`} />
          </Field>
          <Field label="Apellido paterno">{text("ap_paterno", "Apellido paterno", true)}</Field>
          <Field label="Apellido materno">{text("ap_materno", "Apellido materno", true)}</Field>
          <Field label="Nombres" className="lg:col-span-1">{text("nombres", "Nombres", true)}</Field>
          <Field label="Fecha de nacimiento">{text("fecha_nac", "DD/MM/AAAA")}</Field>
          <Field label="Código EsSalud">{text("cod_essalud", "Ej: 8505150DZFSA001", true)}</Field>
        </div>
      </Section>

      <Section icon={<Briefcase className="h-4 w-4 text-blue-600" />} title="Cargo y observaciones">
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Cargo">{text("cargo", "Ej: PSICÓLOGO JEC", true)}</Field>
          <Field label="Cuenta TeleAhorro / Nro. cheque">{text("cuenta_banco", "Ej: 04193284146", false, "font-mono")}</Field>
          <Field label="Leyenda permanente (RD)">{text("leyenda_rd", "Ej: RD N° 1234-2024-UGEL-TSE")}</Field>
          <Field label="Leyenda mensual">{text("leyenda_mensual", "Observación del mes")}</Field>
        </div>
      </Section>

      <Section icon={<ShieldCheck className="h-4 w-4 text-emerald-600" />} title="Régimen pensionario">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <Field label="Régimen">{text("sistema_pensionario", "Ej: ONP o INTEGRA", true)}</Field>
          <Field label="CUSSP">{text("cussp", "Código AFP", true, "font-mono")}</Field>
          <Field label="Fecha de afiliación">{text("fecha_afiliacion", "DD/MM/AAAA")}</Field>
          <Field label="Fecha de devengue">{text("fecha_devengue", "DD/MM/AAAA")}</Field>
        </div>
      </Section>

      <Section
        icon={<Calculator className="h-4 w-4 text-amber-600" />}
        title="Pagos y descuentos (S/.)"
        action={
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={onAutoCalculate}
            className="h-8 text-xs font-semibold"
            title="Total descuentos = suma de aportes; Total líquido = Remuneración − Descuentos"
          >
            <Calculator className="mr-1.5 h-3.5 w-3.5" /> Calcular totales
          </Button>
        }
      >
        <div className="grid gap-3 sm:grid-cols-3">
          <Field label="Remuneración mensual">{money("monto_mensual", "font-bold")}</Field>
          <Field label="Total descuentos">{money("total_dscto", "font-bold")}</Field>
          <Field label="Total líquido a pagar">{money("total_liquido", "border-emerald-400 font-bold text-emerald-800")}</Field>
        </div>

        <details className="mt-4 rounded-md border border-slate-200 bg-slate-50">
          <summary className="cursor-pointer select-none px-3 py-2 text-xs font-semibold text-slate-700">
            Detalle de descuentos de pensión
          </summary>
          <div className="grid grid-cols-2 gap-3 border-t border-slate-200 p-3 sm:grid-cols-3 lg:grid-cols-5">
            <Field label="ONP">{money("onp")}</Field>
            <Field label="AFP Integra">{money("integra")}</Field>
            <Field label="AFP Profuturo">{money("profuturo")}</Field>
            <Field label="AFP Habitat">{money("habitat")}</Field>
            <Field label="AFP Prima">{money("prima")}</Field>
          </div>
        </details>

        <details className="mt-3 rounded-md border border-slate-200 bg-slate-50">
          <summary className="cursor-pointer select-none px-3 py-2 text-xs font-semibold text-slate-700">
            Otros descuentos y retenciones
          </summary>
          <div className="grid grid-cols-1 gap-3 border-t border-slate-200 p-3 sm:grid-cols-3">
            <Field label="Otros descuentos">{money("otros_dsctos")}</Field>
            <Field label="Descuento entidades">{money("dscto_entidades")}</Field>
            <Field label="Descuento judicial">{money("dscto_judicial")}</Field>
          </div>
        </details>
      </Section>
    </div>
  );
}
