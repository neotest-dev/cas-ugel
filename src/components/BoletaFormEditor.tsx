import React from "react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { BoletaFormData } from "@/lib/payrollService";
import {
  User,
  Briefcase,
  Building,
  CreditCard,
  FileSpreadsheet,
  Calculator,
  Calendar,
  ShieldCheck,
} from "lucide-react";

interface BoletaFormEditorProps {
  formData: BoletaFormData;
  onChange: (field: keyof BoletaFormData, value: string) => void;
  onAutoCalculate: () => void;
}

export function BoletaFormEditor({
  formData,
  onChange,
  onAutoCalculate,
}: BoletaFormEditorProps) {
  return (
    <div className="space-y-4 p-4 sm:p-5 bg-white border border-slate-300 rounded shadow-sm text-xs">
      {/* 1. Datos Personales y Filiación */}
      <div className="border border-slate-200 rounded overflow-hidden">
        <div className="bg-[#0b223d] text-white px-3 py-2 flex items-center gap-2 font-bold uppercase tracking-wider text-[11px]">
          <User className="h-3.5 w-3.5 text-amber-400" />
          <span>1. Datos del Servidor (Filiación y Documento)</span>
        </div>

        <div className="p-3 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
          <div>
            <Label className="text-[11px] font-bold text-slate-700">Documento de Identidad (DNI)</Label>
            <Input
              value={formData.dni}
              disabled
              className="h-8 text-xs font-mono font-bold bg-slate-100 border-slate-300 cursor-not-allowed"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Apellido Paterno</Label>
            <Input
              value={formData.ap_paterno}
              onChange={(e) => onChange("ap_paterno", e.target.value.toUpperCase())}
              placeholder="APELLIDO PATERNO"
              className="h-8 text-xs bg-white border-slate-300 uppercase"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Apellido Materno</Label>
            <Input
              value={formData.ap_materno}
              onChange={(e) => onChange("ap_materno", e.target.value.toUpperCase())}
              placeholder="APELLIDO MATERNO"
              className="h-8 text-xs bg-white border-slate-300 uppercase"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Nombres Completos</Label>
            <Input
              value={formData.nombres}
              onChange={(e) => onChange("nombres", e.target.value.toUpperCase())}
              placeholder="NOMBRES"
              className="h-8 text-xs bg-white border-slate-300 uppercase"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-blue-900 flex items-center gap-1">
              <Calendar className="h-3 w-3 text-blue-600" /> Fecha de Nacimiento
            </Label>
            <Input
              value={formData.fecha_nac}
              onChange={(e) => onChange("fecha_nac", e.target.value)}
              placeholder="DD/MM/AAAA (Ej: 15/05/1985)"
              className="h-8 text-xs bg-white border-blue-400 focus:border-blue-600 font-semibold"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Código ESSALUD</Label>
            <Input
              value={formData.cod_essalud}
              onChange={(e) => onChange("cod_essalud", e.target.value.toUpperCase())}
              placeholder="Ej: 8505150DZFSA001"
              className="h-8 text-xs bg-white border-slate-300 uppercase"
            />
          </div>
        </div>
      </div>

      {/* 2. Datos del Puesto y Cargo */}
      <div className="border border-slate-200 rounded overflow-hidden">
        <div className="bg-[#132f50] text-white px-3 py-2 flex items-center gap-2 font-bold uppercase tracking-wider text-[11px]">
          <Briefcase className="h-3.5 w-3.5 text-blue-300" />
          <span>2. Cargo y Ubicación Laboral</span>
        </div>

        <div className="p-3 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
          <div className="sm:col-span-2">
            <Label className="text-[11px] font-bold text-slate-700">Cargo Desempeñado</Label>
            <Input
              value={formData.cargo}
              onChange={(e) => onChange("cargo", e.target.value.toUpperCase())}
              placeholder="Ej: PSICOLO JEC / COORDINADOR ADMIN"
              className="h-8 text-xs bg-white border-slate-300 uppercase font-semibold"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Cta. TeleAhorro / Nro. Cheque</Label>
            <Input
              value={formData.cuenta_banco}
              onChange={(e) => onChange("cuenta_banco", e.target.value)}
              placeholder="Ej: 04193284146"
              className="h-8 text-xs bg-white border-slate-300 font-mono"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Establecimiento</Label>
            <Input
              value="UGEL Nº 04 TRUJILLO SUR ESTE"
              disabled
              className="h-8 text-xs bg-slate-100 border-slate-300 cursor-not-allowed"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Régimen Laboral</Label>
            <Input
              value="D.LEG. Nº 1057 - CAS"
              disabled
              className="h-8 text-xs bg-slate-100 border-slate-300 cursor-not-allowed"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Tipo de Servidor</Label>
            <Input
              value="ADMINISTRATIVO CONTRATADO"
              disabled
              className="h-8 text-xs bg-slate-100 border-slate-300 cursor-not-allowed"
            />
          </div>

          <div className="sm:col-span-2">
            <Label className="text-[11px] font-bold text-slate-700">Leyenda Permanente (RD)</Label>
            <Input
              value={formData.leyenda_rd}
              onChange={(e) => onChange("leyenda_rd", e.target.value)}
              placeholder="Ej: RD N° 1234-2024-UGEL-TSE"
              className="h-8 text-xs bg-white border-slate-300"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">Leyenda Mensual</Label>
            <Input
              value={formData.leyenda_mensual}
              onChange={(e) => onChange("leyenda_mensual", e.target.value)}
              placeholder="Observación del mes"
              className="h-8 text-xs bg-white border-slate-300"
            />
          </div>
        </div>
      </div>

      {/* 3. Régimen Previsional */}
      <div className="border border-slate-200 rounded overflow-hidden">
        <div className="bg-[#1a3d66] text-white px-3 py-2 flex items-center gap-2 font-bold uppercase tracking-wider text-[11px]">
          <ShieldCheck className="h-3.5 w-3.5 text-emerald-400" />
          <span>3. Régimen Pensionario y Previsional</span>
        </div>

        <div className="p-3 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
          <div>
            <Label className="text-[11px] font-bold text-slate-700">Régimen Pensionario</Label>
            <Input
              value={formData.sistema_pensionario}
              onChange={(e) => onChange("sistema_pensionario", e.target.value.toUpperCase())}
              placeholder="Ej: ONP /W o INTEGRA /M"
              className="h-8 text-xs bg-white border-slate-300 uppercase font-semibold"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">CUSSP (Código AFP)</Label>
            <Input
              value={formData.cussp}
              onChange={(e) => onChange("cussp", e.target.value.toUpperCase())}
              placeholder="Ej: 584931JDFL0"
              className="h-8 text-xs bg-white border-slate-300 uppercase font-mono"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">F. Afiliación / Ingreso</Label>
            <Input
              value={formData.fecha_afiliacion}
              onChange={(e) => onChange("fecha_afiliacion", e.target.value)}
              placeholder="DD/MM/AAAA"
              className="h-8 text-xs bg-white border-slate-300"
            />
          </div>

          <div>
            <Label className="text-[11px] font-bold text-slate-700">F. Devengue / Término</Label>
            <Input
              value={formData.fecha_devengue}
              onChange={(e) => onChange("fecha_devengue", e.target.value)}
              placeholder="DD/MM/AAAA"
              className="h-8 text-xs bg-white border-slate-300"
            />
          </div>
        </div>
      </div>

      {/* 4. Importes, Descuentos y Total Líquido */}
      <div className="border border-slate-200 rounded overflow-hidden">
        <div className="bg-[#0b223d] text-white px-3 py-2 flex items-center justify-between font-bold uppercase tracking-wider text-[11px]">
          <div className="flex items-center gap-2">
            <Calculator className="h-3.5 w-3.5 text-amber-400" />
            <span>4. Conceptos Económicos y Liquidación (S/.)</span>
          </div>
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={onAutoCalculate}
            className="h-6 px-2 text-[10px] bg-white text-slate-800 border-slate-300 hover:bg-slate-100 font-bold"
            title="Recalcular automáticamente: Total Dscto = suma de aportes; Total Líquido = Remuneración - Dscto"
          >
            Auto-Calcular Líquido
          </Button>
        </div>

        <div className="p-3 bg-slate-50/50 space-y-3">
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
            <div className="bg-emerald-50 border border-emerald-200 rounded p-2.5">
              <Label className="text-[11px] font-bold text-emerald-900">
                Pago Total Mensual (Remuneración S/.)
              </Label>
              <Input
                value={formData.monto_mensual}
                onChange={(e) => onChange("monto_mensual", e.target.value)}
                placeholder="0.00"
                className="h-8 text-xs font-mono font-bold bg-white border-emerald-300 text-emerald-900 mt-1"
              />
            </div>

            <div className="bg-amber-50 border border-amber-200 rounded p-2.5">
              <Label className="text-[11px] font-bold text-amber-900">
                Total Descuentos (S/.)
              </Label>
              <Input
                value={formData.total_dscto}
                onChange={(e) => onChange("total_dscto", e.target.value)}
                placeholder="0.00"
                className="h-8 text-xs font-mono font-bold bg-white border-amber-300 text-amber-900 mt-1"
              />
            </div>

            <div className="bg-blue-50 border border-blue-200 rounded p-2.5">
              <Label className="text-[11px] font-bold text-blue-900">
                Total Líquido a Pagar (S/.)
              </Label>
              <Input
                value={formData.total_liquido}
                onChange={(e) => onChange("total_liquido", e.target.value)}
                placeholder="0.00"
                className="h-8 text-xs font-mono font-bold bg-white border-blue-300 text-blue-950 mt-1"
              />
            </div>
          </div>

          {/* Desglose de Descuentos Previsionales */}
          <div className="pt-2 border-t border-slate-200">
            <span className="text-[11px] font-bold text-slate-600 block mb-2">
              Desglose Individual de Pensiones (complete según corresponda):
            </span>
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2">
              <div>
                <Label className="text-[10px] text-slate-600">ONP (S/.)</Label>
                <Input
                  value={formData.onp}
                  onChange={(e) => onChange("onp", e.target.value)}
                  placeholder="0.00"
                  className="h-7 text-xs font-mono bg-white border-slate-300"
                />
              </div>

              <div>
                <Label className="text-[10px] text-slate-600">AFP Integra (S/.)</Label>
                <Input
                  value={formData.integra}
                  onChange={(e) => onChange("integra", e.target.value)}
                  placeholder="0.00"
                  className="h-7 text-xs font-mono bg-white border-slate-300"
                />
              </div>

              <div>
                <Label className="text-[10px] text-slate-600">AFP Profuturo (S/.)</Label>
                <Input
                  value={formData.profuturo}
                  onChange={(e) => onChange("profuturo", e.target.value)}
                  placeholder="0.00"
                  className="h-7 text-xs font-mono bg-white border-slate-300"
                />
              </div>

              <div>
                <Label className="text-[10px] text-slate-600">AFP Habitat (S/.)</Label>
                <Input
                  value={formData.habitat}
                  onChange={(e) => onChange("habitat", e.target.value)}
                  placeholder="0.00"
                  className="h-7 text-xs font-mono bg-white border-slate-300"
                />
              </div>

              <div>
                <Label className="text-[10px] text-slate-600">AFP Prima (S/.)</Label>
                <Input
                  value={formData.prima}
                  onChange={(e) => onChange("prima", e.target.value)}
                  placeholder="0.00"
                  className="h-7 text-xs font-mono bg-white border-slate-300"
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
