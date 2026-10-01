import { useState } from "react";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "@/hooks/use-toast";
import { cambiarClaveTrabajador } from "@/lib/payrollService";
import { KeyRound, Loader2, ShieldCheck, Info } from "lucide-react";

interface CambiarClaveDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  defaultDni?: string;
}

export function CambiarClaveDialog({
  open,
  onOpenChange,
  defaultDni = "",
}: CambiarClaveDialogProps) {
  const [dni, setDni] = useState(defaultDni);
  const [ultimos4, setUltimos4] = useState("");
  const [nuevaClave, setNuevaClave] = useState("");
  const [confirmarClave, setConfirmarClave] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!dni || dni.length !== 8) {
      toast({
        title: "DNI inválido",
        description: "El DNI debe contener 8 dígitos numéricos.",
        variant: "destructive",
      });
      return;
    }

    if (!ultimos4 || ultimos4.length !== 4) {
      toast({
        title: "Dígitos de cuenta inválidos",
        description: "Ingrese exactamente los 4 dígitos finales de su cuenta del Banco de la Nación.",
        variant: "destructive",
      });
      return;
    }

    if (!nuevaClave || nuevaClave.length < 4) {
      toast({
        title: "Contraseña muy corta",
        description: "La nueva contraseña debe tener un mínimo de 4 caracteres.",
        variant: "destructive",
      });
      return;
    }

    if (nuevaClave !== confirmarClave) {
      toast({
        title: "Las contraseñas no coinciden",
        description: "Asegúrese de escribir exactamente la misma contraseña en ambos campos.",
        variant: "destructive",
      });
      return;
    }

    setLoading(true);
    try {
      const res = await cambiarClaveTrabajador(dni, ultimos4, nuevaClave);
      if (!res.ok) {
        toast({
          title: "Validación no superada",
          description: res.error || "No se pudo actualizar la contraseña.",
          variant: "destructive",
        });
        return;
      }

      toast({
        title: "Contraseña actualizada exitosamente",
        description: "A partir de ahora puede consultar sus boletas con su nueva clave personal.",
      });

      setUltimos4("");
      setNuevaClave("");
      setConfirmarClave("");
      onOpenChange(false);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md bg-white border border-slate-300 rounded p-0 overflow-hidden shadow-lg">
        {/* Encabezado Modal Bootstrap */}
        <div className="bg-[#0b223d] text-white px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <KeyRound className="h-4 w-4 text-amber-400" />
            <DialogTitle className="text-xs font-bold uppercase tracking-wider text-white">
              Crear o Actualizar Contraseña Personal
            </DialogTitle>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="p-4 sm:p-5 space-y-3.5">
          <div className="bg-[#e7f1ff] border border-[#b6d4fe] text-[#084298] p-2.5 rounded text-xs flex items-start gap-2">
            <Info className="h-4 w-4 shrink-0 mt-0.5 text-[#0d6efd]" />
            <p className="leading-snug text-[11px]">
              Para verificar su identidad, ingrese los <strong>últimos 4 dígitos de su cuenta bancaria</strong> registrada en su boleta de pago de la UGEL 04.
            </p>
          </div>

          <div className="space-y-1">
            <Label htmlFor="dlg-dni" className="text-xs font-bold text-slate-700">
              Número de DNI:
            </Label>
            <Input
              id="dlg-dni"
              type="text"
              inputMode="numeric"
              maxLength={8}
              placeholder="Ejemplo: 41234567"
              value={dni}
              onChange={(e) => setDni(e.target.value.replace(/\D/g, ""))}
              className="h-9 rounded border-slate-300 text-xs bg-white focus:border-blue-600"
              required
            />
          </div>

          <div className="space-y-1">
            <Label htmlFor="dlg-cta" className="text-xs font-bold text-slate-700">
              Últimos 4 dígitos de su Cuenta Banco de la Nación:
            </Label>
            <Input
              id="dlg-cta"
              type="password"
              inputMode="numeric"
              maxLength={4}
              placeholder="••••"
              value={ultimos4}
              onChange={(e) => setUltimos4(e.target.value.replace(/\D/g, ""))}
              className="h-9 rounded border-slate-300 text-xs tracking-widest bg-white focus:border-blue-600"
              required
            />
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
            <div className="space-y-1">
              <Label htmlFor="dlg-pass" className="text-xs font-bold text-slate-700">
                Nueva Contraseña:
              </Label>
              <Input
                id="dlg-pass"
                type="password"
                placeholder="Mínimo 4 caracteres"
                value={nuevaClave}
                onChange={(e) => setNuevaClave(e.target.value)}
                className="h-9 rounded border-slate-300 text-xs bg-white focus:border-blue-600"
                required
              />
            </div>
            <div className="space-y-1">
              <Label htmlFor="dlg-conf" className="text-xs font-bold text-slate-700">
                Confirmar Contraseña:
              </Label>
              <Input
                id="dlg-conf"
                type="password"
                placeholder="Repita la clave"
                value={confirmarClave}
                onChange={(e) => setConfirmarClave(e.target.value)}
                className="h-9 rounded border-slate-300 text-xs bg-white focus:border-blue-600"
                required
              />
            </div>
          </div>

          <DialogFooter className="pt-3 border-t border-slate-200 sm:justify-end gap-2">
            <Button
              type="button"
              variant="outline"
              onClick={() => onOpenChange(false)}
              className="h-8 rounded text-xs border-slate-300 text-slate-700"
            >
              Cancelar
            </Button>
            <Button
              type="submit"
              disabled={loading}
              className="h-8 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white text-xs font-bold shadow-sm"
            >
              {loading ? (
                <>
                  <Loader2 className="mr-1.5 h-3 w-3 animate-spin" /> Guardando...
                </>
              ) : (
                <>
                  <ShieldCheck className="mr-1.5 h-3.5 w-3.5" /> Guardar Contraseña
                </>
              )}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
