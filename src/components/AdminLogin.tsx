import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "@/hooks/use-toast";
import { supabase } from "@/lib/supabase";
import { Lock, Mail, Loader2, ShieldCheck, AlertCircle } from "lucide-react";

interface AdminLoginProps {
  onLoginSuccess: () => void;
}

export function AdminLogin({ onLoginSuccess }: AdminLoginProps) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) {
      toast({
        title: "Campos requeridos",
        description: "Ingrese su correo y contraseña de administrador.",
        variant: "destructive",
      });
      return;
    }

    setLoading(true);
    try {
      const { data, error } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      });

      if (error) {
        toast({
          title: "Acceso denegado",
          description: error.message === "Invalid login credentials"
            ? "Correo o contraseña incorrectos. Verifique sus credenciales."
            : error.message,
          variant: "destructive",
        });
        return;
      }

      if (data.session) {
        toast({
          title: "Acceso autorizado",
          description: `Bienvenido al módulo de planillas, ${data.user?.email}`,
        });
        onLoginSuccess();
      }
    } catch (err) {
      toast({
        title: "Error de conexión",
        description: "No se pudo conectar con el servidor de autenticación.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="mx-auto max-w-md py-6">
      <div className="bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        {/* Encabezado Clásico */}
        <div className="bg-[#0b223d] text-white px-5 py-3 border-b border-slate-300 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Lock className="h-4 w-4 text-amber-400" />
            <h3 className="text-xs font-bold uppercase tracking-wider">
              Acceso a la Oficina de Planillas
            </h3>
          </div>
          <span className="text-[11px] bg-[#1a3d66] text-slate-200 px-2 py-0.5 rounded border border-[#234e80]">
            UGEL 04
          </span>
        </div>

        <div className="p-5 sm:p-6 space-y-4">
          <div className="bg-amber-50 border border-amber-200 text-amber-900 p-2.5 rounded text-xs flex items-start gap-2">
            <AlertCircle className="h-4 w-4 text-amber-700 shrink-0 mt-0.5" />
            <p className="leading-snug">
              Módulo de uso exclusivo para el personal del <strong>Área de Planillas</strong> de la UGEL 04 Trujillo Sur Este.
            </p>
          </div>

          <form onSubmit={handleLogin} className="space-y-3.5">
            <div className="space-y-1">
              <Label htmlFor="admin-email" className="text-xs font-bold text-slate-700">
                Correo Electrónico Institucional:
              </Label>
              <div className="relative">
                <Mail className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
                <Input
                  id="admin-email"
                  type="email"
                  placeholder="usuario@ugel04.pe"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="h-10 rounded border-slate-300 pl-9 text-xs bg-white focus:border-blue-600"
                  required
                />
              </div>
            </div>

            <div className="space-y-1">
              <Label htmlFor="admin-pass" className="text-xs font-bold text-slate-700">
                Contraseña de Acceso:
              </Label>
              <div className="relative">
                <Lock className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
                <Input
                  id="admin-pass"
                  type="password"
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="h-10 rounded border-slate-300 pl-9 text-xs bg-white focus:border-blue-600"
                  required
                />
              </div>
            </div>

            <Button
              type="submit"
              disabled={loading}
              className="w-full h-10 rounded bg-[#0d6efd] hover:bg-[#0b5ed7] text-white font-bold text-xs shadow-sm transition"
            >
              {loading ? (
                <>
                  <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> Verificando...
                </>
              ) : (
                <>
                  <ShieldCheck className="mr-1.5 h-4 w-4" /> Iniciar Sesión en Planillas
                </>
              )}
            </Button>
          </form>
        </div>
      </div>
    </div>
  );
}
