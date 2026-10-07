import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "@/hooks/use-toast";
import { supabase } from "@/lib/supabase";
import { Lock, Mail, Loader2, Eye, EyeOff, LogIn } from "lucide-react";

interface AdminLoginProps {
  onLoginSuccess: () => void;
}

export function AdminLogin({ onLoginSuccess }: AdminLoginProps) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) {
      toast({
        title: "Campos requeridos",
        description: "Ingrese su correo y contraseña.",
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
          description:
            error.message === "Invalid login credentials"
              ? "Correo o contraseña incorrectos."
              : error.message,
          variant: "destructive",
        });
        return;
      }

      if (data.session) {
        onLoginSuccess();
      }
    } catch {
      toast({
        title: "Error de conexión",
        description: "No se pudo conectar con el servidor.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex min-h-[60vh] items-center justify-center px-2 py-6">
      <form
        onSubmit={handleLogin}
        className="w-full max-w-sm space-y-5 rounded-xl border border-slate-200 bg-white p-6 shadow-md sm:p-8"
      >
        <div className="space-y-1 text-center">
          <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-full bg-[#0b223d] text-amber-400">
            <Lock className="h-5 w-5" />
          </div>
          <h2 className="pt-2 text-lg font-bold text-[#0b223d]">Oficina de Planillas</h2>
          <p className="text-xs text-slate-500">Ingresa con tu cuenta de administrador</p>
        </div>

        <div className="space-y-1.5">
          <Label htmlFor="admin-email" className="text-sm font-semibold text-slate-700">
            Correo
          </Label>
          <div className="relative">
            <Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
            <Input
              id="admin-email"
              type="email"
              autoComplete="username"
              placeholder="usuario@ugel04.pe"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="h-11 pl-9 text-base sm:text-sm"
              required
              autoFocus
            />
          </div>
        </div>

        <div className="space-y-1.5">
          <Label htmlFor="admin-pass" className="text-sm font-semibold text-slate-700">
            Contraseña
          </Label>
          <div className="relative">
            <Lock className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
            <Input
              id="admin-pass"
              type={showPassword ? "text" : "password"}
              autoComplete="current-password"
              placeholder="Tu contraseña"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="h-11 pl-9 pr-11 text-base sm:text-sm"
              required
            />
            <button
              type="button"
              onClick={() => setShowPassword((v) => !v)}
              aria-label={showPassword ? "Ocultar contraseña" : "Mostrar contraseña"}
              className="absolute right-1.5 top-1/2 flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded text-slate-500 transition hover:bg-slate-100 hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500"
            >
              {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
            </button>
          </div>
        </div>

        <Button
          type="submit"
          disabled={loading}
          className="h-11 w-full bg-[#0d6efd] text-sm font-bold text-white hover:bg-[#0b5ed7]"
        >
          {loading ? (
            <>
              <Loader2 className="mr-2 h-4 w-4 animate-spin" /> Verificando...
            </>
          ) : (
            <>
              <LogIn className="mr-2 h-4 w-4" /> Ingresar
            </>
          )}
        </Button>
      </form>
    </div>
  );
}
