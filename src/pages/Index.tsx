import { useState, useEffect } from "react";
import { supabase } from "@/lib/supabase";
import { AdminLogin } from "@/components/AdminLogin";
import { AdminUploadExcel } from "@/components/AdminUploadExcel";
import { AdminPlanillas } from "@/components/AdminPlanillas";
import { TutorialDialog } from "@/components/TutorialDialog";
import { Session } from "@supabase/supabase-js";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  LogOut,
  UserRound,
  ChevronDown,
  Upload,
  FolderOpen,
  Building,
} from "lucide-react";
import { toast } from "@/hooks/use-toast";

type AdminSubTab = "planillas" | "upload";

const Index = () => {
  const [adminSubTab, setAdminSubTab] = useState<AdminSubTab>("planillas");
  const [session, setSession] = useState<Session | null>(null);
  const [authLoading, setAuthLoading] = useState(true);

  useEffect(() => {
    document.title = "CAS - UGEL 04 TSE";
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthLoading(false);
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      setAuthLoading(false);
    });

    return () => subscription.unsubscribe();
  }, []);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    toast({
      title: "Sesión finalizada",
      description: "Has salido del sistema de planillas de forma segura.",
    });
  };

  return (
    <div className="min-h-screen bg-[#f4f6f9] text-slate-800 flex flex-col font-sans">
      {/* 1. Barra Institucional Superior (Estilo Gobierno / UGEL) */}
      <header className="no-print bg-[#0b223d] text-white border-b-4 border-[#c59b27] shadow-sm">
        <div className="container mx-auto px-4 py-2.5">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
            {/* Identidad y Escudo */}
            <div className="flex items-center gap-3">
              <div className="h-10 w-10 bg-white rounded p-0.5 flex items-center justify-center shrink-0 shadow-sm border border-slate-200 overflow-hidden">
                <img src="/ugel.jpg" alt="Logo UGEL 04" className="h-full w-full object-contain" />
              </div>
              <div>
                <h1 className="text-base sm:text-lg font-bold tracking-tight text-white leading-tight">
                  UGEL Nº 04 TRUJILLO SUR ESTE
                </h1>
                <p className="text-[11px] text-slate-300">
                  Sistema Integrado de Boletas de Pago CAS
                </p>
              </div>
            </div>

            {/* Estado de Sesión / Botón Salir */}
            <div className="flex items-center gap-2 self-end md:self-center">
              {session && (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <button
                      type="button"
                      className="inline-flex items-center gap-2 rounded-md border border-[#31577e] bg-[#153457] px-3 py-2 text-left text-xs font-semibold text-white shadow-sm transition hover:bg-[#1b4169] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-amber-400"
                      aria-label="Abrir menú de cuenta administrativa"
                    >
                      <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#254d75] text-emerald-400">
                        <UserRound className="h-4 w-4" />
                      </span>
                      <span>ADMIN</span>
                      <ChevronDown className="h-3.5 w-3.5 text-slate-300" />
                    </button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end" sideOffset={8} className="w-72 border-slate-200 bg-white p-1.5 shadow-xl">
                    <DropdownMenuLabel className="px-3 pb-1 pt-2 text-sm font-bold tracking-wide text-[#0b223d]">
                      ADMIN - UGEL04TSE
                    </DropdownMenuLabel>
                    <div className="px-3 pb-2 text-xs text-slate-500 break-all" title={session.user?.email}>
                      {session.user?.email}
                    </div>
                    <DropdownMenuSeparator className="bg-slate-200" />
                    <DropdownMenuItem
                      onSelect={handleLogout}
                      className="mt-1 cursor-pointer gap-2 px-3 py-2 text-sm font-semibold text-red-700 focus:bg-red-50 focus:text-red-800"
                    >
                      <LogOut className="h-4 w-4" />
                      Cerrar sesión
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>
              )}
            </div>
          </div>
        </div>
      </header>

      {/* 3. Contenido Principal */}
      <main className="container mx-auto flex-1 max-w-7xl px-3 py-4 sm:px-4 sm:py-5">
        {/* Oficina de Planillas (atención presencial, solo administradores) */}
        {(
          <div className="space-y-4">
            {!session && !authLoading ? (
              <AdminLogin onLoginSuccess={() => setAdminSubTab("planillas")} />
            ) : (
              <div className="space-y-4">
                {/* Pestañas */}
                <div className="flex gap-1.5">
                  {([
                    { id: "planillas", label: "Planillas", icon: FolderOpen, color: "text-blue-400" },
                    { id: "upload", label: "Importar Excel", icon: Upload, color: "text-emerald-400" },
                  ] as const).map(({ id, label, icon: Icon, color }) => (
                    <button
                      key={id}
                      type="button"
                      onClick={() => setAdminSubTab(id)}
                      className={`inline-flex items-center gap-2 rounded border px-4 py-2 text-xs font-bold transition ${
                        adminSubTab === id
                          ? "border-[#0b223d] bg-[#0b223d] text-white shadow-sm"
                          : "border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
                      }`}
                    >
                      <Icon className={`h-3.5 w-3.5 ${color}`} />
                      {label}
                    </button>
                  ))}
                  <div className="ml-auto">
                    <TutorialDialog />
                  </div>
                </div>

                {adminSubTab === "planillas" && <AdminPlanillas />}
                {adminSubTab === "upload" && (
                  <AdminUploadExcel onPlanillaSaved={() => setAdminSubTab("planillas")} />
                )}
              </div>
            )}
          </div>
        )}
      </main>

      {/* 4. Footer Formal Institucional */}
      <footer className="no-print mt-auto bg-white border-t border-slate-300 py-3 text-xs text-slate-600">
        <div className="container mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <Building className="h-4 w-4 text-slate-400" />
            <p>
              <strong>UGEL Nº 04 Trujillo Sur Este</strong>
            </p>
          </div>
          <p className="text-slate-500 font-mono text-[11px]">
            Versión 3.0 · neotest-dev
          </p>
        </div>
      </footer>
    </div>
  );
};

export default Index;
