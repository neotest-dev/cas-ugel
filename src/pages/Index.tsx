import { useState, useEffect } from "react";
import { supabase } from "@/lib/supabase";
import { TrabajadorConsulta } from "@/components/TrabajadorConsulta";
import { AdminLogin } from "@/components/AdminLogin";
import { AdminVentanilla } from "@/components/AdminVentanilla";
import { AdminUploadExcel } from "@/components/AdminUploadExcel";
import { AdminHistorialCargas } from "@/components/AdminHistorialCargas";
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
  User,
  Shield,
  LogOut,
  UserRound,
  ChevronDown,
  Users,
  Upload,
  History,
  Building,
  FileText,
} from "lucide-react";
import { toast } from "@/hooks/use-toast";

type MainTab = "trabajador" | "admin";
type AdminSubTab = "ventanilla" | "upload" | "historial";

const Index = () => {
  const [mainTab, setMainTab] = useState<MainTab>("trabajador");
  const [adminSubTab, setAdminSubTab] = useState<AdminSubTab>("ventanilla");
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
              {session && mainTab === "admin" && (
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

      {/* 2. Pestañas de Navegación Principales (Estilo Bootstrap Nav-Tabs) */}
      <nav className="no-print bg-[#132f50] text-slate-200 border-b border-slate-300 shadow-sm">
        <div className="container mx-auto px-4">
          <div className="flex w-full space-x-1 sm:w-auto">
            <button
              type="button"
              onClick={() => setMainTab("trabajador")}
              className={`inline-flex min-w-0 flex-1 items-center justify-center gap-1.5 px-2 py-2.5 text-center text-[11px] leading-tight font-semibold border-b-2 transition-all sm:flex-none sm:gap-2 sm:px-5 sm:py-3 sm:text-sm ${
                mainTab === "trabajador"
                  ? "bg-[#f4f6f9] text-[#0b223d] border-[#c59b27] font-bold shadow-inner"
                  : "text-slate-200 border-transparent hover:text-white hover:bg-[#1a3d66]"
              }`}
            >
              <User className="h-4 w-4" />
              <span>Portal del Trabajador</span>
            </button>

            <button
              type="button"
              onClick={() => setMainTab("admin")}
              className={`inline-flex min-w-0 flex-1 items-center justify-center gap-1.5 px-2 py-2.5 text-center text-[11px] leading-tight font-semibold border-b-2 transition-all sm:flex-none sm:gap-2 sm:px-5 sm:py-3 sm:text-sm ${
                mainTab === "admin"
                  ? "bg-[#f4f6f9] text-[#0b223d] border-[#c59b27] font-bold shadow-inner"
                  : "text-slate-200 border-transparent hover:text-white hover:bg-[#1a3d66]"
              }`}
            >
              <Shield className="h-4 w-4 text-amber-300" />
              <span>Oficina de Planillas</span>
            </button>
          </div>
        </div>
      </nav>

      {/* 3. Contenido Principal */}
      <main className="container mx-auto flex-1 max-w-7xl px-3 py-4 sm:px-4 sm:py-5">
        {/* MODO 1: Consulta del Trabajador */}
        {mainTab === "trabajador" && <TrabajadorConsulta />}

        {/* MODO 2: Oficina de Planillas (Admin) */}
        {mainTab === "admin" && (
          <div className="space-y-4">
            {!session && !authLoading ? (
              <AdminLogin onLoginSuccess={() => setAdminSubTab("ventanilla")} />
            ) : (
              <div className="space-y-4">
                {/* Sub-barra de herramientas Admin (Estilo Bootstrap Nav-Pills) */}
                <div className="bg-white border border-slate-300 rounded p-2 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                  <div className="flex flex-wrap gap-1.5">
                    <button
                      type="button"
                      onClick={() => setAdminSubTab("ventanilla")}
                      className={`inline-flex items-center gap-2 px-3.5 py-2 text-xs font-bold rounded border transition ${
                        adminSubTab === "ventanilla"
                          ? "bg-[#0b223d] text-white border-[#0b223d] shadow-sm"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                      }`}
                    >
                      <Users className="h-3.5 w-3.5 text-blue-400" />
                      <span>Consultar</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setAdminSubTab("upload")}
                      className={`inline-flex items-center gap-2 px-3.5 py-2 text-xs font-bold rounded border transition ${
                        adminSubTab === "upload"
                          ? "bg-[#0b223d] text-white border-[#0b223d] shadow-sm"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                      }`}
                    >
                      <Upload className="h-3.5 w-3.5 text-emerald-400" />
                      <span>Importar Planilla Excel</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setAdminSubTab("historial")}
                      className={`inline-flex items-center gap-2 px-3.5 py-2 text-xs font-bold rounded border transition ${
                        adminSubTab === "historial"
                          ? "bg-[#0b223d] text-white border-[#0b223d] shadow-sm"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                      }`}
                    >
                      <History className="h-3.5 w-3.5 text-amber-400" />
                      <span>Historial de Planillas</span>
                    </button>
                  </div>

                  <span className="text-[11px] text-slate-500 font-medium px-2 py-1 bg-slate-100 border border-slate-200 rounded self-start sm:self-auto">
                    Panel Administrativo CAS
                  </span>
                </div>

                {/* Vista Activa */}
                {adminSubTab === "ventanilla" && <AdminVentanilla />}
                {adminSubTab === "upload" && (
                  <AdminUploadExcel onPlanillaSaved={() => setAdminSubTab("historial")} />
                )}
                {adminSubTab === "historial" && <AdminHistorialCargas />}
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
