import { useState, useEffect } from "react";
import { supabase } from "@/lib/supabase";
import { TrabajadorConsulta } from "@/components/TrabajadorConsulta";
import { AdminLogin } from "@/components/AdminLogin";
import { AdminVentanilla } from "@/components/AdminVentanilla";
import { AdminUploadExcel } from "@/components/AdminUploadExcel";
import { AdminHistorialCargas } from "@/components/AdminHistorialCargas";
import { Button } from "@/components/ui/button";
import { Session } from "@supabase/supabase-js";
import {
  User,
  Shield,
  LogOut,
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
                <div className="flex items-center gap-2 bg-[#153457] border border-[#224b7a] px-3 py-1.5 rounded text-xs">
                  <span className="h-2 w-2 rounded-full bg-emerald-400" />
                  <span className="text-slate-200 font-mono font-medium truncate max-w-[200px]" title={session.user?.email}>
                    {session.user?.email}
                  </span>
                  <button
                    type="button"
                    onClick={handleLogout}
                    className="ml-2 text-slate-300 hover:text-red-300 text-xs font-semibold underline flex items-center gap-1"
                  >
                    <LogOut className="h-3.5 w-3.5" />
                    <span>Cerrar sesión</span>
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      {/* 2. Pestañas de Navegación Principales (Estilo Bootstrap Nav-Tabs) */}
      <nav className="no-print bg-[#132f50] text-slate-200 border-b border-slate-300 shadow-sm">
        <div className="container mx-auto px-4">
          <div className="flex space-x-1">
            <button
              type="button"
              onClick={() => setMainTab("trabajador")}
              className={`inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-semibold border-b-2 transition-all ${
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
              className={`inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-semibold border-b-2 transition-all ${
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
      <main className="container mx-auto px-4 py-5 flex-1 max-w-7xl">
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
                      <span>Cargar Planilla Excel</span>
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
