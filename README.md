# CAS UGEL - Boletas de Pago

Sistema de consulta y gestión de boletas para personal CAS (UGEL).

---

## Inicio Rápido (con Bun)

### 1. Clonar y entrar al proyecto
```bash
git clone <URL_DEL_REPOSITORIO>
cd cas-ugel
```

### 2. Instalar dependencias
```bash
bun install
```

### 3. Configurar variables de entorno
```bash
cp .env.example .env
```
Edita `.env` con tus claves de Supabase:
```env
VITE_SUPABASE_URL=https://tu-proyecto.supabase.co
VITE_SUPABASE_ANON_KEY=tu-anon-key-aqui
```

### 4. Iniciar en desarrollo
```bash
bun dev
```
Abre en tu navegador: `http://localhost:5173`

---

## Comandos Útiles

```bash
bun dev        # Iniciar servidor local
bun run build  # Generar build para producción
bun test       # Correr tests unitarios
bun run lint   # Revisar errores de linter
```

---

## Stack
React 18 + TypeScript + Vite + Tailwind CSS + Shadcn UI + Supabase
