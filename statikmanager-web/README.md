# StatikManager Web (Milestone 1)

Separates Mini-Projekt: ASP.NET Core Minimal API + Vite/TypeScript. Kein Bezug zum WPF-StatikManager im übergeordneten Ordner.

## Voraussetzungen

- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- [Node.js](https://nodejs.org/) (LTS, für npm)

## Schnellstart

### Variante A – Entwicklung (API + Vite getrennt)

Zwei Terminals:

**Terminal 1 – Backend**

```powershell
cd statikmanager-web\src\StatikManager.Api
dotnet run
```

API: `http://localhost:5156` (HTTPS: `https://localhost:7156`)

**Terminal 2 – Frontend (Vite, Proxy für `/api`)**

```powershell
cd statikmanager-web\src\StatikManager.Web
npm install
npm run dev
```

Browser: `http://localhost:5173` — ruft `/api/health` über den Proxy auf.

### Variante B – Ein Prozess (Frontend aus `wwwroot`)

Zuerst Frontend bauen (schreibt nach `StatikManager.Api/wwwroot`):

```powershell
cd statikmanager-web\src\StatikManager.Web
npm install
npm run build
```

Dann nur Backend:

```powershell
cd statikmanager-web\src\StatikManager.Api
dotnet run
```

Browser: `http://localhost:5156` — dieselbe Origin für UI und API.

## API (Milestone 1)

| Methode | Pfad           | Antwort        |
|---------|----------------|----------------|
| GET     | `/api/health`  | `{ "ok": true }` |

## Build / Typecheck

```powershell
cd statikmanager-web\src\StatikManager.Api
dotnet build

cd ..\StatikManager.Web
npm install
npm run build
```

## Projektstruktur

```
statikmanager-web/
├── StatikManagerWeb.sln
├── README.md
└── src/
    ├── StatikManager.Api/      # ASP.NET Core 8
    └── StatikManager.Web/       # Vite + TypeScript
```
