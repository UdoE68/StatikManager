# StatikManager Web

Separates Mini-Projekt: ASP.NET Core Minimal API + Vite/TypeScript. Kein Bezug zum WPF-StatikManager im übergeordneten Ordner.

**Stand:** Milestone 2 (Projekt-Root setzen/abfragen, nur Arbeitsspeicher).

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

Browser: `http://localhost:5173` — UI für Projektpfad und Session-API (Proxy `/api`).

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

## API

| Methode | Pfad                 | Antwort / Hinweis |
|---------|----------------------|-------------------|
| GET     | `/api/health`        | `{ "ok": true }` |
| GET     | `/api/session/root`  | `{ "rootPath": "<absolut>" \| null }` — gesetztes Root (nur RAM) |
| POST    | `/api/session/root`  | Body: `{ "rootPath": "..." }`. Erfolg **200** und gleiche JSON wie GET. Fehler **400** mit `{ "error": "..." }` (z. B. leerer Pfad, existiert nicht, ist eine Datei). |

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
