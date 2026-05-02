# StatikManager Web

Separates Mini-Projekt: ASP.NET Core Minimal API + Vite/TypeScript. Kein Bezug zum WPF-StatikManager im übergeordneten Ordner.

**Stand:** Milestone 5+ — nativer **Ordnerdialog** über **Tauri**-Desktop; API ohne WinForms-Dialog.

## Voraussetzungen

- [.NET 8 SDK](https://dotnet.microsoft.com/download) (Windows-TFM `net8.0-windows` für die API)
- [Node.js](https://nodejs.org/) (LTS, für npm)
- Für **Tauri** (`npm run tauri:dev`): [Rust / rustup](https://rustup.rs/) (Cargo im `PATH`)

**Ordner wählen:** Im **Tauri**-Fenster nativer Windows-Dialog; im **Browser** nur Hinweis (Pfad manuell). Siehe `TAURI-DESKTOP.md`.

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

Browser: `http://localhost:5173` — Pfad eintragen, **Öffnen**; **Ordner …** zeigt Hinweis (kein Systemdialog im Browser). Ordnerbaum / Vorschau über Proxy `/api`.

**Desktop (Tauri, nativer Ordnerdialog):** API wie oben starten, dann im Ordner `StatikManager.Web`: `npm run tauri:dev` (siehe `TAURI-DESKTOP.md`).

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
| GET     | `/api/browse`        | Query optional `path`: relativer Unterpfad zum Root (`/` in URLs). Leer oder fehlend = Root. Antwort `{ "entries": [ { "name", "relativePath", "isDirectory", "sizeBytes", "modifiedUtc" } ] }`. Ordner zuerst, dann Dateien, jeweils alphabetisch. Kein Root gesetzt / Pfad außerhalb des Roots / keine Ordner → **400** `{ "error": "..." }`. |
| GET     | `/api/file/meta`     | Query **`path`** (Pflicht): relative Pfadangabe zur Datei. Gleiche Sicherheitsregeln wie Browse. Antwort `{ "relativePath", "name", "kind" }` mit `kind` ∈ `pdf`, `image`, `html`, `json`, `text`, `other`, plus `sizeBytes`, `modifiedUtc`, `mimeType`. Ordner / nicht vorhanden / außerhalb Root → **400**. |
| GET     | `/api/preview/stream`| Query **`path`** (Pflicht): gleiche Regeln wie `/api/file/meta`. Stream der Datei mit passendem `Content-Type` (u. a. UTF-8 bei Text/HTML/JSON). Nur Dateien. Unterstützt Range-Anfragen (`enableRangeProcessing`) für PDF. Fehler **400** `{ "error": "..." }`. |

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
