# StatikManager Web

Separates Mini-Projekt: ASP.NET Core Minimal API + Vite/TypeScript. Kein Bezug zum WPF-StatikManager im übergeordneten Ordner.

**Stand:** Milestone 5 + Ordnerdialog (`POST /api/session/pick-root`, nur Windows-Desktop).

## Voraussetzungen

- [.NET 8 SDK](https://dotnet.microsoft.com/download) (Windows-TFM `net8.0-windows` für die API)
- [Node.js](https://nodejs.org/) (LTS, für npm)

Die API nutzt `FolderBrowserDialog` (WinForms) — **Ordner wählen** funktioniert nur, wenn das Backend auf **Windows** läuft und ein interaktiver Desktop verfügbar ist (kein headless-Szenario).

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

Browser: `http://localhost:5173` — **Ordner wählen** (Dialog), Projekt öffnen oder Pfad eintragen, Ordnerliste und „Nach oben“ (Proxy `/api`).

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
| POST    | `/api/session/pick-root` | **Nur Windows (Desktop):** öffnet einen Ordnerdialog (STA). Antwort `{ "rootPath": "<absolut>" \| null }` — `null` bei Abbruch. Fehler **400** bei Dialog-Fehler. Ändert das Session-Root **nicht** direkt; das Frontend ruft anschließend `POST /api/session/root` auf. |
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
