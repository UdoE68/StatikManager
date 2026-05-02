# StatikManager Web — Desktop-Hülle (Tauri)

## Eignung von Tauri

| Aspekt | Bewertung |
|--------|-----------|
| **Frontend** | Bereits Vite + TypeScript → direkt als Webview nutzbar. |
| **Backend** | ASP.NET Core API bleibt separater Prozess; Tauri lädt nur die UI (Dev: Proxy zu `/api`). |
| **Ordnerdialog** | `plugin-dialog` mit `open({ directory: true })` — nativer Windows-Dialog, kein Server-Thread. |
| **Aufwand** | Ein `src-tauri`-Ordner, Rust-Toolchain, zwei npm-Skripte. |

**Fazit:** Tauri ist hier sinnvoll integrierbar und genau für diese Aufgabe gedacht.

## Umsetzungsplan (kurz)

1. **Tauri-Crate** unter `StatikManager.Web/src-tauri/` mit `tauri-plugin-dialog`.
2. **Dev:** `beforeDevCommand` startet Vite (`5173`), Webview lädt `http://localhost:5173` — Vite leitet `/api` auf `http://localhost:5156` weiter (**StatikManager.Api** muss laufen).
3. **UI:** Button „Ordner …“ ruft nur unter Tauri den Dialog auf; sonst Hinweis (Browser-Fallback).
4. **API:** Unverändert `POST /api/session/root` nach Auswahl (wie „Öffnen“).
5. **ASP.NET:** Ordnerdialog `/api/session/pick-root` entfernt — keine WinForms-Abhängigkeit mehr.

## Build / Start

### Voraussetzungen

- **Node.js** (npm)
- **Rust** + **Cargo** ([rustup](https://rustup.rs/)) für Tauri
- **.NET 8** für die API

### 1) API starten (Terminal 1)

```powershell
cd "c:\KI\HTML Umbau\StatikManager_V2\statikmanager-web\src\StatikManager.Api"
dotnet run
```

Lauscht z. B. auf `http://localhost:5156` (siehe Konsolenausgabe / `launchSettings.json`).

### 2) StatikManager Web + Tauri (Terminal 2)

```powershell
cd "c:\KI\HTML Umbau\StatikManager_V2\statikmanager-web\src\StatikManager.Web"
npm run tauri:dev
```

- Startet Vite (Port **5173**, `strictPort`) und kompiliert die Tauri-Shell.
- Im Fenster: **Ordner …** öffnet den **Windows-Ordnerdialog** → Pfad ins Feld → **automatisch „Projekt öffnen“** (gleiche Logik wie Button „Öffnen“).

### 3) Nur Browser (ohne Tauri)

```powershell
cd "...\StatikManager.Web"
npm run dev
```

Button **Ordner …** zeigt den **Hinweis**; Pfad manuell eintragen und **Öffnen** wählen.

### 4) Releases bauen

- **Web-Asset + API (wwwroot):** in `StatikManager.Web`: `npm run build` (Ausgabe: `StatikManager.Api/wwwroot`).
- **API:** `dotnet build` / `dotnet publish` wie gewohnt.
- **Tauri-Installer:** in `StatikManager.Web` nach Icon-Setup z. B. `npm run tauri:build` (aktuell ist `bundle.active` in `tauri.conf.json` auf **false** gesetzt — für MSI/EXE mit Icons `bundle` aktivieren und unter `src-tauri/icons/` passende `32x32.png`, `128x128.png`, `icon.ico` ablegen, siehe [Tauri Bundle](https://v2.tauri.app/develop/configuration-files/)).

## Projektstruktur (relevant)

```
statikmanager-web/src/StatikManager.Web/
  package.json          # scripts: tauri:dev, tauri:build
  vite.config.ts        # proxy /api → localhost:5156, strictPort
  src/main.ts           # isTauri() + dialog.open / Browser-Hinweis
  src-tauri/            # Rust + tauri.conf.json + capabilities
```

## Hinweise

- **Zwei Prozesse:** Tauri-Fenster allein startet **nicht** die ASP.NET-API — diese vorher oder parallel starten.
- **CORS:** API erlaubt `localhost:5173` bereits für den Vite-Dev-Server.
