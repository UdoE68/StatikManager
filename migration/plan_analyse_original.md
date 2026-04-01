# Planner-Bericht: Analyse /original — StatikManager

**Datum:** 2026-03-31
**Agent:** Planner
**Status:** Analyse abgeschlossen — noch keine Änderungen vorgenommen

---

## 1. Projektüberblick

**StatikManager** ist eine WPF-Desktopanwendung für die Verwaltung und Vorschau von
Statik-Dokumenten (PDFs, Word, Excel, Bilder) in Ingenieurbüros.

| Eigenschaft     | Wert                                         |
|-----------------|----------------------------------------------|
| Framework       | .NET Framework 4.8 (net48)                  |
| UI-Technologie  | WPF (XAML + Code-behind)                    |
| Plattform       | Windows x64                                  |
| Sprache         | C# (Nullable enabled, ImplicitUsings)        |
| Namespace       | `StatikManager`                              |
| Einstiegspunkt  | `App.xaml` / `App.xaml.cs`                  |

---

## 2. Abhängigkeiten (NuGet + COM)

| Typ        | Name                              | Version / Hinweis               |
|------------|-----------------------------------|---------------------------------|
| NuGet      | PdfSharp                          | 1.50.5147                       |
| NuGet      | Docnet.Core                       | 2.4.0 (libpdfium, x64 nötig)   |
| COM-Interop| Microsoft.Office.Interop.Word     | Word 2013+ vorausgesetzt        |
| Framework  | System.Windows.Forms              | Nur für FolderBrowserDialog     |

---

## 3. Architekturanalyse

Das Projekt ist sauber in drei Schichten aufgeteilt:

### 3.1 Core-Schicht (`Core/`)

| Datei              | Klasse / Interface | Funktion                                                           |
|--------------------|--------------------|--------------------------------------------------------------------|
| `IModul.cs`        | `IModul`           | Schnittstelle für alle Module (Id, Name, Panel, Menü, Toolbar, Bereinigen) |
| `ModulManager.cs`  | `ModulManager`     | Modul-Registry (Liste, FindeModul, AllesBereinigen)               |
| `AppZustand.cs`    | `AppZustand`       | Singleton-Appzustand mit Events: `StatusGeändert`, `ProjektGeändert` |
| `Einstellungen.cs` | `Einstellungen`    | XML-persistente Einstellungen in `%APPDATA%\StatikManager\einstellungen.xml` |
| `OrdnerDialog.cs`  | `OrdnerDialog`     | Moderner Windows-Ordner-Dialog via COM `IFileOpenDialog`, Fallback auf WinForms |

**Bewertung:** Sehr sauber. Keine zirkulären Abhängigkeiten. Events statt direkte Referenzen.

### 3.2 Shell-Schicht (Root)

| Datei              | Klasse         | Funktion                                                             |
|--------------------|----------------|----------------------------------------------------------------------|
| `App.xaml.cs`      | `App`          | Globales Error-Handling (AppDomain, Dispatcher, TaskScheduler), Logging nach `StatikManager_Errors.log`, Browser-Emulation Registry |
| `MainWindow.xaml.cs` | `MainWindow` | Fenster-Shell, `ModulManager` instanziiert und integriert Module in Menü + Toolbar + ContentArea |

**Wichtig:** `MainWindow` hat eine direkte Abhängigkeit auf `DokumenteModul` (Zeile 51:
`_modulManager.Registrieren(new DokumenteModul())`). Diese Kopplung ist die einzige
nicht-generische Stelle in der Shell.

### 3.3 Modul-Schicht (`Modules/`)

#### Modul: Dokumente (`Modules/Dokumente/`)

| Datei                    | Klasse                | Funktion                                      |
|--------------------------|-----------------------|-----------------------------------------------|
| `DokumenteModul.cs`      | `DokumenteModul`      | `IModul`-Implementierung, erstellt Panel, Menüeintrag, Toolbar-Button |
| `DokumentePanel.xaml.cs` | `DokumentePanel`      | Hauptpanel: Dokumentenbaum/Liste, PDF-Vorschau, Datei-Filter, Baumtiefe, Word-Export, Cache-Verwaltung, FileSystemWatcher |
| `ProjektLadenDialog.xaml.cs` | `ProjektLadenDialog` | Einfacher Dialog: Standardpfad oder manueller Ordner |

**DokumentePanel ist die komplexeste Klasse** (geschätzt 600–900 Zeilen). Enthält:
- Dokumentenliste als Baum oder flache Liste (umschaltbar)
- Dateitypfilter (Alle, Word, Excel, PDF, Bilder)
- Baumtiefensteuerung (1–3 Ebenen, Alle)
- PDF-Vorschau mit Caching (`%APPDATA%\StatikManager\pdf-cache\<hash>\`, Version `v5`)
- PDF → Word-Klon (via Word COM, STA-Thread im Hintergrund)
- FileSystemWatcher für automatische Aktualisierung
- Word-Vorschau mit Zoom, Render-Cache, CancellationToken-Verwaltung

#### Modul: Werkzeuge (`Modules/Werkzeuge/`)

| Datei                       | Klasse             | Funktion                                          |
|-----------------------------|--------------------|---------------------------------------------------|
| `PdfSchnittEditor.xaml.cs`  | `PdfSchnittEditor` | PDF-Vorschau + verschiebbarer Beschnittrahmen, Zoom (Strg+Rad), Seiten-Rendering, Crop-Export nach Word |
| `PdfZuWordDialog.xaml.cs`   | `PdfZuWordDialog`  | Eigenständiger Dialog: PDF → Word-Konvertierung mit Fortschrittsanzeige und Abbruch-Funktion |

**Hinweis:** `PdfSchnittEditor` ist ein `UserControl` (kein eigenes Modul), vermutlich
eingebettet im `DokumentePanel` (als `PdfEditor`-Referenz in Zeile 124 sichtbar).
`PdfZuWordDialog` ist ein eigenständiges `Window`.

---

## 4. Klassifizierung aller Quelldateien

### Kategorie: Kern-Infrastruktur (agents → Planner/Orchestrator Kontext)
```
Core/IModul.cs              → Modulschnittstelle (Erweiterungspunkt)
Core/ModulManager.cs        → Modulregistrierung
Core/AppZustand.cs          → Gemeinsamer Zustand + Events
Core/Einstellungen.cs       → Persistenz
Core/OrdnerDialog.cs        → Systemdialog-Wrapper
App.xaml / App.xaml.cs      → Einstiegspunkt + Error-Handling
MainWindow.xaml / .cs       → Shell
```

### Kategorie: Dokument-Modul (Kernfunktionalität)
```
Modules/Dokumente/DokumenteModul.cs         → Modul-Registrierung
Modules/Dokumente/DokumentePanel.xaml/.cs   → Hauptpanel (komplex)
Modules/Dokumente/ProjektLadenDialog.xaml/.cs → Hilfsdialog
```

### Kategorie: Werkzeuge-Modul
```
Modules/Werkzeuge/PdfSchnittEditor.xaml/.cs → PDF-Editor UserControl
Modules/Werkzeuge/PdfZuWordDialog.xaml/.cs  → Konvertierungsdialog
```

### Kategorie: Nicht migrieren (Build-Artefakte)
```
bin/           → Kompilate
obj/           → MSBuild-Zwischendateien
.vs/           → Visual Studio Workspace-Daten
*.bak*, *.bakN → Historische Sicherungskopien (bis bak14)
```

### Kategorie: Projektdatei
```
StatikManager.csproj   → Build-Konfiguration, Abhängigkeiten
```

### Kategorie: Claude-Einstellungen (nicht migrieren)
```
.claude/settings.json       → Enthält Permission-Allowlist (teilweise für BildschnittAddin)
.claude/settings.local.json → Lokale Ergänzungen
```

---

## 5. Festgestellte Besonderheiten und Risiken

### R1 — Viele .bak-Dateien (kein Risiko, aber Ballast)
Die Quelldateien haben bis zu 14 Backup-Versionen. Diese enthalten keine
eigenständige Logik und können ignoriert werden. Kein Migrationsbedarf.

### R2 — DokumentePanel.xaml.cs: Sehr groß
Geschätzte Dateigröße: 600–900 Zeilen. Enthält mehrere klar trennbare
Verantwortlichkeiten. Bei einer Refactoring-Migration besteht Risiko der
versehentlichen Auslassung von Logik. → Nolen muss hier besonders genau prüfen.

### R3 — PdfSchnittEditor ist kein eigenes Modul
`PdfSchnittEditor` ist ein `UserControl`, kein `IModul`. Es wird vermutlich direkt
in `DokumentePanel.xaml` eingebettet (Referenz `PdfEditor.ExportierenNachWord()`
in `DokumentePanel.xaml.cs:124`). Dies bedeutet: kein separater
`WerkzeugeModul.cs` existiert bisher — `PdfZuWordDialog` ist ein eigenständiges
Window, das direkt aufgerufen wird.

### R4 — COM-Interop auf STA-Thread
Alle Word-COM-Operationen laufen auf einem eigenen STA-Thread. Das ist korrekt
für COM-Interop. Bei Umstrukturierungen muss diese Thread-Anforderung erhalten bleiben.

### R5 — .claude/settings.json enthält projektfremde Einträge
Die Datei enthält Berechtigungen für `BildschnittAddin` — ein anderes Projekt.
Diese Datei sollte nicht in die neue Struktur migriert werden, sondern neu erstellt werden.

### R6 — Cache-Versionierung
`CacheVersion = "v5"` in `DokumentePanel.xaml.cs`. Bei Änderungen an der
Render-Logik muss die Version erhöht werden, um Stale-Cache-Probleme zu vermeiden.

---

## 6. Vorgeschlagene Zielstruktur

```
StatikManager_V1/
├── CLAUDE.md                          ← Projektregeln (vorhanden)
├── agents/                            ← Multi-Agenten-Definitionen (vorhanden)
│   ├── orchestrator.md
│   ├── planner.md
│   ├── fiona.md
│   └── nolen.md
├── skills/                            ← Wiederverwendbare Fähigkeitsbeschreibungen
├── prompts/                           ← Prompt-Vorlagen
├── memory/                            ← Persistentes Kontextwissen
├── migration/
│   ├── plan_analyse_original.md       ← Dieser Bericht
│   └── plan_migration_schritte.md     ← (noch zu erstellen)
├── original/                          ← UNVERÄNDERTER Quellstand (readonly)
│   └── StatikManager/                 ← Visual Studio Projektordner
└── src/                               ← ZIELSTRUKTUR für neuen/migrierten Code
    └── StatikManager/
        ├── StatikManager.csproj
        ├── App.xaml / App.xaml.cs
        ├── MainWindow.xaml / .cs
        ├── Core/
        │   ├── IModul.cs
        │   ├── ModulManager.cs
        │   ├── AppZustand.cs
        │   ├── Einstellungen.cs
        │   └── OrdnerDialog.cs
        └── Modules/
            ├── Dokumente/
            │   ├── DokumenteModul.cs
            │   ├── DokumentePanel.xaml / .cs
            │   └── ProjektLadenDialog.xaml / .cs
            └── Werkzeuge/
                ├── PdfSchnittEditor.xaml / .cs
                └── PdfZuWordDialog.xaml / .cs
```

**Nicht migrieren:** `bin/`, `obj/`, `.vs/`, `*.bak*`, `.claude/settings.json`

---

## 7. Vorgeschlagene Migrationsphasen

| Phase | Beschreibung                                    | Agent  | Risiko  |
|-------|-------------------------------------------------|--------|---------|
| 1     | `/src`-Struktur anlegen, csproj kopieren        | Fiona  | Niedrig |
| 2     | Core-Dateien migrieren (5 Dateien)              | Fiona  | Niedrig |
| 3     | Shell migrieren (App.xaml, MainWindow)          | Fiona  | Niedrig |
| 4     | Dokumente-Modul migrieren                       | Fiona  | Mittel  |
| 5     | Werkzeuge-Modul migrieren                       | Fiona  | Mittel  |
| 6     | Build-Test (nur Kompilierbarkeit prüfen)        | Nolen  | —       |
| 7     | Funktionsprüfung gegen /original                | Nolen  | —       |

---

## 8. Offene Fragen an den Benutzer

1. **Refactoring oder 1:1-Kopie?**
   Soll die Migration eine saubere 1:1-Kopie sein (kein Refactoring),
   oder sollen dabei Verbesserungen vorgenommen werden (z. B. `DokumentePanel`
   aufteilen, `PdfSchnittEditor` als eigenes Modul)?

2. **Zielordner `/src` oder anders benennen?**
   Die neue Arbeitskopie könnte auch `src/`, `app/` oder direkt als
   `StatikManager/` (neben `/original`) angelegt werden.

3. **`.claude/settings.json` neu erstellen?**
   Die bestehende Datei enthält Fremdprojekt-Einträge. Soll sie neu und
   sauber angelegt werden?

---

**Planner-Fazit:** Das Projekt ist architektonisch sauber und gut strukturiert.
Die Migration ist beherrschbar. Größtes Risiko ist `DokumentePanel.xaml.cs` (Größe).
Bereit für Phase 1, sobald die offenen Fragen geklärt sind.
