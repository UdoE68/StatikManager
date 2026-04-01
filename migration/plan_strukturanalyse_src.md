# Planner-Bericht: Strukturanalyse /src/StatikManager

**Datum:** 2026-03-31
**Agent:** Planner
**Grundlage:** Vollständige Lektüre aller Quelldateien in /src/StatikManager
**Status:** Nur Analyse — keine Änderungen vorgenommen

---

## 1. Zusammenfassung

Das Projekt hat eine gute Makrostruktur (Core / Shell / Modules), leidet aber auf
Dateiebene unter starker Verantwortungsüberlastung — besonders in `DokumentePanel.xaml.cs`.
Mehrere Querschnittsaufgaben (PDF-Rendering, COM-Interop, Cache-Verwaltung) sind
redundant über mehrere Klassen verteilt.

---

## 2. Ist-Zustand: Verantwortlichkeiten pro Datei

### 2.1 Core/ — stabil, keine Probleme

| Datei               | Verantwortung                       | Bewertung      |
|---------------------|-------------------------------------|----------------|
| `IModul.cs`         | Modul-Interface                     | klar, minimal  |
| `ModulManager.cs`   | Modul-Registry                      | klar, minimal  |
| `AppZustand.cs`     | Shared State + Events               | klar, Singleton|
| `Einstellungen.cs`  | Persistenz (XML)                    | klar, Singleton|
| `OrdnerDialog.cs`   | OS-Ordnerdialog-Wrapper             | klar, statisch |

**Befund:** Core ist sauber. Kein Handlungsbedarf.

---

### 2.2 Shell — stabil, eine Abhängigkeit

| Datei              | Verantwortung                        | Befund                              |
|--------------------|--------------------------------------|-------------------------------------|
| `App.xaml.cs`      | Startup, globales Error-Handling, Logging | klar getrennt                  |
| `MainWindow.xaml.cs` | Fenster-Shell, Modul-Integration   | eine direkte Typkopplung (s. u.)    |

**Befund:** `MainWindow.cs:51` instanziiert `new DokumenteModul()` direkt —
eine konkrete Abhängigkeit auf den Modultyp statt auf `IModul`.
Das ist die einzige Bruchstelle in der sonst generischen Shell.

---

### 2.3 DokumentePanel.xaml.cs — KERNPROBLEM

**Geschätzte Zeilenzahl: ~900 Zeilen**

Diese Datei trägt mindestens **8 verschiedene Verantwortlichkeiten**:

| # | Verantwortung                          | Zeilen (ca.) | Kategorie      |
|---|----------------------------------------|--------------|----------------|
| 1 | Dokumentenliste: Baum/Liste anzeigen   | ~120         | UI             |
| 2 | Dateitypfilter + Baumtiefe             | ~60          | UI-Logik       |
| 3 | Datei-Auswahl + Vorschau-Routing       | ~80          | UI-Logik       |
| 4 | PDF-Vorschau (delegiert an PdfEditor)  | ~30          | UI             |
| 5 | Word-Vorschau: Rendern + Zoom + Panel  | ~180         | Preview-Logik  |
| 6 | Word → Basis-PDF (Hintergrundkonv.)    | ~70          | Interop        |
| 7 | PDF → Word-Klon-Export                 | ~80          | Interop        |
| 8 | FileSystemWatcher + Cache-Verwaltung   | ~80          | Infrastruktur  |

**Zusätzlich** enthält sie Hilfsmethoden:
- `PasstZuFilter`, `DateiIcon`, `IstWordDatei`, `IstPdfDatei`, `IstBildDatei`
- `GetWordKlonPfad`, `GetBasisPdfPfad`, `CacheGültig`
- `ZähleDateien`, `EnthältDateien`

---

### 2.4 PdfSchnittEditor.xaml.cs — mittel, isolierbar

**Geschätzte Zeilenzahl: ~700 Zeilen**

Verantwortlichkeiten:

| # | Verantwortung                          | Kategorie        |
|---|----------------------------------------|------------------|
| 1 | PDF-Rendering via Docnet               | Rendering        |
| 2 | Seiten-Layout-Berechnung               | Rendering        |
| 3 | Canvas-Aufbau + Seitenanzeige          | UI               |
| 4 | Crop-Linien: Zeichnen + Hit-Zonen      | UI               |
| 5 | Crop-Linien: Maus-Drag-Interaktion     | UI-Interaktion   |
| 6 | Automatische Rand-Erkennung (Pixel)    | Bildverarbeitung |
| 7 | PDF-Crop + Word-Export via COM         | Interop          |
| 8 | Zoom (Strg+Mausrad)                    | UI               |

PdfSchnittEditor ist ein `UserControl` — **kein eigenes `IModul`**.
Es wird direkt in `DokumentePanel.xaml` eingebettet (`PdfEditor`-Instanz).

---

### 2.5 PdfZuWordDialog.xaml.cs — klar, eigenständig

Eigener Dialog-Typ (`Window`). Enthält ausschließlich:
- Dateiauswahl-UI
- Konvertierungslogik (PDF → DOCX via Word COM, STA-Thread)
- Fortschrittsanzeige + Abbruch

**Befund:** gut gekapselt. Keine strukturellen Probleme.

---

## 3. Identifizierte Probleme

### P1 — DokumentePanel ist ein Monolith (KRITISCH)
~900 Zeilen, 8 Verantwortlichkeiten. Schwer testbar, schwer erweiterbar.
Jede Änderung an Rendering, Interop oder Cache berührt dieselbe Datei.

### P2 — Alpha-Kompositing doppelt implementiert (MITTEL)
Die Umrechnung transparenter PDF-Pixel gegen Weiß ist in
**zwei Klassen** implementiert:
- `PdfSchnittEditor.cs` → `KompositioniereGegenWeiss()` (optimierte Version mit float)
- `DokumentePanel.cs` → inline in `StarteWordVorschauRendern()` (einfachere Version)

Gleiche Logik, unterschiedliche Implementierungen.

### P3 — Word-COM-Interop an drei Stellen (MITTEL)
Word-COM-Operationen verteilen sich auf:
- `DokumentePanel.cs`: `ErstelleUndÖffneWordKlon()` + `VorkonvertierungTask()`
- `PdfSchnittEditor.cs`: `ExportThreadWorker()`
- `PdfZuWordDialog.cs`: `KonvertierenAsync()`

Kein gemeinsamer Wrapper. STA-Thread-Muster wiederholt sich dreifach.

### P4 — Cache-Logik in DokumentePanel eingebettet (GERING)
`CacheGültig()`, `GetBasisPdfPfad()`, `GetWordKlonPfad()`, `CacheVersion`
sind private Methoden/Konstanten in `DokumentePanel`. Nicht wiederverwendbar.

### P5 — Dateitype-Klassifikation verstreut (GERING)
`IstWordDatei()`, `IstPdfDatei()`, `IstBildDatei()`, `PasstZuFilter()`, `DateiIcon()`
sind in `DokumentePanel` definiert, werden aber implizit auch vom
`PdfSchnittEditor` über den Aufrufkontext genutzt.

### P6 — Kein WerkzeugeModul vorhanden (GERING)
`PdfSchnittEditor` und `PdfZuWordDialog` sind logisch dem Werkzeuge-Bereich
zuzuordnen, aber es gibt kein `WerkzeugeModul : IModul`. Das zweite Modul
existiert strukturell nicht — `PdfZuWordDialog` wird direkt aus
`DokumentePanel` aufgerufen (oder ist eigenständig navigierbar).

---

## 4. Vorgeschlagene Zielstruktur

> Basis: 1:1-Migration ist abgeschlossen. Diese Struktur gilt für den
> nachfolgenden Refactoring-Schritt — nicht für jetzt.

```
src/StatikManager/
│
├── Core/                                  (unverändert)
│   ├── IModul.cs
│   ├── ModulManager.cs
│   ├── AppZustand.cs
│   ├── Einstellungen.cs
│   └── OrdnerDialog.cs
│
├── Infrastructure/                        (NEU — Querschnittsaufgaben)
│   ├── PdfRenderer.cs                    → RenderiereAlleSeiten + KompositioniereGegenWeiss
│   ├── PdfCache.cs                       → CacheGültig, GetBasisPdfPfad, CacheVersion
│   ├── DateiTypen.cs                     → IstWordDatei, IstPdfDatei, PasstZuFilter, DateiIcon
│   └── WordInterop/
│       └── WordInteropService.cs         → STA-Thread-Wrapper für alle Word-COM-Ops
│
├── Modules/
│   ├── Dokumente/
│   │   ├── DokumenteModul.cs             (unverändert)
│   │   ├── DokumentePanel.xaml/.cs       → nur noch: Baum/Liste, Filter, Auswahl-Routing
│   │   ├── ProjektLadenDialog.xaml/.cs   (unverändert)
│   │   ├── Preview/
│   │   │   ├── WordVorschauPanel.xaml/.cs → Word-Vorschau + Zoom (aus DokumentePanel auslagern)
│   │   │   └── PreviewRouter.cs           → LadeVorschau-Logik
│   │   └── Export/
│   │       └── WordKlonService.cs         → ErstelleUndÖffneWordKlon (aus DokumentePanel)
│   │
│   └── Werkzeuge/
│       ├── WerkzeugeModul.cs             (NEU — IModul-Implementierung)
│       ├── PdfSchnittEditor.xaml/.cs     → nur noch: Canvas, Crop-UI, Zoom
│       ├── CropLogik.cs                  → ErkenneCropRänderVonBitmap, Rand-Berechnung
│       └── PdfSchnittExport.cs           → ExportThreadWorker (Word-COM-Export)
│
├── Themes/
│   └── ModernTheme.xaml                  (unverändert)
│
└── App.xaml / MainWindow.xaml            (unverändert)
```

---

## 5. Klassifizierung nach Zuständigkeitsbereich

| Zustandsbereich  | Inhalt                                                    |
|------------------|-----------------------------------------------------------|
| **UI**           | DokumentePanel (Baum/Liste), WordVorschauPanel, PdfSchnittEditor (Canvas+Crop-UI) |
| **Logik**        | PreviewRouter, CropLogik, DateiTypen                     |
| **Interop**      | WordInteropService, WordKlonService, PdfSchnittExport    |
| **Export**       | PdfSchnittExport, PdfZuWordDialog                        |
| **Preview**      | WordVorschauPanel, PdfSchnittEditor                      |
| **Infrastruktur**| PdfRenderer, PdfCache, OrdnerDialog (bereits vorhanden)  |

---

## 6. Priorisierung der Refactoring-Schritte

| Priorität | Maßnahme                                      | Aufwand | Nutzen  |
|-----------|-----------------------------------------------|---------|---------|
| 1 (hoch)  | `PdfRenderer.cs` extrahieren (Duplikat)       | Gering  | Hoch    |
| 2 (hoch)  | `WordInteropService.cs` zentralisieren        | Mittel  | Hoch    |
| 3 (mittel)| `WordVorschauPanel` aus DokumentePanel heraus | Mittel  | Mittel  |
| 4 (mittel)| `DateiTypen.cs` extrahieren                   | Gering  | Mittel  |
| 5 (mittel)| `PdfCache.cs` extrahieren                     | Gering  | Mittel  |
| 6 (niedrig)| `WerkzeugeModul.cs` anlegen                  | Gering  | Niedrig |
| 7 (niedrig)| `CropLogik.cs` aus PdfSchnittEditor auslagern| Mittel  | Niedrig |

---

## 7. Was nicht verändert werden sollte

- `Core/` — ist vorbildlich, kein Refactoring-Bedarf
- `App.xaml.cs` — globales Error-Handling ist vollständig und gut strukturiert
- `ProjektLadenDialog.xaml.cs` — minimaler, klarer Dialog
- `PdfZuWordDialog.xaml.cs` — gut gekapselt, eigenständig

---

**Planner-Fazit:** Das Fundament ist solide. Der größte Hebel liegt in der
Zerlegung von `DokumentePanel.xaml.cs` und der Zentralisierung der
Word-COM-Interop-Logik. Alles andere ist optionale Verbesserung.
Bereit zur Übergabe an Orchestrator.
