# StatikManager V1 – Architektur

## Ueberblick
WPF-Desktopanwendung (.NET Framework 4.8, x64) zur Dokumentenverwaltung fuer Statik-Projekte.
Modularer Aufbau ueber IModul/ModulManager. Zentraler Zustand ueber AppZustand-Singleton.

---

## Projektstruktur

```
C:\KI\StatikManager_V1\
  src\StatikManager\
    Core\               – Kern-Infrastruktur
    Modules\            – UI-Module (Dokumente, Werkzeuge, Einstellungen)
    Infrastructure\     – Services (Logger, PDF, FileWatcher)
    Themes\             – WPF ResourceDictionaries
  original\             – Unveraenderter Quellstand (NIEMALS aendern)
  docs\                 – Wissensdatenbank
  agents\               – Legacy Agent-Definitionen
  .claude\agents\       – Claude Code Sub-Agenten
```

---

## Core-Schicht

### AppZustand (Singleton)
Zentraler Anwendungszustand. Alle Komponenten greifen darauf zu.

```csharp
AppZustand.Instanz.SetzeProjekt(pfad)      // Aktives Projekt setzen
AppZustand.Instanz.SetzeStatus(msg)        // Statusleiste aktualisieren
AppZustand.Instanz.RenderSem               // SemaphoreSlim fuer pdfium-Zugriff
AppZustand.Instanz.LadeZustandGeändert     // Event: Ladeanimation an/aus
```

**Kritisch:** Alle pdfium-Zugriffe (Docnet.Core) muessen ueber `RenderSem` serialisiert werden.

### Einstellungen (XML-Singleton)
Persistenz in `%APPDATA%\StatikManager\einstellungen.xml`.

```csharp
Einstellungen.Instanz.ProjektEintraege    // List<ProjektEintrag> mit Pfad/Kurzname/Sichtbar
Einstellungen.Instanz.DokumentAnsicht     // Baum oder Liste
Einstellungen.Instanz.WordVorlagen        // List<WordVorlage> fuer PDF->Word Export
Einstellungen.Instanz.Speichern()         // Schreibt XML
```

### SitzungsZustand
Speichert letztes Projekt + aktive Datei fuer naechsten Start.

### IModul / ModulManager
Plugin-Interface. Aktuell: DokumenteModul (und Einstellungen-Modul).

```csharp
interface IModul {
    string Name { get; }
    FrameworkElement ErstelleUI();
    void SitzungSpeichern(SitzungsZustand);
    void SitzungWiederherstellen(SitzungsZustand);
}
```

---

## Module

### DokumenteModul / DokumentePanel
Hauptmodul. Enthaelt:
- **Linke Seite**: Projektleiste (ComboBox + ⚙), Kopfzeile (Baum/Liste), Filter, DokumentenBaum/DateiListe
- **Rechte Seite**: Vorschau (HtmlToolbar, AbdeckungsPanel, WordInfoPanel, WebBrowser, PdfSchnittEditor)

**Vorschau-Routing** (DocumentRoutingService):
```
.pdf  → VorschauTyp.SchnittEditor → PdfSchnittEditor
.docx → VorschauTyp.WordVorschau  → WordInfoPanel (Bild-Rendering)
.html → VorschauTyp.Browser       → WebBrowser + HtmlToolbar
.jpg  → VorschauTyp.Browser       → WebBrowser (kein Toolbar)
.json → VorschauTyp.JsonVorschau  → WebBrowser (formatierter Text)
sonst → VorschauTyp.KeinVorschau  → Hinweisseite
```

### PdfSchnittEditor (Werkzeuge)
Rendert PDF-Seiten via pdfium, zeigt Crop-Linien. Unabhaengiger UserControl.

---

## Infrastructure

### OrdnerWatcherService
Ueberwacht Projektordner auf strukturelle Aenderungen.
- `NotifyFilter = FileName | DirectoryName` (kein LastWrite!)
- 500ms Debounce via DispatcherTimer
- Feuert `OrdnerGeaendert` → `AktualisiereNurStruktur()` (ohne Vorschau zu unterbrechen)

### FileWatcherService
Ueberwacht einzelne aktive Datei (z.B. fuer Auto-Reload).

### DocumentRoutingService / DateiTypen
Zentrale Dateityp-Erkennung und Routing. Keine Logik in DokumentePanel.

### PdfRenderer / PdfCache
pdfium-basiertes Rendering. Cache pro Projekt (Hash-Unterordner in AppData).

### WordInteropService / WordPdfService
Word COM-Automatisierung auf STA-Threads.

---

## Datenfluss: Dokument laden

```
Benutzer klickt Datei im Baum
  → DokumentenBaum_SelectedItemChanged
  → StarteSelektionDebounce(pfad)         [200ms]
  → LadeVorschau(pfad)
  → DocumentRoutingService.ErmittleVorschauTyp(pfad)
  → ZeigeSchnittEditor() / ZeigeWordInfo() / ZeigeBrowser()
  → Vorschau-Control laden
  → GibUI()                               [entsperrt UI]
```

---

## Datenfluss: Projektverwaltung

```
Einstellungen.xml
  → Einstellungen.Instanz.ProjektEintraege
  → ProjektVerwaltungDialog (Checkboxen, Kurznamen, Reihenfolge)
  → CbProjekte (nur Sichtbar=true)
  → CbProjekte_SelectionChanged
  → LadeProjektPfad(pfad)
  → OrdnerWatcher starten + AktualisiereDokumentListe()
```

---

## Verbindung zu PP_ZoomRahmen

PP_ZoomRahmen (AxisVM-Plugin) erzeugt folgende Struktur:
```
Projektordner/
  Pos_01_Decke_EG/
    position.html    ← Berechnungsausgabe
    position.json    ← Berechnungsdaten
    daten/           ← Eingabedaten
  Pos_02_Traeger/
    ...
```

Der StatikManager:
- Zeigt diese Ordnerstruktur im Baum an
- Klick auf Pos_XX/ oeffnet automatisch position.html
- HTML kann als PDF exportiert werden (Edge Headless)
- StatikManager aendert NIEMALS Dateien in Positionsordnern
