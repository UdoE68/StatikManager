# StatikManager V1

Dokumentenverwaltung und PDF/Word-Vorschau für Statik-Projekte.

## Beschreibung

StatikManager ist eine WPF-Desktopanwendung zur Verwaltung und Vorschau von Projektdokumenten. Sie richtet sich an Ingenieurbüros, die strukturiert auf PDF- und Word-Dateien zugreifen und diese vor dem Export gezielt beschneiden möchten.

## Features

### PDF-Vorschau mit Beschnittrahmen
- Seiten werden hochauflösend gerendert und im integrierten Viewer angezeigt
- Beschnittrahmen (Kopf, Fuß, links, rechts) lassen sich manuell per Drag-and-Drop setzen oder automatisch anhand des Seiteninhalts erkennen
- Randwerte und Auto-Erkennung pro Seite (unterstützt gemischte Hoch-/Querformate)
- Sicherheitsabstand für die automatische Rand-Erkennung einstellbar

### Word-Vorschau
- Word-Dokumente (.docx) werden seitenweise gerendert und im integrierten Panel angezeigt
- Zoom, Seitenbreite und Mehrseiten-Übersicht direkt in der Anwendung

### Mehrseitenansicht (Nebeneinander)
- PDFs können vertikal (übereinander) oder horizontal (nebeneinander) dargestellt werden
- Crop-Linien werden pro Seite unabhängig gesetzt
- Horizontales Scrollen mit dem Mausrad im Nebeneinander-Modus

### Smooth Zoom
- Zoom per Strg + Mausrad oder Toolbar-Schaltflächen
- Animiertes Ease-Out (ca. 110 ms) für ruckelfreies Zoomen
- Der Punkt unter dem Mauszeiger bleibt beim Zoomen stabil
- "Ganzes Fenster" passt eine Seite vollständig in den sichtbaren Bereich ein

### Stabile Lade-Logik
- Laufende Render-Aufträge werden bei schnellem Dokumentwechsel sauber abgebrochen
- Semaphore verhindert parallelen pdfium-Zugriff
- UI wird während des Ladens gesperrt und nach Abschluss automatisch freigegeben

### Dokumentenverwaltung
- Baumansicht und Listenansicht der Projektdateien
- Filter nach Dateityp, einstellbare Baumtiefe
- Mehrfachauswahl und Löschen mit Bestätigungsdialog

## Voraussetzungen

- Windows 10 / 11 (x64)
- .NET Framework 4.8
- Microsoft Word (für Word-Vorschau und Export)

## Starten

1. Repository klonen oder ZIP entpacken
2. Projekt in Visual Studio 2022 öffnen (`src/StatikManager/StatikManager.csproj`)
3. Konfiguration: `Debug | x64`
4. Build starten (`F6` oder `Strg+Umschalt+B`)
5. Anwendung starten (`F5`)

Alternativ direkt die fertige EXE aufrufen:

```
bin\x64\Debug\net48\StatikManager.exe
```

## Projektstruktur

```
src/StatikManager/
  Core/               Zustandsverwaltung, AppZustand, Logging
  Infrastructure/     PdfRenderer (Docnet.Core / pdfium)
  Modules/
    Dokumente/        Dokumentenliste, Vorschau-Panel
    Werkzeuge/        PdfSchnittEditor, Word-Vorschau
original/             Unveränderter Ausgangsstand (Referenz)
backup/               Gesicherte Zwischenstände vor größeren Änderungen
```

## Technologie

| Komponente | Technologie |
|------------|-------------|
| UI-Framework | WPF / .NET Framework 4.8 |
| PDF-Rendering | Docnet.Core 2.6.0 (pdfium) |
| PDF-Verarbeitung | PdfSharp |
| Word-Integration | Microsoft.Office.Interop.Word |
| Sprache | C# |
