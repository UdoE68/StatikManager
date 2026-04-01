# Phase 5 – DokumenteController: Bewusst nicht umgesetzt

**Datum:** 2026-03-31
**Entscheidung:** Refactoring wird nach Phase 4 beendet
**Status:** Abgeschlossen (durch Nicht-Umsetzung)

---

## Ausgangslage

Nach Phasen 1–4 wurde geprüft, ob ein `DokumenteController` sinnvoll eingeführt werden kann,
der folgende Aufgaben übernimmt:

- Reaktion auf Dateiänderungen
- Koordination zwischen FileWatcherService, DocumentRoutingService, WordPdfService, PdfCoverService

---

## Analyseergebnis

### Was extrahierbar wäre (technisch)

| Methode | Anmerkung |
|---|---|
| `StartVorkonvertierung` | Kein UI-Zugriff, nur Thread-Management — extrahierbar |
| `ErstelleUndÖffneWordKlon` | Kein Control-Zugriff, nur Dispatcher-Status — **aber toter Code (nie aufgerufen)** |

### Was nicht extrahierbar ist (ohne Callbacks oder MVVM)

| Methode | Blockierender Zugriff |
|---|---|
| `OnDateiGeändert` | `WordVorschau.Navigate("about:blank")` — XAML-Control direkt |
| `LadeVorschau` + Routing | `ZeigeSchnittEditor()`, `PdfEditor.LadePdf()`, `WordVorschau.Navigate()` |
| `BestimmePdfPfad` | `ChkKopf.IsChecked`, `TxtKopfMm.Text` — UI-Controls direkt |
| `StarteNeuladungImHintergrund` | Dispatcher → `WordVorschau.Navigate(new Uri(...))` |
| `NeuesWordDokument` | Dispatcher → `AktualisiereDokumentListe()` — Panel-Methode |
| Alle Zoom-Methoden | `WordZoomTransform.ScaleX/Y`, `WordScrollViewer.ActualWidth/Height` |

### Zusatzfund: Toter Code

`ErstelleUndÖffneWordKlon(string pdfPfad)` ist im Panel definiert, wird aber **nirgendwo aufgerufen** — weder aus XAML, noch aus anderen Methoden, noch aus anderen Klassen.

---

## Entscheidung des Orchestrators

> „Wir wählen Option B. Refactoring wird hier bewusst beendet."

**Begründung:**
- Die aktuelle Struktur ist bereits sauber getrennt
- Services werden direkt und verständlich genutzt
- Ein Controller ohne echte Entkopplung würde nur Komplexität hinzufügen
- Der verbleibende Panel-Code ist genuine UI-Koordinationslogik und gehört dort

---

## Voraussetzung für echte Controller-Schicht (Phase 5+)

Eine vollständige Entkopplung von UI und Logik erfordert:

1. **MVVM-Pattern:** `DokumentePanelViewModel` mit `INotifyPropertyChanged`
2. **Commands:** `ICommand`-Implementierungen für Datei-Aktionen
3. **Observable State:** Properties statt direkter Control-Zugriffe
4. **Messaging/Events:** Reaktion auf Zustandsänderungen ohne direkte Dispatcher-Aufrufe

Das ist eine eigene Architekturentscheidung, kein Refactoring-Schritt.

---

## Empfehlung für nächste Sitzung

Vor weiteren strukturellen Änderungen:
1. `ErstelleUndÖffneWordKlon` entfernen (toter Code)
2. CS8602-Warnung in DokumentePanel.xaml.cs ~Zeile 105 beheben
3. Dann: Architekturentscheidung MVVM ja/nein treffen

---

## Gesamtbilanz Refactoring Phasen 1–4

| Phase | Ergebnis |
|---|---|
| 1 | `PdfRenderer`, `PdfCache`, `DateiTypen`, `WordInteropService` extrahiert |
| 2 | `FileWatcherService`, `DocumentRoutingService` eingeführt |
| 3 | `PdfCoverService` extrahiert (Option A) |
| 4 | `WordPdfService` mit `ÖffneInWord` + `BerechneNeuenZielPdf` |
| 5 | Bewusst nicht umgesetzt — Architekturgrenze erreicht |

**Build nach Phase 4:** 0 Fehler, 4 bekannte Warnungen (unverändert seit Phase 0).
