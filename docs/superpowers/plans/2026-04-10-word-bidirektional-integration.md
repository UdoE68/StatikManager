# Word-Bidirektional-Integration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wenn ein Nutzer eine `.docx`-Datei in Word bearbeitet und speichert, aktualisiert sich die Vorschau im StatikManager automatisch — ohne Scroll-Verlust, mit sauberem Fehler-Feedback und einmaligem Retry.

**Architecture:** Neuer `WordAutoRefreshService` kapselt die Konvertierungs-Orchestrierung (Cache löschen → STA-Thread → WortDateiZuPdf → Events). `DokumentePanel` leitet Word-Datei-Änderungen an diesen Service weiter und reagiert auf dessen Events. Scroll-Position wird vor/nach Refresh gesichert. Kein neues FileWatching — der bestehende `FileWatcherService` bleibt der einzige Watcher.

**Tech Stack:** C# .NET Framework 4.8, WPF, Microsoft.Office.Interop.Word, docnet.Core (pdfium), MSBuild

---

## Datei-Übersicht

| Datei | Aktion | Verantwortung |
|---|---|---|
| `Infrastructure/WordAutoRefreshService.cs` | **NEU** | Konvertierungs-Orchestrierung, Retry, Events |
| `Modules/Dokumente/DokumentePanel.xaml.cs` | **ÄNDERN** | Service verdrahten, OnDateiGeändert routen, Scroll-Position, neue Event-Handler |

---

## Task 1: `WordAutoRefreshService` erstellen

**Files:**
- Create: `src/StatikManager/Infrastructure/WordAutoRefreshService.cs`

### Kontext

Der Service übernimmt den Word-spezifischen Teil aus `DokumentePanel.OnDateiGeändert`. Er:
- Empfängt `DateiGeändertGemeldet()` vom Panel (UI-Thread)
- Löscht den PDF-Cache für die überwachte Datei
- Startet `.docx` → PDF Konvertierung auf STA-Hintergrundthread via `WordInteropService.WortDateiZuPdf`
- Feuert Events zurück auf den UI-Thread (via `Dispatcher.BeginInvoke`)
- Retry einmal nach 3s bei Fehler (via `DispatcherTimer`)
- Cancelt laufende Konvertierung bei `Starte()` mit neuer Datei oder `Stoppe()`

### Wichtige Abhängigkeiten
- `PdfCache.LöscheCacheFürDatei(pfad, cacheDir)` — Cache invalidieren
- `PdfCache.GetBasisPdfPfad(pfad, cacheDir)` — Ziel-PDF-Pfad ermitteln
- `WordInteropService.WortDateiZuPdf(pfad, zielPdf)` — synchron, STA-Thread required
- `AppZustand.Instanz.SetzeStatus(text, StatusLevel.Info/.Warn)` — Statuszeile
- `Logger.Info/Warn/Fehler(kategorie, nachricht)` — Logging

- [ ] **Schritt 1: Datei anlegen**

Erstelle `src/StatikManager/Infrastructure/WordAutoRefreshService.cs` mit folgendem vollständigen Inhalt:

```csharp
using System;
using System.IO;
using System.Threading;
using System.Windows.Threading;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Orchestriert den Auto-Refresh der Word-Vorschau nach einer Dateiänderung.
    /// Empfängt DateiGeändertGemeldet() vom Panel, löscht den Cache, konvertiert
    /// .docx → PDF auf einem STA-Hintergrundthread und feuert Events zurück auf den UI-Thread.
    /// Bei Fehler: einmaliger Retry nach 3 Sekunden.
    /// </summary>
    internal sealed class WordAutoRefreshService : IDisposable
    {
        private readonly Dispatcher _dispatcher;

        private string? _pfad;
        private string  _cacheDir = "";

        private CancellationTokenSource? _cts;
        private DispatcherTimer?         _retryTimer;
        private bool                     _retryAusstehend;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung startet.</summary>
        public event Action? KonvertierungGestartet;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung erfolgreich war.
        /// Parameter: vollständiger Pfad zum erzeugten Basis-PDF.</summary>
        public event Action<string>? VorschauBereit;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung fehlgeschlagen ist
        /// (nach Retry). Parameter: Fehlermeldung für die Statuszeile.</summary>
        public event Action<string>? KonvertierungFehlgeschlagen;

        public WordAutoRefreshService(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        /// <summary>
        /// Setzt die zu überwachende Datei und das Cache-Verzeichnis.
        /// Bricht eine laufende Konvertierung ab.
        /// </summary>
        public void Starte(string docxPfad, string cacheDir)
        {
            Stoppe();
            _pfad     = docxPfad;
            _cacheDir = cacheDir;
            _retryAusstehend = false;
        }

        /// <summary>
        /// Bricht laufende Konvertierung und ausstehenden Retry ab.
        /// </summary>
        public void Stoppe()
        {
            _cts?.Cancel();
            _cts = null;
            _retryTimer?.Stop();
            _retryTimer = null;
            _retryAusstehend = false;
            _pfad     = null;
        }

        /// <summary>
        /// Wird vom Panel aufgerufen wenn FileWatcherService eine Änderung meldet.
        /// Muss auf dem UI-Thread aufgerufen werden.
        /// </summary>
        public void DateiGeändertGemeldet()
        {
            if (_pfad == null) return;
            _retryTimer?.Stop();
            _retryTimer = null;
            _retryAusstehend = false;
            StarteKonvertierung(_pfad, _cacheDir, istRetry: false);
        }

        public void Dispose() => Stoppe();

        // ── Private ──────────────────────────────────────────────────────────

        private void StarteKonvertierung(string pfad, string cacheDir, bool istRetry)
        {
            // Laufenden Thread abbrechen
            _cts?.Cancel();
            var cts = new CancellationTokenSource();
            _cts = cts;
            var token = cts.Token;

            // Cache löschen damit WortDateiZuPdf neu konvertiert
            PdfCache.LöscheCacheFürDatei(pfad, cacheDir);

            var effektiverCacheDir = string.IsNullOrEmpty(cacheDir)
                ? Path.Combine(Path.GetTempPath(), "StatikManager", "preview")
                : cacheDir;

            var basisPdf = PdfCache.GetBasisPdfPfad(pfad, effektiverCacheDir);
            Directory.CreateDirectory(Path.GetDirectoryName(basisPdf)!);

            if (!istRetry)
            {
                AppZustand.Instanz.SetzeStatus(
                    "Vorschau wird aktualisiert: " + Path.GetFileName(pfad) + " …");
            }
            else
            {
                AppZustand.Instanz.SetzeStatus(
                    "Erneuter Versuch: " + Path.GetFileName(pfad) + " …");
            }

            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (token.IsCancellationRequested) return;
                KonvertierungGestartet?.Invoke();
            }));

            Logger.Info("WordAutoRefresh",
                $"{(istRetry ? "[Retry] " : "")}Starte Konvertierung: {Path.GetFileName(pfad)}");

            var t = new Thread(() =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    WordInteropService.WortDateiZuPdf(pfad, basisPdf);

                    if (token.IsCancellationRequested) return;

                    Logger.Info("WordAutoRefresh",
                        $"Konvertierung erfolgreich: {Path.GetFileName(pfad)}");

                    var basisPdfFinal = basisPdf;
                    _dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested) return;
                        AppZustand.Instanz.SetzeStatus(
                            "Vorschau aktualisiert: " + Path.GetFileName(pfad));
                        VorschauBereit?.Invoke(basisPdfFinal);
                    }));
                }
                catch (Exception ex) when (!(ex is OperationCanceledException))
                {
                    Logger.Warn("WordAutoRefresh",
                        $"Konvertierung fehlgeschlagen ({Path.GetFileName(pfad)}): {ex.Message}");

                    _dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested) return;

                        if (!istRetry)
                        {
                            // Erster Fehler → Retry nach 3s
                            _retryAusstehend = true;
                            AppZustand.Instanz.SetzeStatus(
                                "Konvertierung fehlgeschlagen – erneuter Versuch in 3 s …",
                                StatusLevel.Warn);
                            _retryTimer = new DispatcherTimer
                                { Interval = TimeSpan.FromSeconds(3) };
                            _retryTimer.Tick += (_, _) =>
                            {
                                _retryTimer?.Stop();
                                _retryTimer = null;
                                if (_retryAusstehend && _pfad == pfad)
                                    StarteKonvertierung(pfad, cacheDir, istRetry: true);
                            };
                            _retryTimer.Start();
                        }
                        else
                        {
                            // Retry auch fehlgeschlagen → endgültig
                            _retryAusstehend = false;
                            var fehler = "Konvertierung fehlgeschlagen – Vorschau veraltet ("
                                         + Path.GetFileName(pfad) + ")";
                            AppZustand.Instanz.SetzeStatus(fehler, StatusLevel.Warn);
                            KonvertierungFehlgeschlagen?.Invoke(fehler);
                        }
                    }));
                }
            })
            { IsBackground = true, Name = "WordAutoRefresh" };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }
    }
}
```

- [ ] **Schritt 2: Build ausführen**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' `
  /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
```

Erwartetes Ergebnis: `Build succeeded. 0 Error(s)`

Falls Fehler wegen fehlendem `using`: prüfen ob alle Namespaces korrekt sind (alle Abhängigkeiten liegen in `StatikManager.Infrastructure` bzw. `StatikManager.Core`).

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Infrastructure/WordAutoRefreshService.cs
git commit -m "feat(Infrastructure): WordAutoRefreshService — Auto-Refresh-Orchestrierung für Word-Dateien"
```

---

## Task 2: `DokumentePanel` verdrahten — Service initialisieren

**Files:**
- Modify: `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs`

### Kontext

In `DokumentePanel.xaml.cs` muss:
1. Ein `_wordAutoRefresh`-Feld angelegt werden (Zeile ~56, neben `_fileWatcher`)
2. Im Konstruktor: Service instanziieren und Events subscriben (Zeile ~125)
3. `Dispose`/Cleanup: `_wordAutoRefresh.Stoppe()` bei Projektwechsel

- [ ] **Schritt 1: Feld hinzufügen**

In `DokumentePanel.xaml.cs`, nach Zeile 57 (`private readonly OrdnerWatcherService _ordnerWatcher;`):

```csharp
        private readonly WordAutoRefreshService _wordAutoRefresh;
```

Der Block sieht dann so aus:
```csharp
        private readonly FileWatcherService   _fileWatcher;
        private readonly OrdnerWatcherService _ordnerWatcher;
        private readonly WordAutoRefreshService _wordAutoRefresh;
```

- [ ] **Schritt 2: Konstruktor — Service instanziieren und Events subscriben**

In `DokumentePanel.xaml.cs`, nach Zeile 126 (`_fileWatcher.DateiGeändert += OnDateiGeändert;`):

```csharp
            _wordAutoRefresh = new WordAutoRefreshService(Dispatcher);
            _wordAutoRefresh.KonvertierungGestartet      += OnWordKonvertierungGestartet;
            _wordAutoRefresh.VorschauBereit              += OnWordVorschauBereit;
            _wordAutoRefresh.KonvertierungFehlgeschlagen += OnWordKonvertierungFehler;
```

Der Block sieht dann so aus:
```csharp
            _fileWatcher = new FileWatcherService(Dispatcher);
            _fileWatcher.DateiGeändert += OnDateiGeändert;

            _wordAutoRefresh = new WordAutoRefreshService(Dispatcher);
            _wordAutoRefresh.KonvertierungGestartet      += OnWordKonvertierungGestartet;
            _wordAutoRefresh.VorschauBereit              += OnWordVorschauBereit;
            _wordAutoRefresh.KonvertierungFehlgeschlagen += OnWordKonvertierungFehler;

            _ordnerWatcher = new OrdnerWatcherService(Dispatcher);
            _ordnerWatcher.OrdnerGeändert += AktualisiereNurStruktur;
```

- [ ] **Schritt 3: Build ausführen**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' `
  /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
```

Erwartetes Ergebnis: `Build succeeded. 0 Error(s)`

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs
git commit -m "feat(DokumentePanel): WordAutoRefreshService verdrahten — Feld + Konstruktor"
```

---

## Task 3: `LadeVorschau` — Service bei Dateiauswahl starten

**Files:**
- Modify: `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs` (Zeile 765-769)

### Kontext

Wenn eine Word-Datei ausgewählt wird, muss `_wordAutoRefresh.Starte(pfad, _cacheDir)` aufgerufen werden, damit der Service weiß welche Datei er überwacht. Das passiert im `WordVorschau`-Case von `LadeVorschau`.

- [ ] **Schritt 1: `LadeVorschau` — Word-Case erweitern**

In `DokumentePanel.xaml.cs`, Zeile 765-769. Den bestehenden Code:

```csharp
                    case VorschauTyp.WordVorschau:
                        ZeigeWordInfo(pfad);
                        AppZustand.Instanz.SetzeStatus("Word: " + Path.GetFileName(pfad));
                        // GibUI() kommt aus den Dispatcher-Callbacks in StarteWordVorschauRendern
                        break;
```

ersetzen durch:

```csharp
                    case VorschauTyp.WordVorschau:
                        _wordAutoRefresh.Starte(pfad, _cacheDir);
                        ZeigeWordInfo(pfad);
                        AppZustand.Instanz.SetzeStatus("Word: " + Path.GetFileName(pfad));
                        // GibUI() kommt aus den Dispatcher-Callbacks in StarteWordVorschauRendern
                        break;
```

- [ ] **Schritt 2: Service stoppen wenn andere Datei gewählt wird**

In `LadeVorschau`, im `SchnittEditor`-Case (Zeile 755-763), nach `_wordVorschauCts?.Cancel()`:

```csharp
                    case VorschauTyp.SchnittEditor:
                        _wordZoomCts?.Cancel();
                        _wordVorschauCts?.Cancel();
                        _wordAutoRefresh.Stoppe();           // NEU
                        ZeigeSchnittEditor();
```

Und im `Browser`-Case (Zeile 770), direkt nach `case VorschauTyp.Browser:`:

```csharp
                    case VorschauTyp.Browser:
                        _wordAutoRefresh.Stoppe();           // NEU
                        ZeigeBrowser(mitAbdeckung: false);
```

Und im `JsonVorschau`-Case (Zeile 792):

```csharp
                    case VorschauTyp.JsonVorschau:
                        _wordAutoRefresh.Stoppe();           // NEU
                        ZeigeBrowser(mitAbdeckung: false);
```

Und im `KeinVorschau`-Case (Zeile 797):

```csharp
                    case VorschauTyp.KeinVorschau:
                        _wordAutoRefresh.Stoppe();           // NEU
                        ZeigeBrowser(mitAbdeckung: false);
```

- [ ] **Schritt 3: Build ausführen**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' `
  /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
```

Erwartetes Ergebnis: `Build succeeded. 0 Error(s)`

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs
git commit -m "feat(DokumentePanel): WordAutoRefreshService bei Dateiauswahl starten/stoppen"
```

---

## Task 4: `OnDateiGeändert` — Word-Routing umleiten

**Files:**
- Modify: `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs` (Zeile 1735-1759)

### Kontext

Bisher ruft `OnDateiGeändert` für Word-Dateien direkt `ZeigeWordInfo` auf — das setzt die gesamte Vorschau zurück (Scroll verloren, Bilder gelöscht). Neu: nur `_wordAutoRefresh.DateiGeändertGemeldet()` aufrufen; der Service übernimmt die Konvertierung, Panel reagiert via Events.

- [ ] **Schritt 1: Word-Block in `OnDateiGeändert` ersetzen**

Den bestehenden Code in `OnDateiGeändert` (Zeile 1746-1751):

```csharp
            // Word-Dateien nutzen WordInfoPanel + StarteWordVorschauRendern, nicht den Browser.
            if (DateiTypen.IstWordDatei(Path.GetExtension(_aktiverDateipfad)))
            {
                Logger.Info("AutoRefresh", "Word-Datei → ZeigeWordInfo");
                ZeigeWordInfo(_aktiverDateipfad);
                return;
            }
```

ersetzen durch:

```csharp
            // Word-Dateien: Konvertierung via WordAutoRefreshService (kein Reset der Vorschau)
            if (DateiTypen.IstWordDatei(Path.GetExtension(_aktiverDateipfad)))
            {
                Logger.Info("AutoRefresh", "Word-Datei → WordAutoRefreshService");
                _wordAutoRefresh.DateiGeändertGemeldet();
                return;
            }
```

Außerdem: Die Zeile `PdfCache.LöscheCacheFürDatei` (Zeile 1742) bleibt für Nicht-Word-Dateien. Für Word übernimmt der Service das Cache-Löschen intern. Also nur den Word-Block ersetzen, nicht die allgemeine Cache-Löschung.

- [ ] **Schritt 2: Build ausführen**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' `
  /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
```

Erwartetes Ergebnis: `Build succeeded. 0 Error(s)`

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs
git commit -m "feat(DokumentePanel): OnDateiGeändert leitet Word-Änderungen an WordAutoRefreshService weiter"
```

---

## Task 5: Event-Handler implementieren — Scroll-Position + sanfter Refresh

**Files:**
- Modify: `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs`

### Kontext

Die drei neuen Event-Handler reagieren auf den Service. `OnWordVorschauBereit` ist der kritischste: Er soll die Scroll-Position vor dem Neu-Rendern sichern und danach wiederherstellen, damit der Nutzer nicht nach oben springt. Das Seiten-Rendering läuft auf einem STA-Thread (pdfium + `AppZustand.RenderSem`).

`WordScrollViewer` (ScrollViewer, x:Name in XAML) hat `VerticalOffset` (get) und `ScrollToVerticalOffset(double)` (set).

- [ ] **Schritt 1: Drei neue private Methoden ans Ende des Panels einfügen**

In `DokumentePanel.xaml.cs`, **nach** `OnDateiGeändert()` (nach Zeile 1759), die drei folgenden Methoden einfügen:

```csharp
        // ── WordAutoRefreshService-Events ─────────────────────────────────────

        private void OnWordKonvertierungGestartet()
        {
            // Auf UI-Thread (vom Service per BeginInvoke geliefert)
            TxtWordLadeStatus.Text = "Vorschau wird aktualisiert …";
        }

        private void OnWordVorschauBereit(string basisPdfPfad)
        {
            // Auf UI-Thread. Scroll-Position merken, Seiten neu rendern, Position wiederherstellen.
            if (_aktiverDateipfad == null) return;
            if (WordInfoPanel.Visibility != Visibility.Visible) return;

            var scrollPos = WordScrollViewer.VerticalOffset;

            _wordVorschauCts?.Cancel();
            var cts = new CancellationTokenSource();
            _wordVorschauCts = cts;
            var token   = cts.Token;
            var pfad    = _aktiverDateipfad;
            int myGen   = _ladeGeneration;

            Logger.Info("WordAutoRefresh",
                $"VorschauBereit: {Path.GetFileName(pfad)}, ScrollPos={scrollPos:F0}");

            var t = new Thread(() =>
            {
                try { AppZustand.RenderSem.Wait(token); }
                catch (OperationCanceledException) { return; }
                try
                {
                    if (token.IsCancellationRequested) return;

                    var bilder = new List<System.Windows.Media.Imaging.BitmapSource>();
                    var lib = Docnet.Core.DocLib.Instance;
                    using var docReader = lib.GetDocReader(
                        basisPdfPfad,
                        new Docnet.Core.Models.PageDimensions(WordRenderBreite, WordRenderBreite * 2));
                    int n = docReader.GetPageCount();

                    for (int i = 0; i < n; i++)
                    {
                        if (token.IsCancellationRequested) return;
                        try
                        {
                            using var pageReader = docReader.GetPageReader(i);
                            var raw = pageReader.GetImage();
                            int w = pageReader.GetPageWidth(), h = pageReader.GetPageHeight();
                            if (raw == null || w <= 0 || h <= 0 || raw.Length < w * h * 4) continue;

                            PdfRenderer.KompositioniereGegenWeiss(raw, w, h);
                            var bmp = System.Windows.Media.Imaging.BitmapSource.Create(
                                w, h, 96, 96,
                                System.Windows.Media.PixelFormats.Bgra32, null, raw, w * 4);
                            bmp.Freeze();
                            bilder.Add(bmp);
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn("WordAutoRefresh", $"Seite {i + 1} übersprungen: {ex.Message}");
                        }
                    }

                    if (token.IsCancellationRequested) return;

                    var bilderFinal = bilder;
                    var scrollPosFinal = scrollPos;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                        _wordBasisPdf     = basisPdfPfad;
                        _wordSeitenBilder = bilderFinal;
                        BaueWordSeitenPanel();
                        TxtWordLadeStatus.Text = bilderFinal.Count > 0
                            ? $"{bilderFinal.Count} Seite(n) – aktualisiert"
                            : "Vorschau nicht verfügbar";
                        // Scroll-Position wiederherstellen nach Layout-Aktualisierung
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            WordScrollViewer.ScrollToVerticalOffset(scrollPosFinal);
                        }), System.Windows.Threading.DispatcherPriority.Loaded);
                    }));
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    App.LogFehler("OnWordVorschauBereit", App.GetExceptionKette(ex));
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                        TxtWordLadeStatus.Text = "⚠ Vorschau-Aktualisierung fehlgeschlagen";
                    }));
                }
                finally
                {
                    AppZustand.RenderSem.Release();
                }
            })
            { IsBackground = true, Name = "WordAutoRefreshRender" };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        private void OnWordKonvertierungFehler(string fehler)
        {
            // Auf UI-Thread. Altes Vorschaubild bleibt — nur Status aktualisieren.
            TxtWordLadeStatus.Text = "⚠ Vorschau veraltet";
            Logger.Warn("WordAutoRefresh", $"Fehler gemeldet: {fehler}");
            // AppZustand.SetzeStatus wurde bereits vom Service gesetzt
        }
```

- [ ] **Schritt 2: Build ausführen**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' `
  /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
```

Erwartetes Ergebnis: `Build succeeded. 0 Error(s)`

Falls Build-Fehler bei `Docnet.Core.DocLib` oder `PageDimensions`: prüfe die bestehenden `using`-Statements am Dateianfang von `DokumentePanel.xaml.cs` — alle nötigen Namespaces sind dort bereits vorhanden (sie werden von `StarteWordVorschauRendern` verwendet).

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs
git commit -m "feat(DokumentePanel): Event-Handler für WordAutoRefreshService — sanfter Refresh mit Scroll-Erhalt"
```

---

## Task 6: Manuelle Verifikation + App starten

**Files:** keine Code-Änderungen

### Testplan (in dieser Reihenfolge ausführen)

- [ ] **Schritt 1: Sicherstellen dass StatikManager nicht läuft**

```powershell
taskkill /f /im StatikManager.exe 2>nul; echo "Bereit"
```

- [ ] **Schritt 2: App starten**

```
C:\KI\StatikManager_V2\src\StatikManager\Start_Debug.bat
```

- [ ] **Schritt 3: Test A — Grundlegender Auto-Refresh**

1. Eine `.docx`-Datei im Dateibaum auswählen → Vorschau lädt (mehrere Seiten sichtbar)
2. Schaltfläche „In Word öffnen" klicken → Word öffnet die Datei
3. Im geöffneten Word-Dokument eine kleine Textänderung machen (z.B. Leerzeile einfügen)
4. `Strg+S` in Word drücken
5. **Erwartetes Verhalten:**
   - Nach ca. 2-3 Sekunden: Statuszeile zeigt „Vorschau wird aktualisiert: [Dateiname] …"
   - `TxtWordLadeStatus` zeigt „Vorschau wird aktualisiert …"
   - Nach weiteren 5-30 Sekunden (je nach Word-Geschwindigkeit): Vorschau aktualisiert sich
   - Statuszeile: „Vorschau aktualisiert: [Dateiname]"
   - `TxtWordLadeStatus`: „N Seite(n) – aktualisiert"

- [ ] **Schritt 4: Test B — Scroll-Position bleibt erhalten**

1. Eine mehrseitige `.docx` auswählen (mind. 3 Seiten)
2. In der Vorschau nach unten scrollen (Seite 2 oder 3 sichtbar machen)
3. In Word: Änderung machen → `Strg+S`
4. **Erwartetes Verhalten:** Nach Refresh zeigt die Vorschau noch annähernd dieselbe Scroll-Position (nicht oben)

- [ ] **Schritt 5: Test C — Datei wechseln während Konvertierung**

1. `.docx` auswählen → Word öffnen → speichern (Konvertierung startet)
2. Sofort eine andere Datei im Baum anklicken
3. **Erwartetes Verhalten:** Keine Fehler, keine veraltete Vorschau des alten Dokuments erscheint

- [ ] **Schritt 6: Commit nach erfolgreichem Test**

```bash
git add -A
git commit -m "test: Word-Bidirektional-Integration manuell verifiziert — alle Tests bestanden"
git push
```

---

## Self-Review

**Spec-Abdeckung:**
- ✅ Shell-Execute für `InWordÖffnen` — unverändert (Task 3 stoppt Service für andere Dateitypen)
- ✅ FileWatcher 2s Debounce — unverändert, `OnDateiGeändert` leitet jetzt weiter
- ✅ Retry einmal nach 3s — in `WordAutoRefreshService.StarteKonvertierung` implementiert
- ✅ Statuszeile bei Fehler — `AppZustand.Instanz.SetzeStatus(..., StatusLevel.Warn)`
- ✅ Altes Bild bleibt bei Fehler — `OnWordKonvertierungFehler` macht kein Panel-Clear
- ✅ Scroll-Position — in `OnWordVorschauBereit` gesichert/wiederhergestellt
- ✅ `Stoppe()` bei Projektwechsel — Task 3 fügt `Stoppe()` in alle Nicht-Word-Cases ein
- ✅ Cancellation bei Datei-Wechsel — `Starte()` ruft intern `Stoppe()` auf

**Typ-Konsistenz:**
- `WordAutoRefreshService.Starte(string, string)` → wird in Task 2+3 so aufgerufen ✅
- `WordAutoRefreshService.DateiGeändertGemeldet()` → wird in Task 4 so aufgerufen ✅
- Events `KonvertierungGestartet`, `VorschauBereit(string)`, `KonvertierungFehlgeschlagen(string)` → korrekt in Task 2 subscribed und Task 5 implementiert ✅
- `WordScrollViewer` (XAML x:Name) → existiert im XAML ✅
- `AppZustand.RenderSem.Wait(token)` / `.Release()` — gleiche Pattern wie in `StarteWordVorschauRendern` ✅
