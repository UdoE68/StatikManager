# Design: Word-Bidirektional-Integration

**Datum:** 2026-04-10
**Branch:** feature/word-export-next
**Status:** Freigegeben zur Implementierung

---

## Ziel

Wenn der Nutzer eine `.docx`-Datei im StatikManager auswählt, in Word bearbeitet und speichert, soll die Vorschau im StatikManager automatisch aktualisiert werden — ohne Scroll-Verlust, mit klarem Fehler-Feedback.

---

## Ist-Analyse

Der Kern-Flow ist bereits vorhanden (`FileWatcherService` → `OnDateiGeändert` → `ZeigeWordInfo`), aber mit diesen Schwächen:

| Problem | Datei / Methode | Auswirkung |
|---|---|---|
| Kein Retry bei Konvertierungsfehler | `OnDateiGeändert` | Vorschau bleibt "nicht verfügbar" |
| Kein Statusbar-Update bei Fehler | `StarteWordVorschauRendern` | Nutzer sieht keinen Hinweis |
| Scroll-Position geht bei Refresh verloren | `ZeigeWordInfo` löscht Panel | Jeder Refresh springt nach oben |
| Konvertierungslogik direkt im Panel | `DokumentePanel.xaml.cs` (2438 Zeilen) | Schlecht trennbar, schwer testbar |

---

## Entscheidungen

| Frage | Entscheidung | Begründung |
|---|---|---|
| Word öffnen | Shell-Execute (unveränderter `InWordÖffnen`) | Robust, kein fehleranfälliges COM-Tracking |
| Änderungen erkennen | Bestehender `FileWatcherService` (2s Debounce) | Bereits vorhanden und bewährt |
| Fehler-Feedback | Statuszeile + altes Vorschaubild bleibt | Kein Popup, kein Datenverlust |
| Retry | Einmal nach 3 Sekunden | Word kann beim ersten Event noch schreiben |

---

## Neue Komponente: `WordAutoRefreshService`

**Datei:** `src/StatikManager/Infrastructure/WordAutoRefreshService.cs`

### Verantwortlichkeiten
- Empfängt `DateiGeändertGemeldet()` vom Panel
- Löscht den Cache für die betroffene Datei
- Startet Konvertierung `.docx` → PDF auf STA-Hintergrundthread
- Feuert Events an das Panel
- Setzt Statuszeile via `AppZustand.Instanz.SetzeStatus()`
- Einmaliger Retry nach 3s bei Fehler

### Schnittstelle

```csharp
internal sealed class WordAutoRefreshService : IDisposable
{
    // Zustand
    public bool KonvertierungAktiv { get; }

    // Initialisierung
    public void Starte(string docxPfad, string cacheDir);
    public void Stoppe();

    // Auslöser (wird von DokumentePanel.OnDateiGeändert aufgerufen)
    public void DateiGeändertGemeldet();

    // Ergebnisse
    public event Action<string>  VorschauBereit;           // basisPdfPfad
    public event Action          KonvertierungGestartet;
    public event Action<string>  KonvertierungFehlgeschlagen; // Fehlermeldung
}
```

### Zustandsmaschine

```
Idle
  │  DateiGeändertGemeldet()
  ▼
Konvertiert ──────────────────────────────────────────────────────►  VorschauBereit → Idle
  │
  │  Exception
  ▼
Fehler → StatusBar: "Konvertierung fehlgeschlagen – Vorschau veraltet"
  │
  │  DispatcherTimer 3s → Retry
  ▼
Konvertiert (Retry) ──────────────────────────────────────────────►  VorschauBereit → Idle
  │
  │  Exception (Retry)
  ▼
FehlerEndgültig → StatusBar: "Konvertierung fehlgeschlagen"
                → altes Vorschaubild bleibt sichtbar
```

### Implementierungsdetails
- STA-Thread via `new Thread(...) { ApartmentState = STA }`
- Cancellation via `CancellationTokenSource` bei `Stoppe()` und `Starte()`
- `File.Copy` → Temp-Kopie (verhindert Lock-Konflikt mit Word)
- Retry-Timer via `DispatcherTimer` auf UI-Thread (kein weiterer Thread)

---

## Änderungen in `DokumentePanel`

### 1. Neues Feld

```csharp
private readonly WordAutoRefreshService _wordAutoRefresh;
```

Initialisierung im Konstruktor:
```csharp
_wordAutoRefresh = new WordAutoRefreshService(Dispatcher);
_wordAutoRefresh.KonvertierungGestartet     += OnWordKonvertierungGestartet;
_wordAutoRefresh.VorschauBereit             += OnWordVorschauBereit;
_wordAutoRefresh.KonvertierungFehlgeschlagen += OnWordKonvertierungFehler;
```

### 2. `LadeVorschau` — Word-Fall

```csharp
case VorschauTyp.WordVorschau:
    _wordAutoRefresh.Starte(pfad, _cacheDir);   // NEU
    ZeigeWordInfo(pfad);
    // ...
```

### 3. `OnDateiGeändert` — Word-Routing

```csharp
// vorher:
if (DateiTypen.IstWordDatei(ext))
{
    ZeigeWordInfo(_aktiverDateipfad);
    return;
}

// nachher:
if (DateiTypen.IstWordDatei(ext))
{
    _wordAutoRefresh.DateiGeändertGemeldet();   // Service übernimmt
    return;
}
```

### 4. Neue Event-Handler im Panel

```csharp
private void OnWordKonvertierungGestartet()
    → TxtWordLadeStatus.Text = "Vorschau wird aktualisiert …"

private void OnWordVorschauBereit(string basisPdfPfad)
    → _wordBasisPdf = basisPdfPfad
    → Scroll-Position merken
    → Seiten neu rendern (STA-Thread, pdfium)
    → Scroll-Position wiederherstellen
    → TxtWordLadeStatus.Text = $"{seitenAnzahl} Seite(n) – aktualisiert"

private void OnWordKonvertierungFehler(string fehler)
    → AppZustand.Instanz.SetzeStatus(fehler, StatusLevel.Warn)
    → TxtWordLadeStatus.Text = "⚠ Vorschau veraltet"
    → kein Panel-Clear, altes Bild bleibt
```

### 5. `AktualisiereProjekt` / `Stoppe` bei Projektwechsel

```csharp
_wordAutoRefresh.Stoppe();
```

---

## Was NICHT geändert wird

- `FileWatcherService` — unverändert
- `WordInteropService` — unverändert
- `PdfCache` — unverändert
- `InWordÖffnen()` — unverändert (Shell-Execute bleibt)
- Alle anderen Vorschau-Typen (PDF, HTML, Bilder) — unverändert

---

## Fehlerbehandlung

| Szenario | Verhalten |
|---|---|
| Word hat Datei noch geöffnet beim File.Copy | Klappt (Word öffnet mit FILE_SHARE_READ) |
| WortDateiZuPdf wirft COMException | Retry nach 3s |
| Retry schlägt auch fehl | StatusBar Warn, altes Bild bleibt |
| Nutzer wechselt Datei während Konvertierung | `Stoppe()` cancelt laufenden Thread |
| Kein Projekt geladen (`_cacheDir` leer) | Fallback auf Temp-Verzeichnis (wie bisher) |

---

## Testkriterien

1. `.docx` in StatikManager auswählen → Vorschau lädt
2. „In Word öffnen" → Word öffnet die Datei
3. In Word Änderung machen → Ctrl+S → ca. 2-3s warten → Vorschau im StatikManager aktualisiert sich automatisch
4. Scroll-Position bleibt erhalten wenn man auf Seite 2+ gescrollt hat
5. Statuszeile zeigt während Konvertierung "Vorschau wird aktualisiert …"
6. Nach erfolgreichem Refresh: Statuszeile zeigt Seitenanzahl
7. Bei Fehler: Statuszeile zeigt Warn-Meldung, altes Vorschaubild bleibt sichtbar
