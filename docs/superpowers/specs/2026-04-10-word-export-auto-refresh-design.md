# Design: Auto-Refresh nach Word-Export

**Datum:** 2026-04-10
**Branch:** feature/word-export-next
**Status:** Freigegeben zur Implementierung

---

## Ziel

Wenn der Nutzer im PDF-Schnitt-Editor "Nach Word exportieren" wählt, soll StatikManager nach dem Export automatisch zur exportierten `.docx` wechseln und diese beobachten. Speichert der Nutzer in Word, aktualisiert sich die Vorschau automatisch — identisch zum bestehenden Word-Auto-Refresh-Verhalten.

---

## Ist-Analyse

Der bestehende `WordAutoRefreshService` (aus Word-Bidirektional-Integration) übernimmt Auto-Refresh vollständig — sobald eine `.docx` in `DokumentePanel` als aktive Datei geladen ist. Das Problem: `PdfSchnittEditor` und `DokumentePanel` sind getrennte Module ohne direkte Referenz. Nach dem Export fehlt der Übergang.

---

## Entscheidungen

| Frage | Entscheidung | Begründung |
|---|---|---|
| Kommunikation PdfSchnittEditor → DokumentePanel | Neues Event in `AppZustand` | Bestehende Muster (ProjektGeändert, LadeZustandGeändert), kein Overloading von ModulWechselAngefordert |
| Was passiert nach Export | DokumentePanel wechselt automatisch zur .docx | Nutzer hat gerade exportiert — Vorschau soll den neuen Stand zeigen |
| Wer öffnet Word | PdfSchnittEditor (wie bisher) | Keine Änderung am bestehenden Export-Dialog |

---

## Datenfluss

```
PdfSchnittEditor: ExportThreadWorker abgeschlossen
  → AppZustand.Instanz.MeldeWordExport(zielDocxPfad)
      → Event WordExportAbgeschlossen(zielDocxPfad) feuert (auf Dispatcher-Thread)
          → DokumentePanel.OnWordExportAbgeschlossen(docxPfad)
              → Selektiert Datei im Baum/Liste (visuelles Feedback)
              → LadeVorschau(docxPfad)          ← .docx wird aktive Datei
                  → _wordAutoRefresh.Starte()   ← bestehend, übernimmt automatisch

Word öffnet sich (wie bisher per Dialog in PdfSchnittEditor, unverändert)

Nutzer bearbeitet .docx in Word → Strg+S
  → FileWatcherService (2s Debounce, bestehend)
  → WordAutoRefreshService → Vorschau aktualisiert sich
```

---

## Neue Komponente: Event in `AppZustand`

**Datei:** `src/StatikManager/Core/AppZustand.cs`

```csharp
// Neues Event + Methode nach dem Muster der bestehenden Events:

/// <summary>Wird ausgelöst wenn ein PDF erfolgreich nach Word exportiert wurde.
/// Parameter: vollständiger Pfad der erzeugten .docx-Datei.</summary>
public event Action<string>? WordExportAbgeschlossen;

public void MeldeWordExport(string docxPfad)
    => WordExportAbgeschlossen?.Invoke(docxPfad);
```

---

## Änderung: `PdfSchnittEditor` — Export-Abschluss melden

**Datei:** `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`

Nach erfolgreichem Speichern der `.docx` in `ExportThreadWorker` (ca. Zeile 3193), direkt nach dem Schreiben der Datei und **vor** dem "Öffnen in Word?"-Dialog:

```csharp
// Dispatcher-Invoke nötig da ExportThreadWorker auf STA-Hintergrundthread läuft
Dispatcher.BeginInvoke(new Action(() =>
    AppZustand.Instanz.MeldeWordExport(zielPfad)));
```

**Bedingung:** Nur aufrufen wenn die `.docx` tatsächlich erfolgreich gespeichert wurde (kein Fehler davor).

---

## Änderung: `DokumentePanel` — Event subscriben + Datei laden

**Datei:** `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs`

### Konstruktor — Event subscriben

```csharp
AppZustand.Instanz.WordExportAbgeschlossen += OnWordExportAbgeschlossen;
```

### Neuer Event-Handler

```csharp
private void OnWordExportAbgeschlossen(string docxPfad)
{
    // Nur reagieren wenn die exportierte Datei im aktuellen Projektordner liegt
    if (_projektPfad == null) return;
    if (!docxPfad.StartsWith(_projektPfad, StringComparison.OrdinalIgnoreCase)) return;
    if (!File.Exists(docxPfad)) return;

    // Dateiliste aktualisieren (neue .docx erscheint im Baum)
    AktualisiereNurStruktur();

    // .docx als aktive Vorschau laden → WordAutoRefreshService startet automatisch
    LadeVorschau(docxPfad);

    AppZustand.Instanz.SetzeStatus("Word-Export geöffnet: " + Path.GetFileName(docxPfad));
}
```

---

## Was NICHT geändert wird

- `WordAutoRefreshService` — unverändert
- Export-Logik in `PdfSchnittEditor` (Rendering, Segmentierung, Vorlage) — unverändert
- "Öffnen in Word?"-Dialog nach Export — unverändert
- Alle anderen Vorschau-Typen — unverändert

---

## Fehlerbehandlung

| Szenario | Verhalten |
|---|---|
| Export schlägt fehl | `MeldeWordExport` wird nicht aufgerufen — DokumentePanel reagiert nicht |
| `.docx` liegt außerhalb Projektordner | Guard in `OnWordExportAbgeschlossen` verhindert Navigation |
| `.docx` existiert nicht (Race) | Guard `File.Exists` verhindert LadeVorschau-Aufruf |
| Kein Projekt geladen | Guard `_projektPfad == null` verhindert Navigation |

---

## Testkriterien

1. PDF im StatikManager auswählen → PDF-Editor öffnet
2. "Nach Word exportieren" → Export läuft, `.docx` wird erstellt
3. DokumentePanel wechselt automatisch zur `.docx`-Vorschau
4. Word öffnet sich (wie bisher)
5. In Word Änderung machen → Strg+S → Vorschau aktualisiert sich nach ~2–3s
6. Kein Auto-Refresh wenn exportierte Datei außerhalb des Projektordners liegt
