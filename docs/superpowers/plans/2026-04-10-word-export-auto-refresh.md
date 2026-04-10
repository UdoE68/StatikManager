# Word-Export Auto-Refresh Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Nach einem erfolgreichen Word-Export wechselt DokumentePanel automatisch zur exportierten `.docx` und beobachtet sie — sodass Word-Speicherungen die Vorschau automatisch aktualisieren.

**Architecture:** Neues Event `WordExportAbgeschlossen` in `AppZustand` entkoppelt `PdfSchnittEditor` von `DokumentePanel`. Nach dem `SaveAs2`-Aufruf feuert `PdfSchnittEditor` das Event. `DokumentePanel` reagiert mit `LadeVorschau(docxPfad)`, wodurch der bereits vorhandene `WordAutoRefreshService` automatisch die Datei beobachtet.

**Tech Stack:** C# .NET Framework 4.8, WPF, MSBuild (Debug | x64)

---

## Datei-Übersicht

| Datei | Aktion | Was |
|---|---|---|
| `src/StatikManager/Core/AppZustand.cs` | **ÄNDERN** | Event + Methode hinzufügen |
| `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` | **ÄNDERN** | `MeldeWordExport` nach SaveAs2 aufrufen |
| `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs` | **ÄNDERN** | Event subscriben + Handler implementieren |

---

## Task 1: `WordExportAbgeschlossen` Event in `AppZustand`

**Files:**
- Modify: `src/StatikManager/Core/AppZustand.cs`

### Kontext

`AppZustand.cs` ist die zentrale Kommunikationszentrale des Projekts. Alle Events folgen demselben Muster: `public event Action<T>?` + `public void Methode(T param) => Event?.Invoke(param)`. Die Datei ist ~82 Zeilen, übersichtlich.

- [ ] **Schritt 1: Event + Methode einfügen**

In `AppZustand.cs`, nach dem `ModulWechselAngefordert`-Block (nach Zeile 80), vor der schließenden `}`der Klasse:

```csharp
        // ── Word-Export ───────────────────────────────────────────────────────

        /// <summary>Wird ausgelöst wenn ein PDF erfolgreich nach Word exportiert wurde.
        /// Parameter: vollständiger Pfad der erzeugten .docx-Datei.</summary>
        public event Action<string>? WordExportAbgeschlossen;

        public void MeldeWordExport(string docxPfad)
            => WordExportAbgeschlossen?.Invoke(docxPfad);
```

Der finale Bereich der Klasse sieht dann so aus:
```csharp
        // ── Modul-Wechsel ─────────────────────────────────────────────────────

        public event Action<string, string>? ModulWechselAngefordert;  // (modulId, dateipfad)

        public void FordeModulWechsel(string modulId, string dateipfad)
            => ModulWechselAngefordert?.Invoke(modulId, dateipfad);

        // ── Word-Export ───────────────────────────────────────────────────────

        /// <summary>Wird ausgelöst wenn ein PDF erfolgreich nach Word exportiert wurde.
        /// Parameter: vollständiger Pfad der erzeugten .docx-Datei.</summary>
        public event Action<string>? WordExportAbgeschlossen;

        public void MeldeWordExport(string docxPfad)
            => WordExportAbgeschlossen?.Invoke(docxPfad);
    }
}
```

- [ ] **Schritt 2: Build ausführen**

```powershell
powershell.exe -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:quiet 2>&1 | Select-String -Pattern 'succeeded|FAILED|Error\b' | Select-Object -Last 5; Write-Host 'Exit:' \$LASTEXITCODE"
```

Erwartetes Ergebnis: `Exit: 0` (kein "FAILED", kein "Error")

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Core/AppZustand.cs
git commit -m "feat(AppZustand): WordExportAbgeschlossen Event für Post-Export-Navigation"
```

---

## Task 2: `PdfSchnittEditor` — Export-Abschluss melden

**Files:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` (Zeile ~3196)

### Kontext

In `ExportThreadWorker` wird nach `wordDoc.SaveAs2(zielPfad, ...)` (Zeile 3193) ein `Dispatcher.BeginInvoke`-Block ausgeführt (Zeile 3196–3209). Dieser Block:
1. Setzt `TxtInfo.Text` auf die Erfolgs-Meldung
2. Zeigt einen `MessageBox.Show`-Dialog "Jetzt öffnen?"

`AppZustand.Instanz.MeldeWordExport(zielPfad)` soll **innerhalb dieses bestehenden BeginInvoke-Blocks**, nach `TxtInfo.Text = ...` und **vor** `MessageBox.Show`, eingefügt werden. So lädt DokumentePanel die Datei bevor der Dialog blockiert.

Aktueller Code (Zeilen 3195–3209):
```csharp
                    string dateiName = IO.Path.GetFileName(zielPfad);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = $"Exportiert: {seiteNr} Seite(n) → {dateiName}";
                        if (MessageBox.Show(
                            $"Word-Dokument erstellt:\n{zielPfad}\n\nJetzt öffnen?",
                            "Export fertig", MessageBoxButton.YesNo,
                            MessageBoxImage.Information) == MessageBoxResult.Yes)
                        {
                            try { System.Diagnostics.Process.Start(
                                new System.Diagnostics.ProcessStartInfo(zielPfad)
                                { UseShellExecute = true }); }
                            catch { }
                        }
                    }));
```

- [ ] **Schritt 1: `MeldeWordExport` einfügen**

Ersetze den Block (Zeilen 3195–3209) durch:

```csharp
                    string dateiName = IO.Path.GetFileName(zielPfad);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = $"Exportiert: {seiteNr} Seite(n) → {dateiName}";
                        AppZustand.Instanz.MeldeWordExport(zielPfad);
                        if (MessageBox.Show(
                            $"Word-Dokument erstellt:\n{zielPfad}\n\nJetzt öffnen?",
                            "Export fertig", MessageBoxButton.YesNo,
                            MessageBoxImage.Information) == MessageBoxResult.Yes)
                        {
                            try { System.Diagnostics.Process.Start(
                                new System.Diagnostics.ProcessStartInfo(zielPfad)
                                { UseShellExecute = true }); }
                            catch { }
                        }
                    }));
```

Die einzige Änderung: `AppZustand.Instanz.MeldeWordExport(zielPfad);` nach `TxtInfo.Text = ...`.

- [ ] **Schritt 2: Build ausführen**

```powershell
powershell.exe -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:quiet 2>&1 | Select-String -Pattern 'succeeded|FAILED|Error\b' | Select-Object -Last 5; Write-Host 'Exit:' \$LASTEXITCODE"
```

Erwartetes Ergebnis: `Exit: 0`

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
git commit -m "feat(PdfSchnittEditor): MeldeWordExport nach erfolgreichem Export"
```

---

## Task 3: `DokumentePanel` — Event subscriben + Handler

**Files:**
- Modify: `src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs`

### Kontext

Im Konstruktor von `DokumentePanel` werden alle `AppZustand`-Events subscribed (ca. Zeile 132). `LadeVorschau(pfad)` ist bereits vorhanden und löst den kompletten Lade-Flow aus — inklusive `_wordAutoRefresh.Starte(pfad, _cacheDir)`. `AktualisiereNurStruktur()` aktualisiert Baum/Liste ohne Datei-Selektion zurückzusetzen.

- [ ] **Schritt 1: Subscription im Konstruktor hinzufügen**

Suche im Konstruktor die Zeile:
```csharp
            AppZustand.Instanz.LadeZustandGeändert += aktiv => { if (!aktiv) GibUI(); };
```

Füge danach ein:
```csharp
            AppZustand.Instanz.WordExportAbgeschlossen += OnWordExportAbgeschlossen;
```

- [ ] **Schritt 2: Event-Handler implementieren**

Füge nach der Methode `OnWordKonvertierungFehler` (am Ende des WordAutoRefresh-Bereichs, ca. Zeile 1880) ein:

```csharp
        private void OnWordExportAbgeschlossen(string docxPfad)
        {
            // Nur reagieren wenn exportierte Datei im aktiven Projektordner liegt
            if (_projektPfad == null) return;
            if (!docxPfad.StartsWith(_projektPfad, StringComparison.OrdinalIgnoreCase)) return;
            if (!File.Exists(docxPfad)) return;

            // Dateiliste aktualisieren damit neue .docx im Baum erscheint
            AktualisiereNurStruktur();

            // .docx als aktive Vorschau laden → WordAutoRefreshService startet automatisch
            LadeVorschau(docxPfad);

            AppZustand.Instanz.SetzeStatus("Word-Export geöffnet: " + Path.GetFileName(docxPfad));
        }
```

- [ ] **Schritt 3: Build ausführen**

```powershell
powershell.exe -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V2\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:quiet 2>&1 | Select-String -Pattern 'succeeded|FAILED|Error\b' | Select-Object -Last 5; Write-Host 'Exit:' \$LASTEXITCODE"
```

Erwartetes Ergebnis: `Exit: 0`

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Modules/Dokumente/DokumentePanel.xaml.cs
git commit -m "feat(DokumentePanel): OnWordExportAbgeschlossen — Auto-Navigation zur exportierten .docx"
git push
```

---

## Task 4: Manuelle Verifikation + App starten

- [ ] **Schritt 1: Alte Instanz beenden**

```powershell
powershell.exe -Command "taskkill /f /im StatikManager.exe 2>\$null; Write-Host 'Bereit'"
```

- [ ] **Schritt 2: App starten**

```
C:\KI\StatikManager_V2\src\StatikManager\Start_Debug.bat
```

- [ ] **Schritt 3: Test A — Auto-Navigation nach Export**

1. Eine PDF-Datei im Baum auswählen → PDF-Editor öffnet sich
2. Schnitt-Bereiche wie gewünscht setzen
3. "Nach Word exportieren" klicken → Export läuft
4. **Erwartetes Verhalten:**
   - DokumentePanel wechselt automatisch zur exportierten `.docx` (Vorschau erscheint)
   - Statuszeile: "Word-Export geöffnet: [dateiname].docx"
   - Dialog "Jetzt öffnen?" erscheint danach (wie bisher)

- [ ] **Schritt 4: Test B — Auto-Refresh nach Speichern in Word**

1. Im Dialog "Jetzt öffnen?" → Ja klicken → Word öffnet die `.docx`
2. In Word eine kleine Änderung vornehmen (z.B. Text hinzufügen)
3. `Strg+S` in Word drücken
4. **Erwartetes Verhalten:**
   - Nach ca. 2–3 Sekunden: Statuszeile "Vorschau wird aktualisiert: [dateiname].docx …"
   - Vorschau im StatikManager aktualisiert sich automatisch

- [ ] **Schritt 5: Test C — Kein Effekt außerhalb Projektordner**

1. Eine PDF öffnen die **außerhalb** des aktuellen Projektordners liegt (oder Projekt wechseln nach Export)
2. **Erwartetes Verhalten:** DokumentePanel wechselt NICHT zur .docx (Guard greift)

---

## Self-Review

**Spec-Abdeckung:**
- ✅ Neues Event `WordExportAbgeschlossen` in AppZustand → Task 1
- ✅ `PdfSchnittEditor` ruft `MeldeWordExport` nach SaveAs2 auf → Task 2
- ✅ `DokumentePanel` subscribed Event, ruft `LadeVorschau` auf → Task 3
- ✅ Guard: `_projektPfad == null` → Task 3 Handler
- ✅ Guard: `.docx` außerhalb Projektordner → Task 3 Handler
- ✅ Guard: `File.Exists` → Task 3 Handler
- ✅ `AktualisiereNurStruktur()` vor `LadeVorschau` → Task 3 Handler
- ✅ Manuelle Tests A/B/C → Task 4

**Typ-Konsistenz:**
- `AppZustand.MeldeWordExport(string docxPfad)` → in Task 2 so aufgerufen ✅
- `AppZustand.Instanz.WordExportAbgeschlossen += OnWordExportAbgeschlossen` → Handler-Signatur `private void OnWordExportAbgeschlossen(string docxPfad)` passt zu `Action<string>` ✅
- `LadeVorschau(docxPfad)` → existiert bereits in DokumentePanel ✅
- `AktualisiereNurStruktur()` → existiert bereits in DokumentePanel ✅
