# StatikManager – Fehlversuche

Was nicht funktioniert hat und warum. Damit niemand denselben Fehler zweimal macht.

---

## 2026-04-06 – PdfSchnittEditor: Overlay-Ansatz für Blocktrennung (grundlegend falsch)

**Versuch 1:** `ZeicheSeite(i)` zeichnet vollständige Seite, darüber transparente `Border`-Rechtecke als "Teile".
**Fehler:** Keine echte Trennung. Beim Löschen eines Teils bleibt das Bitmap der ganzen Seite sichtbar (Durchscheinen). Kein echter unabhängiger Block.
**Grund:** Das Altmodell (`_scherenschnitte` als Koordinatenliste + `GetTeilGrenzen()` on-the-fly) produziert niemals physisch getrennte Elemente. Es ist ein Masken-Modell, kein Block-Modell.

**Versuch 2:** `ZeicheSeiteAlsBlöcke()` in `ZeicheSeite()` eingebaut — CroppedBitmap pro Block, aber `_contentBlocks` parallel zu `_scherenschnitte` gepflegt.
**Fehler:** Zwei konkurrierende Wahrheiten. `_contentBlocks` wurde von `SplitContentBlockBeiSchnitt()` korrekt geteilt, aber das Rendering las weiter aus `_seitenBilder[i]` (Altmodell). Kein echter physischer Split im Anzeigebild.
**Grund:** Solange `_scherenschnitte` die führende Datenstruktur für Rendering UND Klick-Zuordnung ist, kann `_contentBlocks` kein echter Ersatz werden — nur ein Begleiter.

**Versuch 3 (ERFOLGREICH — Prototyp):** Vollständig isolierter `BlockEditorPrototype` ohne Altmodell.
- Einzige Wahrheit: `List<ProtoBlock>` + `_originalBitmap`
- `SplitBlock()` entfernt Original, fügt zwei neue ein — keine Koordinatenliste
- `RenderBlocks()` zeichnet nur `CroppedBitmap`-Ausschnitte — kein Gesamtbild
- `DeleteBlock()` setzt `IsDeleted=true` → Block wird von `RenderBlocks()` vollständig übersprungen
**Ergebnis:** PASS — Tests A/B/C bestanden, User-Bestätigung 2026-04-06.

**Lessons Learned:**
1. Ein Overlay ist kein Block. Solange das Ursprungsbild gezeichnet wird, gibt es keinen echten Split.
2. Zwei parallele Datenmodelle (`_scherenschnitte` + `_contentBlocks`) können nicht gleichzeitig führend sein — eines muss gewinnen.
3. Der richtige Migrationspfad: `_contentBlocks` wird die einzige Render-Wahrheit; `_scherenschnitte` dient nur noch als Eingabe für `SplitBlock()`.
4. Kein Umbau am bestehenden Editor, solange der Prototyp nicht bewiesen hat, dass das Modell funktioniert.

---

## 2026-04-05 – PdfSchnittEditor: AutoSpeichern / Datei gesperrt (4 Fehlversuche)

**Versuch 1:** `AutoSpeichern()` schreibt direkt auf `_pdfPfad`
**Fehler:** "Datei wird von einem anderen Prozess verwendet"
**Grund:** pdfium (`DocLib.GetDocReader(pfad, ...)`) haelt die Datei offen.

**Versuch 2:** Temp-Datei + `File.Replace(temp, original, null)`
**Fehler:** Immer noch gesperrt
**Grund:** Datei war immer noch geoeffnet — `File.Replace` scheitert wenn Zieldatei gesperrt.

**Versuch 3:** `_pdfBytes` + `GetDocReader(bytes, ...)` fuer pdfium-Rendering
**Fehler:** Immer noch gesperrt
**Grund:** `PdfReader.Open(_pdfPfad, PdfDocumentOpenMode.Import)` in `SpeicherNachPfad` haelt die Datei offen — war nicht geaendert worden.

**Versuch 4:** `HolePdfSeitenGroesse` auf `MemoryStream` umgestellt
**Fehler:** Immer noch Berichte ueber Fehler (nicht verifiziert ob Fix4 funktioniert)
**Grund:** Unbekannt — moeglicherweise war der Build nicht deployed oder ein weiterer Zugriff existiert.

**Versuch 6 (neu):** Komplette Neustrategie: AutoSpeichern baut PDF in MemoryStream,
schreibt dann mit WriteAllBytes (kein File.Replace), Retry-Loop bei IOException
**Ergebnis:** PASS fuer Datei-Locking, aber AutoSpeichern als Konzept wurde danach aufgegeben.

**Versuch 7 — PARADIGMENWECHSEL (ERFOLGREICH, 2026-04-06):** AutoSpeichern vollstaendig ersetzt durch Dirty-Flag + expliziten Speicher-Dialog.
- `_hatUngespeicherteAenderungen` Feld statt AutoSpeichern-Aufrufen
- `FrageObSpeichern()` zeigt Ja/Nein/Abbrechen-Dialog bei Positionswechsel und App-Schliessen
- `SpeichereAenderungen()` mit Fehler-MessageBox statt lautlosem Scheitern
- `MainWindow.OnClosing()` (nicht OnClosed!) fuer Window-Closing-Interceptor
**Ergebnis:** PASS — Benutzer hat volle Kontrolle, kein unsichtbares Fehlschlagen mehr

**Lehre zum Paradigma:** AutoSpeichern ist prinzipiell fehleranfaellig (Datei gesperrt, Benutzer bemerkt Fehler nicht). Wenn der Benutzer aktiv speichern oder verwerfen kann, ist das robuster und benutzerfreundlicher.

**Versuch 8 — WriteAllBytes auf _pdfPfad (2026-04-06, letzter Datei-Lock-Fehler):**
`SpeichereAenderungen()` und `AutoSpeichern()` verwendeten noch `WriteAllBytes(_pdfPfad, bytes)` als Schreibziel.
**Fehler:** pdfium haelt `_pdfPfad` dauerhaft gesperrt — WriteAllBytes schlaegt immer fehl, auch mit MemoryStream-Aufbau und Retry-Schleife.
**Grund:** Das Problem liegt nicht beim Lesen (bereits auf MemoryStream umgestellt), sondern beim Schreiben: pdfium gibt den Lese-Handle auf `_pdfPfad` nie frei solange das Dokument angezeigt wird. Die Datei kann daher prinzipiell nicht als Schreibziel dienen.
**Loesung:** `BearbeitetPfadFuer(_pdfPfad)` Muster — Original-PDF ist permanent read-only, alle Schreibvorgaenge gehen auf `<name>_bearbeitet.pdf`. Siehe LEARNINGS.md 2026-04-06.
**Verifikation:** `grep "WriteAllBytes(_pdfPfad"` → 0 Treffer. PASS.

**Versuch 5 (ERFOLGREICH):** Alle Zugriffe auf MemoryStream umgestellt
**Was hat funktioniert:**
- `SpeicherNachPfad` (Zeile ~4903): `PdfReader.Open(new MemoryStream(_pdfBytes), Import)` — war bereits gefixt
- `HolePdfSeitenGroesse` (Zeile ~2776): `PdfReader.Open(ms, ReadOnly)` mit byte[]-Ueberladung — war bereits gefixt
- `GetDocReader` in PdfRenderer: `lib.GetDocReader(bytes, ...)` — war bereits gefixt
- **NEU (Versuch 5):** `RendereTeileExportieren` (Zeile ~3599): `PdfReader.Open(_pdfPfad!, Import)` → `PdfReader.Open(new MemoryStream(_pdfBytes), Import)` — das war die letzte verbliebene Datei-Sperr-Stelle
- `AutoSpeichern`: `_pdfBytes = File.ReadAllBytes(_pdfPfad)` nach `File.Replace` — bereits vorhanden
**Beweis:** Build 0 Fehler, Commit 0c22632, EXE-Zeitstempel 05.04.2026 23:51:47

**Lehre:**
- Mit `grep -r "PdfReader.Open\|File.Open\|FileStream\|GetDocReader" --include="*.cs"` ALLE Dateizugriffe finden
- JEDER `PdfReader.Open(pfad, ...)` muss zu `PdfReader.Open(new MemoryStream(_pdfBytes), ...)` werden
- KEIN Commit ohne @tester-Verifikation dass Datei tatsaechlich frei ist
- Tester-Check: `[IO.File]::OpenWrite("datei.pdf").Close()` — kein Fehler = frei

---

## 2026-04-05 – PdfSchnittEditor: Off-by-one beim Loeschen (3+ Fehlversuche)

**Symptom:** Wenn Teil T markiert und geloescht wird, verschwindet der falsche Bereich.

**Versuch 1:** ErzeugeKompositBild — Quelle `_seitenBilder` vs `_kompositBilder` geprueft
**Fehler:** Hat nichts geaendert

**Versuch 2:** MouseLeftButtonUp — Null-Drag-Erkennung
**Fehler:** Hat nichts geaendert

**Versuch 3:** Weitere Versuche ohne Root-Cause-Analyse
**Fehler:** Hat nichts geaendert

**Vermutete Ursache (nicht bewiesen):** Segment-Index und Schnittlinien-Index werden unterschiedlich gezaehlt. Wenn `_kompositBilder[si]` existiert und `_scherenschnitte` fuer si vorhanden sind, gelten die Fraktionen im Komposit-Raum. Aber nach `SchiebeTeileZusammen` werden die alten Schnitte entfernt und neue hinzugefuegt — moeglicherweise mit falschen Fraktionswerten.

**Lehre:**
- Root-Cause zuerst beweisen via Debug.WriteLine BEVOR Code geaendert wird
- Minimal-Diagnose: `Debug.WriteLine($"[LOESCHEN] si={si} t={t} grenzen=[{string.Join(",", GetTeilGrenzen(si).Select(g => $"{g.Oben:F3}-{g.Unten:F3}"))}]")`
- Ohne bewiesene Ursache kein Fix-Versuch

---

## 2026-04-05 – PdfSchnittEditor: Leerzeile erzeugt Leerseiten (2+ Fehlversuche)

**Versuch 1:** `FuegeLeerzeileEin` mit `newH = sourceH + 30`, Ueberlauf wenn `newH > origH`
**Fehler:** Immer Ueberlauf weil `ErzeugeKompositBild` auf `origH` paddert → `sourceBmp.PixelHeight == origH` → `newH = origH + 30` immer Ueberlauf → immer neue Seite → nach 2x sind nur noch Leerseiten da (Inhalt verloren)

**Versuch 2:** Schnittfraktionen auf `origH` normiert (`oldYPx / origH`)
**Fehler:** Teilweise besser aber Leerseiten-Problem ungeloest

**Root Cause:** `ErzeugeKompositBild` paddert immer auf `origH`. Daher ist `sourceBmp.PixelHeight` immer `origH`, unabhaengig davon wieviel echter Inhalt vorhanden ist. `inhaltH` kann NICHT aus der Bitmap-Hoehe abgeleitet werden.

**Korrekte Loesung (nicht vollstaendig implementiert):**
- Echte Inhalt-Hoehe tracken via `_seitenBelegteHoehe: Dictionary<int, int>`
- Init: `_seitenBelegteHoehe[si] = origH`
- Nach Loeschen (SchiebeTeileZusammen): `_seitenBelegteHoehe[si] = summe der sichtbaren Teile in Pixel`
- In FuegeLeerzeileEin: `belegtH = _seitenBelegteHoehe.GetValueOrDefault(si, origH)`, Ueberlauf = `belegtH + 30 > origH`

---

## 2026-04-05 – Agenten: Commits ohne Tester-Verifikation

**Muster:** Entwickler-Agent meldet "fertig" und committed, ohne dass die Funktion getestet wurde.
**Folge:** Bugs werden committed und deployed, Nutzer testet und findet sie kaputt.
**Lehre:** KEIN Commit ohne explizites PASS vom @tester-Agenten. Der @orchestrator-Agent muss diesen Workflow erzwingen.

---

## 2026-04-05 – Build nicht deployed (Start_Debug.bat falscher Pfad)

**Versuch:** Build ausgefuehrt, EXE-Zeitstempel war aktuell, aber App zeigte altes Datum.
**Fehler:** `Start_Debug.bat` startete `c:\Projekte\StatikManager\bin\Debug\net48\StatikManager.exe` (alter Pfad, nicht aktualisiert).
**Lehre:** Nach Build IMMER EXE-Zeitstempel der gestarteten Instanz pruefen: `Get-Process StatikManager | Get-Item` oder Titelleiste pruefen.
**Gefixt in:** Commit d60f88a

---

## 2026-04-05 – WebView2 PrintToPdfAsync

**Versuch:** HTML-zu-PDF Export ueber `WebView2.CoreWebView2.PrintToPdfAsync()`.

**Fehler:** Die App nutzt `WebBrowser` (IE-Engine), nicht WebView2. `PrintToPdfAsync` existiert nicht.

**Grund:** WebView2 und WebBrowser sind verschiedene Controls. StatikManager wurde mit dem alten WPF `WebBrowser`-Control entwickelt. Migration waere ein grosser Umbau.

**Alternative:** Edge Headless (`msedge.exe --headless --print-to-pdf`). Laeuft im Hintergrundprozess, kein UI-Control noetig.

---

## 2026-04-05 – FileSystemWatcher Changed + LastWrite

**Versuch:** `NotifyFilter = FileName | DirectoryName | LastWrite` und `Changed`-Event abonnieren um auch Dateiinhalt-Aenderungen zu erkennen.

**Fehler:** Debounce-Timer wird bei jedem Speichervorgang (Word, AutoSave etc.) hunderte Male zurueckgesetzt und feuert nie. UI aktualisiert sich nicht.

**Grund:** `LastWrite` + `Changed` feuert bei JEDEM Schreibzugriff auf eine Datei. Bei einem Word-Speichervorgang koennen das >100 Events/Sekunde sein.

**Alternative:** Nur `FileName | DirectoryName`. `Changed` gar nicht abonnieren. Strukturaenderungen (neue Dateien, Loeschungen) sind ausreichend fuer Baum-Aktualisierung.

---

## 2026-04-05 – Positionsverwaltung im StatikManager

**Versuch:** Neue Position erstellen, Unterordner `daten/` anlegen, position.html generieren im StatikManager.

**Fehler:** Falsche Zustaendigkeit. Der StatikManager ist ein read-only Dokumenten-Viewer.

**Grund:** Positionen werden vom AxisVM-Plugin PP_ZoomRahmen erstellt. StatikManager zeigt nur an. Klare Trennung der Verantwortlichkeiten.

**Alternative:** PP_ZoomRahmen fuer Positionserstellung nutzen. StatikManager nur fuer Anzeige und PDF-Export.

---

## 2026-04-05 – init-Accessor in .NET 4.8

**Versuch:** `public string Pfad { get; init; }` in ProjektVerwaltungDialog.xaml.cs.

**Fehler:** `CS0518: Der vordefinierte Typ System.Runtime.CompilerServices.IsExternalInit ist nicht definiert`.

**Grund:** Der `init`-Accessor ist ein C# 9-Feature das auf .NET 5+ ausgelegt ist. .NET 4.8 kennt den `IsExternalInit`-Typ nicht, der intern benoetigt wird.

**Alternative:** `set` statt `init` verwenden. Fuer .NET 4.8 ist `set` korrekt und ausreichend.

---

## 2026-04-05 – Projektverwaltung mit festem Basispfad

**Versuch:** Einen festen "Projektbasis-Pfad" als Stammordner erzwingen und alle Unterordner als Projekte scannen.

**Fehler:** Zu unflexibel. Projekte koennen auf verschiedenen Laufwerken und in beliebiger Tiefe liegen.

**Grund:** Reale Arbeitsumgebung hat Projekte verteilt: `D:\Projekte\Kunde1\`, `E:\Archiv\2024\`, etc.

**Alternative:** Freie Liste von Projektpfaden. Jeder Pfad wird einzeln per Ordner-Dialog hinzugefuegt. Keine Einschraenkung auf Unterordner eines Stamms.

---

## 2026-04-05 – AktualisiereDokumentListe vom OrdnerWatcher aufrufen

**Versuch:** `_ordnerWatcher.OrdnerGeaendert += AktualisiereDokumentListe` – bei jeder Datei-Aenderung die volle Liste neu laden.

**Fehler:** `AktualisiereDokumentListe()` setzt `_aktiverDateipfad = null` und navigiert zu `about:blank`. Aktive Vorschau wird bei jeder Ordneraenderung geloescht.

**Grund:** `AktualisiereDokumentListe` war urspruenglich fuer manuelle Projektwechsel designed und loescht dabei den gesamten Zustand.

**Alternative:** Separate `AktualisiereNurStruktur()`-Methode die nur Baum/Liste neu baut ohne aktive Vorschau anzufassen. Watcher ruft nur diese auf.
