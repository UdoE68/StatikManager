# StatikManager – Learnings

Chronologische Erkenntnisse aus der Entwicklung. Was funktioniert hat und warum.

---

## 2026-04-06 – _bearbeitet.pdf hat Vorrang beim Laden (Komposit-Persistenz)

**Problem:** Nach "Lücke schließen" + "Teile getrennt lassen" wurde beim nächsten Laden die Änderung scheinbar verworfen — der gelöschte Bereich erschien wieder.

**Root Cause:** `SchiebeTeileZusammen(verschmelzen=false)` erzeugt `_kompositBilder[si]` (zusammengeschobene Bitmap ohne Lücke). `SpeichereGesamtzustand()` schreibt diese korrekt in `_bearbeitet.pdf`. Aber `SpeichereSchnittState()` persistiert `_kompositBilder` NICHT im JSON (JSON kennt nur `_scherenschnitte` + `_gelöschteParts`). Beim nächsten Laden: JSON vorhanden → Original wurde geladen → LadeSchnittState → Schnittlinien wieder da, aber `_kompositBilder` leer → Rendering zeigt Original-Seite mit Schnittlinien → Lücke erscheint wieder.

**Lösung:**
1. `LadePdf`: Lade-Entscheidung geändert — `_bearbeitet.pdf` wird IMMER geladen wenn vorhanden, unabhängig davon ob JSON existiert.
2. `LadeSchnittState(bool nurSchnittlinien)`: Neuer Parameter. Wenn `_bearbeitet.pdf` geladen wurde (`bearbeitetVorhanden=true`), werden nur `_scherenschnitte` wiederhergestellt (für Interaktivität), aber NICHT `_gelöschteParts` und `_gelöschteSeiten` — diese sind bereits physisch in der `_bearbeitet.pdf` eingearbeitet.

**Goldene Regel:** `_bearbeitet.pdf` ist die physische Wahrheit. Sie enthält den vollständig eingearbeiteten Zustand (Komposit-Bitmaps, zusammengeschobene Teile, gelöschte Seiten). JSON ist nur Metadaten für Interaktivität (Schnittlinien zum erneuten Schneiden). Immer `_bearbeitet.pdf` bevorzugen.

**Beweis:** Build 0 Fehler, Commit f587f9f, 2026-04-06. Tester-Verifikation ausstehend.

**Dateien:** `Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`

---

## 2026-04-06 – Physische Blocktrennung: CroppedBitmap als alleinige Render-Wahrheit

**Problem:** Im PdfSchnittEditor wurde die vollständige Originalseite als Bitmap gezeichnet, darüber Overlay-Rechtecke für "Teile". Damit war kein echter Split möglich — die Blöcke waren visuelle Masken, nicht physisch getrennte Elemente.

**Lösung (bewiesen im Prototyp BlockEditorPrototype):**
- `_originalBitmap` wird NIEMALS als Ganzes auf den Canvas gezeichnet
- Jeder `ProtoBlock` (FracTop / FracBottom) wird als eigener `CroppedBitmap`-Ausschnitt gerendert:
  ```csharp
  new CroppedBitmap(originalBitmap, new Int32Rect(0, pixelTop, srcW, pixelHeight))
  ```
- `SplitBlock()` entfernt den Originalblock aus `_blocks` und ersetzt ihn durch zwei neue — kein Schnitt-Koordinaten-Eintrag, kein Overlay
- `DeleteBlock()` setzt `IsDeleted = true`; `RenderBlocks()` überspringt diese Blöcke komplett
- Ergebnis: kein Durchscheinen, kein Overlay-Rest, volle physische Trennung

**Warum das alte Modell scheiterte:**
`_scherenschnitte` speicherte nur Koordinatenpunkte. `GetTeilGrenzen()` berechnete daraus Grenzen on-the-fly. `ZeicheSeite()` zeichnete stets die volle Seite — die "Teile" lagen nur als transparente Rechtecke darüber. Echter Split war strukturell nicht möglich.

**Goldene Regel:** `_originalBitmap` (oder `_seitenBilder[i]`) darf im Split-Renderpfad **niemals vollständig** auf den Canvas gezeichnet werden. Nur `CroppedBitmap`-Ausschnitte sind erlaubt.

**Beweis:** BlockEditorPrototype, Tests A/B/C, 2026-04-06. User-Bestätigung: kein Durchscheinen, kein Overlay, Delete funktioniert sauber.

**Dateien:** `Modules/Werkzeuge/BlockEditorPrototype.xaml.cs`

---

## 2026-04-06 – Paradigmenwechsel: Nie die Original-PDF ueberschreiben (BearbeitetPfad-Muster)

**Problem:** pdfium haelt die geoeffnete PDF-Datei gesperrt solange sie angezeigt wird. `File.WriteAllBytes(_pdfPfad, ...)` schlaegt daher immer mit "Datei wird von einem anderen Prozess verwendet" fehl — unabhaengig von Retry-Schleifen oder MemoryStream-Tricks.

**Loesung:** Die Original-PDF wird NIEMALS als Schreibziel verwendet. Stattdessen wird eine Geschwister-Datei angelegt:
- `BearbeitetPfadFuer(string originalPfad)` → gibt `<name>_bearbeitet.pdf` im selben Ordner zurueck
- `SpeichereAenderungen()` → schreibt ausschliesslich auf `bearbeitetPfad`, nie auf `_pdfPfad`
- `AutoSpeichern()` → ebenfalls auf `bearbeitetPfad` umgestellt (auch wenn nicht aktiv aufgerufen)
- `LadePdf()` → prueft ob `_bearbeitet.pdf` neben dem Original existiert; wenn ja, wird diese geladen
- `_pdfPfad` bleibt IMMER der Original-Pfad und dient nur der Namensgebung von `_bearbeitet.pdf`
- `SpeichereMetadaten()` → schreibt Schnittlinien und geloeschte Teile als `<original>.edit.json`

**Goldene Regel:** `_pdfPfad` darf NIEMALS als Schreibziel verwendet werden. Immer `BearbeitetPfadFuer(_pdfPfad)`.

**Grund:** pdfium-Handles sind nicht schliessbar ohne das gesamte Dokument zu entladen. Das Nebeneinander von Original (read-only, immer gesperrt) und Bearbeitet-Datei (write-target, nie gesperrt) umgeht das Problem grundsaetzlich statt es zu kaempfen.

**Beweis:** Tester PASS, 2026-04-06. `grep "WriteAllBytes(_pdfPfad"` → 0 Treffer. Build: 0 Errors.

**Dateien:** `Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`

---

## 2026-04-06 – Speicher-Dialog beim Positionswechsel (statt AutoSpeichern)

**Problem:** AutoSpeichern-Ansatz hatte mehrere Versagensmodi (Datei-Locking, fehlgeschlagene Retries, kein Feedback an Benutzer). Der Benutzer wusste nie ob seine Aenderungen gespeichert wurden oder verloren gingen.

**Loesung:** Dirty-Flag + expliziter Speicher-Dialog:
- `_hatUngespeicherteAenderungen = true` bei jeder Aenderung (ersetzt alle 10 AutoSpeichern()-Aufrufe)
- `FrageObSpeichern()`: zeigt Ja/Nein/Abbrechen-Dialog wenn Dirty-Flag gesetzt
- `SpeichereAenderungen()`: speichert via `SpeicherInStream` + `WriteAllBytes` + 3 Retries, mit sichtbarer Fehler-MessageBox, setzt Dirty-Flag auf false
- `LadePdf()`: ruft `FrageObSpeichern()` am Anfang auf, bricht bei Abbrechen ab (return), setzt `_hatUngespeicherteAenderungen = false` nach erfolgreichem Laden
- `MainWindow.OnClosing()` (nicht OnClosed!): prueft `panel.PdfEditor.FrageObSpeichern()`, setzt `e.Cancel = true` bei Abbrechen

**Architektur-Entscheidungen:**
- Dialog in `LadePdf()` selbst, nicht im aufrufenden Code — LadePdf ist der einzige Eintrittspunkt fuer neue PDFs
- `OnClosing` (nicht `OnClosed`) weil nur `Closing` das `e.Cancel`-Flag hat um das Schliessen abzubrechen
- `AutoSpeichern()` bleibt als Methode erhalten (kein Code-Break), wird aber nicht mehr aufgerufen
- `SpeichereAenderungen()` ist von `AutoSpeichern()` unabhaengig — zeigt MessageBox bei Fehler statt lautlos zu scheitern

**Beweis:** Tester PASS, feature/word-export-next branch

**Dateien:** `Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`, `MainWindow.xaml.cs`

---

## 2026-04-06 – AutoSpeichern via WriteAllBytes statt File.Replace

**Problem:** File.Replace scheitert wenn Datei von irgend etwas gesperrt ist (z.B. Explorer-Vorschau, Antivirus)

**Loesung:** SpeicherInStream() baut PDF in MemoryStream, File.WriteAllBytes() schreibt direkt + 3x Retry bei IOException. Fallback auf _autosave.pdf wenn alle Versuche scheitern. Keine MessageBox mehr (blockiert UI nicht).

**Warum besser:** WriteAllBytes ist atomarer als Replace, kein Zwischen-Handle, Retry-Schleife faengt temporaere Sperren ab

**Dateien:** PdfSchnittEditor.xaml.cs (AutoSpeichern + neue Methode SpeicherInStream)

---

## 2026-04-05 – AutoSpeichern ohne Datei-Locking

**Problem:** pdfium (`DocLib.GetDocReader`) und PdfSharp (`PdfReader.Open`) halten PDF-Dateien offen, was `File.Replace` im AutoSpeichern blockiert ("Datei wird von einem anderen Prozess verwendet").

**Loesung:** `_pdfBytes`-Feld haelt die PDF als byte[] im Speicher. ALLE Lese-Zugriffe gehen ueber `new MemoryStream(_pdfBytes)` statt ueber den Dateipfad:
- `PdfRenderer.RenderiereAlleSeiten(byte[] bytes, ...)` — kein Datei-Handle
- `PdfReader.Open(new MemoryStream(_pdfBytes), PdfDocumentOpenMode.Import)` in SpeicherNachPfad
- `PdfReader.Open(new MemoryStream(_pdfBytes), PdfDocumentOpenMode.Import)` im Scheren-Export
- `HolePdfSeitenGroesse(byte[] bytes)` — PdfReader.Open via MemoryStream
- `_pdfBytes = File.ReadAllBytes(_pdfPfad)` nach erfolgreichem AutoSpeichern (aktuell halten)

**Beweis:** Build 0 Fehler, Commit 0c22632, PDFs FREI im Sperr-Test

**Dateien:** `Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`, `Infrastructure/PdfRenderer.cs`

---

## 2026-04-05 – FileSystemWatcher Thundering Herd vermeiden

**Problem:** OrdnerWatcher-Debounce-Timer wurde staendig zurueckgesetzt und feuerte nie, weil `Changed+LastWrite` bei jedem Speichervorgang hunderte Events erzeugte.

**Loesung:** `NotifyFilter` auf `FileName | DirectoryName` reduziert. `Changed`-Event gar nicht abonniert. `InternalBufferSize = 65536` fuer grosse Ordner.

**Grund:** `LastWrite` feuert bei JEDEM Schreibvorgang auf eine Datei. Bei einem Speichern von Word kann das >100 Events pro Sekunde sein.

**Dateien:** `Infrastructure/OrdnerWatcherService.cs`

---

## 2026-04-05 – AktualisiereNurStruktur statt AktualisiereDokumentListe fuer Watcher

**Problem:** `OrdnerGeaendert`-Event rief `AktualisiereDokumentListe()` auf, das `_aktiverDateipfad = null` setzte und die Vorschau loeschte.

**Loesung:** Neue Methode `AktualisiereNurStruktur()`: baut nur Baum/Liste neu ohne aktive Vorschau zu unterbrechen. Watcher feuert nur diese Methode.

**Dateien:** `Modules/Dokumente/DokumentePanel.xaml.cs`

---

## 2026-04-05 – NavigateToString Umlaute (IE-Engine)

**Problem:** Deutsche Umlaute wurden als `verfuegbar` statt `verfügbar` angezeigt. `NavigateToString` uebergibt UTF-16 aber IE interpretiert es als ANSI.

**Loesung:**
1. `HtmlEncode()`: konvertiert alle Zeichen > 127 zu `&#NNN;` numerischen Entities
2. `HtmlSeite()`: fuegt `<meta http-equiv='Content-Type' content='text/html; charset=utf-16'>` ein

**Grund:** WPF WebBrowser nutzt IE-Engine. `NavigateToString` uebergibt intern UTF-16, aber IE ignoriert das ohne explizites charset.

**Dateien:** `Modules/Dokumente/DokumentePanel.xaml.cs`

---

## 2026-04-05 – WPF TreeView Multi-Select

**Problem:** WPF TreeView unterstuetzt kein natives Multi-Select.

**Loesung:** Manuell mit:
- `HashSet<string> _baumMehrfachAuswahl` fuer ausgewaehlte Pfade
- `PreviewMouseLeftButtonDown` fuer Ctrl+Klick (Toggle) und Shift+Klick (Range)
- `AktualisiereTreeViewHervorhebung()`: setzt `item.Background` fuer visuelle Rueckmeldung
- `OrdneTreeItemsFlach()`: rekursive Hilfe fuer Range-Selektion

**Dateien:** `Modules/Dokumente/DokumentePanel.xaml.cs`

---

## 2026-04-05 – DataGridCheckBoxColumn braucht 2 Klicks

**Problem:** `DataGridCheckBoxColumn` erfordert immer zwei Klicks: Erst Zeile selektieren, dann Checkbox toggeln.

**Loesung:** `DataGridTemplateColumn` mit `CheckBox` und `UpdateSourceTrigger=PropertyChanged`:
```xml
<DataGridTemplateColumn>
    <DataGridTemplateColumn.CellTemplate>
        <DataTemplate>
            <CheckBox IsChecked="{Binding Sichtbar, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
        </DataTemplate>
    </DataGridTemplateColumn.CellTemplate>
</DataGridTemplateColumn>
```

---

## 2026-04-05 – C# init-Accessor nicht in .NET 4.8

**Problem:** `public string Pfad { get; init; }` → `CS0518: IsExternalInit ist nicht definiert`.

**Loesung:** `init` durch `set` ersetzen. .NET 4.8 kennt den init-only Accessor nicht.

---

## 2026-04-05 – HTML-zu-PDF ohne WebView2

**Problem:** App nutzt `WebBrowser` (IE-Engine), nicht WebView2. `PrintToPdfAsync()` nicht verfuegbar.

**Loesung:** Microsoft Edge Headless:
```
msedge.exe --headless --no-sandbox --print-to-pdf="output.pdf" "file:///path/to/file.html"
```
Edge ist auf Windows 10/11 immer installiert. Timeout 30s. Ausgabepfad neben HTML-Datei.

**Dateien:** `Modules/Dokumente/DokumentePanel.xaml.cs` (SucheEdgePfad, ErzeugeHtmlPdfImHintergrund)

---

## 2026-04-05 – Ordner-Klick oeffnet position.html automatisch

**Problem:** Wenn Benutzer Pos_XX-Ordner anklickt, passiert nichts (kein Vorschau-Routing fuer Ordner).

**Loesung:** In `DokumentenBaum_SelectedItemChanged`: wenn `Directory.Exists(pfad)`, pruefen ob `position.html` im Ordner liegt. Falls ja: `StarteSelektionDebounce(posHtmlPfad)`.

**Dateien:** `Modules/Dokumente/DokumentePanel.xaml.cs`
