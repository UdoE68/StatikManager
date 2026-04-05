# StatikManager – Learnings

Chronologische Erkenntnisse aus der Entwicklung. Was funktioniert hat und warum.

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
