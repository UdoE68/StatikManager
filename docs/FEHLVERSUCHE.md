# StatikManager – Fehlversuche

Was nicht funktioniert hat und warum. Damit niemand denselben Fehler zweimal macht.

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
