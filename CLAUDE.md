# StatikManager V1 — Projektregeln

## Projektbeschreibung
StatikManager ist eine modulare WPF-Desktopanwendung (.NET Framework 4.8, x64) zur Dokumentenverwaltung fuer Statik-Projekte. Sie zeigt Projektordner als Dateibaum an und ermoeglicht die Vorschau von PDF-, Word-, HTML-, Bild- und JSON-Dateien.

**Verbindung zu PP_ZoomRahmen**: Das AxisVM-Plugin PP_ZoomRahmen erzeugt in Positionsordnern `position.html` und `position.json`. Der StatikManager zeigt diese Dateien in der Vorschau an. Der StatikManager ist read-only – er erstellt keine Positionen, das tut ausschliesslich PP_ZoomRahmen.

---

## Agenten-Struktur

### .claude/agents/ (Claude Code Sub-Agenten)

| Agent        | Rolle                                                          |
|--------------|----------------------------------------------------------------|
| orchestrator | Projektleiter: plant, delegiert, prueft. KEIN Commit ohne @tester |
| bibliothekar | Wissensverwalter: LEARNINGS, PATTERNS, FEHLVERSUCHE            |
| entwickler   | WPF/C# Implementierung, Build, Git                             |
| tester       | Qualitaetssicherung: verifiziert konkret, meldet PASS/FAIL     |

Wissensdatenbank: `docs/` (ARCHITEKTUR.md, LEARNINGS.md, PATTERNS.md, FEHLVERSUCHE.md)

### Pflicht-Workflow fuer jede Aufgabe

```
1. User → @orchestrator: Aufgabe beschreiben
2. @orchestrator → @bibliothekar: "Was wissen wir zu diesem Thema?"
3. @bibliothekar liefert: Vorwissen, Fehlversuche, Patterns, Fallen
4. @orchestrator → @entwickler: Aufgabe + Bibliothekar-Wissen
5. @entwickler: Lesen → Analysieren → Code → Bauen
6. @orchestrator → @tester: "Verifiziere Fix X"
7. @tester prueft KONKRET (Zeitstempel, Debug-Output, Datei-Lock)
8. Bei FAIL → zurueck zu Schritt 4 mit Tester-Feedback
9. Bei PASS → git commit + push
10. @orchestrator → @bibliothekar: "Dokumentiere neue Erkenntnisse"
```

**KEIN COMMIT OHNE TESTER-OK.**

---

## Allgemeine Arbeitsregeln

### Reihenfolge: Erst analysieren, dann aendern
Vor jeder Aenderung wird der betroffene Code vollstaendig gelesen und verstanden.
Kein blindes Umschreiben. Kein Raten. Immer zuerst den Ist-Zustand erfassen.

### /original bleibt unberuehrt
Der Ordner `/original` enthaelt den unveraenderten Quellstand des Projekts.
Keine Datei in `/original` wird jemals veraendert, verschoben oder geloescht.
Er dient als Referenz und Rueckfallposition.

### Aenderungen nur ausserhalb von /original
Alle Anpassungen, Refactorings, neuen Dateien und Migrationen landen ausschliesslich
im Arbeitsbereich ausserhalb von `/original` (z. B. `/src`, `/migration`, etc.).

### Backup-Prinzip vor groesseren Aenderungen
Vor strukturell bedeutsamen Aenderungen wird ein Snapshot oder Zwischenstand gesichert.
Idealerweise als Kopie im `/migration`-Ordner mit Zeitstempel-Suffix.

---

## Build-Anleitung

**Konfiguration:** Debug | x64
**Voraussetzung:** Visual Studio 2022, .NET Framework 4.8, Microsoft Word installiert

```powershell
# Build:
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"

# Starten:
C:\KI\StatikManager_V1\src\StatikManager\Start_Debug.bat
```

**Wichtig vor dem Build:** StatikManager.exe beenden (pdfium.dll wird sonst gesperrt).

---

## Technologie-Stack

| Technologie              | Verwendung                                    |
|--------------------------|-----------------------------------------------|
| WPF / XAML               | UI-Framework                                  |
| C# .NET Framework 4.8    | Sprache und Runtime                           |
| Docnet.Core 2.6.0        | PDF-Rendering (pdfium)                        |
| PdfSharp                 | PDF-Manipulation                              |
| Microsoft.Office.Interop | Word COM-Automatisierung                      |
| WebBrowser (IE-Engine)   | HTML/Bild-Vorschau (kein WebView2!)           |
| Edge Headless            | HTML -> PDF Export (msedge.exe --headless)    |
| XmlSerializer            | Einstellungen (AppData\StatikManager\)        |

---

## Bekannte Probleme und Fallstricke

| Problem                          | Ursache                          | Loesung                               |
|----------------------------------|----------------------------------|---------------------------------------|
| Umlaute in NavigateToString      | IE liest UTF-16 falsch           | HtmlEncode() + charset=utf-16 Meta    |
| FileSystemWatcher Thundering Herd| Changed+LastWrite bei Speichern  | Nur FileName\|DirectoryName           |
| WPF TreeView Multi-Select        | Nicht nativ unterstuetzt         | HashSet + PreviewMouseDown manuell    |
| DataGridCheckBoxColumn 2 Klicks  | WPF-Bug                          | DataGridTemplateColumn mit CheckBox   |
| C# init-Accessor                 | .NET 4.8 unbekannt               | set statt init verwenden              |
| pdfium paralleler Zugriff        | Native DLL nicht thread-safe     | AppZustand.RenderSem (Semaphore)      |
| Word-COM im UI-Thread            | COM STA-Anforderung              | Thread.SetApartmentState(STA)         |

---

## Git-Regeln

Nach jeder erfolgreichen Aenderung IMMER automatisch:

1. `git add [Dateien]`
2. `git commit -m "Beschreibung"`
3. `git push`

- Keine Aenderung ohne Commit + Push abschliessen
- Vor Commit immer Build pruefen (0 Fehler)
- Keine Nachfrage notwendig
- Branch: `feature/word-export-next`

---

## Kommunikationsformat zwischen Agenten

Jeder Auftrag zwischen Agenten enthaelt:
- **Ziel**: Was soll erreicht werden?
- **Kontext**: Welche Dateien / Abhaengigkeiten sind relevant?
- **Einschraenkungen**: Was darf nicht veraendert werden?
- **Erfolgskriterium**: Woran erkennt man, dass die Aufgabe erledigt ist?
