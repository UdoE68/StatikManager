# Phase 1 – Infrastruktur-Extraktion: Abschlussbericht

**Datum:** 2026-03-31
**Status:** Abgeschlossen — Build erfolgreich, 0 Fehler
**Verantwortlich:** Fiona (Implementierung), Nolen (Verifikation)

---

## Ziel

Infrastruktur-Logik aus der monolithischen `DokumentePanel.xaml.cs` extrahieren.
Keinerlei funktionale Änderungen. Kein neues Verhalten. Nur Verschiebung in dedizierte Klassen.

---

## Neue Dateien (erstellt)

### `src/StatikManager/Infrastructure/PdfRenderer.cs`
- `RenderiereAlleSeiten(string pfad, int breite, int höhe, CancellationToken)` — rendert alle Seiten einer PDF als BitmapSource-Liste via Docnet.Core
- `KompositioniereGegenWeiss(byte[] raw, int w, int h)` — Alpha-Kompositionierung gegen Weiß (float-Variante)
- **Quelle:** `PdfSchnittEditor.xaml.cs` (exakte Kopie)

### `src/StatikManager/Infrastructure/PdfCache.cs`
- `CacheBasis` — Basispfad unter `%APPDATA%\StatikManager\pdf-cache`
- `CacheVersion = "v5"` — Cache-Invalidierungsschlüssel
- `GetBasisPdfPfad`, `GetCoveredPdfPfad`, `GetWordKlonPfad` — Pfadberechnung per Hash
- `CacheGültig` — LastWriteTime-Vergleich
- `LöscheCacheFürDatei` — Löscht alle Cache-Einträge zu einer Quelldatei
- **Quelle:** `DokumentePanel.xaml.cs`

### `src/StatikManager/Infrastructure/DateiTypen.cs`
- `IstWordDatei`, `IstPdfDatei`, `IstBildDatei` — Erweiterungs-Klassifikation
- `DateiIcon` — Emoji-Icon je Dateityp
- **Quelle:** `DokumentePanel.xaml.cs`

### `src/StatikManager/Infrastructure/WordInteropService.cs`
- `WortDateiZuPdf(string quellPfad, string zielpfad)` — Einzeldatei Word→PDF via COM-Interop
- `WortDateienBatchZuPdf(DirectoryInfo root, string cacheDir, CancellationToken)` — Batch-Konvertierung mit Cache-Prüfung
- **Quelle:** `DokumentePanel.xaml.cs` (Methoden `ErstelleBasisPdf`, `VorkonvertierungTask`)
- **Hinweis:** STA-Thread-Verantwortung liegt beim Aufrufer (unverändert)

---

## Geänderte Dateien

### `DokumentePanel.xaml.cs`
- `using StatikManager.Infrastructure;` hinzugefügt
- Folgende lokale Methoden entfernt: `ErstelleBasisPdf`, `VorkonvertierungTask`, `GetBasisPdfPfad`, `GetCoveredPdfPfad`, `GetWordKlonPfad`, `CacheGültig`, `LöscheCacheFürDatei`, `IstWordDatei`, `IstPdfDatei`, `IstBildDatei`, `DateiIcon`
- Felder `CacheBasis`, `CacheVersion` entfernt
- Inline-Alpha-Loop in `RendereWordSeiten` ersetzt durch `PdfRenderer.KompositioniereGegenWeiss(...)`
- Alle Aufrufe auf `PdfCache.*`, `DateiTypen.*`, `WordInteropService.*`, `PdfRenderer.*` umgeleitet

### `PdfSchnittEditor.xaml.cs`
- `using StatikManager.Infrastructure;` hinzugefügt
- Lokale Methoden `RenderiereAlleSeiten`, `KompositioniereGegenWeiss` entfernt
- Aufrufe auf `PdfRenderer.RenderiereAlleSeiten(...)` umgeleitet

---

## Build-Verifikation (Nolen)

```
MSBuild 17.14.8 für .NET Framework
Konfiguration: Debug | x64
Ziel:          net48
```

| Kategorie | Anzahl | Details |
|---|---|---|
| Fehler | 0 | — |
| Warnungen | 4 | siehe unten |
| Ausgabe | ✓ | `bin/x64/Debug/net48/StatikManager.exe` |

### Bekannte Restwarnungen (alle vor Phase 1 vorhanden)

| Code | Anzahl | Beschreibung |
|---|---|---|
| NU1603 | 3 | Docnet.Core 2.6.0 aufgelöst statt 2.4.0 — NuGet-Versionskonflikt, nicht Phase-1-bedingt |
| CS8602 | 1 | Nullable-Dereferenzierung, DokumentePanel.xaml.cs Zeile 103 — war vor Phase 1 vorhanden |

---

## Zwischenfälle

| Problem | Ursache | Lösung |
|---|---|---|
| CS0106 / CS0538 in DokumentePanel.xaml.cs | `replace_all` für `GetWordKlonPfad(` traf auch die Methodendefinition und erzeugte `private static string PdfCache.GetWordKlonPfad(...)` (ungültige explizite Interface-Syntax) | Methode vollständig entfernt (war bereits in PdfCache extrahiert) |

---

## Offene Punkte für Phase 2

1. **DokumentePanel.xaml.cs** (~900 Zeilen): UI, Zustand und Events noch nicht getrennt — kandidiert für weitere Zerlegung
2. **PdfSchnittEditor.xaml.cs**: Layout- und Export-Hilfslogik noch inline
3. **CS8602-Warnung** (Zeile 103): Nullable-Absicherung prüfen und ggf. beheben
4. **Kein Unit-Test-Projekt**: Die neuen Infrastructure-Klassen sind jetzt isoliert testbar — Test-Setup wäre sinnvoll
5. **Weitere Module** (z. B. `PositionenPanel`, `ProjektVerwaltung`): Strukturanalyse aus `plan_strukturanalyse_src.md` noch nicht adressiert
