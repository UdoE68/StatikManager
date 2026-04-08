# Design: Word-Auto-Einfügen im StatikManager

**Datum:** 2026-04-08
**Status:** Genehmigt

---

## Zusammenfassung

Das WordExportPanel erhält einen "Auto-Einfügen"-Toggle. Wenn er aktiv ist, überwacht ein `FileSystemWatcher` den Projektordner auf neue PNG-Exporte von PP_ZoomRahmen. Jede neue PNG wird automatisch an der aktuellen Word-Cursorposition eingefügt — ohne dass der Nutzer in den StatikManager wechseln muss.

---

## Architektur

Keine neue Klasse, keine Änderung an PP_ZoomRahmen oder `WordEinfuegenService`. Alle Änderungen liegen in:

- `WordExportPanel.xaml` — Checkbox + Statuszeile
- `WordExportPanel.xaml.cs` — Watcher-Logik
- `Einstellungen.cs` — neue Property `WordExportAutoEinfuegen`

---

## UI-Änderungen

Neuer Bereich am Ende der Einfüge-Optionen:

```
┌─────────────────────────────────────┐
│  Einfüge-Optionen                   │
│  [x] Überschrift  [x] Maßstab       │
│  Bildbreite: [Seitenbreite ▼]       │
│  ─────────────────────────────────  │
│  [ ] Auto-Einfügen bei neuem Export │
│  Zuletzt: Ansicht vorne — 14:32     │
└─────────────────────────────────────┘
```

- **Checkbox** `ChkAutoEinfuegen` — schaltet Watcher an/aus, Zustand wird in `Einstellungen` gespeichert
- **Statuszeile** `TxtAutoStatus` — einzeiliger grauer Text, kein Popup, kein Fokuswechsel

---

## Ablauf

```
PP_ZoomRahmen schreibt neue PNG in {Projekt}/Statik/Pos_*/daten/
    ↓
FileSystemWatcher (Created-Event auf *.png)
    ↓
300ms warten (Datei fertig schreiben)
    ↓
Gleichnamige .json lesen → Überschrift + Maßstab
    ↓
WordEinfuegenService.EinfuegenAnCursor(pngPfad, überschrift, massstab, ...)
    ↓
TxtAutoStatus aktualisieren: "Zuletzt: {Überschrift} — {HH:mm}"
```

Der Watcher-Callback läuft auf einem Worker-Thread. Der `EinfuegenAnCursor`-Aufruf und die UI-Aktualisierung werden via `Dispatcher.BeginInvoke` auf den UI-Thread geleitet (COM-Interop erfordert STA).

---

## Watcher-Lifecycle

| Ereignis | Aktion |
|----------|--------|
| Toggle eingeschaltet | Watcher starten (falls Projekt geladen) |
| Toggle ausgeschaltet | Watcher stoppen |
| Projekt wechselt (ProjektGeändert) | Watcher stoppen, bei Toggle=an neu starten auf neuen Pfad |
| Panel `Bereinigen()` | Watcher stoppen + disposen |

Der Watcher läuft **unabhängig von der Panel-Sichtbarkeit** — Auto-Einfügen funktioniert auch wenn der Nutzer gerade auf einem anderen Tab arbeitet.

---

## Fehlerbehandlung

| Situation | Verhalten |
|-----------|-----------|
| Word nicht offen | `TxtAutoStatus`: "Fehlgeschlagen: Word nicht verbunden" |
| Kein Dokument offen | `TxtAutoStatus`: "Fehlgeschlagen: Kein Dokument offen" |
| JSON fehlt | Leere Überschrift/Maßstab — Bild wird trotzdem eingefügt |
| PNG noch gesperrt nach 300ms | Exception → Status "Fehlgeschlagen: {Meldung}" |
| Projekt wechselt während Einfügen | Watcher bereits gestoppt, laufendes Einfügen läuft zu Ende |

Kein Crash, kein Popup, kein Fokuswechsel in allen Fehlerfällen.

---

## Einstellungen

Neue Property in `Einstellungen.cs`:

```csharp
public bool WordExportAutoEinfuegen { get; set; } = false;
```

Standard: `false` (opt-in, nicht standardmäßig aktiv).

---

## Abgrenzung (nicht im Scope)

- Kein Warteschlangen-Mechanismus für mehrere gleichzeitige Exporte
- Kein Undo für automatisch eingefügte Grafiken (Word-Undo bleibt nutzbar)
- Keine Benachrichtigung außerhalb des Panels (kein Systemtray, kein Toast)
