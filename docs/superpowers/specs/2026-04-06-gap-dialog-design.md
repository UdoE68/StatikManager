# Gap-Dialog: Lösch-Varianten im BlockEditorPrototype

**Datum:** 2026-04-06
**Branch:** feature/word-export-next
**Betroffene Datei:** `src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml(.cs)`

---

## Ziel

Wenn der Nutzer einen Block löscht, soll ein Dialog erscheinen mit drei Varianten für den entstehenden Lückenabstand. Bestehende Lücken sollen per Rechtsklick-Kontextmenü erneut bearbeitet werden können.

---

## Datenmodell

`ProtoBlock` wird um zwei Felder erweitert (nur relevant wenn `IsDeleted == true`):

```csharp
public enum GapModus { OriginalAbstand, KundenAbstand, KeinAbstand }

public GapModus GapModus { get; set; } = GapModus.OriginalAbstand;
public double   GapMm    { get; set; } = 0.0;   // nur für KundenAbstand
```

---

## GapDialog (neues Fenster)

Datei: `GapDialog.xaml` + `GapDialog.xaml.cs`

### Inhalt

- **Variante A** (Radio): „Originalabstand behalten"
  → Lücke behält die ursprüngliche Pixelhöhe des gelöschten Blocks
- **Variante B** (Radio): „Abstand festlegen:" + TextBox in mm
  → TextBox nur aktiv wenn B gewählt; Eingabe wird auf ≥ 0 validiert
- **Variante C** (Radio): „Kein Abstand (0 mm)"
  → Lücke ist physisch vorhanden aber hat Höhe 0 → nahtloser Übergang

Buttons: **OK** (IsDefault) | **Abbrechen** (IsCancel)

### Rückgabe

```csharp
public bool      Bestätigt { get; }
public GapModus  GewählterModus { get; }
public double    EingabeGapMm  { get; }   // 0.0 wenn nicht B
```

---

## Lücken-Rendering in RenderBlocks()

Gelöschte Blöcke werden nicht mehr übersprungen, sondern als **visueller Platzhalter** gerendert:

- **Höhenberechnung:**
  - A: `(FracBottom - FracTop) * srcH`  (Original-Pixelhöhe)
  - B: `GapMm * bitmap.DpiY / 25.4`
  - C: `0.0` → kein Element gerendert

- **Aussehen** (wenn Höhe > 0):
  - Hellgrauer Hintergrund (`#E8E8E8`) mit gestricheltem Rand
  - Label: `↕ [Originalabstand]` / `↕ 12,5 mm` / keine Anzeige bei C
  - Kein Schatten-Effekt (visuell schwächer als echte Inhaltsblöcke)

- **Interaktion:**
  - Linksklick: selektiert die Lücke (zeigt sie in der BlockList)
  - Rechtsklick: ContextMenu mit Eintrag „Abstand bearbeiten …"
    → öffnet GapDialog vorausgefüllt mit aktuellem GapModus/GapMm
    → bei OK: ProtoBlock aktualisieren + RenderBlocks()

---

## Änderungen am bestehenden Code

| Stelle | Änderung |
|--------|----------|
| `BtnDelete_Click` | Statt `DeleteBlock()` direkt → erst GapDialog öffnen, bei OK dann `DeleteBlock()` + Modus setzen |
| `DeleteBlock()` | Nimmt zusätzlich `GapModus` + `GapMm` entgegen |
| `RenderBlocks()` | Lücken-Rendering für `IsDeleted == true` mit Höhe > 0 |
| `RefreshBlockList()` | ToString() zeigt Modus: `[0mm]`, `[orig]`, `[12.5mm]` |

---

## Nicht im Scope

- Persistenz der Lückeneinstellungen (kein Speichern/Laden)
- Undo/Redo
- Übertragung in das ReflowModel (`ContentBlock`) — folgt in separatem Schritt
