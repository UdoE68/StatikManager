# Design: Seitenüberlauf mit lokaler neuer Seite

**Datum:** 2026-04-06
**Branch:** feature/word-export-next
**Status:** Approved

---

## Ziel

Wenn durch Lücken (gelöschte Blöcke) die verbleibenden Blöcke einer Seite nicht mehr auf die Seitenhöhe passen, wird direkt nach dieser Seite eine neue Seite eingefügt. Keine anderen Seiten werden verändert (kein Domino-Effekt).

---

## Regeln

### Regel 1 — Feste Seitengröße
Die Seitenhöhe (von der Quell-PDF übernommen) ist konstant. Sie verändert sich nie. `OutputPage.MaxHeightPx` bleibt unverändert.

### Regel 2 — Unterster Block gelöscht
Wenn der letzte nicht-gelöschte Block einer Seite gelöscht wird, erscheint kein Gap-Dialog. Der entstehende Leerraum füllt stumm den verbleibenden Platz bis zum Seitenende. Kein Gap-Placeholder wird gerendert.

### Regel 3 — Überlauf → neue Seite direkt dahinter
Wenn nach Abzug der Lücken die verbleibenden Blöcke nicht mehr auf die Seitenhöhe passen, werden so viele Blöcke wie nötig auf eine neue Seite direkt nach der aktuellen Seite geschoben. Die neue Seite ist eine `OutputPage` mit `IsOverflowPage = true`.

### Regel 4 — Kein Domino-Effekt
Jede Quellseite wird unabhängig berechnet. Der Überlauf einer Seite erzeugt maximal eine neue Seite direkt danach. Alle anderen Seiten bleiben vollständig unverändert.

### Regel 5 — Reaktive Berechnung
Der Überlauf wird nach jeder Änderung (Block löschen, Gap ändern) sofort neu berechnet. Wenn eine Lücke verkleinert wird und der Überlauf wegfällt, verschwindet die Überlauf-Seite automatisch wieder.

---

## Beispiel

```
Ausgangszustand:
  Seite 1: Block A, Block B, Block C

Aktion: B wird gelöscht mit 80mm Lücke
  → A + 80mm Lücke + C > Seitenhöhe
  → C rutscht auf neue Seite 1b (direkt nach Seite 1)
  → Seite 2, 3, 4: unverändert

Aktion: Gap B von 80mm auf 10mm verkleinert
  → A + 10mm Lücke + C ≤ Seitenhöhe
  → Seite 1b verschwindet, C kehrt auf Seite 1 zurück
```

---

## Architektur

### Ansatz: Pro-Seiten-Reflow

Statt eines globalen Reflow-Passes über alle Blöcke berechnet jede Quellseite ihren eigenen Output unabhängig.

#### Warum dieser Ansatz
- Entspricht exakt Regel 4 (kein Domino)
- Isoliert und testbar
- `RunReflow()` ist bereits pure/stateless — Umbau ist überschaubar
- Kein Freeze-Flag-Chaos, keine parallelen Modelle

---

## Datenmodell-Änderungen

### `OutputPage` — 2 neue Properties
```csharp
int SourcePageIdx   // Welche Quellseite hat diese OutputPage erzeugt
bool IsOverflowPage // true = durch Überlauf entstanden
```

### Neue Hilfsstruktur in PdfSchnittEditor
```csharp
Dictionary<int, List<OutputPage>> _seitenOutput
// Key = SourcePageIdx → Value = 1 oder 2 OutputPages
```

`ContentBlock`, `PlacedBlock` bleiben unverändert.

---

## Neue Methoden

### `RunReflowFürSeite()`
```
Signatur:
  List<OutputPage> RunReflowFürSeite(
      int sourcePageIdx,
      IReadOnlyList<ContentBlock> alleBlöcke,
      double pageHeightPx,
      double pageWidthPx)

Logik:
  1. Filtere Blöcke: SourcePageIdx == sourcePageIdx && !IsDeleted
  2. Berechne Höhen inkl. Gaps
  3. Prüfe Regel 2: Ist letzter nicht-gelöschter Block = physisch letzter?
     → ja: kein Gap-Placeholder für ihn, Leerraum stumm bis Seitenende
  4. Fülle OutputPage 1 (gleiche MinSplitHeightPx-Logik wie heute)
  5. Blöcke übrig → OutputPage 2 (IsOverflowPage=true) mit restlichen Blöcken
  6. Gib 1 oder 2 OutputPages zurück
```

### `BaueReflowResult()`
```
Logik:
  für jede SourcePageIdx in Reihenfolge:
      _seitenOutput[idx] = RunReflowFürSeite(idx, ...)
  ReflowResult.Pages = _seitenOutput.Values.SelectMany(x => x).ToList()
```

### Nach jeder Änderung
```
Nur RunReflowFürSeite() für die betroffene Seite aufrufen
→ _seitenOutput[idx] aktualisieren
→ Canvas neu zeichnen
```

---

## Gap-Dialog-Logik (Regel 2)

Beim Löschen eines Blocks wird vor dem Öffnen des Gap-Dialogs geprüft:

```
Ist der zu löschende Block der letzte nicht-gelöschte Block auf seiner Seite?
  → ja: kein Gap-Dialog, Block direkt löschen (GapArt = KeinAbstand)
  → nein: Gap-Dialog wie bisher
```

---

## Begleitender Bugfix

**Problem:** Wenn GapArt = KeinAbstand (Gap = 0mm), funktioniert der Rechtsklick auf den Gap-Placeholder nicht.

**Ursache:** Der Placeholder hat bei 0px Höhe keinen trefferbasierten Bereich für Mausereignisse.

**Fix:** Gap-Placeholder mit GapArt=KeinAbstand bekommt eine Mindest-Klickfläche (z.B. 4px unsichtbare Trefferfläche) auch wenn die visuelle Höhe 0 ist.

---

## Was nicht geändert wird

- `ContentBlock`, `PlacedBlock` — keine Änderungen
- Bestehende Gap-Dialog-Logik für nicht-unterste Blöcke — unverändert
- Undo-Stack — bestehende Mechanik bleibt
- Export/Word-Logik — nicht betroffen
- Kein automatisches Verschmelzen von Leerblöcken

---

## Erfolgskriterien

1. Block löschen mit großer Lücke → Überlauf-Seite erscheint sofort
2. Lücke verkleinern → Überlauf-Seite verschwindet, Block kehrt zurück
3. Seiten hinter der betroffenen Seite bleiben unverändert
4. Letzter Block gelöscht → kein Gap-Dialog, kein Placeholder
5. Rechtsklick auf 0mm-Gap funktioniert
