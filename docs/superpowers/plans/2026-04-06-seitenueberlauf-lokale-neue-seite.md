# Seitenüberlauf mit lokaler neuer Seite — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wenn gelöschte Blöcke mit Lücken dazu führen, dass verbleibende Blöcke nicht mehr auf die Seitenhöhe passen, wird genau eine neue Überlauf-Seite direkt nach der betroffenen Quellseite eingefügt — ohne Domino-Effekt auf andere Seiten.

**Architecture:** Pro-Seiten-Reflow: jede Quellseite berechnet ihren eigenen Output unabhängig. Ergebnis wird in `_seitenOutput: Dictionary<int, List<OutputPage>>` gehalten. `ZeicheCanvas` liest `_seitenOutput` zur Y-Positionierung; `ZeicheSeiteAlsBlöcke` rendert Überlauf-Blöcke auf einer visuellen Überlauf-Seite direkt darunter.

**Tech Stack:** C# .NET Framework 4.8, WPF/XAML, `ReflowModel.cs`, `PdfSchnittEditor.xaml.cs`

---

## Dateien

| Aktion   | Datei | Was |
|----------|-------|-----|
| Modify   | `src/StatikManager/Modules/Werkzeuge/ReflowModel.cs` | `OutputPage`: 2 neue Properties; `NeueSeite` Overload |
| Modify   | `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` | `_seitenOutput` Feld; `RunReflowFürSeite()`; `BerechneSeitenOutput()`; `IstLetzterAktiverBlock()`; `LöscheAusgewählteParts()` Regel 2; `ZeicheCanvas()` Y-Layout; `ZeicheSeiteAlsBlöcke()` Überlauf-Rendering; 0mm-Bugfix |

---

## Task 1: Bugfix — 0mm Gap Rechtsklick

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` (ZeicheSeiteAlsBlöcke, ~Zeile 1163)

**Problem:** Wenn `GapArt == KeinAbstand` ist `gapH == 0`. Der `if (gapH > 0)`-Guard verhindert die Erstellung des Placeholder-Elements → kein Rechtsklick möglich.

- [ ] **Step 1: Abschnitt in ZeicheSeiteAlsBlöcke lesen und verstehen**

  Lese Zeilen 1163–1204 in `PdfSchnittEditor.xaml.cs`. Ziel: den `if (gapH > 0)`-Block verstehen.

- [ ] **Step 2: Placeholder-Erstellung aus dem gapH>0-Guard herauslösen**

  Ersetze den gesamten `if (block.IsDeleted)`-Block (Zeilen 1163–1205) mit dieser Version:

  ```csharp
  // Gelöschter Block → Lücken-Platzhalter mit konfigurierter Höhe
  if (block.IsDeleted)
  {
      double gapH = BerechneGapHöhe(block, sourceBmp.DpiY, originalDisplayH);
      double blockY = currentY;
      currentY += gapH; // bei KeinAbstand = 0, currentY wächst nicht
      int capturedBlockId = block.BlockId;

      // Immer einen Placeholder erstellen — bei gapH==0 als unsichtbare Klickfläche (4px)
      double renderH = gapH > 0 ? gapH : 4.0;
      var placeholder = new Border
      {
          Width           = bmpPixelW,
          Height          = renderH,
          Background      = gapH > 0
                              ? new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8))
                              : Brushes.Transparent,
          BorderBrush     = gapH > 0
                              ? new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0))
                              : Brushes.Transparent,
          BorderThickness = new Thickness(gapH > 0 ? 1 : 0),
      };
      if (gapH > 0)
      {
          placeholder.Child = new TextBlock
          {
              Text                = block.GapArt == GapModus.KundenAbstand
                                      ? $"↕  {block.GapMm:F1} mm"
                                      : "↕  Originalgröße",
              Foreground          = new SolidColorBrush(Color.FromRgb(0xA0, 0xA0, 0xA0)),
              FontStyle           = FontStyles.Italic,
              FontSize            = 11,
              HorizontalAlignment = HorizontalAlignment.Center,
              VerticalAlignment   = VerticalAlignment.Center
          };
      }

      var cm = new ContextMenu();
      var mi = new MenuItem { Header = "↕  Abstand bearbeiten …" };
      mi.Click += (_, __) => BearbeiteBlockGap(capturedBlockId);
      cm.Items.Add(mi);
      placeholder.ContextMenu = cm;

      Canvas.SetLeft(placeholder, setX);
      Canvas.SetTop(placeholder,  blockY);
      PdfCanvas.Children.Add(placeholder);
      continue;
  }
  ```

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "fix: 0mm Gap-Placeholder bekommt 4px Klickfläche für Rechtsklick"
  ```

---

## Task 2: OutputPage Datenmodell erweitern

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/ReflowModel.cs` (Klassen OutputPage und ReflowEngine)

- [ ] **Step 1: Zwei Properties zu OutputPage hinzufügen**

  In `ReflowModel.cs`, in der Klasse `OutputPage` (nach `WidthPx`), diese zwei Properties einfügen:

  ```csharp
  /// <summary>Index der Quell-PDF-Seite, aus der diese OutputPage entstammt. -1 = unbekannt.</summary>
  public int SourcePageIdx { get; set; } = -1;

  /// <summary>True wenn diese Seite durch Überlauf der vorigen Quellseite entstanden ist.</summary>
  public bool IsOverflowPage { get; set; }
  ```

- [ ] **Step 2: NeueSeite-Overload hinzufügen**

  In `ReflowEngine` (direkt nach der bestehenden `private static OutputPage NeueSeite(double maxH, double w)` Methode), diesen Overload einfügen:

  ```csharp
  private static OutputPage NeueSeite(double maxH, double w, int srcIdx, bool isOverflow)
      => new OutputPage { MaxHeightPx = maxH, WidthPx = w, SourcePageIdx = srcIdx, IsOverflowPage = isOverflow };
  ```

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/ReflowModel.cs
  git commit -m "feat: OutputPage bekommt SourcePageIdx und IsOverflowPage"
  ```

---

## Task 3: RunReflowFürSeite — Pro-Seiten-Reflow-Methode

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`

Diese Methode berechnet das Layout einer einzelnen Quellseite (1 oder 2 OutputPages) und berücksichtigt Gap-Höhen bei der Overflow-Erkennung.

- [ ] **Step 1: Methode in PdfSchnittEditor einfügen**

  In `PdfSchnittEditor.xaml.cs`, direkt nach `BerechneGapHöhe` (nach Zeile ~1329), diese private Methode einfügen:

  ```csharp
  /// <summary>
  /// Berechnet das Layout einer einzelnen Quellseite. Berücksichtigt Gap-Höhen bei der
  /// Überlauf-Erkennung. Gibt 1 oder 2 OutputPages zurück (Original + optional Überlauf).
  /// </summary>
  private List<OutputPage> RunReflowFürSeite(int sourcePageIdx)
  {
      if (_contentBlocks == null || _seitenBilder == null) return new List<OutputPage>();

      double pageMaxH = sourcePageIdx < _seitenHöhe.Length
          ? Math.Max(1.0, _seitenHöhe[sourcePageIdx])
          : 1.0;
      double pageW = sourcePageIdx < _seitenBilder.Count
          ? Math.Max(1.0, (double)_seitenBilder[sourcePageIdx].PixelWidth)
          : 1.0;
      double dpiY = (sourcePageIdx < _seitenBilder.Count && _seitenBilder[sourcePageIdx] != null)
          ? _seitenBilder[sourcePageIdx].DpiY
          : 96.0;

      var page1 = new OutputPage { MaxHeightPx = pageMaxH, WidthPx = pageW, SourcePageIdx = sourcePageIdx, IsOverflowPage = false };
      var result = new List<OutputPage> { page1 };

      var blöckeDeserSeite = _contentBlocks
          .Where(b => b.SourcePageIdx == sourcePageIdx)
          .ToList();

      if (blöckeDeserSeite.Count == 0) return result;

      var aktuelleSeite = page1;
      double currentY    = 0.0;
      bool   hatÜberlauf = false;

      foreach (var block in blöckeDeserSeite)
      {
          if (block.IsDeleted)
          {
              if (!hatÜberlauf)
              {
                  // Gap-Höhe trägt zu currentY auf Seite 1 bei
                  double gapH = BerechneGapHöhe(block, dpiY, (block.FracUnten - block.FracOben) * pageMaxH);
                  currentY += gapH;
              }
              // Nach Überlauf: Gaps bleiben auf Seite 1, beeinflussen Überlauf-Y nicht
              continue;
          }

          double blockH = block.ContentHeightPx(pageMaxH);
          if (blockH <= 0.0) continue;

          // Würde dieser Block die Seite überschreiten?
          if (!hatÜberlauf && currentY + blockH > pageMaxH)
          {
              // Überlauf-Seite erstellen (maximal eine pro Quellseite, Regel 4)
              var overflow = new OutputPage
              {
                  MaxHeightPx  = pageMaxH,
                  WidthPx      = pageW,
                  SourcePageIdx = sourcePageIdx,
                  IsOverflowPage = true
              };
              result.Add(overflow);
              aktuelleSeite = overflow;
              currentY      = 0.0;
              hatÜberlauf   = true;
          }

          aktuelleSeite.Blocks.Add(new PlacedBlock
          {
              Block        = block,
              YOffset      = currentY,
              HeightPx     = blockH,
              SrcFracOben  = block.FracOben,
              SrcFracUnten = block.FracUnten
          });
          currentY += blockH;
      }

      return result;
  }
  ```

- [ ] **Step 2: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 3: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "feat: RunReflowFürSeite – pro-Seiten-Overflow-Erkennung mit Gap-Höhen"
  ```

---

## Task 4: BerechneSeitenOutput + _seitenOutput Feld

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs`

- [ ] **Step 1: _seitenOutput Feld in den Feldern-Block einfügen**

  In `PdfSchnittEditor.xaml.cs`, bei den privaten Feldern (dort wo `_seitenHöhe` und `_seitenBilder` deklariert sind, ca. Zeile 55-57), diese Zeile hinzufügen:

  ```csharp
  private Dictionary<int, List<OutputPage>> _seitenOutput = new Dictionary<int, List<OutputPage>>();
  ```

- [ ] **Step 2: BerechneSeitenOutput Methode einfügen**

  Direkt nach `RunReflowFürSeite` (aus Task 3), diese Methode einfügen:

  ```csharp
  /// <summary>
  /// Berechnet _seitenOutput für alle sichtbaren Quellseiten neu.
  /// Muss vor ZeicheCanvas aufgerufen werden.
  /// </summary>
  private void BerechneSeitenOutput()
  {
      _seitenOutput.Clear();
      if (_contentBlocks == null || _seitenBilder == null) return;

      var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
      var sichtbar    = reihenfolge.Where(i => i < _seitenBilder.Count && !_gelöschteSeiten.Contains(i)).ToList();

      foreach (int si in sichtbar)
      {
          _seitenOutput[si] = RunReflowFürSeite(si);
      }
  }
  ```

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "feat: BerechneSeitenOutput + _seitenOutput Feld"
  ```

---

## Task 5: Regel 2 — Kein Gap-Dialog für letzten aktiven Block

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` (LöscheAusgewählteParts ~Zeile 4312)

Regel 2: Wenn der letzte nicht-gelöschte Block einer Seite gelöscht wird → kein Gap-Dialog, Gap wird stumm auf `KeinAbstand` gesetzt.

- [ ] **Step 1: Hilfsmethode IstLetzterAktiverBlock einfügen**

  In `PdfSchnittEditor.xaml.cs`, direkt nach `LöscheAusgewählteParts` (nach Zeile ~4352), diese Methode einfügen:

  ```csharp
  /// <summary>
  /// Gibt true zurück wenn der t-te Bitmap-Block von Seite si der einzige
  /// nicht-gelöschte Bitmap-Block auf dieser Seite ist.
  /// Wird für Regel 2 verwendet: kein Gap-Dialog für letzten aktiven Block.
  /// </summary>
  private bool IstLetzterAktiverBlock(int si, int t)
  {
      if (_contentBlocks == null) return false;

      // Alle nicht-gelöschten Bitmap-Blöcke der Seite zählen
      int aktiveGesamt = _contentBlocks.Count(
          b => b.SourcePageIdx == si && !b.IsLeerzeile && !b.IsDeleted);

      if (aktiveGesamt != 1) return false;

      // Prüfen ob der t-te Bitmap-Block genau dieser eine aktive Block ist
      int bitmapCount = 0;
      foreach (var b in _contentBlocks)
      {
          if (b.SourcePageIdx != si || b.IsLeerzeile) continue;
          if (bitmapCount == t) return !b.IsDeleted;
          bitmapCount++;
      }
      return false;
  }
  ```

- [ ] **Step 2: LöscheAusgewählteParts anpassen**

  In `LöscheAusgewählteParts` (Zeile ~4324), den Block der Gap-Dialog-Logik so ersetzen:

  Aktueller Code (Zeilen 4318–4330):
  ```csharp
  // GapDialog nur zeigen wenn Schnitte vorhanden (sonst kein sinnvoller Lücken-Kontext)
  bool hatSchnitte = _ausgewählteParts.Any(p => GetTeilGrenzen(p.Seite).Count > 1);

  GapModus gewählterModus = GapModus.OriginalAbstand;
  double   eingabeMm     = 0.0;

  if (hatSchnitte)
  {
      var dlg = new GapDialog { Owner = Window.GetWindow(this) };
      if (dlg.ShowDialog() != true) return;
      gewählterModus = dlg.GewählterModus;
      eingabeMm      = dlg.EingabeGapMm;
  }
  ```

  Neuer Code:
  ```csharp
  // GapDialog nur zeigen wenn:
  // (a) Schnitte vorhanden UND
  // (b) nicht alle ausgewählten Teile der letzte aktive Block ihrer Seite sind (Regel 2)
  bool hatSchnitte     = _ausgewählteParts.Any(p => GetTeilGrenzen(p.Seite).Count > 1);
  bool sindLetzteBlöcke = _ausgewählteParts.All(p => IstLetzterAktiverBlock(p.Seite, p.Teil));

  GapModus gewählterModus = GapModus.KeinAbstand;  // Default für letzten Block: kein Abstand
  double   eingabeMm     = 0.0;

  if (hatSchnitte && !sindLetzteBlöcke)
  {
      gewählterModus = GapModus.OriginalAbstand;    // Reset auf sinnvollen Default für Dialog
      var dlg = new GapDialog { Owner = Window.GetWindow(this) };
      if (dlg.ShowDialog() != true) return;
      gewählterModus = dlg.GewählterModus;
      eingabeMm      = dlg.EingabeGapMm;
  }
  ```

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "feat: Regel 2 – kein Gap-Dialog wenn letzter aktiver Block gelöscht wird"
  ```

---

## Task 6: ZeicheCanvas — Y-Layout mit Überlauf-Seiten

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` (ZeicheCanvas ~Zeile 824)

**Ziel:** `BerechneSeitenOutput()` aufrufen und Y-Positionen der Quellseiten so berechnen, dass Überlauf-Seiten Platz finden.

- [ ] **Step 1: Vertikales Layout mit Überlauf-Extra-Höhe vorbereiten**

  Hinweis: `BerechneSeitenOutput()` wird in Step 2 direkt im Code-Block integriert (nach `BerechneLayoutStatic`). Kein separater Aufruf nötig.

- [ ] **Step 2: Vertikales Layout mit Überlauf-Extra-Höhe**

  Im vertikalen Layout-Zweig von `ZeicheCanvas()` (Zeilen 869–883, der `else`-Zweig nach `if (_layoutHorizontal)`), den Block so ersetzen:

  Aktueller Code:
  ```csharp
  else
  {
      BerechneLayoutStatic(sichtbarBmps, SeitenAbstand, out var ordYStart, out var ordHöhe);
      for (int di = 0; di < sichtbar.Count; di++)
      {
          int oi = sichtbar[di];
          _seitenYStart[oi] = ordYStart[di];
          _seitenHöhe[oi]   = ordHöhe[di];
      }
      int lastOi = sichtbar[sichtbar.Count - 1];
      double gesamtH = _seitenYStart[lastOi] + _seitenHöhe[lastOi] + SeitenAbstand;
      double maxBmpW = sichtbarBmps.Max(b => (double)b.PixelWidth);
      PdfCanvas.Width  = maxBmpW + SeiteX * 2;
      PdfCanvas.Height = Math.Max(gesamtH, 1);
  }
  ```

  Neuer Code:
  ```csharp
  else
  {
      // Basis-Layout ohne Überlauf
      BerechneLayoutStatic(sichtbarBmps, SeitenAbstand, out var ordYStart, out var ordHöhe);
      for (int di = 0; di < sichtbar.Count; di++)
      {
          int oi = sichtbar[di];
          _seitenYStart[oi] = ordYStart[di];
          _seitenHöhe[oi]   = ordHöhe[di];
      }

      BerechneSeitenOutput(); // nach BerechneLayoutStatic – _seitenHöhe ist jetzt befüllt

      // Überlauf-Verschiebung: Seiten nach einer Überlauf-Quelle nach unten verschieben
      double extraYAkkumuliert = 0.0;
      for (int di = 0; di < sichtbar.Count; di++)
      {
          int oi = sichtbar[di];
          _seitenYStart[oi] += extraYAkkumuliert;

          // Hat diese Quellseite eine Überlauf-Seite?
          if (_seitenOutput.TryGetValue(oi, out var outPages) && outPages.Count > 1)
          {
              // Platz für Überlauf-Seite: gleiche Höhe wie Quellseite + Abstand
              extraYAkkumuliert += _seitenHöhe[oi] + SeitenAbstand;
          }
      }

      int lastOi2  = sichtbar[sichtbar.Count - 1];
      double gesamtH = _seitenYStart[lastOi2] + _seitenHöhe[lastOi2] + SeitenAbstand;
      // Extra: falls die letzte Seite selbst überläuft
      if (_seitenOutput.TryGetValue(lastOi2, out var lastOut) && lastOut.Count > 1)
          gesamtH += _seitenHöhe[lastOi2] + SeitenAbstand;

      double maxBmpW = sichtbarBmps.Max(b => (double)b.PixelWidth);
      PdfCanvas.Width  = maxBmpW + SeiteX * 2;
      PdfCanvas.Height = Math.Max(gesamtH, 1);
  }
  ```

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "feat: ZeicheCanvas – Y-Layout reserviert Platz für Überlauf-Seiten"
  ```

---

## Task 7: ZeicheSeiteAlsBlöcke — Überlauf-Seite visuell rendern

**Dateien:**
- Modify: `src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs` (ZeicheSeiteAlsBlöcke ~Zeile 1127)

**Ziel:** Nicht-gelöschte Blöcke, die laut `_seitenOutput` auf der Überlauf-Seite liegen, werden auf einer neuen visuellen Seite direkt unter der Quellseite gerendert.

- [ ] **Step 1: Hilfsmethode ZeicheÜberlaufSeite einfügen**

  Direkt nach `ZeicheSeiteAlsBlöcke` (nach Zeile ~1308), diese Methode einfügen:

  ```csharp
  /// <summary>
  /// Zeichnet den weißen Seiten-Hintergrund für die Überlauf-Seite einer Quellseite.
  /// </summary>
  private void ZeicheÜberlaufSeitenHintergrund(int seitenIdx, double yBase, double w, double h)
  {
      DropShadowEffect? shadow = null;
      try
      {
          shadow = new DropShadowEffect
          {
              BlurRadius = 20, ShadowDepth = 7,
              Direction  = 280, Color = Colors.Black, Opacity = 0.85
          };
      }
      catch { /* ohne Schatten */ }

      var blatt = new Border
      {
          Tag             = $"SEITE_OVERFLOW_{seitenIdx}",
          Width           = w,
          Height          = h,
          Background      = Brushes.White,
          BorderBrush     = new SolidColorBrush(Color.FromRgb(180, 160, 160)),
          BorderThickness = new Thickness(2),
          Effect          = shadow
      };
      Canvas.SetLeft(blatt, _layoutHorizontal ? _seitenXStart[seitenIdx] : SeiteX);
      Canvas.SetTop(blatt,  yBase);
      PdfCanvas.Children.Add(blatt);
  }
  ```

- [ ] **Step 2: ZeicheSeiteAlsBlöcke für Überlauf erweitern**

  In `ZeicheSeiteAlsBlöcke`, direkt nach der Zeile `bool ersterBlock = true;` (ca. Zeile 1148) und **vor** der `foreach`-Schleife, diesen Setup-Code einfügen:

  ```csharp
  // Überlauf-Blöcke ermitteln
  bool hatÜberlauf = _seitenOutput.TryGetValue(seitenIdx, out var outputSeiten) && outputSeiten.Count > 1;
  var überflussIds = hatÜberlauf
      ? new HashSet<int>(outputSeiten[1].Blocks.Select(pb => pb.Block.BlockId))
      : new HashSet<int>();

  double overflowBaseY = yBase + pageH + SeitenAbstand;
  double overflowCurrentY = 0.0;

  if (hatÜberlauf)
      ZeicheÜberlaufSeitenHintergrund(seitenIdx, overflowBaseY, bmpPixelW, pageH);
  ```

  Dann in der `foreach`-Schleife, den Teil für nicht-gelöschte Blöcke (nach `double blockDisplayH = ...`):

  Aktueller Code (ab Zeile ~1207 bis ~1305):
  ```csharp
  double blockDisplayH = Math.Max(1, originalDisplayH);
  double blockY2       = currentY;
  currentY            += blockDisplayH;
  ```

  Neuer Code:
  ```csharp
  double blockDisplayH = Math.Max(1, originalDisplayH);
  bool   istÜberlauf   = überflussIds.Contains(block.BlockId);
  double blockY2;

  if (istÜberlauf)
  {
      blockY2          = overflowBaseY + overflowCurrentY;
      overflowCurrentY += blockDisplayH;
      // currentY auf Seite 1 wächst nicht — Überlauf-Block erscheint nicht dort
  }
  else
  {
      blockY2   = currentY;
      currentY += blockDisplayH;
  }
  ```

  Hinweis: Die Variable `blockY2` wird bereits weiter unten für `Canvas.SetTop(img, blockY2)` verwendet — diese Zeile bleibt unverändert.

- [ ] **Step 3: Build**

  ```powershell
  & 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1
  ```
  Erwartetes Ergebnis: `0 Error(s)`

- [ ] **Step 4: Commit**

  ```bash
  git add src/StatikManager/Modules/Werkzeuge/PdfSchnittEditor.xaml.cs
  git commit -m "feat: ZeicheSeiteAlsBlöcke rendert Überlauf-Blöcke auf eigener visueller Seite"
  ```

---

## Task 8: Manuelle Verifikation

- [ ] **Step 1: App starten**

  ```powershell
  taskkill /f /im StatikManager.exe 2>nul
  Start-Sleep -Seconds 1
  & 'C:\KI\StatikManager_V1\src\StatikManager\Start_Debug.bat'
  ```

- [ ] **Step 2: Überlauf-Test**

  1. PDF mit geschnittenen Blöcken (mind. 3 Teile auf einer Seite) öffnen
  2. Mittleren Block löschen, Gap auf großen mm-Wert setzen (z.B. 200mm)
  3. **Erwartet:** Letzter Block erscheint auf einer neuen visuellen Seite direkt darunter
  4. **Erwartet:** Seiten dahinter unverändert
  5. Gap verkleinern bis kein Überlauf mehr → **Erwartet:** Überlauf-Seite verschwindet

- [ ] **Step 3: Regel-2-Test**

  1. Seite mit 2 Teilen öffnen
  2. Ersten Teil löschen (gap dialog erscheint normal) — PASS prüfen
  3. Zweiten Teil löschen (= letzter aktiver Block) → **Erwartet:** kein Gap-Dialog

- [ ] **Step 4: 0mm-Bugfix-Test**

  1. Block löschen, im Gap-Dialog "Kein Abstand" wählen
  2. Rechtsklick auf die (unsichtbare) Fläche wo der Block war
  3. **Erwartet:** Kontextmenü "Abstand bearbeiten …" erscheint

- [ ] **Step 5: Abschluss-Commit**

  Nach erfolgreichem Test:
  ```bash
  git push
  ```
