# Gap-Dialog: Lösch-Varianten im BlockEditorPrototype — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Beim Löschen eines Blocks erscheint ein Dialog mit drei Lücken-Varianten (Originalabstand / mm-Eingabe / 0 mm); Lücken werden visuell auf dem Canvas dargestellt und können per Rechtsklick-Kontextmenü nachbearbeitet werden.

**Architecture:** `GapModus`-Enum wird auf Namespace-Ebene deklariert. `ProtoBlock` bekommt zwei neue Felder. Ein neues `GapDialog`-Fenster kapselt die drei Varianten. `RenderBlocks()` rendert gelöschte Blöcke als sichtbare Lücken-Platzhalter mit Kontextmenü.

**Tech Stack:** WPF, C# .NET Framework 4.8, MSBuild x64 Debug

---

## Dateiübersicht

| Aktion   | Datei                                                                      | Verantwortung                          |
|----------|----------------------------------------------------------------------------|----------------------------------------|
| Modify   | `src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs`         | ProtoBlock-Modell, Rendering, Events   |
| Create   | `src/StatikManager/Modules/Werkzeuge/GapDialog.xaml`                       | Dialog-UI (3 Radio-Buttons + TextBox)  |
| Create   | `src/StatikManager/Modules/Werkzeuge/GapDialog.xaml.cs`                    | Dialog-Logik, Validierung, Rückgabe    |

---

## Task 1: GapModus-Enum + ProtoBlock-Felder

**Files:**
- Modify: `src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs`

- [ ] **Schritt 1: `GapModus`-Enum auf Namespace-Ebene einfügen**

Direkt vor `public partial class BlockEditorPrototype` (nach den `using`-Statements) einfügen:

```csharp
namespace StatikManager.Modules.Werkzeuge
{
    public enum GapModus { OriginalAbstand, KundenAbstand, KeinAbstand }

    public partial class BlockEditorPrototype : Window
    {
```

- [ ] **Schritt 2: Zwei neue Felder in `ProtoBlock` ergänzen**

In der `sealed class ProtoBlock` nach `public bool IsDeleted { get; set; }` einfügen:

```csharp
/// <summary>Wie groß die Lücke nach dem Löschen sein soll (GapModus). Nur relevant wenn IsDeleted == true.</summary>
/// <remarks>Heißt GapArt (nicht GapModus) um Naming Collision mit dem Enum-Typ zu vermeiden.</remarks>
public GapModus GapArt { get; set; } = GapModus.OriginalAbstand;

/// <summary>Lückengröße in mm. Nur relevant für GapModus.KundenAbstand. Muss >= 0 sein.</summary>
public double GapMm { get; set; } = 0.0;
```

- [ ] **Schritt 3: `ToString()` in `ProtoBlock` aktualisieren**

Bestehende `ToString()`-Methode ersetzen:

```csharp
public override string ToString()
{
    if (IsDeleted)
    {
        string gap = GapArt == GapModus.OriginalAbstand ? "[orig]"
                   : GapArt == GapModus.KundenAbstand   ? $"[{GapMm:F1}mm]"
                   : "[0mm]";
        return $"[DEL {gap}]  B{Id}  {FracTop:F4} – {FracBottom:F4}";
    }
    return $"       B{Id}  {FracTop:F4} – {FracBottom:F4}";
}
```

- [ ] **Schritt 4: Build — prüfen ob 0 Fehler**

```powershell
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"
```

Erwartet: `Build succeeded.` / `0 Error(s)`

- [ ] **Schritt 5: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs
git commit -m "feat: GapModus-Enum + ProtoBlock-Felder für Lückenabstand"
```

---

## Task 2: GapDialog XAML

**Files:**
- Create: `src/StatikManager/Modules/Werkzeuge/GapDialog.xaml`

- [ ] **Schritt 1: `GapDialog.xaml` anlegen**

```xml
<Window x:Class="StatikManager.Modules.Werkzeuge.GapDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Lückenabstand festlegen"
        Width="360" Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False">

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Variante A -->
        <RadioButton x:Name="RbOriginal" Grid.Row="0"
                     Content="Originalabstand behalten"
                     IsChecked="True"
                     Checked="Rb_Checked"/>

        <!-- Variante B -->
        <RadioButton x:Name="RbKunden" Grid.Row="2"
                     Checked="Rb_Checked">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Abstand festlegen:" VerticalAlignment="Center"/>
                <TextBox x:Name="TxtMm" Width="65" Height="22" Margin="8,0,4,0"
                         IsEnabled="False" VerticalContentAlignment="Center"
                         Text="0"/>
                <TextBlock Text="mm" VerticalAlignment="Center"/>
            </StackPanel>
        </RadioButton>

        <!-- Variante C -->
        <RadioButton x:Name="RbKein" Grid.Row="4"
                     Content="Kein Abstand (0 mm — nahtlos)"
                     Checked="Rb_Checked"/>

        <!-- Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal"
                    HorizontalAlignment="Right">
            <Button x:Name="BtnOk" Content="OK"
                    Width="80" Height="26" Margin="0,0,8,0"
                    IsDefault="True"
                    Click="BtnOk_Click"/>
            <Button Content="Abbrechen"
                    Width="90" Height="26"
                    IsCancel="True"
                    Click="BtnAbbrechen_Click"/>
        </StackPanel>
    </Grid>
</Window>
```

- [ ] **Schritt 2: Build — prüfen ob 0 Fehler**

```powershell
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"
```

Erwartet: `Build succeeded.` / `0 Error(s)`

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/GapDialog.xaml
git commit -m "feat: GapDialog XAML — 3 Radio-Buttons + mm-TextBox"
```

---

## Task 3: GapDialog Code-Behind

**Files:**
- Create: `src/StatikManager/Modules/Werkzeuge/GapDialog.xaml.cs`

- [ ] **Schritt 1: `GapDialog.xaml.cs` anlegen**

```csharp
using System.Globalization;
using System.Windows;

namespace StatikManager.Modules.Werkzeuge
{
    public partial class GapDialog : Window
    {
        // ── Rückgabewerte ─────────────────────────────────────────────────────
        public bool     Bestätigt      { get; private set; }
        public GapModus GewählterModus { get; private set; }
        public double   EingabeGapMm   { get; private set; }

        // ── Konstruktor ───────────────────────────────────────────────────────
        /// <summary>
        /// Öffnet den Dialog. Optionale Parameter füllen ihn für "Bearbeiten"-Modus vor.
        /// </summary>
        public GapDialog(GapModus aktuellModus = GapModus.OriginalAbstand,
                         double   aktuellMm   = 0.0)
        {
            InitializeComponent();

            switch (aktuellModus)
            {
                case GapModus.KundenAbstand:
                    RbKunden.IsChecked = true;
                    TxtMm.Text = aktuellMm.ToString("F1", CultureInfo.CurrentCulture);
                    break;
                case GapModus.KeinAbstand:
                    RbKein.IsChecked = true;
                    break;
                default:
                    RbOriginal.IsChecked = true;
                    break;
            }

            AktualisiereTextBox();
        }

        // ── Events ────────────────────────────────────────────────────────────
        private void Rb_Checked(object sender, RoutedEventArgs e)
            => AktualisiereTextBox();

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (RbKunden.IsChecked == true)
            {
                // Komma und Punkt akzeptieren
                string raw = TxtMm.Text.Replace(",", ".");
                if (!double.TryParse(raw, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double mm) || mm < 0)
                {
                    MessageBox.Show("Bitte eine gültige Zahl ≥ 0 eingeben.",
                        "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Warning);
                    TxtMm.Focus();
                    return;
                }
                GewählterModus = GapModus.KundenAbstand;
                EingabeGapMm   = mm;
            }
            else if (RbKein.IsChecked == true)
            {
                GewählterModus = GapModus.KeinAbstand;
                EingabeGapMm   = 0.0;
            }
            else
            {
                GewählterModus = GapModus.OriginalAbstand;
                EingabeGapMm   = 0.0;
            }

            Bestätigt    = true;
            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            Bestätigt    = false;
            DialogResult = false;
        }

        // ── Hilfsmethoden ─────────────────────────────────────────────────────
        private void AktualisiereTextBox()
        {
            if (TxtMm == null) return;
            bool aktivieren = RbKunden.IsChecked == true;
            TxtMm.IsEnabled = aktivieren;
            if (aktivieren) TxtMm.Focus();
        }
    }
}
```

- [ ] **Schritt 2: Build — prüfen ob 0 Fehler**

```powershell
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"
```

Erwartet: `Build succeeded.` / `0 Error(s)`

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/GapDialog.xaml.cs
git commit -m "feat: GapDialog Code-Behind — Validierung + Rückgabewerte"
```

---

## Task 4: BtnDelete_Click + DeleteBlock() anpassen

**Files:**
- Modify: `src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs`

- [ ] **Schritt 1: `DeleteBlock()` Signatur erweitern**

Bestehende Methode ersetzen:

```csharp
/// <summary>
/// Markiert den Block als gelöscht und speichert den gewünschten Lückenabstand.
/// RenderBlocks() stellt die Lücke dann als visuellen Platzhalter dar.
/// </summary>
private void DeleteBlock(int blockId, GapModus modus, double gapMm)
{
    var block = _blocks.FirstOrDefault(b => b.Id == blockId);
    if (block == null) return;
    block.IsDeleted = true;
    block.GapArt    = modus;
    block.GapMm     = gapMm;
    _selectedId     = -1;
    RenderBlocks();
}
```

- [ ] **Schritt 2: `BtnDelete_Click` Dialog öffnen**

Bestehenden Handler ersetzen:

```csharp
private void BtnDelete_Click(object sender, RoutedEventArgs e)
{
    if (_selectedId < 0) return;

    var dlg = new GapDialog { Owner = this };
    if (dlg.ShowDialog() != true) return;

    DeleteBlock(_selectedId, dlg.GewählterModus, dlg.EingabeGapMm);
}
```

- [ ] **Schritt 3: Build — prüfen ob 0 Fehler**

```powershell
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"
```

Erwartet: `Build succeeded.` / `0 Error(s)`

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs
git commit -m "feat: Löschen-Button öffnet GapDialog statt direktem Delete"
```

---

## Task 5: Lücken-Rendering in RenderBlocks()

**Files:**
- Modify: `src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs`

- [ ] **Schritt 1: Hilfsmethode `BerechneLückenHöhePx()` hinzufügen**

Nach `CanvasYToFrac()` einfügen:

```csharp
/// <summary>
/// Berechnet die anzuzeigende Lückenhöhe in Pixeln anhand des GapModus.
/// DPI wird direkt aus _originalBitmap.DpiY ausgelesen.
/// </summary>
private double BerechneLückenHöhePx(ProtoBlock block)
{
    if (_originalBitmap == null) return 0.0;

    switch (block.GapArt)
    {
        case GapModus.OriginalAbstand:
            return (block.FracBottom - block.FracTop) * _originalBitmap.PixelHeight;
        case GapModus.KundenAbstand:
            double dpi = _originalBitmap.DpiY > 0 ? _originalBitmap.DpiY : 96.0;
            return block.GapMm * dpi / 25.4;
        case GapModus.KeinAbstand:
        default:
            return 0.0;
    }
}
```

- [ ] **Schritt 2: Hilfsmethode `RenderLücke()` hinzufügen**

Nach `BerechneLückenHöhePx()` einfügen:

```csharp
/// <summary>
/// Zeichnet einen visuellen Lücken-Platzhalter auf dem Canvas.
/// Enthält ein Kontextmenü zum Nachbearbeiten des Abstands.
/// </summary>
private void RenderLücke(ProtoBlock block, double gapH)
{
    if (_originalBitmap == null) return;

    int    srcW       = _originalBitmap.PixelWidth;
    double displayTop = CanvasPad + block.FracTop * _originalBitmap.PixelHeight;
    int    capturedId = block.Id;

    string label = block.GapArt == GapModus.OriginalAbstand
        ? "↕  Originalabstand"
        : block.GapArt == GapModus.KundenAbstand
            ? $"↕  {block.GapMm:F1} mm"
            : "";

    var gapBorder = new Border
    {
        Tag             = capturedId,
        Width           = srcW,
        Height          = gapH,
        Background      = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
        BorderBrush     = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
        BorderThickness = new Thickness(1),
        Child           = new TextBlock
        {
            Text                = label,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment   = VerticalAlignment.Center,
            Foreground          = new SolidColorBrush(Color.FromRgb(150, 150, 150)),
            FontSize            = 11,
            FontStyle           = FontStyles.Italic
        }
    };

    // Kontextmenü: Abstand nachbearbeiten
    var cm = new ContextMenu();
    var mi = new MenuItem { Header = "Abstand bearbeiten …" };
    mi.Click += (_, __) => BearbeiteGap(capturedId);
    cm.Items.Add(mi);
    gapBorder.ContextMenu = cm;

    Canvas.SetLeft(gapBorder, CanvasPad);
    Canvas.SetTop(gapBorder,  displayTop);
    EditorCanvas.Children.Add(gapBorder);
}

/// <summary>
/// Öffnet den GapDialog vorausgefüllt für den angegebenen gelöschten Block.
/// Wird vom Kontextmenü des Lücken-Platzhalters aufgerufen.
/// </summary>
private void BearbeiteGap(int blockId)
{
    var block = _blocks.FirstOrDefault(b => b.Id == blockId && b.IsDeleted);
    if (block == null) return;

    var dlg = new GapDialog(block.GapArt, block.GapMm) { Owner = this };
    if (dlg.ShowDialog() != true) return;

    block.GapArt = dlg.GewählterModus;
    block.GapMm    = dlg.EingabeGapMm;
    RenderBlocks();
}
```

- [ ] **Schritt 3: Lücken-Rendering in `RenderBlocks()` einbauen**

Den Beginn der `foreach`-Schleife in `RenderBlocks()` anpassen. Die bestehende Zeile:

```csharp
if (block.IsDeleted) continue;
```

ersetzen durch:

```csharp
if (block.IsDeleted)
{
    double gapH = BerechneLückenHöhePx(block);
    if (gapH > 0) RenderLücke(block, gapH);
    continue;
}
```

- [ ] **Schritt 4: Build — prüfen ob 0 Fehler**

```powershell
powershell -Command "& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\KI\StatikManager_V1\src\StatikManager\StatikManager.csproj' /p:Configuration=Debug /p:Platform=x64 /t:Build /v:minimal 2>&1"
```

Erwartet: `Build succeeded.` / `0 Error(s)`

- [ ] **Schritt 5: Commit**

```bash
git add src/StatikManager/Modules/Werkzeuge/BlockEditorPrototype.xaml.cs
git commit -m "feat: Lücken als visueller Platzhalter auf Canvas + Rechtsklick-Kontextmenü"
```

---

## Selbstreview — Spec-Abgleich

| Spec-Anforderung | Task |
|------------------|------|
| Variante A: Originalabstand | Task 1 (GapModus) + Task 5 (BerechneLückenHöhePx) |
| Variante B: mm-Eingabe | Task 3 (GapDialog Validierung) + Task 5 (DPI-Berechnung) |
| Variante C: 0 mm | Task 1 (GapModus.KeinAbstand) + Task 5 (gapH=0→kein Element) |
| Dialog beim Löschen | Task 4 (BtnDelete_Click) |
| Kontextmenü auf Lücke | Task 5 (RenderLücke + BearbeiteGap) |
| Dialog vorausgefüllt beim Bearbeiten | Task 3 (Konstruktor-Parameter) + Task 5 (BearbeiteGap) |
| DPI aus Bitmap auslesen | Task 5 (BerechneLückenHöhePx: `_originalBitmap.DpiY`) |
