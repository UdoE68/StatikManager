# Word-Export im StatikManager – Implementierungsplan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** AxisVM-Ausschnitte (PNG + Metadaten) per Klick an die aktuelle Cursorposition in einem geöffneten Word-Dokument einfügen.

**Architecture:** Neues Modul `WordExport` mit Panel und `IModul`-Implementierung. Ein neuer `WordEinfuegenService` kapselt alle interaktiven Word-COM-Aufrufe (getrennt vom bestehenden `WordInteropService`, der nur stille PDF-Konvertierung macht). Das Panel liest AxisVM-Positionen aus dem Projektordner (`{Projektpfad}/Statik/Pos_NN_*/position.json`) und zeigt sie als aufklappbare Liste.

**Tech Stack:** C# / WPF (.NET 4.8), Microsoft.Office.Interop.Word (COM-Verweis bereits im .csproj), System.Text.Json / Newtonsoft.Json (wie im Rest des Projekts)

---

## Dateistruktur

| Aktion | Pfad | Zweck |
|--------|------|-------|
| Neu | `Modules/WordExport/WordExportModul.cs` | IModul-Implementierung, Modul-Registrierung |
| Neu | `Modules/WordExport/WordExportPanel.xaml` | WPF-UI: Dokument-Leiste, Positions-Liste, Optionen |
| Neu | `Modules/WordExport/WordExportPanel.xaml.cs` | Code-behind: Daten laden, Einfügen auslösen |
| Neu | `Infrastructure/WordEinfuegenService.cs` | Interaktive Word-Interop (Singleton-App, InsertAtCursor) |
| Ändern | `Core/Einstellungen.cs` | Neue Felder: Bildbreite, MitUeberschrift, MitMassstab |
| Ändern | `Core/SitzungsZustand.cs` | Neues Feld: WordExportLetztesDokument |
| Ändern | `MainWindow.xaml.cs` | Neues Modul registrieren |

---

## Task 1: WordEinfuegenService

**Files:**
- Neu: `src/StatikManager/Infrastructure/WordEinfuegenService.cs`

Dieser Service kapselt alle COM-Aufrufe für das interaktive Word-Editing. Er ist komplett unabhängig vom bestehenden `WordInteropService` (der für stille PDF-Konvertierung zuständig ist).

- [ ] **Schritt 1: Datei anlegen**

```csharp
// src/StatikManager/Infrastructure/WordEinfuegenService.cs
using System;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Interaktiver Word-Service: verbindet sich mit dem laufenden Word,
    /// öffnet/erstellt Dokumente und fügt Inhalte an der Cursorposition ein.
    /// Alle Methoden müssen auf dem UI-Thread aufgerufen werden (STA).
    /// </summary>
    internal static class WordEinfuegenService
    {
        private static Word.Application? _wordApp;

        // ── Verbindung ─────────────────────────────────────────────────────

        /// <summary>
        /// Gibt true zurück wenn Word läuft und ein Dokument geöffnet ist.
        /// </summary>
        public static bool IstWordBereit()
        {
            try
            {
                var app = HoleWordApp();
                return app != null && app.Documents.Count > 0;
            }
            catch { return false; }
        }

        /// <summary>
        /// Gibt den Pfad des aktiven Word-Dokuments zurück, oder null.
        /// </summary>
        public static string? GetAktiveDokumentPfad()
        {
            try
            {
                var app = HoleWordApp();
                if (app == null || app.Documents.Count == 0) return null;
                var doc = app.ActiveDocument;
                return string.IsNullOrEmpty(doc.FullName) ? null : doc.FullName;
            }
            catch { return null; }
        }

        // ── Dokument öffnen / erstellen ────────────────────────────────────

        /// <summary>
        /// Öffnet eine bestehende .docx-Datei in Word (sichtbar).
        /// </summary>
        public static void OeffneDokument(string pfad)
        {
            var app = HoleOderStarteWord();
            app.Documents.Open(
                FileName: pfad,
                ReadOnly: false,
                AddToRecentFiles: true,
                Visible: true);
            app.Visible = true;
            app.Activate();
        }

        /// <summary>
        /// Erstellt ein neues Dokument, optional auf Basis einer .dotx-Vorlage.
        /// </summary>
        public static void ErstelleDokument(string? vorlagePfad)
        {
            var app = HoleOderStarteWord();
            if (!string.IsNullOrEmpty(vorlagePfad) && File.Exists(vorlagePfad))
                app.Documents.Add(Template: vorlagePfad, Visible: true);
            else
                app.Documents.Add(Visible: true);
            app.Visible = true;
            app.Activate();
        }

        // ── Einfügen ──────────────────────────────────────────────────────

        /// <summary>
        /// Fügt ein PNG-Bild mit optionaler Beschriftungszeile an der aktuellen
        /// Cursor-Position im aktiven Word-Dokument ein.
        /// </summary>
        /// <param name="pngPfad">Absoluter Pfad zur PNG-Datei.</param>
        /// <param name="ueberschrift">Titel des Ausschnitts (z.B. "Ansicht vorne").</param>
        /// <param name="massstab">Maßstab-Text (z.B. "1:100").</param>
        /// <param name="bildbreiteOption">Wie breit das Bild eingefügt wird.</param>
        /// <param name="mitUeberschrift">Beschriftungszeile einfügen.</param>
        /// <param name="mitMassstab">Maßstab in Beschriftung aufnehmen.</param>
        public static void EinfuegenAnCursor(
            string pngPfad,
            string ueberschrift,
            string massstab,
            BildbreiteOption bildbreiteOption,
            bool mitUeberschrift,
            bool mitMassstab)
        {
            if (!File.Exists(pngPfad))
                throw new FileNotFoundException($"PNG nicht gefunden: {pngPfad}");

            var app = HoleWordApp()
                ?? throw new InvalidOperationException("Word ist nicht geöffnet.");
            if (app.Documents.Count == 0)
                throw new InvalidOperationException("Kein Word-Dokument geöffnet.");

            // Aktive Cursorposition (Selection) holen
            Word.Range cursor = app.Selection.Range;

            // Bild einfügen
            Word.InlineShape shape = cursor.InlineShapes.AddPicture(
                FileName: pngPfad,
                LinkToFile: false,
                SaveWithDocument: true);

            // Bildbreite setzen
            float breitePixel = BerechneBildbreite(app, bildbreiteOption);
            if (breitePixel > 0)
            {
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Width = breitePixel;
            }

            // Beschriftungszeile
            if (mitUeberschrift || mitMassstab)
            {
                shape.Range.InsertParagraphAfter();

                // Cursor hinter das eingefügte Bild bewegen
                Word.Range nachBild = shape.Range;
                nachBild.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                nachBild.MoveEnd(Word.WdUnits.wdParagraph, 1);
                nachBild.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                string beschriftung = BaueBeschriftung(ueberschrift, massstab, mitUeberschrift, mitMassstab);
                nachBild.InsertAfter(beschriftung);
            }

            // Cursor hinter den eingefügten Block bewegen
            app.Selection.EndOf(Word.WdUnits.wdParagraph);
        }

        // ── Hilfsmethoden ─────────────────────────────────────────────────

        private static Word.Application? HoleWordApp()
        {
            try
            {
                if (_wordApp != null)
                {
                    // Prüfen ob die bestehende Referenz noch gültig ist
                    _ = _wordApp.Version;
                    return _wordApp;
                }
            }
            catch
            {
                // COM-Objekt ungültig (Word abgestürzt o.ä.) → neu versuchen
                _wordApp = null;
            }

            try
            {
                _wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                return _wordApp;
            }
            catch
            {
                return null; // Word läuft nicht
            }
        }

        private static Word.Application HoleOderStarteWord()
        {
            var app = HoleWordApp();
            if (app != null) return app;

            _wordApp = new Word.Application { Visible = true };
            return _wordApp;
        }

        private static float BerechneBildbreite(Word.Application app, BildbreiteOption option)
        {
            // Word arbeitet intern in Punkt (1 cm = 28.3465 Punkt)
            const float CmZuPunkt = 28.3465f;
            switch (option)
            {
                case BildbreiteOption.Seitenbreite:
                {
                    var doc = app.ActiveDocument;
                    float seitenBreite = (float)doc.PageSetup.PageWidth;
                    float randLinks    = (float)doc.PageSetup.LeftMargin;
                    float randRechts   = (float)doc.PageSetup.RightMargin;
                    return seitenBreite - randLinks - randRechts;
                }
                case BildbreiteOption.HalbeSeitenbreite:
                {
                    var doc = app.ActiveDocument;
                    float seitenBreite = (float)doc.PageSetup.PageWidth;
                    float randLinks    = (float)doc.PageSetup.LeftMargin;
                    float randRechts   = (float)doc.PageSetup.RightMargin;
                    return (seitenBreite - randLinks - randRechts) / 2f;
                }
                case BildbreiteOption.Manuell_14cm:
                    return 14f * CmZuPunkt;
                case BildbreiteOption.Manuell_10cm:
                    return 10f * CmZuPunkt;
                default:
                    return 0; // 0 = keine Skalierung, Originalgröße
            }
        }

        private static string BaueBeschriftung(string ueberschrift, string massstab,
            bool mitUeberschrift, bool mitMassstab)
        {
            if (mitUeberschrift && mitMassstab && !string.IsNullOrEmpty(massstab))
                return $"{ueberschrift}   |   Maßstab {massstab}";
            if (mitUeberschrift)
                return ueberschrift;
            if (mitMassstab && !string.IsNullOrEmpty(massstab))
                return $"Maßstab {massstab}";
            return "";
        }
    }

    /// <summary>Wie breit das Bild in Word eingefügt wird.</summary>
    public enum BildbreiteOption
    {
        Seitenbreite,
        HalbeSeitenbreite,
        Manuell_14cm,
        Manuell_10cm,
        Original
    }
}
```

- [ ] **Schritt 2: Kompilieren**

```
Im Projektordner C:\KI\StatikManager_V1\src\StatikManager:
  dotnet build StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded` ohne Fehler.

- [ ] **Schritt 3: Commit**

```bash
cd "C:/KI/StatikManager_V1"
git add src/StatikManager/Infrastructure/WordEinfuegenService.cs
git commit -m "feat: WordEinfuegenService – interaktiver Word-COM-Service"
```

---

## Task 2: Einstellungen + SitzungsZustand erweitern

**Files:**
- Ändern: `src/StatikManager/Core/Einstellungen.cs`
- Ändern: `src/StatikManager/Core/SitzungsZustand.cs`

- [ ] **Schritt 1: Einstellungen.cs erweitern**

In `Einstellungen.cs` nach der bestehenden `WordVorlagen`-Eigenschaft einfügen:

```csharp
// ── Word-Export-Einstellungen ─────────────────────────────────────────────

/// <summary>Standard-Bildbreite beim Einfügen in Word.</summary>
public BildbreiteOption WordExportBildbreite { get; set; } = BildbreiteOption.Seitenbreite;

/// <summary>Überschrift unter eingefügtes Bild schreiben.</summary>
public bool WordExportMitUeberschrift { get; set; } = true;

/// <summary>Maßstab in Beschriftung aufnehmen.</summary>
public bool WordExportMitMassstab { get; set; } = true;
```

Außerdem am Anfang der Datei ergänzen:
```csharp
using StatikManager.Infrastructure; // für BildbreiteOption
```

- [ ] **Schritt 2: SitzungsZustand.cs erweitern**

In `SitzungsZustand.cs` nach `AktiveDatei` einfügen:

```csharp
/// <summary>Zuletzt in Word geöffnetes/verwendetes Dokument.</summary>
public string? WordExportLetztesDokument { get; set; }
```

- [ ] **Schritt 3: Kompilieren**

```
dotnet build src/StatikManager/StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded`.

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Core/Einstellungen.cs src/StatikManager/Core/SitzungsZustand.cs
git commit -m "feat: Einstellungen + SitzungsZustand um Word-Export-Felder erweitern"
```

---

## Task 3: AxisVM-Datenmodelle (ViewModels)

**Files:**
- Neu: `src/StatikManager/Modules/WordExport/PositionViewModel.cs`

Diese ViewModels werden nur intern im WordExportPanel verwendet.

- [ ] **Schritt 1: PositionViewModel.cs anlegen**

```csharp
// src/StatikManager/Modules/WordExport/PositionViewModel.cs
using System.Collections.Generic;

namespace StatikManager.Modules.WordExport
{
    /// <summary>Repräsentiert eine AxisVM-Position (Pos_01_Fundament/).</summary>
    internal sealed class PositionViewModel
    {
        public string Id         { get; set; } = "";
        public string Name       { get; set; } = "";
        public string OrdnerPfad { get; set; } = "";

        public List<AusschnittViewModel> Ausschnitte { get; set; } = new();
    }

    /// <summary>Repräsentiert einen einzelnen Ausschnitt (PNG + Metadaten).</summary>
    internal sealed class AusschnittViewModel
    {
        public int    Nr          { get; set; }
        public string Ueberschrift { get; set; } = "";
        public string PngPfad     { get; set; } = "";
        public string Massstab    { get; set; } = "";

        /// <summary>Anzeigename in der Liste: "001 – Ansicht vorne  [1:100]"</summary>
        public string Anzeigename =>
            string.IsNullOrEmpty(Massstab)
                ? $"{Nr:000}  –  {Ueberschrift}"
                : $"{Nr:000}  –  {Ueberschrift}   [{Massstab}]";
    }
}
```

- [ ] **Schritt 2: Kompilieren**

```
dotnet build src/StatikManager/StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded`.

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/WordExport/PositionViewModel.cs
git commit -m "feat: PositionViewModel + AusschnittViewModel für WordExportPanel"
```

---

## Task 4: WordExportPanel XAML

**Files:**
- Neu: `src/StatikManager/Modules/WordExport/WordExportPanel.xaml`

- [ ] **Schritt 1: XAML-Datei anlegen**

```xml
<!-- src/StatikManager/Modules/WordExport/WordExportPanel.xaml -->
<UserControl x:Class="StatikManager.Modules.WordExport.WordExportPanel"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <Grid>
        <Grid.RowDefinitions>
            <!-- Abschnitt 1: Word-Dokument -->
            <RowDefinition Height="Auto"/>
            <!-- Trennlinie -->
            <RowDefinition Height="1"/>
            <!-- Abschnitt 2: Positionen-Liste -->
            <RowDefinition Height="*"/>
            <!-- Trennlinie -->
            <RowDefinition Height="1"/>
            <!-- Abschnitt 3: Einfüge-Optionen -->
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- ── 1. Word-Dokument ─────────────────────────────── -->
        <DockPanel Grid.Row="0" Margin="8,8,8,6">
            <TextBlock Text="Word-Dokument"
                       FontWeight="SemiBold" FontSize="13"
                       VerticalAlignment="Center"
                       DockPanel.Dock="Top"
                       Margin="0,0,0,6"/>

            <StackPanel Orientation="Horizontal" DockPanel.Dock="Top">
                <Button x:Name="BtnNeuErstellen"
                        Content="Neu erstellen"
                        Padding="8,3" Margin="0,0,4,0"
                        Click="BtnNeuErstellen_Click"/>
                <Button x:Name="BtnOeffnen"
                        Content="Öffnen …"
                        Padding="8,3" Margin="0,0,4,0"
                        Click="BtnOeffnen_Click"/>
                <Border Width="1" Background="#CCC" Margin="4,2,4,2"/>
                <Ellipse x:Name="StatusKreis"
                         Width="10" Height="10"
                         Fill="Red"
                         VerticalAlignment="Center"
                         ToolTip="Word-Status"
                         Margin="0,0,4,0"/>
                <TextBlock x:Name="TxtWordStatus"
                           Text="Word nicht verbunden"
                           FontSize="11" Foreground="Gray"
                           VerticalAlignment="Center"/>
            </StackPanel>

            <TextBlock x:Name="TxtDokumentPfad"
                       Text=""
                       FontSize="10" Foreground="Gray"
                       TextTrimming="CharacterEllipsis"
                       DockPanel.Dock="Top"
                       Margin="0,4,0,0"
                       ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}"/>
        </DockPanel>

        <!-- Trennlinie -->
        <Border Grid.Row="1" Background="#DDD"/>

        <!-- ── 2. Positionen aus AxisVM ─────────────────────── -->
        <DockPanel Grid.Row="2" Margin="8,8,8,0">
            <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
                <TextBlock Text="Positionen aus AxisVM"
                           FontWeight="SemiBold" FontSize="13"
                           VerticalAlignment="Center"/>
                <Button x:Name="BtnAktualisieren"
                        Content="↻"
                        FontSize="13"
                        Width="24" Height="24"
                        DockPanel.Dock="Right"
                        ToolTip="Liste neu einlesen"
                        Click="BtnAktualisieren_Click"/>
            </DockPanel>

            <TextBlock x:Name="TxtKeinePositionen"
                       Text="Kein Projekt geladen."
                       Foreground="Gray" FontStyle="Italic"
                       FontSize="11"
                       DockPanel.Dock="Top"
                       Visibility="Visible"/>

            <TreeView x:Name="PositionenTree"
                      DockPanel.Dock="Top"
                      Visibility="Collapsed"
                      BorderThickness="0"
                      Background="Transparent">
                <TreeView.ItemTemplate>
                    <HierarchicalDataTemplate ItemsSource="{Binding Ausschnitte}">
                        <!-- Positions-Ebene -->
                        <TextBlock FontWeight="SemiBold" FontSize="12">
                            <Run Text="{Binding Id}"/>
                            <Run Text=" – "/>
                            <Run Text="{Binding Name}"/>
                        </TextBlock>

                        <!-- Ausschnitt-Ebene -->
                        <HierarchicalDataTemplate.ItemTemplate>
                            <DataTemplate>
                                <DockPanel Margin="0,1">
                                    <Button Content="Einfügen"
                                            Padding="5,1"
                                            FontSize="10"
                                            DockPanel.Dock="Right"
                                            Tag="{Binding}"
                                            Click="BtnEinfuegen_Click"
                                            ToolTip="An Cursor-Position in Word einfügen"/>
                                    <TextBlock Text="{Binding Anzeigename}"
                                               FontSize="11"
                                               VerticalAlignment="Center"
                                               TextTrimming="CharacterEllipsis"/>
                                </DockPanel>
                            </DataTemplate>
                        </HierarchicalDataTemplate.ItemTemplate>
                    </HierarchicalDataTemplate>
                </TreeView.ItemTemplate>
            </TreeView>
        </DockPanel>

        <!-- Trennlinie -->
        <Border Grid.Row="3" Background="#DDD" Margin="0,8,0,0"/>

        <!-- ── 3. Einfüge-Optionen ───────────────────────────── -->
        <StackPanel Grid.Row="4" Margin="8,8,8,12">
            <TextBlock Text="Einfüge-Optionen"
                       FontWeight="SemiBold" FontSize="12"
                       Margin="0,0,0,6"/>

            <DockPanel Margin="0,0,0,4">
                <TextBlock Text="Bildbreite:" Width="80" VerticalAlignment="Center"/>
                <ComboBox x:Name="CbBildbreite"
                          Width="160"
                          SelectionChanged="CbBildbreite_SelectionChanged">
                    <ComboBoxItem Content="Seitenbreite"       Tag="Seitenbreite"/>
                    <ComboBoxItem Content="1/2 Seite"          Tag="HalbeSeitenbreite"/>
                    <ComboBoxItem Content="Manuell 14 cm"      Tag="Manuell_14cm"/>
                    <ComboBoxItem Content="Manuell 10 cm"      Tag="Manuell_10cm"/>
                    <ComboBoxItem Content="Originalgröße"      Tag="Original"/>
                </ComboBox>
            </DockPanel>

            <CheckBox x:Name="ChkMitUeberschrift"
                      Content="Überschrift einfügen"
                      Margin="0,0,0,3"
                      Checked="Optionen_Geaendert"
                      Unchecked="Optionen_Geaendert"/>
            <CheckBox x:Name="ChkMitMassstab"
                      Content="Maßstab in Beschriftung"
                      Checked="Optionen_Geaendert"
                      Unchecked="Optionen_Geaendert"/>
        </StackPanel>

    </Grid>
</UserControl>
```

- [ ] **Schritt 2: Kompilieren (XAML-Syntax prüfen)**

```
dotnet build src/StatikManager/StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded` (Code-behind fehlt noch → Fehler erwartet, nur XAML-Syntax prüfen ist hier das Ziel; Fehler über fehlende Klasse ist OK).

- [ ] **Schritt 3: Commit**

```bash
git add src/StatikManager/Modules/WordExport/WordExportPanel.xaml
git commit -m "feat: WordExportPanel XAML – Layout mit 3 Abschnitten"
```

---

## Task 5: WordExportPanel Code-Behind

**Files:**
- Neu: `src/StatikManager/Modules/WordExport/WordExportPanel.xaml.cs`

- [ ] **Schritt 1: Code-behind anlegen**

```csharp
// src/StatikManager/Modules/WordExport/WordExportPanel.xaml.cs
using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace StatikManager.Modules.WordExport
{
    public partial class WordExportPanel : System.Windows.Controls.UserControl
    {
        private readonly DispatcherTimer _statusTimer;
        private string? _projektPfad;

        public WordExportPanel()
        {
            InitializeComponent();

            // Status-Poll-Timer: alle 3 Sekunden Word-Status prüfen
            _statusTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            _statusTimer.Tick += (_, _) => AktualisiereWordStatus();
            _statusTimer.Start();

            // Auf Projektwechsel reagieren
            AppZustand.Instanz.ProjektGeändert += pfad =>
            {
                _projektPfad = pfad;
                LadePositionen();
            };

            // Einstellungen laden
            var ein = Einstellungen.Instanz;
            ChkMitUeberschrift.IsChecked = ein.WordExportMitUeberschrift;
            ChkMitMassstab.IsChecked     = ein.WordExportMitMassstab;
            WaehleBildbreiteComboBox(ein.WordExportBildbreite);

            AktualisiereWordStatus();
        }

        // ── Word-Status ───────────────────────────────────────────────────

        private void AktualisiereWordStatus()
        {
            bool bereit = WordEinfuegenService.IstWordBereit();
            StatusKreis.Fill   = bereit ? Brushes.Green : Brushes.Red;
            TxtWordStatus.Text = bereit ? "Word verbunden" : "Word nicht verbunden";

            var pfad = WordEinfuegenService.GetAktiveDokumentPfad();
            TxtDokumentPfad.Text = pfad != null
                ? System.IO.Path.GetFileName(pfad)
                : "";
        }

        // ── Dokument-Aktionen ─────────────────────────────────────────────

        private void BtnNeuErstellen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var vorlage = Einstellungen.Instanz.WordVorlagen
                    .Find(v => v.Standard && v.PfadGültig)?.Pfad;
                WordEinfuegenService.ErstelleDokument(vorlage);
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Erstellen: {ex.Message}",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BtnOeffnen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Word-Dokument öffnen",
                Filter = "Word-Dokumente|*.docx;*.doc|Alle Dateien|*.*"
            };

            var sitzung = SitzungsZustand.Laden();
            if (!string.IsNullOrEmpty(sitzung.WordExportLetztesDokument))
                dlg.InitialDirectory = Path.GetDirectoryName(sitzung.WordExportLetztesDokument);

            if (dlg.ShowDialog() != true) return;

            try
            {
                WordEinfuegenService.OeffneDokument(dlg.FileName);

                sitzung.WordExportLetztesDokument = dlg.FileName;
                sitzung.Speichern();
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Öffnen: {ex.Message}",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Positionen laden ──────────────────────────────────────────────

        private void BtnAktualisieren_Click(object sender, RoutedEventArgs e)
            => LadePositionen();

        private void LadePositionen()
        {
            var positionen = new List<PositionViewModel>();

            if (!string.IsNullOrEmpty(_projektPfad))
            {
                string statistikOrdner = Path.Combine(_projektPfad, "Statik");
                if (Directory.Exists(statistikOrdner))
                {
                    foreach (var posOrdner in Directory.GetDirectories(statistikOrdner, "Pos_*"))
                    {
                        var pos = LesePosition(posOrdner);
                        if (pos != null) positionen.Add(pos);
                    }
                }
            }

            if (positionen.Count == 0)
            {
                TxtKeinePositionen.Visibility = Visibility.Visible;
                PositionenTree.Visibility     = Visibility.Collapsed;
                TxtKeinePositionen.Text       = string.IsNullOrEmpty(_projektPfad)
                    ? "Kein Projekt geladen."
                    : "Keine Positionen gefunden.";
            }
            else
            {
                TxtKeinePositionen.Visibility = Visibility.Collapsed;
                PositionenTree.Visibility     = Visibility.Visible;
                PositionenTree.ItemsSource    = positionen;
            }
        }

        private PositionViewModel? LesePosition(string ordnerPfad)
        {
            try
            {
                string jsonPfad = Path.Combine(ordnerPfad, "position.json");
                if (!File.Exists(jsonPfad)) return null;

                var doc = JsonDocument.Parse(File.ReadAllText(jsonPfad));
                var root = doc.RootElement;

                var pos = new PositionViewModel
                {
                    Id         = root.GetProperty("id").GetString() ?? "",
                    Name       = root.GetProperty("name").GetString() ?? "",
                    OrdnerPfad = ordnerPfad
                };

                if (root.TryGetProperty("ausschnitte", out var ausschnitteEl))
                {
                    foreach (var a in ausschnitteEl.EnumerateArray())
                    {
                        string dateiname = a.TryGetProperty("dateiname", out var fn)
                            ? fn.GetString() ?? "" : "";
                        string pngPfad   = Path.Combine(ordnerPfad, "daten", dateiname);

                        // Maßstab aus Ausschnitt-JSON lesen
                        string massstab = "";
                        string jsonDatei = Path.ChangeExtension(pngPfad, ".json");
                        if (File.Exists(jsonDatei))
                        {
                            try
                            {
                                var aDoc  = JsonDocument.Parse(File.ReadAllText(jsonDatei));
                                if (aDoc.RootElement.TryGetProperty("massstab", out var ms))
                                    massstab = ms.GetString() ?? "";
                            }
                            catch { }
                        }

                        pos.Ausschnitte.Add(new AusschnittViewModel
                        {
                            Nr           = a.TryGetProperty("nr", out var nr) ? nr.GetInt32() : 0,
                            Ueberschrift = a.TryGetProperty("ueberschrift", out var ue)
                                ? ue.GetString() ?? "" : "",
                            PngPfad  = pngPfad,
                            Massstab = massstab
                        });
                    }
                }

                return pos;
            }
            catch (Exception ex)
            {
                Logger.Warn("WordExport", $"Position lesen fehlgeschlagen: {ex.Message}");
                return null;
            }
        }

        // ── Einfügen ──────────────────────────────────────────────────────

        private void BtnEinfuegen_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Button btn) return;
            if (btn.Tag is not AusschnittViewModel ausschnitt) return;

            if (!WordEinfuegenService.IstWordBereit())
            {
                MessageBox.Show("Bitte zuerst ein Word-Dokument öffnen.",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!File.Exists(ausschnitt.PngPfad))
            {
                MessageBox.Show($"PNG-Datei nicht gefunden:\n{ausschnitt.PngPfad}",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                WordEinfuegenService.EinfuegenAnCursor(
                    pngPfad:          ausschnitt.PngPfad,
                    ueberschrift:     ausschnitt.Ueberschrift,
                    massstab:         ausschnitt.Massstab,
                    bildbreiteOption: Einstellungen.Instanz.WordExportBildbreite,
                    mitUeberschrift:  ChkMitUeberschrift.IsChecked == true,
                    mitMassstab:      ChkMitMassstab.IsChecked == true);

                AppZustand.Instanz.SetzeStatus(
                    $"Eingefügt: {ausschnitt.Ueberschrift}");
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Einfügen:\n{ex.Message}",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Optionen ──────────────────────────────────────────────────────

        private void CbBildbreite_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CbBildbreite.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            if (!Enum.TryParse<BildbreiteOption>(item.Tag?.ToString(), out var option)) return;

            Einstellungen.Instanz.WordExportBildbreite = option;
            Einstellungen.Instanz.Speichern();
        }

        private void Optionen_Geaendert(object sender, RoutedEventArgs e)
        {
            var ein = Einstellungen.Instanz;
            ein.WordExportMitUeberschrift = ChkMitUeberschrift.IsChecked == true;
            ein.WordExportMitMassstab     = ChkMitMassstab.IsChecked == true;
            ein.Speichern();
        }

        private void WaehleBildbreiteComboBox(BildbreiteOption option)
        {
            foreach (System.Windows.Controls.ComboBoxItem item in CbBildbreite.Items)
            {
                if (item.Tag?.ToString() == option.ToString())
                {
                    CbBildbreite.SelectedItem = item;
                    return;
                }
            }
            CbBildbreite.SelectedIndex = 0; // Fallback: Seitenbreite
        }

        // ── Cleanup ───────────────────────────────────────────────────────

        public void Bereinigen()
        {
            _statusTimer.Stop();
        }
    }
}
```

- [ ] **Schritt 2: Kompilieren**

```
dotnet build src/StatikManager/StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded`.

- [ ] **Schritt 3: Manuell testen**

1. StatikManager starten (`Start_Debug.bat`)
2. Word-Export-Tab öffnen (kommt in Task 6)
3. Prüfen: Status-Kreis rot, "Word nicht verbunden"
4. Word öffnen, neues Dokument erstellen
5. Prüfen: nach max. 3 Sek. Status-Kreis grün
6. Projekt mit AxisVM-Daten laden
7. Prüfen: Positionen erscheinen in der Liste

- [ ] **Schritt 4: Commit**

```bash
git add src/StatikManager/Modules/WordExport/WordExportPanel.xaml.cs
git commit -m "feat: WordExportPanel Code-behind – Positionen laden, Einfügen, Status-Poll"
```

---

## Task 6: WordExportModul + Registrierung

**Files:**
- Neu: `src/StatikManager/Modules/WordExport/WordExportModul.cs`
- Ändern: `src/StatikManager/MainWindow.xaml.cs`

- [ ] **Schritt 1: WordExportModul.cs anlegen**

```csharp
// src/StatikManager/Modules/WordExport/WordExportModul.cs
using StatikManager.Core;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Modules.WordExport
{
    public class WordExportModul : IModul
    {
        private WordExportPanel? _panel;

        public string Id   => "wordexport";
        public string Name => "Word-Export";

        public UIElement ErstellePanel()
        {
            _panel = new WordExportPanel();
            return _panel;
        }

        public MenuItem? ErzeugeMenüEintrag()
        {
            var menü = new MenuItem { Header = "_Word-Export" };

            var itemNeu = new MenuItem { Header = "Neues Word-Dokument …" };
            itemNeu.Click += (_, _) => _panel?.BtnNeuErstellen_Click(itemNeu, new RoutedEventArgs());

            var itemOeffnen = new MenuItem { Header = "Word-Dokument öffnen …" };
            itemOeffnen.Click += (_, _) => _panel?.BtnOeffnen_Click(itemOeffnen, new RoutedEventArgs());

            menü.Items.Add(itemNeu);
            menü.Items.Add(itemOeffnen);
            return menü;
        }

        public IEnumerable<FrameworkElement> ErzeugeWerkzeugleistenEinträge()
        {
            yield break;
        }

        public void Bereinigen() => _panel?.Bereinigen();
    }
}
```

Damit die Menü-Methoden von außen aufrufbar sind, müssen `BtnNeuErstellen_Click` und `BtnOeffnen_Click` in `WordExportPanel.xaml.cs` von `private` auf `internal` geändert werden:

```csharp
// WordExportPanel.xaml.cs – diese beiden Methoden auf internal ändern:
internal void BtnNeuErstellen_Click(object sender, RoutedEventArgs e) { ... }
internal void BtnOeffnen_Click(object sender, RoutedEventArgs e) { ... }
```

- [ ] **Schritt 2: In MainWindow.xaml.cs registrieren**

In `MainWindow.xaml.cs` ganz oben den using-Eintrag hinzufügen:
```csharp
using StatikManager.Modules.WordExport;
```

Im Konstruktor nach dem BildschnittModul registrieren:
```csharp
_modulManager.Registrieren(new DokumenteModul());
_modulManager.Registrieren(new BildschnittModul());
_modulManager.Registrieren(new WordExportModul()); // NEU
```

- [ ] **Schritt 3: Kompilieren**

```
dotnet build src/StatikManager/StatikManager.csproj
```
Erwartetes Ergebnis: `Build succeeded`.

- [ ] **Schritt 4: Vollständiger manueller Test**

1. StatikManager starten
2. "Word-Export"-Tab in der Modul-Leiste anklicken
3. **Test: Neu erstellen** → Word öffnet sich mit leerem Dokument, Status-Kreis wird grün
4. **Test: Öffnen** → Datei-Dialog erscheint, Dokument öffnet sich in Word
5. **Test: Positionen laden** → AxisVM-Projekt laden, Positionen erscheinen in der Liste
6. **Test: Einfügen** → Cursor in Word setzen, dann "Einfügen" klicken → Bild + Beschriftung erscheint
7. **Test: Bildbreite** → Seitenbreite wechseln, erneut einfügen → Bildgröße ändert sich
8. **Test: ohne Word** → Word schließen, "Einfügen" → Fehlermeldung erscheint

- [ ] **Schritt 5: Commit**

```bash
git add src/StatikManager/Modules/WordExport/WordExportModul.cs \
        src/StatikManager/MainWindow.xaml.cs
git commit -m "feat: WordExportModul registrieren – Word-Export vollständig integriert"
```

---

## Abschluss

Nach Abschluss aller Tasks ist das Feature vollständig:
- Neues "Word-Export"-Modul im StatikManager
- AxisVM-Positionen werden aus dem Projektordner gelesen und aufgeklappt angezeigt
- Ein Klick auf "Einfügen" fügt das PNG mit Beschriftung an der aktuellen Cursor-Position in Word ein
- Word-Status wird alle 3 Sekunden aktualisiert
- Einstellungen (Bildbreite, Überschrift, Maßstab) werden persistent gespeichert
