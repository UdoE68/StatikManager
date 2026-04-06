using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace StatikManager.Modules.Werkzeuge
{
    /// <summary>
    /// Isolierter Minimal-Prototyp für physische Blocktrennung.
    /// Kein Altmodell, kein _scherenschnitte, kein _gelöschteParts.
    /// Nur: Split / Delete / Render auf echter ContentBlock-Basis.
    /// </summary>
    public partial class BlockPrototyp : Window
    {
        // ── Datenmodell ───────────────────────────────────────────────────────

        /// <summary>
        /// Atomarer Block: ein vertikaler Ausschnitt aus der Quell-Bitmap.
        /// FracOben und FracUnten sind Anteile (0.0–1.0) der Originalhöhe.
        /// </summary>
        private sealed class ProtoBlock
        {
            public int     Id        { get; set; }
            public double  FracOben  { get; set; }
            public double  FracUnten { get; set; }
            public bool    IsDeleted { get; set; }

            /// <summary>Lückenabstand-Modus. Nur relevant wenn IsDeleted == true.</summary>
            public GapModus GapArt { get; set; } = GapModus.OriginalAbstand;

            /// <summary>Lückengröße in mm. Nur relevant für GapModus.KundenAbstand. Muss >= 0 sein.</summary>
            public double GapMm { get; set; } = 0.0;

            public string Beschreibung
            {
                get
                {
                    if (IsDeleted)
                    {
                        string gap = GapArt == GapModus.OriginalAbstand ? "[orig]"
                                   : GapArt == GapModus.KundenAbstand   ? $"[{GapMm:F1}mm]"
                                   : "[0mm]";
                        return $"[DEL {gap}] B{Id}  {FracOben:F3} – {FracUnten:F3}";
                    }
                    return $"      B{Id}  {FracOben:F3} – {FracUnten:F3}";
                }
            }
        }

        // ── Zustand ───────────────────────────────────────────────────────────

        private BitmapSource?    _quellBitmap;
        private List<ProtoBlock> _blöcke   = new List<ProtoBlock>();
        private int              _nextId   = 0;
        private int              _selId    = -1;    // ID des ausgewählten Blocks, -1 = keine Auswahl
        private bool             _renderLäuft = false;  // Verhindert Rekursion in LstBlöcke_SelectionChanged

        public BlockPrototyp()
        {
            InitializeComponent();
        }

        // ═════════════════════════════════════════════════════════════════════
        //  KERN-METHODEN
        // ═════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Lädt eine Bilddatei und initialisiert einen einzigen Block (0.0–1.0).
        /// Jeder vorherige Zustand wird verworfen.
        /// </summary>
        private void LadeBitmap(string pfad)
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.UriSource   = new Uri(pfad, UriKind.Absolute);
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.EndInit();
            bmp.Freeze();

            _quellBitmap = bmp;
            _blöcke.Clear();
            _nextId = 0;
            _selId  = -1;

            _blöcke.Add(new ProtoBlock
            {
                Id       = _nextId++,
                FracOben = 0.0,
                FracUnten= 1.0,
                IsDeleted= false
            });

            Render();
        }

        /// <summary>
        /// Teilt den Block mit der gegebenen ID an yFrac physisch auf.
        /// Der Originalblock wird aus der Liste entfernt.
        /// An seiner Stelle entstehen zwei neue, unabhängige Blöcke:
        ///   Block A: FracOben  → yFrac
        ///   Block B: yFrac     → FracUnten
        /// </summary>
        private void SplitBlock(int id, double yFrac)
        {
            int idx = _blöcke.FindIndex(b => b.Id == id);
            if (idx < 0) return;

            var orig = _blöcke[idx];
            // Schnitt muss innerhalb des Blocks liegen
            if (yFrac <= orig.FracOben || yFrac >= orig.FracUnten) return;

            var blockA = new ProtoBlock
            {
                Id       = _nextId++,
                FracOben = orig.FracOben,
                FracUnten= yFrac,
                IsDeleted= false
            };
            var blockB = new ProtoBlock
            {
                Id       = _nextId++,
                FracOben = yFrac,
                FracUnten= orig.FracUnten,
                IsDeleted= false
            };

            // Original entfernen, zwei neue an gleicher Stelle einfügen
            _blöcke.RemoveAt(idx);
            _blöcke.Insert(idx, blockB);
            _blöcke.Insert(idx, blockA);

            _selId = -1;
            Render();
        }

        /// <summary>
        /// Markiert den Block als gelöscht und speichert den Lückenabstand.
        /// Render() zeigt einen visuellen Platzhalter abhängig von GapArt/GapMm.
        /// </summary>
        private void DeleteBlock(int id, GapModus modus, double gapMm)
        {
            var b = _blöcke.FirstOrDefault(x => x.Id == id);
            if (b == null) return;
            b.IsDeleted = true;
            b.GapArt    = modus;
            b.GapMm     = gapMm;
            _selId = -1;
            Render();
        }

        /// <summary>Berechnet die Lückenhöhe in Pixeln anhand des GapArt.</summary>
        private double BerechneLückenHöhePx(ProtoBlock block)
        {
            if (_quellBitmap == null) return 0.0;
            switch (block.GapArt)
            {
                case GapModus.OriginalAbstand:
                    return (block.FracUnten - block.FracOben) * _quellBitmap.PixelHeight;
                case GapModus.KundenAbstand:
                    double dpi = _quellBitmap.DpiY > 0 ? _quellBitmap.DpiY : 96.0;
                    return block.GapMm * dpi / 25.4;
                default:
                    return 0.0;
            }
        }

        /// <summary>Zeichnet einen grauen Lücken-Platzhalter mit Kontextmenü auf dem Canvas.</summary>
        private void RenderLücke(ProtoBlock block, double gapH)
        {
            if (_quellBitmap == null) return;
            const double Rand = 20.0;
            int    bmpW       = _quellBitmap.PixelWidth;
            double displayY   = Rand + block.FracOben * _quellBitmap.PixelHeight;
            int    capturedId = block.Id;

            string label = block.GapArt == GapModus.OriginalAbstand
                ? "↕  Originalabstand"
                : block.GapArt == GapModus.KundenAbstand
                    ? $"↕  {block.GapMm:F1} mm"
                    : "";

            var lücke = new Border
            {
                Tag             = capturedId,
                Width           = bmpW,
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

            var cm = new ContextMenu();
            var mi = new MenuItem { Header = "Abstand bearbeiten …" };
            mi.Click += (_, __) => BearbeiteGap(capturedId);
            cm.Items.Add(mi);
            lücke.ContextMenu = cm;

            Canvas.SetLeft(lücke, Rand);
            Canvas.SetTop(lücke,  displayY);
            ProtoCanvas.Children.Add(lücke);
        }

        /// <summary>Öffnet den GapDialog vorausgefüllt für eine bestehende Lücke.</summary>
        private void BearbeiteGap(int blockId)
        {
            var block = _blöcke.FirstOrDefault(b => b.Id == blockId && b.IsDeleted);
            if (block == null) return;

            var dlg = new GapDialog(block.GapArt, block.GapMm) { Owner = this };
            if (dlg.ShowDialog() != true) return;

            block.GapArt = dlg.GewählterModus;
            block.GapMm  = dlg.EingabeGapMm;
            Render();
        }

        /// <summary>
        /// Zeichnet alle aktiven (nicht gelöschten) Blöcke als eigenständige
        /// CroppedBitmap-Elemente auf dem Canvas.
        /// Gelöschte Blöcke werden vollständig übersprungen — kein Bitmap, kein Border.
        /// </summary>
        private void Render()
        {
            ProtoCanvas.Children.Clear();

            if (_quellBitmap == null)
            {
                AktualisiereListeUndInfo();
                return;
            }

            int    bmpW = _quellBitmap.PixelWidth;
            int    bmpH = _quellBitmap.PixelHeight;
            const double Rand = 20.0;

            ProtoCanvas.Width  = bmpW + Rand * 2;
            ProtoCanvas.Height = bmpH + Rand * 2;

            foreach (var block in _blöcke)
            {
                if (block.IsDeleted)
                {
                    double gapH = BerechneLückenHöhePx(block);
                    if (gapH > 0) RenderLücke(block, gapH);
                    continue;
                }

                double fracO = Math.Max(0.0, Math.Min(1.0, block.FracOben));
                double fracU = Math.Max(fracO, Math.Min(1.0, block.FracUnten));
                if (fracU <= fracO) continue;

                // Pixel-Koordinaten im Quell-Bitmap
                int pixelY = (int)Math.Round(fracO * bmpH);
                int pixelH = (int)Math.Round((fracU - fracO) * bmpH);
                pixelY = Math.Max(0, Math.Min(pixelY, bmpH - 1));
                pixelH = Math.Max(1, Math.Min(pixelH, bmpH - pixelY));
                if (pixelH <= 0) continue;

                // Eigenständiger Bildausschnitt — keine Referenz auf Originalblock
                BitmapSource cropped;
                try
                {
                    cropped = new CroppedBitmap(
                        _quellBitmap,
                        new Int32Rect(0, pixelY, bmpW, pixelH));
                }
                catch { continue; }

                double displayY = Rand + fracO * bmpH;
                double displayH = Math.Max(1.0, (fracU - fracO) * bmpH);
                bool   isSelected = (block.Id == _selId);

                DropShadowEffect? shadow = null;
                try
                {
                    shadow = new DropShadowEffect
                    {
                        BlurRadius  = 8,
                        ShadowDepth = 3,
                        Direction   = 270,
                        Color       = Colors.Black,
                        Opacity     = 0.55
                    };
                }
                catch { /* ohne Schatten weiterzeichnen */ }

                int capturedId = block.Id;
                var blatt = new Border
                {
                    Tag             = capturedId,
                    Width           = bmpW,
                    Height          = displayH,
                    Background      = Brushes.White,
                    BorderBrush     = isSelected
                        ? new SolidColorBrush(Color.FromRgb(0, 100, 210))
                        : new SolidColorBrush(Color.FromRgb(140, 140, 140)),
                    BorderThickness = new Thickness(isSelected ? 3 : 1),
                    Child = new Image
                    {
                        Source              = cropped,
                        Width               = bmpW,
                        Height              = displayH,
                        Stretch             = Stretch.Fill,
                        SnapsToDevicePixels = true
                    },
                    Effect              = shadow,
                    SnapsToDevicePixels = true
                };

                // Klick im Auswahl-Modus: Block selektieren
                blatt.MouseLeftButtonDown += (_, ev) =>
                {
                    if (BtnSchnittModus.IsChecked == true) return; // Schnitt-Modus: Ereignis hochblasen lassen
                    _selId = capturedId;
                    Render();
                    ev.Handled = true;
                };

                Canvas.SetLeft(blatt, Rand);
                Canvas.SetTop(blatt,  displayY);
                ProtoCanvas.Children.Add(blatt);
            }

            AktualisiereListeUndInfo();
        }

        /// <summary>
        /// Berechnet aus einer Canvas-Y-Koordinate den yFrac-Wert (0.0–1.0)
        /// und gibt die ID des getroffenen aktiven Blocks zurück.
        /// Gibt (-1, 0) zurück wenn kein aktiver Block getroffen wurde.
        /// </summary>
        private (int blockId, double yFrac) KlickZuYFrac(double canvasY)
        {
            if (_quellBitmap == null) return (-1, 0.0);

            const double Rand = 20.0;
            double relY  = canvasY - Rand;
            double yFrac = Math.Max(0.0, Math.Min(1.0, relY / _quellBitmap.PixelHeight));

            foreach (var block in _blöcke)
            {
                if (block.IsDeleted) continue;
                if (yFrac > block.FracOben && yFrac < block.FracUnten)
                    return (block.Id, yFrac);
            }
            return (-1, yFrac);
        }

        // ═════════════════════════════════════════════════════════════════════
        //  UI-HILFSMETHODEN
        // ═════════════════════════════════════════════════════════════════════

        private void AktualisiereListeUndInfo()
        {
            _renderLäuft = true;
            try
            {
                LstBlöcke.Items.Clear();
                int selListIdx = -1;
                for (int i = 0; i < _blöcke.Count; i++)
                {
                    LstBlöcke.Items.Add(_blöcke[i].Beschreibung);
                    if (_blöcke[i].Id == _selId) selListIdx = i;
                }
                if (selListIdx >= 0) LstBlöcke.SelectedIndex = selListIdx;
            }
            finally { _renderLäuft = false; }

            int aktiv    = _blöcke.Count(b => !b.IsDeleted);
            int gelöscht = _blöcke.Count(b => b.IsDeleted);
            string selInfo   = _selId >= 0 ? $"Auswahl: B{_selId}" : "keine Auswahl";
            string modus     = BtnSchnittModus.IsChecked == true ? "  |  ✂ SCHNITT-MODUS" : "";
            TxtInfo.Text     = $"Blöcke: {aktiv} aktiv, {gelöscht} gelöscht  |  {selInfo}{modus}";
            BtnDelete.IsEnabled = _selId >= 0 && _blöcke.Any(b => b.Id == _selId && !b.IsDeleted);
        }

        // ═════════════════════════════════════════════════════════════════════
        //  EVENT-HANDLER
        // ═════════════════════════════════════════════════════════════════════

        private void BtnLaden_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title  = "Bilddatei laden",
                Filter = "Bilder|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif|Alle Dateien|*.*"
            };
            if (dlg.ShowDialog(this) == true)
                LadeBitmap(dlg.FileName);
        }

        private void BtnSchnittModus_Checked(object sender, RoutedEventArgs e)
        {
            ProtoCanvas.Cursor = Cursors.Cross;
            AktualisiereListeUndInfo();
        }

        private void BtnSchnittModus_Unchecked(object sender, RoutedEventArgs e)
        {
            ProtoCanvas.Cursor = Cursors.Arrow;
            AktualisiereListeUndInfo();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_selId < 0) return;

            var dlg = new GapDialog { Owner = this };
            if (dlg.ShowDialog() != true) return;

            DeleteBlock(_selId, dlg.GewählterModus, dlg.EingabeGapMm);
        }

        private void ProtoCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (BtnSchnittModus.IsChecked != true) return;

            var pos = e.GetPosition(ProtoCanvas);
            var (blockId, yFrac) = KlickZuYFrac(pos.Y);

            if (blockId >= 0)
            {
                SplitBlock(blockId, yFrac);
                BtnSchnittModus.IsChecked = false;  // nach Schnitt Modus beenden
            }
            e.Handled = true;
        }

        private void LstBlöcke_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_renderLäuft) return;  // Rekursionsschutz

            int selIdx = LstBlöcke.SelectedIndex;
            if (selIdx >= 0 && selIdx < _blöcke.Count)
            {
                _selId = _blöcke[selIdx].Id;
                Render();
            }
        }
    }
}
