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
    public enum GapModus { OriginalAbstand, KundenAbstand, KeinAbstand }

    // ══════════════════════════════════════════════════════════════════════════
    //  BlockEditorPrototype
    //
    //  Isolierter Prototyp für physische Blocktrennung.
    //  Keine Altlogik. Kein Overlay. Kein _scherenschnitte.
    //  Einzige Wahrheit: List<ProtoBlock> + OriginalBitmap.
    // ══════════════════════════════════════════════════════════════════════════
    public partial class BlockEditorPrototype : Window
    {
        // ── Datenmodell ───────────────────────────────────────────────────────

        private sealed class ProtoBlock
        {
            public int    Id              { get; set; }
            public int    SourcePageIndex { get; set; }   // immer 0 im Einzelseiten-Prototyp
            public double FracTop         { get; set; }   // 0.0 = Seitenanfang
            public double FracBottom      { get; set; }   // 1.0 = Seitenende
            public bool   IsDeleted       { get; set; }

            /// <summary>Wie groß die Lücke nach dem Löschen sein soll (GapModus). Nur relevant wenn IsDeleted == true.</summary>
            public GapModus GapArt { get; set; } = GapModus.OriginalAbstand;

            /// <summary>Lückengröße in mm. Nur relevant für GapModus.KundenAbstand. Muss >= 0 sein.</summary>
            public double GapMm { get; set; } = 0.0;

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
        }

        // ── Zustand ───────────────────────────────────────────────────────────

        private BitmapSource?    _originalBitmap;
        private List<ProtoBlock> _blocks         = new List<ProtoBlock>();
        private int              _nextId         = 0;
        private int              _selectedId     = -1;   // -1 = kein Block ausgewählt
        private bool             _suppressEvents = false; // Rekursionsschutz für BlockList

        private const double CanvasPad = 24.0; // Rand um das Bild

        public BlockEditorPrototype()
        {
            InitializeComponent();
        }

        // ══════════════════════════════════════════════════════════════════════
        //  PFLICHTMETHODEN
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Lädt eine Bilddatei und erzeugt den initialen Einzelblock (0.0–1.0).
        /// Jeder vorherige Zustand wird vollständig verworfen.
        /// </summary>
        private void LoadBitmap(string path)
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.UriSource   = new Uri(path, UriKind.Absolute);
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.EndInit();
            bmp.Freeze();

            _originalBitmap = bmp;
            _blocks.Clear();
            _nextId     = 0;
            _selectedId = -1;

            _blocks.Add(new ProtoBlock
            {
                Id              = _nextId++,
                SourcePageIndex = 0,
                FracTop         = 0.0,
                FracBottom      = 1.0,
                IsDeleted       = false
            });

            RenderBlocks();
        }

        /// <summary>
        /// Zeichnet ausschließlich die nicht-gelöschten Blöcke als je eigenen
        /// CroppedBitmap-Ausschnitt auf dem Canvas.
        ///
        /// Die vollständige Ursprungsseite wird im Renderpfad NICHT gezeichnet.
        /// Jeder Block ist ein physisch eigenständiges Bildfragment.
        /// </summary>
        private void RenderBlocks()
        {
            EditorCanvas.Children.Clear();

            if (_originalBitmap == null)
            {
                RefreshBlockList();
                return;
            }

            int    srcW = _originalBitmap.PixelWidth;
            int    srcH = _originalBitmap.PixelHeight;

            EditorCanvas.Width  = srcW + CanvasPad * 2;
            EditorCanvas.Height = srcH + CanvasPad * 2;

            foreach (var block in _blocks)
            {
                if (block.IsDeleted)
                {
                    double gapH = BerechneLückenHöhePx(block);
                    if (gapH > 0) RenderLücke(block, gapH);
                    continue;
                }

                // Fraktionen auf gültige Pixel-Grenzen abbilden
                double fracTop    = Math.Max(0.0, Math.Min(1.0, block.FracTop));
                double fracBottom = Math.Max(fracTop, Math.Min(1.0, block.FracBottom));
                if (fracBottom <= fracTop) continue;

                int pixelTop    = (int)Math.Round(fracTop    * srcH);
                int pixelHeight = (int)Math.Round(fracBottom * srcH) - pixelTop;
                pixelTop    = Math.Max(0, Math.Min(pixelTop, srcH - 1));
                pixelHeight = Math.Max(1, Math.Min(pixelHeight, srcH - pixelTop));
                if (pixelHeight <= 0) continue;

                // Eigenständiger Ausschnitt — keine Verbindung zum Originalblock mehr
                BitmapSource cropped;
                try
                {
                    cropped = new CroppedBitmap(
                        _originalBitmap,
                        new Int32Rect(0, pixelTop, srcW, pixelHeight));
                }
                catch
                {
                    continue;
                }

                double displayTop    = CanvasPad + fracTop * srcH;
                double displayHeight = (fracBottom - fracTop) * srcH;
                bool   isSelected    = (block.Id == _selectedId);

                DropShadowEffect? shadow = null;
                try
                {
                    shadow = new DropShadowEffect
                    {
                        BlurRadius  = 10,
                        ShadowDepth = 4,
                        Direction   = 270,
                        Color       = Colors.Black,
                        Opacity     = 0.6
                    };
                }
                catch { }

                int capturedId = block.Id;

                var blockBorder = new Border
                {
                    Tag             = capturedId,
                    Width           = srcW,
                    Height          = displayHeight,
                    Background      = Brushes.White,
                    BorderBrush     = isSelected
                                        ? new SolidColorBrush(Color.FromRgb(0, 100, 220))
                                        : new SolidColorBrush(Color.FromRgb(130, 130, 130)),
                    BorderThickness = new Thickness(isSelected ? 3 : 1),
                    Effect          = shadow,
                    SnapsToDevicePixels = true,
                    Child = new Image
                    {
                        Source              = cropped,
                        Width               = srcW,
                        Height              = displayHeight,
                        Stretch             = Stretch.Fill,
                        SnapsToDevicePixels = true
                    }
                };

                // Klick im Auswahl-Modus selektiert diesen Block
                blockBorder.MouseLeftButtonDown += (_, ev) =>
                {
                    if (BtnCutMode.IsChecked == true) return; // Schnitt-Modus: bubbling an Canvas
                    _selectedId = capturedId;
                    RenderBlocks();
                    ev.Handled = true;
                };

                // Rechtsklick: "Teil löschen …"
                var blockCm = new ContextMenu();
                var blockMi = new MenuItem { Header = "Teil löschen \u2026" };
                int capturedDeleteId = capturedId;
                blockMi.Click += (_, __) =>
                {
                    var dlg = new GapDialog { Owner = this };
                    if (dlg.ShowDialog() != true) return;
                    DeleteBlock(capturedDeleteId, dlg.GewählterModus, dlg.EingabeGapMm);
                };
                blockCm.Items.Add(blockMi);
                blockBorder.ContextMenu = blockCm;

                Canvas.SetLeft(blockBorder, CanvasPad);
                Canvas.SetTop(blockBorder,  displayTop);
                EditorCanvas.Children.Add(blockBorder);
            }

            RefreshBlockList();
        }

        /// <summary>
        /// Sucht den aktiven (nicht gelöschten) Block, der den gegebenen
        /// Fraktionswert enthält. Gibt null zurück wenn keiner passt.
        /// </summary>
        private ProtoBlock? FindBlockAtFraction(double frac)
        {
            foreach (var block in _blocks)
            {
                if (block.IsDeleted) continue;
                if (frac > block.FracTop && frac < block.FracBottom)
                    return block;
            }
            return null;
        }

        /// <summary>
        /// Teilt den Block mit der gegebenen ID physisch bei splitFrac auf.
        ///
        /// Der Originalblock wird aus der Liste ENTFERNT.
        /// An seiner Stelle entstehen zwei neue unabhängige Blöcke:
        ///   Block A: FracTop    → splitFrac
        ///   Block B: splitFrac  → FracBottom
        ///
        /// Nach diesem Aufruf existiert der ursprüngliche Block nicht mehr.
        /// </summary>
        private void SplitBlock(int blockId, double splitFrac)
        {
            int idx = _blocks.FindIndex(b => b.Id == blockId);
            if (idx < 0) return;

            var orig = _blocks[idx];
            if (splitFrac <= orig.FracTop || splitFrac >= orig.FracBottom) return;

            var blockA = new ProtoBlock
            {
                Id              = _nextId++,
                SourcePageIndex = orig.SourcePageIndex,
                FracTop         = orig.FracTop,
                FracBottom      = splitFrac,
                IsDeleted       = false
            };

            var blockB = new ProtoBlock
            {
                Id              = _nextId++,
                SourcePageIndex = orig.SourcePageIndex,
                FracTop         = splitFrac,
                FracBottom      = orig.FracBottom,
                IsDeleted       = false
            };

            _blocks.RemoveAt(idx);        // Original weg
            _blocks.Insert(idx, blockB);  // B an Stelle des Originals
            _blocks.Insert(idx, blockA);  // A davor

            _selectedId = -1;
            RenderBlocks();
        }

        /// <summary>
        /// Markiert den Block als gelöscht und speichert den gewünschten Lückenabstand.
        /// RenderBlocks() rendert für diesen Block einen Lücken-Platzhalter (Darstellung abhängig von GapArt/GapMm).
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

        // ══════════════════════════════════════════════════════════════════════
        //  HILFSMETHODEN
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>Rechnet Canvas-Y-Koordinate in Block-Fraktion (0.0–1.0) um.</summary>
        private double CanvasYToFrac(double canvasY)
        {
            if (_originalBitmap == null) return 0.0;
            double relY = canvasY - CanvasPad;
            return Math.Max(0.0, Math.Min(1.0, relY / _originalBitmap.PixelHeight));
        }

        /// <summary>
        /// Berechnet die anzuzeigende Lückenhöhe in Pixeln anhand des GapArt.
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
            block.GapMm  = dlg.EingabeGapMm;
            RenderBlocks();
        }

        private void RefreshBlockList()
        {
            _suppressEvents = true;
            try
            {
                BlockList.Items.Clear();
                int selIdx = -1;
                for (int i = 0; i < _blocks.Count; i++)
                {
                    BlockList.Items.Add(_blocks[i].ToString());
                    if (_blocks[i].Id == _selectedId) selIdx = i;
                }
                if (selIdx >= 0) BlockList.SelectedIndex = selIdx;
            }
            finally
            {
                _suppressEvents = false;
            }

            int active  = _blocks.Count(b => !b.IsDeleted);
            int deleted = _blocks.Count(b => b.IsDeleted);
            string sel  = _selectedId >= 0 ? $"B{_selectedId} gewählt" : "kein Block gewählt";
            string mode = BtnCutMode.IsChecked == true ? "  |  ✂ SCHNITT-MODUS" : "";
            TxtStatus.Text      = $"{active} aktiv / {deleted} gelöscht  |  {sel}{mode}";
        }

        // ══════════════════════════════════════════════════════════════════════
        //  EVENT-HANDLER
        // ══════════════════════════════════════════════════════════════════════

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title  = "Testbild laden (PNG / JPG / BMP)",
                Filter = "Bilder|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif|Alle Dateien|*.*"
            };
            if (dlg.ShowDialog(this) == true)
                LoadBitmap(dlg.FileName);
        }

        private void BtnCutMode_Checked(object sender, RoutedEventArgs e)
        {
            EditorCanvas.Cursor = Cursors.Cross;
            RefreshBlockList();
        }

        private void BtnCutMode_Unchecked(object sender, RoutedEventArgs e)
        {
            EditorCanvas.Cursor = Cursors.Arrow;
            RefreshBlockList();
        }

        private void EditorCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (BtnCutMode.IsChecked != true) return;

            double canvasY = e.GetPosition(EditorCanvas).Y;
            double frac    = CanvasYToFrac(canvasY);
            var    block   = FindBlockAtFraction(frac);

            if (block != null)
            {
                SplitBlock(block.Id, frac);
                BtnCutMode.IsChecked = false;
            }

            e.Handled = true;
        }

        private void BlockList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressEvents) return;

            int idx = BlockList.SelectedIndex;
            if (idx >= 0 && idx < _blocks.Count)
            {
                _selectedId = _blocks[idx].Id;
                RenderBlocks();
            }
        }
    }
}
