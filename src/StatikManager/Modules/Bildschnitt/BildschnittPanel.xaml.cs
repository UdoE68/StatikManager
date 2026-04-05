using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WpfLine = System.Windows.Shapes.Line;
using WpfRect = System.Windows.Shapes.Rectangle;
using DrawRect = System.Drawing.Rectangle;

namespace StatikManager.Modules.Bildschnitt
{
    public partial class BildschnittPanel : UserControl
    {
        // Felder
        private System.Drawing.Bitmap? _quellBitmap;
        private RandRechteck? _rand;
        private RandRechteck? _basisRand;
        private readonly List<UIElement> _overlayElemente = new List<UIElement>();
        private WpfRect? _overlayRahmen;
        private bool _rahmenSichtbar = true;
        private readonly List<SchnittLinie> _schnittLinien = new List<SchnittLinie>();
        private List<SegmentRechteck> _segmente = new List<SegmentRechteck>();
        private readonly List<UIElement> _linienElemente = new List<UIElement>();
        private readonly List<UIElement> _segmentElemente = new List<UIElement>();

        // ── Scheren-Werkzeug ──────────────────────────────────────────────────────
        private bool _scherenModus;
        private readonly List<int> _scherenschnitte = new List<int>(); // Y-Positionen in Bitmap-Pixel
        private WpfLine? _scherenVorschauLinie;
        private readonly List<UIElement> _scherenElemente = new List<UIElement>(); // Parts-Images + Rahmen + Labels
        private const int ScherenAbstand = 8; // px Abstand zwischen Teilen auf dem Canvas

        // Drag-State Crop
        private bool _isDragging;
        private string? _dragEdge;
        private System.Windows.Point _dragLastPos;

        // Drag-State Schnittlinien
        private bool _isDraggingLinie;
        private SchnittLinie? _dragSchnittLinie;
        private WpfLine? _dragLinieVis;
        private WpfRect? _dragLinieHit;

        // Platzierungs-Modus
        private bool _platzierungsModus;
        private SchnittRichtung _platzierungsRichtung;
        private WpfLine? _platzierungsVorschau;

        // Auswahl Schnittlinie (für Entf-Taste)
        private SchnittLinie? _selectedSchnittLinie;

        // Zoom
        private double _zoomFaktor = 1.0;
        private const double ZoomMin = 0.1, ZoomMax = 5.0, ZoomStep = 0.15;

        // PDF-Felder (Phase 4)
        private string?  _quellPdfPfad;
        private double   _pdfRenderBreite;
        private double   _pdfRenderHöhe;
        private double   _pdfSeitenBreitePt;
        private double   _pdfSeitenHöhePt;
        private int      _pdfSeitenNummer;

        private const int GriffBreite = 14;
        private const int MinGröße = 20;
        private const int LinieGriffBreite = 14;
        private double _sicherheitsabstandMm = 0.0;

        private float Dpi => _quellBitmap?.HorizontalResolution > 0 ? _quellBitmap.HorizontalResolution : 96f;
        private double PxZuMm(int px) => Math.Round(px / (Dpi / 25.4), 1);
        private int MmZuPx(double mm) => (int)Math.Round(mm * Dpi / 25.4);
        private static int Klemme(int v, int min, int max) => v < min ? min : v > max ? max : v;

        public BildschnittPanel()
        {
            InitializeComponent();
            VorschauCanvas.Visibility = Visibility.Collapsed;
            // Fokus setzen wenn auf Canvas geklickt wird (für Entf-Taste)
            VorschauCanvas.MouseDown += (_, _) => Focus();
        }

        // ── Öffentliche Methoden ───────────────────────────────────────────────

        /// <summary>Lädt eine Bild- oder PDF-Datei (für Integration mit DokumenteModul).</summary>
        public void LadeDatei(string pfad)
        {
            var ext = System.IO.Path.GetExtension(pfad).ToLowerInvariant();
            if (ext == ".pdf")
                LadePdfDatei(pfad);
            else
                LadeBildDatei(pfad);
        }

        private void LadeBildDatei(string pfad)
        {
            try
            {
                _quellPdfPfad = null;
                var bmp = new System.Drawing.Bitmap(pfad);
                LadeBitmap(bmp, pfad);
            }
            catch (Exception ex) { SetzeStatus("Fehler: " + ex.Message); }
        }

        /// <summary>Setzt ein Bitmap direkt (für Integration mit DokumenteModul).</summary>
        public void LadeBitmap(System.Drawing.Bitmap bitmap, string? quelle = null)
        {
            _quellBitmap?.Dispose();

            // Scheren-State zurücksetzen
            _scherenschnitte.Clear();
            foreach (var el in _scherenElemente)
                VorschauCanvas.Children.Remove(el);
            _scherenElemente.Clear();
            if (_scherenVorschauLinie != null)
            {
                VorschauCanvas.Children.Remove(_scherenVorschauLinie);
                _scherenVorschauLinie = null;
            }
            VorschauBild.Visibility = Visibility.Visible;
            if (_scherenModus)
            {
                _scherenModus = false;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (BtnSchereToggle.IsChecked == true)
                        BtnSchereToggle.IsChecked = false;
                }));
            }

            _quellBitmap = bitmap;
            _schnittLinien.Clear();
            _segmente.Clear();
            ZeigeBild();
            AutoErkennungDurchführen();
            AktualisiereButtons();
            if (quelle != null)
                SetzeStatus("Bild geladen: " + System.IO.Path.GetFileName(quelle));
        }

        public void Bereinigen()
        {
            _quellBitmap?.Dispose();
            _quellBitmap = null;
        }

        private void SetzeStatus(string text)
        {
            TxtStatus.Text = text;
            AppZustand.Instanz.SetzeStatus(text);
        }

        private void AktualisiereButtons()
        {
            bool hatBild = _quellBitmap != null;
            BtnAutoErkennen.IsEnabled       = hatBild;
            BtnHLinie.IsEnabled             = hatBild && _rand != null;
            BtnVLinie.IsEnabled             = hatBild && _rand != null;
            BtnLinienZurücksetzen.IsEnabled  = hatBild && _schnittLinien.Count > 0;
            BtnAlsPngSpeichern.IsEnabled    = hatBild && _rand != null;
            BtnSegmenteExportieren.IsEnabled = hatBild && _segmente.Count(s => s.Aktiv) > 0;
            BtnAlsPdfSpeichern.IsEnabled    = _quellPdfPfad != null && _rand != null
                                              && System.IO.File.Exists(_quellPdfPfad ?? "");
            bool hatSchnitte = _scherenschnitte.Count > 0;
            BtnSchereToggle.IsEnabled                = hatBild;
            BtnScherenschnitteZurücksetzen.IsEnabled = hatSchnitte;
            BtnTeileExportieren.IsEnabled            = hatSchnitte;
        }

        // ── Zoom (Strg+Mausrad) ────────────────────────────────────────────────

        private void BildScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) return;
            e.Handled = true;
            double delta = e.Delta > 0 ? ZoomStep : -ZoomStep;
            _zoomFaktor = Math.Max(ZoomMin, Math.Min(ZoomMax, _zoomFaktor + delta));
            ZoomTransform.ScaleX = _zoomFaktor;
            ZoomTransform.ScaleY = _zoomFaktor;
        }

        // ── Entf-Taste löscht ausgewählte Schnittlinie ─────────────────────────

        private void UserControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && _selectedSchnittLinie != null)
            {
                _schnittLinien.Remove(_selectedSchnittLinie);
                _selectedSchnittLinie = null;
                AktualisiereLinienUndSegmente();
                AktualisiereButtons();
                SetzeStatus("Schnittlinie gelöscht");
                e.Handled = true;
            }

            // Strg+Z – Letzten Scheren-Schnitt rückgängig
            if (e.Key == Key.Z &&
                (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                if (_scherenschnitte.Count > 0)
                {
                    _scherenschnitte.RemoveAt(_scherenschnitte.Count - 1);
                    AktualisiereScherenCanvas();
                    SetzeStatus($"Schnitt rückgängig. {_scherenschnitte.Count} Schnitt(e) verbleiben.");
                    e.Handled = true;
                }
            }
            // Esc – Scheren-Modus beenden
            if (e.Key == Key.Escape && _scherenModus)
            {
                BtnSchereToggle.IsChecked = false;
                e.Handled = true;
            }
        }

        // ── Bild anzeigen ──────────────────────────────────────────────────────

        private void ZeigeBild()
        {
            if (_quellBitmap == null) return;
            int b = _quellBitmap.Width, h = _quellBitmap.Height;
            VorschauCanvas.Width  = b;
            VorschauCanvas.Height = h;
            VorschauBild.Width    = b;
            VorschauBild.Height   = h;
            VorschauBild.Source   = BitmapZuBitmapSource(_quellBitmap);
            PnlPlatzhalter.Visibility = Visibility.Collapsed;
            VorschauCanvas.Visibility  = Visibility.Visible;
            TxtBildInfo.Text = $"{b} × {h} px  |  {Dpi:0} DPI";
        }

        // ── Auto-Erkennung ─────────────────────────────────────────────────────

        private void AutoErkennungDurchführen()
        {
            if (_quellBitmap == null) return;
            try
            {
                SetzeStatus("Analysiere Ränder …");
                LeseSicherheitsabstand();
                _basisRand = RandErkenner.ErkenneRand(_quellBitmap);
                WendeSicherheitsabstandAn();
                SetzeStatus($"Rand erkannt – Inhalt: {_rand!.Breite} × {_rand.Höhe} px");
            }
            catch (Exception ex)
            {
                SetzeStatus("Fehler bei Randerkennung: " + ex.Message);
            }
        }

        private void WendeSicherheitsabstandAn()
        {
            if (_basisRand == null || _quellBitmap == null) return;
            KlemmeSicherheitsabstand();
            _rand = _basisRand.Clone();
            if (_sicherheitsabstandMm != 0)
            {
                int sp = MmZuPx(_sicherheitsabstandMm);
                _rand.Oben   = Klemme(_rand.Oben   + sp, 0, _quellBitmap.Height - 1);
                _rand.Unten  = Klemme(_rand.Unten  - sp, _rand.Oben + 1, _quellBitmap.Height);
                _rand.Links  = Klemme(_rand.Links  + sp, 0, _quellBitmap.Width - 1);
                _rand.Rechts = Klemme(_rand.Rechts - sp, _rand.Links + 1, _quellBitmap.Width);
            }
            AktualisiereOverlay();
            AktualisiereLinienUndSegmente();
            AktualisiereRandInfo();
            AktualisiereButtons();
        }

        private void LeseSicherheitsabstand()
        {
            string raw = TxtSicherheitsabstand.Text.Replace(",", ".");
            if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double mm))
                _sicherheitsabstandMm = mm;
        }

        private void KlemmeSicherheitsabstand()
        {
            if (_basisRand == null || _quellBitmap == null) return;
            double maxAußenMm = Math.Min(
                Math.Min(PxZuMm(_basisRand.Oben), PxZuMm(_quellBitmap.Height - _basisRand.Unten)),
                Math.Min(PxZuMm(_basisRand.Links), PxZuMm(_quellBitmap.Width - _basisRand.Rechts)));
            double minMm = -maxAußenMm;
            double maxInnenMmH = PxZuMm((_basisRand.Unten - _basisRand.Oben) / 2 - MinGröße / 2);
            double maxInnenMmW = PxZuMm((_basisRand.Rechts - _basisRand.Links) / 2 - MinGröße / 2);
            double maxMm = Math.Max(0, Math.Min(maxInnenMmH, maxInnenMmW));
            _sicherheitsabstandMm = Math.Max(minMm, Math.Min(maxMm, _sicherheitsabstandMm));
        }

        // ── Overlay (Crop-Rahmen + Griffe) ─────────────────────────────────────

        private void AktualisiereOverlay()
        {
            foreach (var el in _overlayElemente)
                VorschauCanvas.Children.Remove(el);
            _overlayElemente.Clear();
            _overlayRahmen = null;

            if (_rand == null || !_rahmenSichtbar) return;

            double l = _rand.Links, o = _rand.Oben, r = _rand.Rechts, u = _rand.Unten;
            double w = r - l, h = u - o;

            _overlayRahmen = new WpfRect
            {
                Width = w, Height = h,
                Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 230, 50)),
                StrokeThickness = 2,
                StrokeDashArray = new DoubleCollection(new[] { 8.0, 4.0 }),
                Fill = System.Windows.Media.Brushes.Transparent,
                IsHitTestVisible = false
            };
            Canvas.SetLeft(_overlayRahmen, l); Canvas.SetTop(_overlayRahmen, o);
            OverlayHinzufügen(_overlayRahmen);

            OverlayHinzufügen(ErstelleCropGriff(l,                     o - GriffBreite / 2.0, w, GriffBreite, "Oben",   Cursors.SizeNS));
            OverlayHinzufügen(ErstelleCropGriff(l,                     u - GriffBreite / 2.0, w, GriffBreite, "Unten",  Cursors.SizeNS));
            OverlayHinzufügen(ErstelleCropGriff(l - GriffBreite / 2.0, o, GriffBreite, h,          "Links",  Cursors.SizeWE));
            OverlayHinzufügen(ErstelleCropGriff(r - GriffBreite / 2.0, o, GriffBreite, h,          "Rechts", Cursors.SizeWE));
        }

        private WpfRect ErstelleCropGriff(double left, double top, double width, double height,
                                           string edge, System.Windows.Input.Cursor cursor)
        {
            var g = new WpfRect
            {
                Width = Math.Max(1, width), Height = Math.Max(1, height),
                Fill = System.Windows.Media.Brushes.Transparent,
                Cursor = cursor, Tag = edge, IsHitTestVisible = true
            };
            Canvas.SetLeft(g, left); Canvas.SetTop(g, top);
            g.MouseEnter += CropGriff_MouseEnter;
            g.MouseLeave += CropGriff_MouseLeave;
            g.MouseLeftButtonDown += CropGriff_MouseDown;
            return g;
        }

        private void OverlayHinzufügen(UIElement el)
        {
            Canvas.SetZIndex(el, 10);
            VorschauCanvas.Children.Add(el);
            _overlayElemente.Add(el);
        }

        private void CropGriff_MouseEnter(object sender, MouseEventArgs e)
        {
            if (_overlayRahmen != null)
                _overlayRahmen.Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 255, 100));
        }

        private void CropGriff_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!_isDragging && _overlayRahmen != null)
                _overlayRahmen.Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 230, 50));
        }

        private void CropGriff_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed) return;
            _dragEdge = (string)((WpfRect)sender).Tag;
            _dragLastPos = e.GetPosition(VorschauCanvas);
            _isDragging = true;
            VorschauCanvas.CaptureMouse();
            VorschauCanvas.PreviewMouseMove += CropDrag_MouseMove;
            VorschauCanvas.PreviewMouseLeftButtonUp += CropDrag_MouseUp;
            e.Handled = true;
        }

        private void CropDrag_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging || _rand == null || _quellBitmap == null) return;
            var pos = e.GetPosition(VorschauCanvas);
            var delta = pos - _dragLastPos;
            _dragLastPos = pos;
            int bW = _quellBitmap.Width, bH = _quellBitmap.Height;
            switch (_dragEdge)
            {
                case "Oben":   _rand.Oben   = Klemme((int)(_rand.Oben   + delta.Y), 0, _rand.Unten  - MinGröße); break;
                case "Unten":  _rand.Unten  = Klemme((int)(_rand.Unten  + delta.Y), _rand.Oben + MinGröße, bH); break;
                case "Links":  _rand.Links  = Klemme((int)(_rand.Links  + delta.X), 0, _rand.Rechts - MinGröße); break;
                case "Rechts": _rand.Rechts = Klemme((int)(_rand.Rechts + delta.X), _rand.Links + MinGröße, bW); break;
            }
            AktualisiereOverlay();
            AktualisiereLinienUndSegmente();
            AktualisiereRandInfo();
        }

        private void CropDrag_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDragging) return;
            _isDragging = false;
            VorschauCanvas.ReleaseMouseCapture();
            VorschauCanvas.PreviewMouseMove -= CropDrag_MouseMove;
            VorschauCanvas.PreviewMouseLeftButtonUp -= CropDrag_MouseUp;
            if (_overlayRahmen != null)
                _overlayRahmen.Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 230, 50));
            _basisRand = _rand?.Clone();
            _sicherheitsabstandMm = 0;
            TxtSicherheitsabstand.Text = "0";
            e.Handled = true;
        }

        // ── Schnittlinien + Segmente ───────────────────────────────────────────

        private void AktualisiereLinienUndSegmente()
        {
            foreach (var el in _linienElemente) VorschauCanvas.Children.Remove(el);
            _linienElemente.Clear();
            foreach (var el in _segmentElemente) VorschauCanvas.Children.Remove(el);
            _segmentElemente.Clear();

            if (_rand == null) return;
            _segmente = SegmentierungsEngine.Segmentiere(_rand, _schnittLinien);

            foreach (var seg in _segmente)
            {
                var rect = new WpfRect
                {
                    Width = seg.Bereich.Width, Height = seg.Bereich.Height,
                    Fill = SegmentFarbe(seg.Aktiv),
                    Stroke = System.Windows.Media.Brushes.Transparent,
                    Tag = seg, Cursor = Cursors.Hand, IsHitTestVisible = true
                };
                Canvas.SetLeft(rect, seg.Bereich.X); Canvas.SetTop(rect, seg.Bereich.Y);
                Canvas.SetZIndex(rect, 2);
                rect.MouseLeftButtonDown += Segment_Geklickt;
                _segmentElemente.Add(rect);
                VorschauCanvas.Children.Add(rect);
            }

            foreach (var linie in _schnittLinien)
                ZeichneSchnittLinie(linie);

            AktualisiereSegmentInfo();
        }

        private static System.Windows.Media.Brush SegmentFarbe(bool aktiv) =>
            aktiv
                ? new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 150, 255))
                : new SolidColorBrush(System.Windows.Media.Color.FromArgb(120, 0, 0, 0));

        private void ZeichneSchnittLinie(SchnittLinie linie)
        {
            if (_rand == null) return;
            bool horiz = linie.Richtung == SchnittRichtung.Horizontal;
            double pos = linie.Position;

            // Sichtbare Linie (orange)
            var vis = new WpfLine
            {
                Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 140, 0)),
                StrokeThickness = 1.5,
                IsHitTestVisible = false,
                SnapsToDevicePixels = true
            };
            if (horiz)
            { vis.X1 = _rand.Links; vis.Y1 = pos; vis.X2 = _rand.Rechts; vis.Y2 = pos; }
            else
            { vis.X1 = pos; vis.Y1 = _rand.Oben; vis.X2 = pos; vis.Y2 = _rand.Unten; }
            Canvas.SetZIndex(vis, 5);
            VorschauCanvas.Children.Add(vis);
            _linienElemente.Add(vis);

            // Hit-Bereich (transparent, breiter als Linie)
            var hit = new WpfRect
            {
                Fill = System.Windows.Media.Brushes.Transparent,
                Cursor = horiz ? Cursors.SizeNS : Cursors.SizeWE,
                Tag = linie, IsHitTestVisible = true
            };
            if (horiz)
            {
                hit.Width = _rand.Breite; hit.Height = LinieGriffBreite;
                Canvas.SetLeft(hit, _rand.Links); Canvas.SetTop(hit, pos - LinieGriffBreite / 2.0);
            }
            else
            {
                hit.Width = LinieGriffBreite; hit.Height = _rand.Höhe;
                Canvas.SetLeft(hit, pos - LinieGriffBreite / 2.0); Canvas.SetTop(hit, _rand.Oben);
            }
            Canvas.SetZIndex(hit, 6);
            hit.MouseLeftButtonDown += SchnittLinieGriff_MouseDown;
            VorschauCanvas.Children.Add(hit);
            _linienElemente.Add(hit);
        }

        private void Segment_Geklickt(object sender, MouseButtonEventArgs e)
        {
            if (_platzierungsModus) return;
            if (sender is WpfRect r && r.Tag is SegmentRechteck seg)
            {
                seg.Aktiv = !seg.Aktiv;
                r.Fill = SegmentFarbe(seg.Aktiv);
                AktualisiereSegmentInfo();
                AktualisiereButtons();
            }
        }

        private void SchnittLinieGriff_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _platzierungsModus) return;
            _dragSchnittLinie = (SchnittLinie)((WpfRect)sender).Tag;
            _selectedSchnittLinie = _dragSchnittLinie;
            _dragLinieVis = _linienElemente.OfType<WpfLine>()
                .FirstOrDefault(l => {
                    bool h = _dragSchnittLinie.Richtung == SchnittRichtung.Horizontal;
                    return h ? Math.Abs(l.Y1 - _dragSchnittLinie.Position) < 2
                             : Math.Abs(l.X1 - _dragSchnittLinie.Position) < 2;
                });
            _dragLinieHit = (WpfRect)sender;
            _isDraggingLinie = true;
            VorschauCanvas.CaptureMouse();
            VorschauCanvas.PreviewMouseMove += LinieDrag_MouseMove;
            VorschauCanvas.PreviewMouseLeftButtonUp += LinieDrag_MouseUp;
            e.Handled = true;
        }

        private void LinieDrag_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingLinie || _dragSchnittLinie == null || _rand == null) return;
            var pos = e.GetPosition(VorschauCanvas);
            bool horiz = _dragSchnittLinie.Richtung == SchnittRichtung.Horizontal;
            int newPos;
            if (horiz) newPos = Klemme((int)pos.Y, _rand.Oben + 1, _rand.Unten - 1);
            else        newPos = Klemme((int)pos.X, _rand.Links + 1, _rand.Rechts - 1);
            _dragSchnittLinie.Position = newPos;

            // Live-Update der sichtbaren Linie
            if (_dragLinieVis != null)
            {
                if (horiz) { _dragLinieVis.Y1 = newPos; _dragLinieVis.Y2 = newPos; }
                else       { _dragLinieVis.X1 = newPos; _dragLinieVis.X2 = newPos; }
            }
            if (_dragLinieHit != null)
            {
                if (horiz) Canvas.SetTop(_dragLinieHit, newPos - LinieGriffBreite / 2.0);
                else       Canvas.SetLeft(_dragLinieHit, newPos - LinieGriffBreite / 2.0);
            }
        }

        private void LinieDrag_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingLinie) return;
            _isDraggingLinie = false;
            VorschauCanvas.ReleaseMouseCapture();
            VorschauCanvas.PreviewMouseMove -= LinieDrag_MouseMove;
            VorschauCanvas.PreviewMouseLeftButtonUp -= LinieDrag_MouseUp;
            AktualisiereLinienUndSegmente();
        }

        // ── Platzierungs-Modus ────────────────────────────────────────────────

        private void StartePlatzierungsModus(SchnittRichtung richtung)
        {
            if (_rand == null) return;
            _platzierungsModus = true;
            _platzierungsRichtung = richtung;
            bool horiz = richtung == SchnittRichtung.Horizontal;

            _platzierungsVorschau = new WpfLine
            {
                Stroke = new SolidColorBrush(System.Windows.Media.Color.FromArgb(160, 255, 140, 0)),
                StrokeThickness = 1.5, StrokeDashArray = new DoubleCollection(new[] { 6.0, 3.0 }),
                IsHitTestVisible = false
            };
            if (horiz)
            { _platzierungsVorschau.X1 = _rand.Links; _platzierungsVorschau.X2 = _rand.Rechts; }
            else
            { _platzierungsVorschau.Y1 = _rand.Oben; _platzierungsVorschau.Y2 = _rand.Unten; }
            Canvas.SetZIndex(_platzierungsVorschau, 20);
            VorschauCanvas.Children.Add(_platzierungsVorschau);

            VorschauCanvas.Cursor = horiz ? Cursors.SizeNS : Cursors.SizeWE;
            VorschauCanvas.MouseMove += Platzierung_MouseMove;
            VorschauCanvas.MouseLeftButtonDown += Platzierung_MouseDown;
            VorschauCanvas.MouseRightButtonDown += Platzierung_Abbrechen;
        }

        private void Platzierung_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_platzierungsModus || _platzierungsVorschau == null || _rand == null) return;
            var pos = e.GetPosition(VorschauCanvas);
            bool horiz = _platzierungsRichtung == SchnittRichtung.Horizontal;
            if (horiz)
            { double y = Klemme((int)pos.Y, _rand.Oben + 1, _rand.Unten - 1);
              _platzierungsVorschau.Y1 = y; _platzierungsVorschau.Y2 = y; }
            else
            { double x = Klemme((int)pos.X, _rand.Links + 1, _rand.Rechts - 1);
              _platzierungsVorschau.X1 = x; _platzierungsVorschau.X2 = x; }
        }

        private void Platzierung_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!_platzierungsModus || _rand == null) return;
            var pos = e.GetPosition(VorschauCanvas);
            bool horiz = _platzierungsRichtung == SchnittRichtung.Horizontal;
            int newPos;
            if (horiz) newPos = Klemme((int)pos.Y, _rand.Oben + 1, _rand.Unten - 1);
            else        newPos = Klemme((int)pos.X, _rand.Links + 1, _rand.Rechts - 1);
            _schnittLinien.Add(new SchnittLinie(_platzierungsRichtung, newPos));
            BeendePlatzierungsModus();
            AktualisiereLinienUndSegmente();
            AktualisiereButtons();
            e.Handled = true;
        }

        private void Platzierung_Abbrechen(object sender, MouseButtonEventArgs e)
        {
            if (_platzierungsModus) { BeendePlatzierungsModus(); e.Handled = true; }
        }

        private void BeendePlatzierungsModus()
        {
            _platzierungsModus = false;
            if (_platzierungsVorschau != null)
            {
                VorschauCanvas.Children.Remove(_platzierungsVorschau);
                _platzierungsVorschau = null;
            }
            VorschauCanvas.Cursor = Cursors.Arrow;
            VorschauCanvas.MouseMove -= Platzierung_MouseMove;
            VorschauCanvas.MouseLeftButtonDown -= Platzierung_MouseDown;
            VorschauCanvas.MouseRightButtonDown -= Platzierung_Abbrechen;
        }

        // ── Info-Texte ─────────────────────────────────────────────────────────

        private void AktualisiereRandInfo()
        {
            if (_rand == null || _quellBitmap == null) { TxtRandInfo.Text = "–"; return; }
            TxtRandInfo.Text = $"L: {_rand.Links} px  R: {_rand.Rechts} px\n" +
                               $"O: {_rand.Oben} px   U: {_rand.Unten} px\n" +
                               $"Breite: {PxZuMm(_rand.Breite):0.#} mm\n" +
                               $"Höhe:   {PxZuMm(_rand.Höhe):0.#} mm";
        }

        private void AktualisiereSegmentInfo()
        {
            int gesamt = _segmente.Count;
            int aktiv  = _segmente.Count(s => s.Aktiv);
            TxtSegmentInfo.Text = gesamt == 0
                ? "Keine Schnittlinien"
                : $"{gesamt} Segment(e)\n{aktiv} aktiv";
        }

        // ── Button-Handler ─────────────────────────────────────────────────────

        public void BtnBildLaden_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Bild laden",
                Filter = "Bilder|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|Alle Dateien|*.*"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                _quellPdfPfad = null;
                var bmp = new System.Drawing.Bitmap(dlg.FileName);
                LadeBitmap(bmp, dlg.FileName);
            }
            catch (Exception ex)
            {
                SetzeStatus("Fehler beim Laden: " + ex.Message);
            }
        }

        private void BtnAutoErkennen_Click(object sender, RoutedEventArgs e)
            => AutoErkennungDurchführen();

        private void BtnHLinie_Click(object sender, RoutedEventArgs e)
        {
            if (_rand == null) return;
            StartePlatzierungsModus(SchnittRichtung.Horizontal);
            SetzeStatus("Klick auf Bild, um H-Linie zu platzieren (Rechtsklick: abbrechen)");
        }

        private void BtnVLinie_Click(object sender, RoutedEventArgs e)
        {
            if (_rand == null) return;
            StartePlatzierungsModus(SchnittRichtung.Vertikal);
            SetzeStatus("Klick auf Bild, um V-Linie zu platzieren (Rechtsklick: abbrechen)");
        }

        private void BtnLinienZurücksetzen_Click(object sender, RoutedEventArgs e)
        {
            _schnittLinien.Clear();
            _selectedSchnittLinie = null;
            AktualisiereLinienUndSegmente();
            AktualisiereButtons();
            SetzeStatus("Alle Schnittlinien gelöscht");
        }

        private void BtnAlsPngSpeichern_Click(object sender, RoutedEventArgs e)
        {
            if (_quellBitmap == null || _rand == null) return;
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Zugeschnittenes Bild speichern",
                Filter = "PNG-Bild|*.png",
                DefaultExt = ".png"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                var cropRect = _rand.AlsDrawingRect();
                using var cropped = _quellBitmap.Clone(cropRect, _quellBitmap.PixelFormat);
                cropped.Save(dlg.FileName, ImageFormat.Png);
                SetzeStatus("Gespeichert: " + dlg.FileName);
            }
            catch (Exception ex)
            {
                SetzeStatus("Fehler beim Speichern: " + ex.Message);
            }
        }

        private void BtnSegmenteExportieren_Click(object sender, RoutedEventArgs e)
        {
            if (_quellBitmap == null) return;
            var aktiveSegmente = _segmente.Where(s => s.Aktiv).ToList();
            if (aktiveSegmente.Count == 0) { SetzeStatus("Keine aktiven Segmente zum Exportieren"); return; }

            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Ordner für Segment-Export wählen",
                ShowNewFolderButton = true
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            int erfolgreich = 0;
            foreach (var seg in aktiveSegmente)
            {
                try
                {
                    using var segBmp = _quellBitmap.Clone(seg.Bereich, _quellBitmap.PixelFormat);
                    string dateiName = $"Segment_{seg.Zeile + 1}_{seg.Spalte + 1}.png";
                    string zielPfad  = System.IO.Path.Combine(dlg.SelectedPath, dateiName);
                    segBmp.Save(zielPfad, ImageFormat.Png);
                    erfolgreich++;
                }
                catch (Exception ex) { Logger.Fehler("SegmentExport", ex.Message); }
            }
            SetzeStatus($"{erfolgreich} von {aktiveSegmente.Count} Segment(e) exportiert → {dlg.SelectedPath}");
        }

        private void BtnAlsPdfSpeichern_Click(object sender, RoutedEventArgs e)
        {
            if (_quellPdfPfad == null || _rand == null || _quellBitmap == null) return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Gecroptes PDF speichern",
                Filter = "PDF-Dokument|*.pdf",
                DefaultExt = ".pdf",
                FileName = System.IO.Path.GetFileNameWithoutExtension(_quellPdfPfad) + "_crop.pdf"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                // Skalierungsfaktor: Pixel → PDF-Punkte
                double scale = _pdfSeitenBreitePt / _pdfRenderBreite;

                // PDF Y-Achse ist bottom-up (Ursprung unten-links)
                double x1 = _rand.Links  * scale;
                double y1 = (_pdfRenderHöhe - _rand.Unten)  * scale;   // flip Y
                double x2 = _rand.Rechts * scale;
                double y2 = (_pdfRenderHöhe - _rand.Oben)   * scale;   // flip Y

                // Sicherstellen dass x1<x2 und y1<y2
                if (x1 > x2) { double t = x1; x1 = x2; x2 = t; }
                if (y1 > y2) { double t = y1; y1 = y2; y2 = t; }

                using var pdfIn = PdfReader.Open(_quellPdfPfad, PdfDocumentOpenMode.Import);
                var pdfOut = new PdfDocument();
                var seite  = pdfOut.AddPage(pdfIn.Pages[_pdfSeitenNummer]);
                seite.CropBox = new PdfSharp.Pdf.PdfRectangle(
                    new XPoint(x1, y1),
                    new XPoint(x2, y2));

                pdfOut.Save(dlg.FileName);
                SetzeStatus("PDF gespeichert: " + dlg.FileName);
            }
            catch (Exception ex)
            {
                SetzeStatus("Fehler: " + ex.Message);
                Logger.Fehler("BtnAlsPdfSpeichern", ex.Message);
            }
        }

        private void BtnSpinUp_Click(object sender, RoutedEventArgs e)
        {
            LeseSicherheitsabstand();
            _sicherheitsabstandMm = Math.Round(_sicherheitsabstandMm + 1.0, 1);
            KlemmeSicherheitsabstand();
            TxtSicherheitsabstand.Text = _sicherheitsabstandMm.ToString("0.#");
            WendeSicherheitsabstandAn();
        }

        private void BtnSpinDown_Click(object sender, RoutedEventArgs e)
        {
            LeseSicherheitsabstand();
            _sicherheitsabstandMm = Math.Round(_sicherheitsabstandMm - 1.0, 1);
            KlemmeSicherheitsabstand();
            TxtSicherheitsabstand.Text = _sicherheitsabstandMm.ToString("0.#");
            WendeSicherheitsabstandAn();
        }

        // ── Phase 4: PDF laden ────────────────────────────────────────────────

        private void LadePdfDatei(string pfad)
        {
            _quellPdfPfad = pfad;
            _pdfSeitenNummer = 0;
            SetzeStatus("Rendere PDF-Seite …");
            AppZustand.Instanz.SetzeStatus("Rendere PDF …");

            var thread = new Thread(() =>
            {
                System.Drawing.Bitmap? bmp = null;
                string? fehler = null;
                double seitenBreitePt = 0, seitenHöhePt = 0;

                // Seitengröße aus PdfSharp ermitteln (kein pdfium → kein Semaphore nötig)
                try
                {
                    var pdfDoc = PdfReader.Open(pfad, PdfDocumentOpenMode.InformationOnly);
                    if (pdfDoc.PageCount > 0)
                    {
                        seitenBreitePt = pdfDoc.Pages[0].Width.Point;
                        seitenHöhePt   = pdfDoc.Pages[0].Height.Point;
                    }
                }
                catch (Exception ex) { fehler = ex.Message; }

                // Rendern mit PdfRenderer (1200px Breite)
                try
                {
                    AppZustand.RenderSem.Wait();
                    try
                    {
                        const int renderBreite = 1200;
                        var seiten = PdfRenderer.RenderiereAlleSeiten(pfad, renderBreite);
                        if (seiten.Count > 0)
                            bmp = BitmapSourceZuBitmap(seiten[0]);
                    }
                    finally { AppZustand.RenderSem.Release(); }
                }
                catch (Exception ex) { fehler = ex.Message; }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (bmp != null)
                    {
                        _pdfSeitenBreitePt = seitenBreitePt > 0 ? seitenBreitePt : bmp.Width * 72.0 / 96.0;
                        _pdfSeitenHöhePt   = seitenHöhePt   > 0 ? seitenHöhePt   : bmp.Height * 72.0 / 96.0;
                        _pdfRenderBreite = bmp.Width;
                        _pdfRenderHöhe   = bmp.Height;
                        LadeBitmap(bmp, pfad);
                    }
                    else
                    {
                        SetzeStatus("Fehler beim PDF-Rendern: " + (fehler ?? "unbekannt"));
                    }
                }));
            });
            thread.IsBackground = true;
            thread.Name = "BildschnittPdfLaden";
            thread.Start();
        }

        private static System.Drawing.Bitmap BitmapSourceZuBitmap(BitmapSource bmpSrc)
        {
            using var ms = new MemoryStream();
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmpSrc));
            encoder.Save(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return new System.Drawing.Bitmap(ms);
        }

        // ── Scheren-Werkzeug ──────────────────────────────────────────────────────

        // Gibt für jeden Teil: (BitmapY_Start, Hoehe_Pixel, CanvasY_Start) zurück.
        private List<(int BmpY, int Hoehe, double CanvasY)> BerechneTeilLayout()
        {
            var liste = new List<(int, int, double)>();
            if (_quellBitmap == null) return liste;

            var grenzen = new List<int> { 0 };
            grenzen.AddRange(_scherenschnitte.OrderBy(y => y).Where(y => y > 0 && y < _quellBitmap.Height));
            grenzen.Add(_quellBitmap.Height);

            double canvasY = 0;
            for (int i = 0; i < grenzen.Count - 1; i++)
            {
                int bmpY  = grenzen[i];
                int hoehe = grenzen[i + 1] - bmpY;
                if (hoehe <= 0) continue;
                liste.Add((bmpY, hoehe, canvasY));
                canvasY += hoehe + ScherenAbstand;
            }
            return liste;
        }

        private int CanvasYZuBitmapY(double canvasY)
        {
            var layout = BerechneTeilLayout();
            if (layout.Count == 0) return (int)canvasY;

            foreach (var (bmpY, hoehe, teilCanvasY) in layout)
            {
                if (canvasY >= teilCanvasY && canvasY < teilCanvasY + hoehe)
                    return bmpY + (int)(canvasY - teilCanvasY);
            }
            // In Abstand-Lücke: nearest boundary
            for (int i = 0; i < layout.Count - 1; i++)
            {
                var cur  = layout[i];
                var next = layout[i + 1];
                double gapMitte = cur.CanvasY + cur.Hoehe + ScherenAbstand / 2.0;
                if (canvasY < gapMitte) return cur.BmpY + cur.Hoehe - 1;
                if (canvasY < next.CanvasY) return next.BmpY;
            }
            var last = layout[layout.Count - 1];
            return last.BmpY + last.Hoehe - 1;
        }

        private void AktualisiereScherenCanvas()
        {
            // Alte Scheren-Elemente entfernen
            foreach (var el in _scherenElemente)
                VorschauCanvas.Children.Remove(el);
            _scherenElemente.Clear();

            if (_quellBitmap == null) return;

            if (_scherenschnitte.Count == 0)
            {
                // Kein Schnitt: Original-Bild wieder sichtbar, Canvas-Größe zurücksetzen
                VorschauBild.Visibility   = Visibility.Visible;
                VorschauCanvas.Width  = _quellBitmap.Width;
                VorschauCanvas.Height = _quellBitmap.Height;
                TxtScherenInfo.Text = "Keine Schnitte";
                return;
            }

            // Original-Bild ausblenden (Teile werden separat gerendert)
            VorschauBild.Visibility = Visibility.Collapsed;

            var layout = BerechneTeilLayout();
            double totalHöhe = layout.Sum(t => (double)t.Hoehe) + ScherenAbstand * (layout.Count - 1);
            VorschauCanvas.Width  = _quellBitmap.Width;
            VorschauCanvas.Height = totalHöhe;

            for (int i = 0; i < layout.Count; i++)
            {
                var (bmpY, hoehe, canvasY) = layout[i];
                var cropRect = new DrawRect(0, bmpY, _quellBitmap.Width, hoehe);

                BitmapSource src;
                using (var teilBmp = _quellBitmap.Clone(cropRect, _quellBitmap.PixelFormat))
                    src = BitmapZuBitmapSource(teilBmp);

                // Bild-Element
                var img = new System.Windows.Controls.Image
                {
                    Width   = _quellBitmap.Width,
                    Height  = hoehe,
                    Source  = src,
                    Stretch = Stretch.None,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(img, 0); Canvas.SetTop(img, canvasY);
                Canvas.SetZIndex(img, 1);
                VorschauCanvas.Children.Add(img);
                _scherenElemente.Add(img);

                // Dünner Rahmen
                var rahmen = new WpfRect
                {
                    Width           = _quellBitmap.Width,
                    Height          = hoehe,
                    Stroke          = new SolidColorBrush(System.Windows.Media.Color.FromArgb(160, 180, 180, 180)),
                    StrokeThickness = 1,
                    Fill            = System.Windows.Media.Brushes.Transparent,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(rahmen, 0); Canvas.SetTop(rahmen, canvasY);
                Canvas.SetZIndex(rahmen, 3);
                VorschauCanvas.Children.Add(rahmen);
                _scherenElemente.Add(rahmen);

                // Nummern-Label
                var label = new System.Windows.Controls.TextBlock
                {
                    Text       = (i + 1).ToString(),
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(180, 30, 30, 30)),
                    FontSize   = 11,
                    FontWeight = FontWeights.SemiBold,
                    Padding    = new Thickness(5, 1, 5, 1)
                };
                Canvas.SetLeft(label, 5); Canvas.SetTop(label, canvasY + 4);
                Canvas.SetZIndex(label, 4);
                VorschauCanvas.Children.Add(label);
                _scherenElemente.Add(label);
            }

            TxtScherenInfo.Text = $"{layout.Count} Teile\n{_scherenschnitte.Count} Schnitt(e)";
            AktualisiereButtons();
        }

        private void BtnSchereToggle_Checked(object sender, RoutedEventArgs e)
        {
            // Platzierungs-Modus beenden falls aktiv
            if (_platzierungsModus) BeendePlatzierungsModus();

            _scherenModus = true;
            VorschauCanvas.Cursor = Cursors.Cross;
            VorschauCanvas.MouseMove            += Schere_MouseMove;
            VorschauCanvas.MouseLeftButtonDown  += Schere_MouseDown;
            VorschauCanvas.MouseRightButtonDown += Schere_RechtsklickAbbrechen;
            SetzeStatus("Scheren-Modus aktiv – Klick setzt Schnitt, Rechtsklick/Esc beendet");
        }

        private void BtnSchereToggle_Unchecked(object sender, RoutedEventArgs e)
            => BeendeScherenModus();

        private void BeendeScherenModus()
        {
            _scherenModus = false;
            if (_scherenVorschauLinie != null)
            {
                VorschauCanvas.Children.Remove(_scherenVorschauLinie);
                _scherenVorschauLinie = null;
            }
            VorschauCanvas.Cursor = Cursors.Arrow;
            VorschauCanvas.MouseMove            -= Schere_MouseMove;
            VorschauCanvas.MouseLeftButtonDown  -= Schere_MouseDown;
            VorschauCanvas.MouseRightButtonDown -= Schere_RechtsklickAbbrechen;

            int anzahl = _scherenschnitte.Count;
            SetzeStatus(anzahl > 0 ? $"Scheren-Modus beendet. {anzahl} Schnitt(e) gesetzt." : "Scheren-Modus beendet.");
        }

        private void Schere_MouseMove(object sender, MouseEventArgs e)
        {
            if (_quellBitmap == null) return;
            var pos = e.GetPosition(VorschauCanvas);

            if (_scherenVorschauLinie == null)
            {
                _scherenVorschauLinie = new WpfLine
                {
                    Stroke          = new SolidColorBrush(System.Windows.Media.Color.FromArgb(210, 220, 40, 40)),
                    StrokeThickness = 1.5,
                    StrokeDashArray = new DoubleCollection(new[] { 10.0, 5.0 }),
                    IsHitTestVisible = false
                };
                Canvas.SetZIndex(_scherenVorschauLinie, 50);
                VorschauCanvas.Children.Add(_scherenVorschauLinie);
            }

            double canvasW = VorschauCanvas.Width > 0 ? VorschauCanvas.Width : _quellBitmap.Width;
            _scherenVorschauLinie.X1 = 0;       _scherenVorschauLinie.Y1 = pos.Y;
            _scherenVorschauLinie.X2 = canvasW; _scherenVorschauLinie.Y2 = pos.Y;
        }

        private void Schere_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_quellBitmap == null || e.LeftButton != MouseButtonState.Pressed) return;
            var pos     = e.GetPosition(VorschauCanvas);
            int bitmapY = CanvasYZuBitmapY(pos.Y);

            // Außerhalb des Bitmaps → ignorieren
            if (bitmapY <= 0 || bitmapY >= _quellBitmap.Height - 1)
            {
                SetzeStatus("Schnitt außerhalb des Bildes – ignoriert");
                return;
            }
            // Zu nahe an bestehendem Schnitt (< 8px) → ignorieren
            if (_scherenschnitte.Any(s => Math.Abs(s - bitmapY) < 8))
            {
                SetzeStatus("Zu nahe an bestehendem Schnitt – ignoriert");
                return;
            }

            _scherenschnitte.Add(bitmapY);
            AktualisiereScherenCanvas();
            SetzeStatus($"Schnitt {_scherenschnitte.Count} bei Y={bitmapY} px gesetzt");
            e.Handled = true;
        }

        private void Schere_RechtsklickAbbrechen(object sender, MouseButtonEventArgs e)
        {
            if (_scherenModus)
            {
                BtnSchereToggle.IsChecked = false; // löst Unchecked-Event aus → BeendeScherenModus()
                e.Handled = true;
            }
        }

        private void BtnScherenschnitteZurücksetzen_Click(object sender, RoutedEventArgs e)
        {
            if (_scherenModus)
            {
                BtnSchereToggle.IsChecked = false; // BeendeScherenModus()
            }
            _scherenschnitte.Clear();
            foreach (var el in _scherenElemente)
                VorschauCanvas.Children.Remove(el);
            _scherenElemente.Clear();

            if (_quellBitmap != null)
            {
                VorschauBild.Visibility   = Visibility.Visible;
                VorschauCanvas.Width  = _quellBitmap.Width;
                VorschauCanvas.Height = _quellBitmap.Height;
            }
            TxtScherenInfo.Text = "Keine Schnitte";
            AktualisiereButtons();
            SetzeStatus("Alle Schnitte entfernt");
        }

        private void BtnTeileExportieren_Click(object sender, RoutedEventArgs e)
        {
            if (_quellBitmap == null || _scherenschnitte.Count == 0) return;

            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Ordner für Teil-Export wählen",
                ShowNewFolderButton = true
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            var layout   = BerechneTeilLayout();
            string basis = _quellPdfPfad != null
                ? System.IO.Path.GetFileNameWithoutExtension(_quellPdfPfad)
                : "bild";

            int ok = 0;
            for (int i = 0; i < layout.Count; i++)
            {
                try
                {
                    var (bmpY, hoehe, _) = layout[i];
                    var cropRect = new DrawRect(0, bmpY, _quellBitmap.Width, hoehe);
                    using var teilBmp = _quellBitmap.Clone(cropRect, _quellBitmap.PixelFormat);
                    string datei = System.IO.Path.Combine(dlg.SelectedPath, $"{basis}_teil{i + 1}.png");
                    teilBmp.Save(datei, ImageFormat.Png);
                    ok++;
                }
                catch (Exception ex) { Logger.Fehler("TeileExport", ex.Message); }
            }
            SetzeStatus($"{ok}/{layout.Count} Teile exportiert → {dlg.SelectedPath}");
        }

        // ── Hilfsmethoden ──────────────────────────────────────────────────────

        private static BitmapSource BitmapZuBitmapSource(System.Drawing.Bitmap bmp)
        {
            using var ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Bmp);
            ms.Seek(0, SeekOrigin.Begin);
            var bs = new BitmapImage();
            bs.BeginInit();
            bs.CacheOption = BitmapCacheOption.OnLoad;
            bs.StreamSource = ms;
            bs.EndInit();
            bs.Freeze();
            return bs;
        }
    }
}
