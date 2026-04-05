using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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
        }

        // Öffentliche Methode: Bild von außen setzen (für Integration mit DokumenteModul)
        public void LadeBitmap(System.Drawing.Bitmap bitmap, string? quelle = null)
        {
            _quellBitmap?.Dispose();
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
            }
        }

        private void SchnittLinieGriff_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _platzierungsModus) return;
            _dragSchnittLinie = (SchnittLinie)((WpfRect)sender).Tag;
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
