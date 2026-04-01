using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using IO = System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Modules.Werkzeuge
{
    /// <summary>
    /// PDF-Vorschau mit Beschnittrahmen und Word-Export.
    /// Zoom per Strg+Mausrad, verschiebbarer Beschnittrahmen, präziser Word-Export.
    /// </summary>
    public partial class PdfSchnittEditor : UserControl
    {
        // ── Render-Konstanten ─────────────────────────────────────────────────
        private const int    RenderBreite  = 900;
        private const double ExportDpi     = 150.0;  // DPI für Export-Render
        private const double SeitenAbstand = 30;
        private const int    SeiteX        = 20;

        // Zoom-Grenzen
        private const double ZoomMin  = 0.20;
        private const double ZoomMax  = 4.00;
        private const double ZoomStep = 0.15;

        public PdfSchnittEditor() => InitializeComponent();

        // ── Zustand ───────────────────────────────────────────────────────────
        private string?            _pdfPfad;
        private List<BitmapSource> _seitenBilder  = new();
        private double[]           _seitenYStart   = Array.Empty<double>();
        private double[]           _seitenHöhe     = Array.Empty<double>();

        private double _zoomFaktor = 1.0;

        // Crop-Ränder als Bruchteil der Seitenbreite/-höhe (0 = kein Rand, 0.5 = halbe Seite)
        private double _cropLinks  = 0.0;
        private double _cropRechts = 0.0;
        private double _cropOben   = 0.0;
        private double _cropUnten  = 0.0;

        // Sicherheitsabstand für automatische Rand-Erkennung (Vorschau-Pixel, min 0, max 50)
        private double _cropSicherheitMm  = 2.0;   // Sicherheitsabstand fuer Auto-Rand (mm)
        private double _pxPerMm           = 4.0;   // Pixel pro mm (beim PDF-Laden berechnet)
        private bool   _autoRandAktiv;              // true wenn Auto-Rand zuletzt angewendet

        // Drag-State für Crop-Linien (Tag-Wert der gezogenen Linie)
        private string? _gezogeneCropSeite;

        // Export-Sperre: verhindert Doppelstart
        private volatile bool _exportLäuft;

        // Abbruch laufender Lade-Aufträge
        private CancellationTokenSource? _ladeCts;

        // ── Fehler-Hilfsmethoden ──────────────────────────────────────────────

        private static void LogException(Exception ex, string kontext)
            => App.LogFehler(kontext, App.GetExceptionKette(ex));

        private bool SafeExecute(Action aktion, string kontext)
        {
            try { aktion(); return true; }
            catch (Exception ex) { LogException(ex, kontext); return false; }
        }

        // ── Öffentliche API ───────────────────────────────────────────────────

        public void LadePdf(string pfad)
        {
            if (string.IsNullOrEmpty(pfad) || !IO.File.Exists(pfad))
            {
                TxtInfo.Text = "Datei nicht gefunden.";
                return;
            }

            _ladeCts?.Cancel();
            var cts = new CancellationTokenSource();
            _ladeCts = cts;
            var token = cts.Token;

            _pdfPfad = pfad;
            PdfCanvas.Children.Clear();
            _seitenBilder.Clear();
            _seitenYStart = Array.Empty<double>();
            _seitenHöhe   = Array.Empty<double>();
            _cropLinks = _cropRechts = _cropOben = _cropUnten = 0.0;
            TxtInfo.Text        = "Lade PDF …";
            BtnExport.IsEnabled = false;

            LogException(new Exception($"[DEBUG] Starte PDF-Laden: {IO.Path.GetFileName(pfad)}"), "LadePdf");

            string pfadKopie = pfad;

            Task.Run(() =>
            {
                List<BitmapSource>? bilder = null;
                double[]? yStart = null, höhe = null;
                string? fehler = null;

                try
                {
                    if (token.IsCancellationRequested) return;
                    bilder = RenderiereAlleSeiten(pfadKopie, RenderBreite, token: token);
                    if (token.IsCancellationRequested) return;
                    if (bilder.Count == 0) { fehler = "PDF enthält keine lesbaren Seiten."; }
                    else
                    {
                        BerechneLayoutStatic(bilder, SeitenAbstand, out yStart, out höhe);
                    }
                }
                catch (OperationCanceledException) { return; }
                catch (Exception ex) { LogException(ex, "LadePdf"); fehler = App.GetExceptionKette(ex); }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (token.IsCancellationRequested) return;
                    try
                    {
                        if (fehler != null)
                        {
                            TxtInfo.Text = "Ladefehler. Siehe Log.";
                            MessageBox.Show(fehler, "PDF-Schnitt-Editor",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        _seitenBilder = bilder!;
                        _seitenYStart = yStart!;
                        _seitenHöhe   = höhe!;
                        ZeicheCanvas();
                        // Pixel-pro-mm fuer Sicherheitsabstand-Konvertierung
                        if (_pdfPfad != null && bilder!.Count > 0)
                        {
                            var (wPts, _) = HolePdfSeitenGrösse(_pdfPfad);
                            _pxPerMm = wPts > 0 ? bilder![0].PixelWidth / (wPts / 72.0 * 25.4) : 4.0;
                        }
                        BtnExport.IsEnabled = true;
                        TxtInfo.Text = $"{bilder!.Count} Seite(n) geladen";
                    }
                    catch (Exception ex) { LogException(ex, "LadePdf/UI"); }
                }));
            });
        }

        public void ExportierenNachWord()
        {
            // Doppelstart verhindern: kein paralleler Word-Vorgang erlaubt
            if (_exportLäuft)
            {
                MessageBox.Show(
                    "Solange Word geöffnet wird oder ein Word-Vorgang aktiv ist,\n" +
                    "kann kein weiteres Word-Dokument erstellt werden.",
                    "Word beschäftigt",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (_seitenBilder.Count == 0 || _pdfPfad == null) return;

            string zielPfad = IO.Path.Combine(
                IO.Path.GetDirectoryName(_pdfPfad)!,
                IO.Path.GetFileNameWithoutExtension(_pdfPfad) + ".docx");

            // Überschreiben-Abfrage VOR dem Sperren
            if (IO.File.Exists(zielPfad))
            {
                var antwort = MessageBox.Show(
                    $"Die Datei existiert bereits:\n{IO.Path.GetFileName(zielPfad)}\n\nMöchten Sie sie überschreiben?",
                    "Datei überschreiben?",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (antwort != MessageBoxResult.Yes) return;
            }

            // Ab hier gesperrt
            _exportLäuft = true;
            BtnExport.IsEnabled = false;
            TxtInfo.Text = "Export gestartet …";

            var yStartK  = (double[])_seitenYStart.Clone();
            var höheK    = (double[])_seitenHöhe.Clone();
            string pdfK  = _pdfPfad;
            double cropL = _cropLinks, cropR = _cropRechts, cropO = _cropOben, cropU = _cropUnten;

            var thread = new Thread(() =>
                ExportThreadWorker(zielPfad, pdfK, yStartK, höheK, cropL, cropR, cropO, cropU))
            { IsBackground = true, Name = "PdfSchnittExport" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        // ── Rendering ────────────────────────────────────────────────────────

        private static List<BitmapSource> RenderiereAlleSeiten(string pfad, int breite, int höhe = 0,
                                                               CancellationToken token = default)
        {
            if (höhe <= 0) höhe = breite * 2;
            var result = new List<BitmapSource>();
            try
            {
                using var lib       = DocLib.Instance;
                // Docnet erfordert dimOne <= dimTwo (Breite <= Höhe).
                // Bei Querformat-PDFs tauschen, damit der Constraint nicht verletzt wird.
                // Die tatsächliche Render-Größe liefert GetPageWidth/Height().
                int dimMin = Math.Min(breite, höhe);
                int dimMax = Math.Max(breite, höhe);
                using var docReader = lib.GetDocReader(pfad, new PageDimensions(dimMin, dimMax));
                int n = docReader.GetPageCount();
                for (int i = 0; i < n; i++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        using var pageReader = docReader.GetPageReader(i);
                        byte[]? raw = pageReader.GetImage();
                        int w = pageReader.GetPageWidth(), h = pageReader.GetPageHeight();
                        if (raw == null || w <= 0 || h <= 0 || raw.Length < w * h * 4) continue;

                        // Transparente Pixel gegen Weiß kompositionieren
                        KompositioniereGegenWeiss(raw, w, h);

                        var bmp = BitmapSource.Create(w, h, 96, 96,
                            PixelFormats.Bgra32, null, raw, w * 4);
                        bmp.Freeze();
                        result.Add(bmp);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { LogException(ex, $"Seite[{i}]"); }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { LogException(ex, "RenderiereAlleSeiten"); }
            return result;
        }

        private static void KompositioniereGegenWeiss(byte[] raw, int w, int h)
        {
            int maxOff = Math.Min(w * h * 4, raw.Length - 3);
            for (int p = 0; p < maxOff; p += 4)
            {
                byte a = raw[p + 3];
                if (a == 255) continue;
                if (a == 0) { raw[p] = raw[p+1] = raw[p+2] = raw[p+3] = 255; }
                else
                {
                    float af = a / 255f, inv = 1f - af;
                    raw[p]   = (byte)(raw[p]   * af + 255f * inv);
                    raw[p+1] = (byte)(raw[p+1] * af + 255f * inv);
                    raw[p+2] = (byte)(raw[p+2] * af + 255f * inv);
                    raw[p+3] = 255;
                }
            }
        }

        // ── Layout ────────────────────────────────────────────────────────────

        private static void BerechneLayoutStatic(
            List<BitmapSource> bilder, double abstand,
            out double[] yStart, out double[] höhe)
        {
            yStart = new double[bilder.Count];
            höhe   = new double[bilder.Count];
            double y = 0;
            for (int i = 0; i < bilder.Count; i++)
            {
                yStart[i] = y;
                höhe[i]   = Math.Max(1, bilder[i].PixelHeight);
                y += höhe[i] + abstand;
            }
        }
        // ── Crop-Linien ───────────────────────────────────────────────────────


        // ── Canvas zeichnen ───────────────────────────────────────────────────

        private void ZeicheCanvas()
        {
            try
            {
                PdfCanvas.Children.Clear();
                if (_seitenBilder.Count == 0) return;
                if (_seitenYStart.Length != _seitenBilder.Count ||
                    _seitenHöhe.Length   != _seitenBilder.Count) return;

                int    last = _seitenYStart.Length - 1;
                double gesamtH  = _seitenYStart[last] + _seitenHöhe[last] + SeitenAbstand;
                double maxBmpW  = _seitenBilder.Max(b => (double)b.PixelWidth);
                PdfCanvas.Width  = maxBmpW + SeiteX * 2;
                PdfCanvas.Height = Math.Max(gesamtH, 1);

                for (int i = 0; i < _seitenBilder.Count; i++)
                    SafeExecute(() => ZeicheSeite(i), $"ZeicheSeite[{i}]");

                ZeicheCropLinien();
            }
            catch (Exception ex) { LogException(ex, "ZeicheCanvas"); }
        }

        private void ZeicheSeite(int i)
        {
            if (_seitenBilder[i] == null) return;
            // Tatsächliche Bitmap-Abmessungen verwenden – kein Strecken auf RenderBreite
            double bmpW = _seitenBilder[i].PixelWidth;
            double bmpH = Math.Max(1, _seitenHöhe[i]); // = PixelHeight

            DropShadowEffect? shadow = null;
            try
            {
                shadow = new DropShadowEffect
                {
                    BlurRadius  = 20, ShadowDepth = 7,
                    Direction   = 280, Color       = Colors.Black, Opacity = 0.85
                };
            }
            catch { /* ohne Schatten weiterzeichnen */ }

            var blatt = new Border
            {
                Width               = bmpW,
                Height              = bmpH,
                Background          = Brushes.White,
                BorderBrush         = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                BorderThickness     = new Thickness(2),
                Child               = new Image
                {
                    Source  = _seitenBilder[i],
                    Width   = bmpW,
                    Height  = bmpH,
                    Stretch = Stretch.Uniform,   // Seitenverhältnis bleibt erhalten
                    SnapsToDevicePixels = true
                },
                Effect              = shadow,
                SnapsToDevicePixels = true
            };
            Canvas.SetLeft(blatt, SeiteX);
            Canvas.SetTop(blatt, _seitenYStart[i]);
            PdfCanvas.Children.Add(blatt);
        }
        // Crop-Linien-Tags: visuelle Linie = "CROP_XXX", Hit-Zone = "CROP_XXX_HIT"
        private static bool IstCropLinie(Line l)
            => l.Tag is string s && s.StartsWith("CROP_");

        private static string BasisTag(string tag)
            => tag.EndsWith("_HIT") ? tag.Substring(0, tag.Length - 4) : tag;

        private void ZeicheCropLinien()
        {
            if (_seitenBilder.Count == 0 || _seitenYStart.Length == 0) return;
            bool sichtbar = BtnRandAnzeigen.IsChecked == true;
            double canvasH = PdfCanvas.Height;

            double xL = SeiteX + _cropLinks  * RenderBreite;
            double xR = SeiteX + (1.0 - _cropRechts) * RenderBreite;
            MacheCropLinie(xL, 0, xL, canvasH, "CROP_LINKS",  sichtbar);
            MacheCropLinie(xR, 0, xR, canvasH, "CROP_RECHTS", sichtbar);

            for (int i = 0; i < _seitenBilder.Count; i++)
            {
                double yO = _seitenYStart[i] + _cropOben  * _seitenHöhe[i];
                double yU = _seitenYStart[i] + (1.0 - _cropUnten) * _seitenHöhe[i];
                MacheCropLinie(SeiteX, yO, SeiteX + RenderBreite, yO, $"CROP_OBEN_{i}",  sichtbar);
                MacheCropLinie(SeiteX, yU, SeiteX + RenderBreite, yU, $"CROP_UNTEN_{i}", sichtbar);
            }
        }

        // Fügt zwei übereinanderliegende Linien zum Canvas hinzu:
        //   visLine  – dünne grüne Linie (Anzeige, kein HitTest)
        //   hitLine  – breite transparente Zone (HitTest, nimmt Mausereignisse)
        // So bleibt die Anzeige sauber und die Greiffläche ist groß genug.
        private void MacheCropLinie(
            double x1, double y1, double x2, double y2, string tag, bool sichtbar)
        {
            double dickeVis = 2.0 / _zoomFaktor;
            double dickeHit = Math.Max(14.0, 14.0 / _zoomFaktor); // ≥14 px Greifbereich
            bool istVertikal = (tag == "CROP_LINKS" || tag == "CROP_RECHTS");
            var cursor     = istVertikal ? Cursors.SizeWE : Cursors.SizeNS;
            var visibility = sichtbar ? Visibility.Visible : Visibility.Collapsed;

            // Visuelle (dünne) Linie – nur Anzeige
            var visLine = new Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke           = Brushes.LimeGreen,
                StrokeThickness  = dickeVis,
                Tag              = tag,
                IsHitTestVisible = false,
                Visibility       = visibility
            };
            Panel.SetZIndex(visLine, 100);
            PdfCanvas.Children.Add(visLine);

            // Transparente Hit-Zone – breit, leicht zu greifen
            var hitLine = new Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke           = Brushes.Transparent,
                StrokeThickness  = dickeHit,
                Tag              = tag + "_HIT",
                Cursor           = cursor,
                IsHitTestVisible = true,
                Visibility       = visibility
            };
            Panel.SetZIndex(hitLine, 101);

            // Hover: visuelle Linie hervorheben
            hitLine.MouseEnter += (_, __) => visLine.Stroke = Brushes.Yellow;
            hitLine.MouseLeave += (_, __) => visLine.Stroke = Brushes.LimeGreen;

            hitLine.MouseLeftButtonDown += CropLinie_MouseDown;

            var menu     = new ContextMenu();
            var itemEdit = new MenuItem { Header = "Rand bearbeiten …" };
            itemEdit.Click += (_, __) => BtnRandBearbeiten_Click(this, new RoutedEventArgs());
            menu.Items.Add(itemEdit);
            hitLine.ContextMenu = menu;

            PdfCanvas.Children.Add(hitLine);
        }

        private void AktualisiereCropLinien()
        {
            if (_seitenBilder.Count == 0) return;
            double canvasH = PdfCanvas.Height;
            foreach (var l in PdfCanvas.Children.OfType<Line>())
            {
                // Basis-Tag ermitteln (visuelle Linie "CROP_X" und Hit-Zone "CROP_X_HIT" gemeinsam)
                string raw  = l.Tag?.ToString() ?? "";
                string tag  = BasisTag(raw);

                if (tag == "CROP_LINKS")
                    { l.X1 = l.X2 = SeiteX + _cropLinks * RenderBreite; l.Y1 = 0; l.Y2 = canvasH; }
                else if (tag == "CROP_RECHTS")
                    { l.X1 = l.X2 = SeiteX + (1.0 - _cropRechts) * RenderBreite; l.Y1 = 0; l.Y2 = canvasH; }
                else if (tag.StartsWith("CROP_OBEN_") &&
                         int.TryParse(tag.Substring(10), out int iO) && iO < _seitenYStart.Length)
                    l.Y1 = l.Y2 = _seitenYStart[iO] + _cropOben * _seitenHöhe[iO];
                else if (tag.StartsWith("CROP_UNTEN_") &&
                         int.TryParse(tag.Substring(11), out int iU) && iU < _seitenYStart.Length)
                    l.Y1 = l.Y2 = _seitenYStart[iU] + (1.0 - _cropUnten) * _seitenHöhe[iU];
            }
        }

        // ── Automatische Rand-Erkennung ───────────────────────────────────────

        private static (double Links, double Rechts, double Oben, double Unten)
            ErkenneCropRänderVonBitmap(BitmapSource bmp, int sicherheitsPx = 4)
        {
            int w = bmp.PixelWidth, h = bmp.PixelHeight;
            if (w <= 0 || h <= 0) return (0, 0, 0, 0);
            int stride = w * 4;
            var pixels = new byte[(long)h * stride];
            try { bmp.CopyPixels(pixels, stride, 0); }
            catch { return (0, 0, 0, 0); }

            // Dichte-basierte Erkennung: Pixel deutlich dunkler als Weiß = Inhalt.
            // Schwelle 245 (statt 240) damit hellgraue Hintergründe ignoriert werden.
            // Dichte 0.15 % (statt 0.3 %) damit auch leichter Inhalt erkannt wird.
            // Mindestens 2 aufeinanderfolgende Zeilen/Spalten mit Inhalt → robuster
            // gegen einzelne Rauschpixel am Blattrand.
            const byte   BgSchwelle     = 245;
            const double DichteSchwelle = 0.0015;
            const int    MinKonsekutiv  = 2;

            bool ZeileHatInhalt(int y)
            {
                int off = y * stride, dunkel = 0;
                for (int x = 0; x < w; x++)
                {
                    int idx = off + x * 4;
                    if (pixels[idx] < BgSchwelle || pixels[idx+1] < BgSchwelle || pixels[idx+2] < BgSchwelle)
                        dunkel++;
                }
                return (double)dunkel / w >= DichteSchwelle;
            }

            bool SpalteHatInhalt(int x)
            {
                int dunkel = 0;
                for (int y = 0; y < h; y++)
                {
                    int idx = y * stride + x * 4;
                    if (pixels[idx] < BgSchwelle || pixels[idx+1] < BgSchwelle || pixels[idx+2] < BgSchwelle)
                        dunkel++;
                }
                return (double)dunkel / h >= DichteSchwelle;
            }

            // Erste Position suchen wo MinKonsekutiv aufeinanderfolgende Zeilen/Spalten
            // Inhalt haben → verhindert, dass einzelne Rauschpixel den Rand zu früh setzen
            int FindeErste(int von, int bis, int schritt, System.Func<int, bool> hatInhalt)
            {
                int konsek = 0, erster = -1;
                for (int p = von; p != bis; p += schritt)
                {
                    if (hatInhalt(p)) { if (konsek == 0) erster = p; konsek++; if (konsek >= MinKonsekutiv) return erster; }
                    else              { konsek = 0; erster = -1; }
                }
                return -1;
            }

            int oben   = FindeErste(0,   h,    1,  ZeileHatInhalt);
            if (oben < 0) return (0, 0, 0, 0);
            int unten  = FindeErste(h-1, -1,  -1,  ZeileHatInhalt);
            int links  = FindeErste(0,   w,    1,  SpalteHatInhalt);
            int rechts = FindeErste(w-1, -1,  -1,  SpalteHatInhalt);

            if (links < 0 || rechts < 0 || unten < 0) return (0, 0, 0, 0);

            // Sicherheitsabstand
            oben   = Math.Max(0,   oben  - sicherheitsPx);
            unten  = Math.Min(h-1, unten + sicherheitsPx);
            links  = Math.Max(0,   links - sicherheitsPx);
            rechts = Math.Min(w-1, rechts + sicherheitsPx);

            return (
                (double)links        / w,
                (double)(w-1-rechts) / w,
                (double)oben         / h,
                (double)(h-1-unten)  / h
            );
        }

        // ── Crop-Linien Drag & Drop ───────────────────────────────────────────

        private void CropLinie_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is not Line l || !IstCropLinie(l)) return;
                // Basis-Tag speichern (ohne "_HIT") damit MouseMove-Handler ihn vergleichen kann
                _gezogeneCropSeite = BasisTag(l.Tag?.ToString() ?? "");
                // Capture auf Canvas – robuster als auf der schmalen Linie selbst
                PdfCanvas.CaptureMouse();
                PdfCanvas.MouseMove         += CropCanvas_MouseMove;
                PdfCanvas.MouseLeftButtonUp += CropCanvas_MouseUp;
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropLinie_MouseDown"); }
        }

        private void CropCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (_gezogeneCropSeite == null || _seitenBilder.Count == 0) return;
                Point pos = e.GetPosition(PdfCanvas);

                switch (_gezogeneCropSeite)
                {
                    case "CROP_LINKS":
                        _cropLinks = Math.Max(0, Math.Min(0.49 - _cropRechts,
                            (pos.X - SeiteX) / RenderBreite));
                        break;
                    case "CROP_RECHTS":
                        _cropRechts = Math.Max(0, Math.Min(0.49 - _cropLinks,
                            (SeiteX + RenderBreite - pos.X) / RenderBreite));
                        break;
                    default:
                        if (_gezogeneCropSeite.StartsWith("CROP_OBEN_") &&
                            int.TryParse(_gezogeneCropSeite.Substring(10), out int iO) &&
                            iO < _seitenYStart.Length && _seitenHöhe[iO] > 0)
                        {
                            double frac = (pos.Y - _seitenYStart[iO]) / _seitenHöhe[iO];
                            _cropOben = Math.Max(0, Math.Min(0.49 - _cropUnten, frac));
                        }
                        else if (_gezogeneCropSeite.StartsWith("CROP_UNTEN_") &&
                                 int.TryParse(_gezogeneCropSeite.Substring(11), out int iU) &&
                                 iU < _seitenYStart.Length && _seitenHöhe[iU] > 0)
                        {
                            double bot  = _seitenYStart[iU] + _seitenHöhe[iU];
                            double frac = (bot - pos.Y) / _seitenHöhe[iU];
                            _cropUnten = Math.Max(0, Math.Min(0.49 - _cropOben, frac));
                        }
                        break;
                }

                AktualisiereCropLinien();
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropCanvas_MouseMove"); }
        }

        private void CropCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (_gezogeneCropSeite == null) return;
                PdfCanvas.ReleaseMouseCapture();
                PdfCanvas.MouseMove         -= CropCanvas_MouseMove;
                PdfCanvas.MouseLeftButtonUp -= CropCanvas_MouseUp;
                _gezogeneCropSeite = null;
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropCanvas_MouseUp"); _gezogeneCropSeite = null; }
        }

        // ── Toolbar-Handler: Rand ─────────────────────────────────────────────

        private void BtnRandAnzeigen_Checked(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                foreach (var l in PdfCanvas.Children.OfType<Line>().Where(IstCropLinie))
                    l.Visibility = Visibility.Visible;
            }, "BtnRandAnzeigen_Checked");

        private void BtnRandAnzeigen_Unchecked(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                foreach (var l in PdfCanvas.Children.OfType<Line>().Where(IstCropLinie))
                    l.Visibility = Visibility.Collapsed;
            }, "BtnRandAnzeigen_Unchecked");

        private void BtnRandAuto_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => { _autoRandAktiv = true; AktualisiereAutoRand(); }, "BtnRandAuto_Click");

        private void AktualisiereAutoRand()
        {
            if (_seitenBilder.Count == 0) return;
            TxtInfo.Text = "Erkenne Rand …";
            var bilder = _seitenBilder.ToList();
            int sicherheitPx = Math.Max(0, (int)Math.Round(_cropSicherheitMm * _pxPerMm));

            Task.Run(() =>
            {
                double minL = double.MaxValue, minR = double.MaxValue;
                double minO = double.MaxValue, minU = double.MaxValue;
                foreach (var bmp in bilder)
                {
                    var (l, r, o, u) = ErkenneCropRänderVonBitmap(bmp, sicherheitPx);
                    if (l < minL) minL = l;
                    if (r < minR) minR = r;
                    if (o < minO) minO = o;
                    if (u < minU) minU = u;
                }
                if (minL == double.MaxValue) minL = 0;
                if (minR == double.MaxValue) minR = 0;
                if (minO == double.MaxValue) minO = 0;
                if (minU == double.MaxValue) minU = 0;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _cropLinks  = minL; _cropRechts = minR;
                    _cropOben   = minO; _cropUnten  = minU;
                    AktualisiereCropLinien();
                    if (BtnRandAnzeigen.IsChecked != true) BtnRandAnzeigen.IsChecked = true;
                    TxtInfo.Text = $"Rand · L:{minL*100:F1}%  O:{minO*100:F1}%  R:{minR*100:F1}%  U:{minU*100:F1}%";
                }));
            });
        }

        private void BtnRandBearbeiten_Click(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0)
                { MessageBox.Show("Kein PDF geladen.", "Rand bearbeiten"); return; }

                double refH    = _seitenHöhe.Length > 0 ? _seitenHöhe[0] : 1000;
                double refW    = _seitenBilder.Count > 0 ? _seitenBilder[0].PixelWidth : RenderBreite;
                double pxPerMm = _pxPerMm > 0 ? _pxPerMm : 4.0;

                // Crop-Fractions -> Pixel -> mm
                double obenMm   = Math.Round(_cropOben   * refH / pxPerMm, 1);
                double untenMm  = Math.Round(_cropUnten  * refH / pxPerMm, 1);
                double linksMm  = Math.Round(_cropLinks  * refW / pxPerMm, 1);
                double rechtsMm = Math.Round(_cropRechts * refW / pxPerMm, 1);

                TextBox MkTxt(double v) => new TextBox
                {
                    Text = v.ToString("F1"), Width = 72,
                    TextAlignment = TextAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center
                };
                TextBlock MkLbl(string s) => new TextBlock
                {
                    Text = s, VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 6, 0), Width = 60
                };

                var txtOben   = MkTxt(obenMm);
                var txtUnten  = MkTxt(untenMm);
                var txtLinks  = MkTxt(linksMm);
                var txtRechts = MkTxt(rechtsMm);

                var sp = new StackPanel { Margin = new Thickness(14) };
                sp.Children.Add(new TextBlock
                {
                    Text = "Rand in mm (0 = kein Rand):",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 8)
                });

                void AddRow(string label, TextBox tb)
                {
                    var row = new DockPanel { Margin = new Thickness(0, 3, 0, 3) };
                    var lbl = MkLbl(label); DockPanel.SetDock(lbl, Dock.Left); row.Children.Add(lbl);
                    var mm  = new TextBlock { Text = " mm", VerticalAlignment = VerticalAlignment.Center };
                    DockPanel.SetDock(mm, Dock.Right); row.Children.Add(mm);
                    row.Children.Add(tb);
                    sp.Children.Add(row);
                }

                AddRow("Oben:",   txtOben);
                AddRow("Unten:",  txtUnten);
                AddRow("Links:",  txtLinks);
                AddRow("Rechts:", txtRechts);

                bool ok = false;
                Window? dlg = null;
                var btnOk = new Button
                    { Content = "OK", Width = 70, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
                var btnCancel = new Button { Content = "Abbrechen", Width = 80, IsCancel = true };
                btnOk.Click     += (_, __) => { ok = true; dlg!.Close(); };
                btnCancel.Click += (_, __) => dlg!.Close();

                var btnRow = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 12, 0, 0)
                };
                btnRow.Children.Add(btnOk);
                btnRow.Children.Add(btnCancel);
                sp.Children.Add(btnRow);

                dlg = new Window
                {
                    Title = "Rand bearbeiten", Content = sp,
                    SizeToContent = SizeToContent.WidthAndHeight,
                    MinWidth = 260,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = Window.GetWindow(this),
                    ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
                };
                dlg.ShowDialog();
                if (!ok) return;

                if (!double.TryParse(txtOben.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nO) || nO < 0) return;
                if (!double.TryParse(txtUnten.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nU) || nU < 0) return;
                if (!double.TryParse(txtLinks.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nL) || nL < 0) return;
                if (!double.TryParse(txtRechts.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nR) || nR < 0) return;

                double newObenPx   = nO * pxPerMm;
                double newUntenPx  = nU * pxPerMm;
                double newLinksPx  = nL * pxPerMm;
                double newRechtsPx = nR * pxPerMm;

                if (newObenPx + newUntenPx >= refH || newLinksPx + newRechtsPx >= refW)
                {
                    MessageBox.Show("Rand-Summe überschreitet die Seitengröße.",
                        "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _cropOben  = newObenPx  / refH; _cropUnten  = newUntenPx  / refH;
                _cropLinks = newLinksPx / refW; _cropRechts = newRechtsPx / refW;
                _autoRandAktiv = false;

                AktualisiereCropLinien();
                if (BtnRandAnzeigen.IsChecked != true) BtnRandAnzeigen.IsChecked = true;
            }, "BtnRandBearbeiten_Click");
        }

        // ── Zoom ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Strg + Mausrad → zentriert auf Mausposition zoomen.
        /// Normales Mausrad → Scrollen (Standard-Verhalten des ScrollViewer).
        /// </summary>
        private void ScrollView_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0) return;
            e.Handled = true;

            try
            {
                // Mausposition in Canvas-Koordinaten (vor dem Zoom)
                Point mouseInCanvas = e.GetPosition(PdfCanvas);
                // Mausposition relativ zum ScrollViewer-Viewport
                Point mouseInView   = e.GetPosition(ScrollView);

                double faktor   = e.Delta > 0 ? 1.0 + ZoomStep : 1.0 / (1.0 + ZoomStep);
                double neuerZoom = Math.Max(ZoomMin, Math.Min(ZoomMax, _zoomFaktor * faktor));
                if (Math.Abs(neuerZoom - _zoomFaktor) < 0.001) return;

                SetzeZoom(neuerZoom);

                // Scrollposition so anpassen, dass der Punkt unter dem Mauszeiger stabil bleibt:
                // newOffset = mouseInCanvas * newZoom - mouseInView
                Dispatcher.BeginInvoke(
                    new Action(() =>
                    {
                        ScrollView.ScrollToHorizontalOffset(mouseInCanvas.X * neuerZoom - mouseInView.X);
                        ScrollView.ScrollToVerticalOffset  (mouseInCanvas.Y * neuerZoom - mouseInView.Y);
                    }),
                    System.Windows.Threading.DispatcherPriority.Loaded);
            }
            catch (Exception ex) { LogException(ex, "Zoom/MouseWheel"); }
        }

        private void BtnZoomMinus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(Math.Max(ZoomMin, _zoomFaktor / (1.0 + ZoomStep))), "ZoomMinus");

        private void BtnZoomPlus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(Math.Min(ZoomMax, _zoomFaktor * (1.0 + ZoomStep))), "ZoomPlus");

        private void BtnZoomReset_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(1.0), "ZoomReset");

        private void SetzeZoom(double zoom)
        {
            _zoomFaktor   = zoom;
            CanvasZoom.ScaleX = zoom;
            CanvasZoom.ScaleY = zoom;
            TxtZoom.Text  = $"{zoom * 100:0}%";

            // Liniendicken anpassen: visuelle Linien dünn, Hit-Zonen breit
            foreach (var l in PdfCanvas.Children.OfType<Line>())
            {
                string tag = l.Tag?.ToString() ?? "";
                if (tag.EndsWith("_HIT"))
                    l.StrokeThickness = Math.Max(14.0, 14.0 / zoom);
                else if (tag.StartsWith("CROP_"))
                    l.StrokeThickness = 2.0 / zoom;
                l.StrokeDashArray = null;
            }
        }

        private void BtnZoomFit_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                double sichtbar  = ScrollView.ActualWidth;
                double canvBreite = RenderBreite + 2.0 * SeiteX;
                if (canvBreite <= 0 || sichtbar <= 0) return;
                SetzeZoom(Math.Max(ZoomMin, Math.Min(ZoomMax, sichtbar / canvBreite)));
            }, "BtnZoomFit_Click");
        private void BtnExport_Click(object sender, RoutedEventArgs e)
            => SafeExecute(ExportierenNachWord, "BtnExport_Click");

        private void BtnSicherheitMinus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_cropSicherheitMm <= 0) return;
                _cropSicherheitMm = Math.Round(Math.Max(0.0, _cropSicherheitMm - 1.0), 1);
                TxtSicherheit.Text = _cropSicherheitMm.ToString("F0");
                TxtInfo.Text = $"Sicherheitsabstand: {_cropSicherheitMm} mm";
                if (_autoRandAktiv) AktualisiereAutoRand();
            }, "BtnSicherheitMinus_Click");

        private void BtnSicherheitPlus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_cropSicherheitMm >= 20) return;
                _cropSicherheitMm = Math.Round(Math.Min(20.0, _cropSicherheitMm + 1.0), 1);
                TxtSicherheit.Text = _cropSicherheitMm.ToString("F0");
                TxtInfo.Text = $"Sicherheitsabstand: {_cropSicherheitMm} mm";
                if (_autoRandAktiv) AktualisiereAutoRand();
            }, "BtnSicherheitPlus_Click");
        // ── Word-Export ───────────────────────────────────────────────────────

        private void ExportThreadWorker(
            string zielPfad, string pdfPfad,
            double[] seitenYStart, double[] seitenHöhe,
            double cropLinks, double cropRechts, double cropOben, double cropUnten)
        {
            var alleTemp = new List<string>();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                // ── 1. PDF-Seitengröße und effektive Maße nach Beschnitt ──────────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Lese PDF-Maße …"));
                var (nativeW_pts, nativeH_pts) = HolePdfSeitenGrösse(pdfPfad);
                int exportPixW = Math.Max(1, (int)Math.Round(nativeW_pts / 72.0 * ExportDpi));
                int exportPixH = Math.Max(1, (int)Math.Round(nativeH_pts / 72.0 * ExportDpi));

                double nativeW_eff = nativeW_pts * Math.Max(0.01, 1.0 - cropLinks - cropRechts);
                double nativeH_eff = nativeH_pts * Math.Max(0.01, 1.0 - cropOben  - cropUnten);

                App.LogFehler("Export/Maße",
                    $"PDF: {nativeW_pts:F1}×{nativeH_pts:F1} pt | " +
                    $"Render: {exportPixW}×{exportPixH} px | DPI={ExportDpi} | " +
                    $"Eff: {nativeW_eff:F1}×{nativeH_eff:F1} pt | " +
                    $"Crop L:{cropLinks:P1} R:{cropRechts:P1} O:{cropOben:P1} U:{cropUnten:P1}");

                // ── 2. Hochauflösend rendern ───────────────────────────────────────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Rendere in hoher Auflösung …"));
                long t1 = sw.ElapsedMilliseconds;
                List<BitmapSource> hochRes;
                try { hochRes = RenderiereAlleSeiten(pdfPfad, exportPixW, exportPixH); }
                catch (Exception ex) { LogException(ex, "Export/Render"); hochRes = new List<BitmapSource>(); }
                App.LogFehler("Export/Timing", $"Rendern: {sw.ElapsedMilliseconds - t1} ms ({hochRes.Count} Seiten, {exportPixW}×{exportPixH} px)");

                if (hochRes.Count == 0)
                {
                    Dispatcher.BeginInvoke(new Action(() => {
                        BtnExport.IsEnabled = true;
                        TxtInfo.Text = "Export abgebrochen: Render fehlgeschlagen.";
                    }));
                    return;
                }

                // ── 3. Seitenanzahl prüfen ────────────────────────────────────────
                if (hochRes.Count == 0)
                {
                    Dispatcher.BeginInvoke(new Action(() => {
                        BtnExport.IsEnabled = true;
                        TxtInfo.Text = "Export abgebrochen: Keine Seiten gerendert.";
                    }));
                    return;
                }

                // ── 4. Word öffnen, Schreibbereich lesen, Skalierungsdialog ─────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Word wird geöffnet …"));
                Word.Application? wordApp = null;
                Word.Document?    wordDoc = null;
                int seiteNr = 0;
                try
                {
                    wordApp = new Word.Application { Visible = false };
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wordDoc = wordApp.Documents.Add();
                    NormiereAbsatzstil(wordDoc);

                    // Seitenformat in Word passend zur PDF-Seite einstellen
                    SetzeWordSeitenFormat(wordDoc, nativeW_pts, nativeH_pts);

                    var (availW_pt, availH_pt) = HoleSchreibbereich(wordDoc);
                    App.LogFehler("Export/Schreibbereich",
                        $"Vorlage: {availW_pt:F1} × {availH_pt:F1} pt | " +
                        $"Druckbereich: {nativeW_eff:F1} × {nativeH_eff:F1} pt");

                    double scaleW      = nativeW_eff > 0 ? availW_pt / nativeW_eff : 1.0;
                    double scaleH      = nativeH_eff > 0 ? availH_pt / nativeH_eff : 1.0;
                    double globalScale = Math.Min(1.0, Math.Min(scaleW, scaleH));
                    bool   doScale     = true;

                    App.LogFehler("Export/Skalierung-Detail",
                        $"nativeW_eff={nativeW_eff:F2} pt | nativeH_eff={nativeH_eff:F2} pt | " +
                        $"availW_pt={availW_pt:F2} pt | availH_pt={availH_pt:F2} pt | " +
                        $"scaleW={scaleW:F4} ({scaleW*100:F1}%) | scaleH={scaleH:F4} ({scaleH*100:F1}%) | " +
                        $"globalScale={globalScale:F4} ({globalScale*100:F1}%)");

                    if (globalScale < 0.99)
                    {
                        int prozent = (int)Math.Round(globalScale * 100.0);
                        string engpassAchse = scaleW <= scaleH
                            ? $"Breite: PDF {nativeW_eff:F0} pt > Schreibbereich {availW_pt:F0} pt"
                            : $"Höhe: PDF {nativeH_eff:F0} pt > Schreibbereich {availH_pt:F0} pt";
                        Dispatcher.Invoke(new Action(() =>
                        {
                            var antwort = MessageBox.Show(
                                $"Der Druckbereich des PDFs ({nativeW_eff:F0} × {nativeH_eff:F0} pt)\n" +
                                $"ist größer als der Schreibbereich der Word-Vorlage ({availW_pt:F0} × {availH_pt:F0} pt).\n" +
                                $"  Engpass: {engpassAchse}\n\n" +
                                $"Hinweis: Die Word-Seitenränder reduzieren den Schreibbereich.\n" +
                                $"Proportionale Verkleinerung auf {prozent} % passt den Inhalt vollständig ein.\n\n" +
                                $"Möchten Sie alle Seiten auf {prozent} % verkleinern?\n" +
                                $"(Nein = Überstand rechts/unten abschneiden, Originalgröße bleibt erhalten)",
                                "Proportionale Verkleinerung",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Question);
                            doScale = (antwort == MessageBoxResult.Yes);
                        }));
                    }
                    else
                    {
                        globalScale = 1.0;
                    }

                    App.LogFehler("Export/Skalierung",
                        $"globalScale={globalScale:F4} ({(int)Math.Round(globalScale * 100)} %) | doScale={doScale}");

                    int availW_px = Math.Max(1, (int)Math.Round(availW_pt / 72.0 * ExportDpi));
                    int availH_px = Math.Max(1, (int)Math.Round(availH_pt / 72.0 * ExportDpi));
                    int gesamtGruppen = hochRes.Count;

                    // ── 5. Pro Original-PDF-Seite genau ein Bild → eine Word-Seite ──
                    long t4 = sw.ElapsedMilliseconds;
                    for (int gi = 0; gi < hochRes.Count; gi++)
                    {
                        seiteNr++;
                        int sNr = seiteNr, gGes = gesamtGruppen;
                        Dispatcher.BeginInvoke(new Action(() =>
                            TxtInfo.Text = $"Erstelle Seite {sNr} von {gGes} …"));

                        BitmapSource? segBmp = null;
                        try
                        {
                            segBmp = RendereBlockHochRes(hochRes, seitenYStart, seitenHöhe,
                                    seitenYStart[gi], seitenYStart[gi] + seitenHöhe[gi],
                                    cropLinks, cropRechts, cropOben, cropUnten);
                        }
                        catch (Exception ex) { LogException(ex, $"Export/Block[{gi}]"); }
                        if (segBmp == null) continue;

                        BitmapSource finalBmp;
                        float finalW_pt, finalH_pt;
                        double segW_pt = segBmp.PixelWidth  / ExportDpi * 72.0;
                        double segH_pt = segBmp.PixelHeight / ExportDpi * 72.0;

                        if (doScale)
                        {
                            finalBmp  = segBmp;
                            finalW_pt = (float)Math.Min(availW_pt, segW_pt * globalScale);
                            finalH_pt = (float)Math.Min(availH_pt, segH_pt * globalScale);
                        }
                        else
                        {
                            int clipW = Math.Min(segBmp.PixelWidth,  availW_px);
                            int clipH = Math.Min(segBmp.PixelHeight, availH_px);
                            BitmapSource clipped = segBmp;
                            try
                            {
                                var cb = new CroppedBitmap(segBmp, new Int32Rect(0, 0, clipW, clipH));
                                cb.Freeze();
                                clipped = cb;
                            }
                            catch (Exception ex) { LogException(ex, $"Export/Clip[{gi}]"); }
                            finalBmp  = clipped;
                            finalW_pt = (float)(finalBmp.PixelWidth  / ExportDpi * 72.0);
                            finalH_pt = (float)(finalBmp.PixelHeight / ExportDpi * 72.0);
                        }

                        App.LogFehler("Export/Seite",
                            $"[{gi}] {segBmp.PixelWidth}x{segBmp.PixelHeight} px " +
                            $"Shape: {finalW_pt:F1}x{finalH_pt:F1} pt | doScale={doScale}");

                        string png = IO.Path.Combine(IO.Path.GetTempPath(), $"sm_{Guid.NewGuid():N}.png");
                        try
                        {
                            var enc = new PngBitmapEncoder();
                            enc.Frames.Add(BitmapFrame.Create(finalBmp));
                            using var fs = IO.File.Create(png);
                            enc.Save(fs);
                            alleTemp.Add(png);
                        }
                        catch (Exception ex) { LogException(ex, $"Export/PNG[{gi}]"); continue; }

                        bool seitenumbruch = (seiteNr > 1);
                        try { EinfügenSegment(wordDoc, png, finalW_pt, finalH_pt,
                                neuerAbsatz: false, seitenumbruchVorher: seitenumbruch); }
                        catch (Exception ex) { LogException(ex, $"Export/Word[{gi}]"); }
                    }
                    App.LogFehler("Export/Timing",
                        $"Bilder+Einfügen: {sw.ElapsedMilliseconds - t4} ms ({seiteNr} Seiten)");

                    BereinigeDokumentAbsätze(wordDoc);

                    // Leerseite am Ende verhindern: Word-Dokumente haben immer einen
                    // abschließenden leeren Absatz. Wenn der letzte Inhalt die Seite
                    // komplett füllt, landet dieser Absatz auf einer neuen (leeren) Seite.
                    // Lösung: letzten leeren Absatz auf 1 pt Schriftgröße setzen.
                    try
                    {
                        int absNr = wordDoc.Paragraphs.Count;
                        if (absNr > 0)
                        {
                            var letzter = wordDoc.Paragraphs[absNr];
                            var txt     = letzter.Range.Text;
                            // Absatz gilt als leer wenn er nur Absatzmarke enthält
                            if (txt == "\r" || (txt != null && txt.Replace("\r","").Replace("\n","").Trim().Length == 0))
                            {
                                letzter.Range.Font.Size    = 1f;
                                letzter.Format.SpaceBefore = 0f;
                                letzter.Format.SpaceAfter  = 0f;
                                App.LogFehler("Export/LetzterAbsatz", "Auf 1pt minimiert (war leer)");
                            }
                        }
                    }
                    catch (Exception ex) { LogException(ex, "Export/LetzterAbsatz"); }

                    // Datei vorher löschen → kein Word-Überschreiben-Dialog
                    if (IO.File.Exists(zielPfad))
                        try { IO.File.Delete(zielPfad); } catch { }

                    Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Speichere Word-Dokument …"));
                    wordDoc.SaveAs2(zielPfad, Word.WdSaveFormat.wdFormatXMLDocument);

                    string dateiName = IO.Path.GetFileName(zielPfad);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = $"Exportiert: {seiteNr} Seite(n) → {dateiName}";
                        if (MessageBox.Show(
                            $"Word-Dokument erstellt:\n{zielPfad}\n\nJetzt öffnen?",
                            "Export fertig", MessageBoxButton.YesNo,
                            MessageBoxImage.Information) == MessageBoxResult.Yes)
                        {
                            try { System.Diagnostics.Process.Start(
                                new System.Diagnostics.ProcessStartInfo(zielPfad)
                                { UseShellExecute = true }); }
                            catch { }
                        }
                    }));
                }
                catch (Exception ex)
                {
                    LogException(ex, "Export/Word");
                    var inner = App.EntpackeTIE(ex);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = "Export fehlgeschlagen. Siehe Log.";
                        MessageBox.Show(
                            $"Word-Export Fehler:\n[{inner.GetType().Name}] {inner.Message}\n\n[Log: {App.LogDatei}]",
                            "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                    }));
                }
                finally
                {
                    try { wordDoc?.Close(false); }  catch { }
                    try { wordApp?.Quit(); }          catch { }
                    if (wordDoc != null) try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc); } catch { }
                    if (wordApp != null) try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
            finally
            {
                App.LogFehler("Export/Timing", $"Gesamt: {sw.ElapsedMilliseconds} ms");
                foreach (var f in alleTemp) try { IO.File.Delete(f); } catch { }

                // Export-Sperre immer zurücksetzen (auch im Fehlerfall)
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _exportLäuft = false;
                    BtnExport.IsEnabled = true;
                }));
            }
        }
        /// <summary>
        /// Liest die nativen Seitenmaße der ersten PDF-Seite in Points (72 pt = 1 inch).
        /// Wird für die exakte 300-DPI-Render-Berechnung verwendet.
        /// Fallback: A4 (595 × 842 pt).
        /// </summary>
        private static (double widthPts, double heightPts) HolePdfSeitenGrösse(string pfad)
        {
            try
            {
                using var doc = PdfReader.Open(pfad, PdfDocumentOpenMode.ReadOnly);
                if (doc.PageCount > 0)
                {
                    var page = doc.Pages[0];
                    double w = page.Width.Point;
                    double h = page.Height.Point;
                    if (w > 0 && h > 0) return (w, h);
                }
            }
            catch (Exception ex) { LogException(ex, "HolePdfSeitenGrösse"); }
            return (595.0, 842.0); // A4-Fallback
        }

        /// <summary>
        /// Liest den tatsächlich verfügbaren Schreibbereich des Word-Dokuments
        /// aus den vorhandenen Seiteneinstellungen (Ränder bleiben unverändert).
        /// Rückgabe in Points (72 pt = 1 inch).
        /// Fallback: A4 mit 2,5 cm Rand ≈ 451 × 694 pt.
        /// </summary>
        private static (double BreiteP, double HöheP) HoleSchreibbereich(Word.Document doc)
        {
            try
            {
                var ps = doc.PageSetup;
                double w = ps.PageWidth  - ps.LeftMargin - ps.RightMargin;
                double h = ps.PageHeight - ps.TopMargin  - ps.BottomMargin;
                return (Math.Max(10.0, w), Math.Max(10.0, h));
            }
            catch (Exception ex)
            {
                LogException(ex, "HoleSchreibbereich");
                return (451.0, 694.0); // A4, 2,5 cm Rand
            }
        }

        /// <summary>
        /// Setzt das Word-Seitenformat (Größe, Ausrichtung, Ränder) passend zur PDF-Seite.
        /// Ränder werden auf 0 gesetzt, damit das Bild die volle Seite füllt.
        /// </summary>
        private static void SetzeWordSeitenFormat(Word.Document doc, double pdfW_pts, double pdfH_pts)
        {
            try
            {
                var ps = doc.PageSetup;

                // Orientation ZUERST setzen: Word begrenzt PageWidth im Hochformat auf
                // ≤ aktuelle PageHeight. Bei Querformat-PDFs (z.B. A3: pdfW=1190 pt)
                // würde PageWidth ohne diesen Schritt auf ~842 pt gekürzt → falsches Format.
                // Nach dem Orientation-Swap werden die exakten PDF-Maße gesetzt.
                ps.Orientation = pdfW_pts > pdfH_pts
                    ? Word.WdOrientation.wdOrientLandscape
                    : Word.WdOrientation.wdOrientPortrait;

                ps.PageWidth  = (float)pdfW_pts;
                ps.PageHeight = (float)pdfH_pts;

                // Minimale Ränder (1 pt ≈ 0,35 mm) statt 0 – verhindert COM-Ausnahmen
                // bei Druckertreibern die keine 0-Ränder akzeptieren
                const float RandPt = 1f;
                ps.TopMargin    = RandPt;
                ps.BottomMargin = RandPt;
                ps.LeftMargin   = RandPt;
                ps.RightMargin  = RandPt;

                App.LogFehler("Export/SeitenFormat",
                    $"Word-Seite: {pdfW_pts:F1} × {pdfH_pts:F1} pt " +
                    $"({(pdfW_pts > pdfH_pts ? "Quer" : "Hoch")})");
            }
            catch (Exception ex) { LogException(ex, "SetzeWordSeitenFormat"); }
        }

        /// <summary>
        /// Setzt den Normal-Stil des Dokuments auf 0 Abstand / einfachen Zeilenabstand.
        /// </summary>
        private static void NormiereAbsatzstil(Word.Document doc)
        {
            try
            {
                var normal = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                var pf     = normal.ParagraphFormat;
                pf.SpaceBefore         = 0f;
                pf.SpaceAfter          = 0f;
                pf.SpaceBeforeAuto     = 0;  // False: nicht automatisch
                pf.SpaceAfterAuto      = 0;
                pf.LineSpacingRule     = Word.WdLineSpacing.wdLineSpaceSingle;
                pf.LeftIndent          = 0f;
                pf.RightIndent         = 0f;
                pf.FirstLineIndent     = 0f;
            }
            catch (Exception ex) { LogException(ex, "NormiereAbsatzstil"); }
        }

        /// <summary>
        /// Fügt ein Bildsegment in das Word-Dokument ein.
        ///
        /// Geometrie:  pts = px / ExportDpi * 72  (× skalierung)
        ///
        /// seitenumbruchVorher = true:
        ///   Vor diesem Segment wird ein neuer Absatz eingefügt und
        ///   PageBreakBefore = wdTrue gesetzt, damit die neue PDF-Seite
        ///   sauber auf einer neuen Word-Seite beginnt.
        ///
        /// neuerAbsatz = true:
        ///   Folgesegment innerhalb derselben PDF-Seite — neuer Absatz,
        ///   kein Seitenumbruch.
        /// </summary>
        private static void EinfügenSegment(
            Word.Document doc,
            string bildPfad,
            float imgW_pt, float imgH_pt,
            bool neuerAbsatz,
            bool seitenumbruchVorher = false)
        {
            if (seitenumbruchVorher || neuerAbsatz)
            {
                Word.Range rNew = doc.Content;
                rNew.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                rNew.InsertParagraphAfter();
            }

            Word.Range r = doc.Content;
            r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            SetzeAbsatzFormat(r.ParagraphFormat);
            try { r.ParagraphFormat.PageBreakBefore = seitenumbruchVorher ? -1 : 0; } catch { }

            var shape = r.InlineShapes.AddPicture(bildPfad, false, true);
            shape.Width  = imgW_pt;
            shape.Height = imgH_pt;
        }

        private static void SetzeAbsatzFormat(Word.ParagraphFormat pf)
        {
            try
            {
                pf.SpaceBefore      = 0f;
                pf.SpaceAfter       = 0f;
                pf.SpaceBeforeAuto  = 0;
                pf.SpaceAfterAuto   = 0;
                pf.LineSpacingRule  = Word.WdLineSpacing.wdLineSpaceSingle;
                pf.LeftIndent       = 0f;
                pf.RightIndent      = 0f;
                pf.FirstLineIndent  = 0f;
            }
            catch { /* ParagraphFormat-Fehler nicht propagieren */ }
        }

        /// <summary>
        /// Letzte Sicherung: alle Absätze im fertigen Dokument nochmals bereinigen.
        /// Verhindert, dass Word Abstände nachträglich einfügt.
        /// </summary>
        private static void BereinigeDokumentAbsätze(Word.Document doc)
        {
            try
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                    SetzeAbsatzFormat(para.Format);
            }
            catch (Exception ex) { LogException(ex, "BereinigeDokumentAbsätze"); }
        }

        // ── Hochauflösung: Block zusammensetzen ──────────────────────────────

        private static BitmapSource? RendereBlockHochRes(
            List<BitmapSource> hochRes,
            double[] seitenYStart, double[] seitenHöhe,
            double yStart, double yEnd,
            double cropLinks = 0, double cropRechts = 0,
            double cropOben = 0, double cropUnten = 0)
        {
            if (hochRes == null || hochRes.Count == 0) return null;
            if (seitenYStart == null || seitenHöhe  == null) return null;
            if (yEnd <= yStart) return null;

            var teile = new List<(BitmapSource Bmp, int H)>();
            int maxW  = 0;
            int n     = Math.Min(hochRes.Count, Math.Min(seitenYStart.Length, seitenHöhe.Length));

            for (int i = 0; i < n; i++)
            {
                try
                {
                    if (hochRes[i] == null || seitenHöhe[i] <= 0) continue;

                    double sTop    = seitenYStart[i];
                    double sBottom = sTop + seitenHöhe[i];

                    // Effektive Seitengrenzen nach Beschnitt (in Vorschau-Koordinaten)
                    double effTop    = sTop    + cropOben  * seitenHöhe[i];
                    double effBottom = sBottom - cropUnten * seitenHöhe[i];

                    double übTop = Math.Max(yStart, effTop);
                    double übBot = Math.Min(yEnd,   effBottom);
                    if (übBot <= übTop) continue;

                    double scY = hochRes[i].PixelHeight / seitenHöhe[i];
                    int pixT   = (int)Math.Round((übTop - sTop) * scY);  // ab Seiten-Top (inkl. Crop-Offset)
                    int pixH   = (int)Math.Round((übBot - übTop) * scY);
                    int w      = hochRes[i].PixelWidth;

                    // Horizontaler Beschnitt in Export-Pixeln
                    int cropL_px = (int)Math.Round(cropLinks  * w);
                    int cropR_px = (int)Math.Round(cropRechts * w);
                    int cropW    = Math.Max(1, w - cropL_px - cropR_px);
                    if (cropL_px + cropW > w) cropW = w - cropL_px;
                    if (cropW <= 0) continue;

                    pixT = Math.Max(0, pixT);
                    int maxH = hochRes[i].PixelHeight - pixT;
                    if (maxH <= 0) continue;
                    pixH = Math.Min(maxH, Math.Max(1, pixH));
                    if (pixT + pixH > hochRes[i].PixelHeight) continue;

                    var crop = new CroppedBitmap(hochRes[i], new Int32Rect(cropL_px, pixT, cropW, pixH));
                    crop.Freeze();
                    teile.Add((crop, pixH));
                    if (cropW > maxW) maxW = cropW;
                }
                catch (Exception ex) { LogException(ex, $"RendereBlockHochRes/Seite[{i}]"); }
            }

            if (teile.Count == 0 || maxW == 0) return null;

            int totalH = teile.Sum(t => t.H);
            if (totalH <= 0) return null;

            try
            {
                var visual = new DrawingVisual();
                using (var ctx = visual.RenderOpen())
                {
                    double y = 0;
                    foreach (var (bmp, h) in teile)
                    {
                        ctx.DrawImage(bmp, new Rect(0, y, maxW, h));
                        y += h;
                    }
                }
                // 96 DPI: 1 WPF-Unit = 1 Pixel, damit Koordinaten direkt als Pixel gelten.
                var rtb = new RenderTargetBitmap(
                    maxW, totalH, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(visual);
                rtb.Freeze();
                return rtb;
            }
            catch (Exception ex) { LogException(ex, "RendereBlockHochRes/Render"); return null; }
        }
    }
}
