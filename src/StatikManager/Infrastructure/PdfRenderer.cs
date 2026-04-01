using Docnet.Core;
using Docnet.Core.Models;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Zentraler PDF-Rendering-Service.
    /// Rendert PDF-Seiten via Docnet.Core als WPF-BitmapSource-Objekte.
    /// Extrahiert aus PdfSchnittEditor.xaml.cs.
    /// </summary>
    internal static class PdfRenderer
    {
        /// <summary>
        /// Rendert alle Seiten einer PDF-Datei als Bitmap-Liste.
        /// Transparente Pixel werden gegen Weiß kompositioniert.
        /// </summary>
        public static List<BitmapSource> RenderiereAlleSeiten(string pfad, int breite, int höhe = 0,
                                                               CancellationToken token = default)
        {
            if (höhe <= 0) höhe = breite * 2;
            var result = new List<BitmapSource>();
            try
            {
                var lib             = DocLib.Instance;
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
                        token.ThrowIfCancellationRequested();
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
                    catch (Exception ex) { App.LogFehler($"Seite[{i}]", App.GetExceptionKette(ex)); }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { App.LogFehler("RenderiereAlleSeiten", App.GetExceptionKette(ex)); }
            return result;
        }

        /// <summary>
        /// Kompositioniert transparente Pixel eines BGRA32-Puffers gegen Weiß (in-place).
        /// </summary>
        public static void KompositioniereGegenWeiss(byte[] raw, int w, int h)
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
    }
}
