using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Brennt weiße Abdeckbänder an Kopf und/oder Fuß jeder PDF-Seite ein.
    /// Die Bänder werden in den Seiteninhalt gezeichnet (XGraphicsPdfPageOptions.Append)
    /// und erscheinen in der Vorschau präzise auf der Seite – nicht im Browser-Fenster verankert.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal static class PdfCoverService
    {
        /// <summary>
        /// Öffnet <paramref name="quellPdf"/>, zeichnet weiße Rechtecke an Kopf und/oder Fuß
        /// jeder Seite und speichert das Ergebnis unter <paramref name="zielpdf"/>.
        /// </summary>
        /// <param name="kopfMm">Höhe des Kopfbandes in Millimetern (0 = kein Band).</param>
        /// <param name="fussMm">Höhe des Fußbandes in Millimetern (0 = kein Band).</param>
        public static void AbdeckePdfSeiten(string quellPdf, string zielpdf,
                                             double kopfMm, double fussMm)
        {
            const double mmToPt = 72.0 / 25.4;   // Millimeter → PDF-Punkte

            using var doc = PdfReader.Open(quellPdf, PdfDocumentOpenMode.Modify);

            foreach (PdfPage page in doc.Pages)
            {
                double w = page.Width.Point;
                double h = page.Height.Point;

                // XGraphicsPdfPageOptions.Append: Inhalt wird ÜBER dem vorhandenen
                // Seiteninhalt gezeichnet (deckt Kopf-/Fußzeilen-Text ab).
                using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

                if (kopfMm > 0)
                    gfx.DrawRectangle(XBrushes.White, 0, 0, w, kopfMm * mmToPt);

                if (fussMm > 0)
                    gfx.DrawRectangle(XBrushes.White, 0, h - fussMm * mmToPt, w, fussMm * mmToPt);
            }

            doc.Save(zielpdf);
        }
    }
}
