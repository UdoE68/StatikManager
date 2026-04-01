using System;
using System.IO;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Hilfsmethoden für Word/PDF-Verarbeitung ohne UI-Abhängigkeiten.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal static class WordPdfService
    {
        /// <summary>
        /// Öffnet eine Datei mit der systemweit zugewiesenen Standardanwendung (Shell Execute).
        /// </summary>
        public static void ÖffneInWord(string pfad)
        {
            try
            {
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo(pfad) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Datei konnte nicht geöffnet werden:\n" + ex.Message,
                    "Fehler",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Berechnet den anzuzeigenden PDF-Pfad für eine Quelldatei (Word oder PDF),
        /// optional mit eingebrannten Abdeckbändern.
        /// Kein Cache-Check für Word-Dateien — konvertiert immer neu (für Auto-Refresh nach Dateiänderung).
        /// Keine UI-Referenzen. Aufrufer ist für Thread-Verwaltung (STA) verantwortlich.
        /// </summary>
        /// <returns>Pfad zum anzuzeigenden PDF, oder <c>null</c> bei Fehler.</returns>
        public static string? BerechneNeuenZielPdf(string pfad, string cacheDir,
                                                    bool kopf, bool fuss,
                                                    double kopfMm, double fussMm)
        {
            Logger.Info("PDF-Vorschau",
                $"Berechne Ziel-PDF: {Path.GetFileName(pfad)}" +
                (kopf || fuss ? $" [Kopf={kopfMm:0}mm Fuß={fussMm:0}mm]" : ""));
            try
            {
                if (DateiTypen.IstWordDatei(Path.GetExtension(pfad)))
                {
                    var basePdf = PdfCache.GetBasisPdfPfad(pfad, cacheDir);
                    WordInteropService.WortDateiZuPdf(pfad, basePdf);

                    if (kopf || fuss)
                    {
                        var covPdf = PdfCache.GetCoveredPdfPfad(pfad, cacheDir, kopfMm, fussMm);
                        PdfCoverService.AbdeckePdfSeiten(basePdf, covPdf, kopfMm, fussMm);
                        return covPdf;
                    }
                    return basePdf;
                }
                else
                {
                    // PDF direkt
                    if (kopf || fuss)
                    {
                        var covPdf = PdfCache.GetCoveredPdfPfad(pfad, cacheDir, kopfMm, fussMm);
                        PdfCoverService.AbdeckePdfSeiten(pfad, covPdf, kopfMm, fussMm);
                        return covPdf;
                    }
                    return pfad;
                }
            }
            catch (Exception ex)
            {
                Logger.Fehler("PDF-Vorschau", $"Berechnung fehlgeschlagen für '{Path.GetFileName(pfad)}': {ex.Message}");
                return null;
            }
        }
    }
}
