using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Wrapper für Word-COM-Interop-Operationen.
    /// Alle Methoden sind synchron und müssen auf einem STA-Thread aufgerufen werden.
    /// Der Aufrufer ist verantwortlich für Thread-Konfiguration und Fehlerbehandlung.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal static class WordInteropService
    {
        /// <summary>
        /// Öffnet eine Word-Datei unsichtbar und exportiert sie als PDF.
        /// </summary>
        public static void WortDateiZuPdf(string quellPfad, string zielpfad)
        {
            Logger.Info("Word→PDF", $"Starte: {System.IO.Path.GetFileName(quellPfad)}");
            Word.Application? wordApp = null;
            Word.Document?    doc     = null;
            try
            {
                wordApp = new Word.Application { Visible = false };
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;  // keine Dialoge auch wenn Datei bereits offen
                doc     = wordApp.Documents.Open(
                              FileName: quellPfad, ReadOnly: true,
                              AddToRecentFiles: false, Visible: false);
                doc.ExportAsFixedFormat(
                    OutputFileName: zielpfad,
                    ExportFormat:   Word.WdExportFormat.wdExportFormatPDF);
                Logger.Info("Word→PDF", $"Fertig: {System.IO.Path.GetFileName(zielpfad)}");
            }
            finally
            {
                doc?.Close(SaveChanges: false);
                wordApp?.Quit();
                if (doc     != null) Marshal.ReleaseComObject(doc);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        /// <summary>
        /// Konvertiert alle Word-Dateien in einem Verzeichnis nach PDF (Batch, im Hintergrund).
        /// Nur Dateien ohne gültigen Cache werden konvertiert.
        /// </summary>
        public static void WortDateienBatchZuPdf(System.IO.DirectoryInfo root, string cacheDir,
                                                  CancellationToken token)
        {
            var zuKonvertieren = root
                .EnumerateFiles("*.*", SearchOption.AllDirectories)
                .Where(f => DateiTypen.IstWordDatei(f.Extension) &&
                            !PdfCache.CacheGültig(PdfCache.GetBasisPdfPfad(f.FullName, cacheDir), f.FullName))
                .ToArray();

            if (zuKonvertieren.Length == 0 || token.IsCancellationRequested) return;

            Word.Application? wordApp = null;
            try
            {
                wordApp = new Word.Application { Visible = false };
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                foreach (var file in zuKonvertieren)
                {
                    if (token.IsCancellationRequested) break;

                    var pdfPfad = PdfCache.GetBasisPdfPfad(file.FullName, cacheDir);
                    if (PdfCache.CacheGültig(pdfPfad, file.FullName)) continue;

                    try
                    {
                        var tmpPfad = pdfPfad + ".tmp";
                        var doc = wordApp.Documents.Open(
                            FileName: file.FullName, ReadOnly: true,
                            AddToRecentFiles: false, Visible: false);
                        doc.ExportAsFixedFormat(
                            OutputFileName: tmpPfad,
                            ExportFormat:   Word.WdExportFormat.wdExportFormatPDF);
                        doc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(doc);

                        if (File.Exists(pdfPfad)) File.Delete(pdfPfad);
                        File.Move(tmpPfad, pdfPfad);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn("BatchPdf", $"{file.Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Fehler("BatchPdf", $"Batch-Konvertierung abgebrochen: {ex.Message}");
            }
            finally
            {
                try { wordApp?.Quit(); } catch { }
                if (wordApp != null)
                    try { Marshal.ReleaseComObject(wordApp); } catch { }
            }
        }
    }
}
