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
        /// Die Quelldatei wird in eine temporäre Kopie geöffnet, damit das Original
        /// nicht gesperrt wird (verhindert COMException wenn Nutzer die Datei gleichzeitig in
        /// Word geöffnet hat und speichern möchte).
        /// </summary>
        public static void WortDateiZuPdf(string quellPfad, string zielpfad)
        {
            Logger.Info("Word→PDF", $"Starte: {System.IO.Path.GetFileName(quellPfad)}");
            Word.Application? wordApp = null;
            Word.Document?    doc     = null;
            string tempKopie = Path.Combine(
                Path.GetTempPath(),
                "sm_pdf_" + Guid.NewGuid().ToString("N") + Path.GetExtension(quellPfad));
            try
            {
                File.Copy(quellPfad, tempKopie, overwrite: true);
                wordApp = new Word.Application { Visible = false };
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                doc     = wordApp.Documents.Open(
                              FileName: tempKopie, ReadOnly: true,
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
                try { File.Delete(tempKopie); } catch { }
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

                    // Temporäre Kopie der Word-Datei verwenden, damit das Original
                    // während der Konvertierung nicht gesperrt wird.
                    string tempKopie = Path.Combine(
                        Path.GetTempPath(),
                        "sm_batch_" + Guid.NewGuid().ToString("N") + file.Extension);
                    Word.Document? doc = null;
                    try
                    {
                        var tmpPfad = pdfPfad + ".tmp";
                        File.Copy(file.FullName, tempKopie, overwrite: true);
                        doc = wordApp.Documents.Open(
                            FileName: tempKopie, ReadOnly: true,
                            AddToRecentFiles: false, Visible: false);
                        doc.ExportAsFixedFormat(
                            OutputFileName: tmpPfad,
                            ExportFormat:   Word.WdExportFormat.wdExportFormatPDF);

                        if (File.Exists(pdfPfad)) File.Delete(pdfPfad);
                        File.Move(tmpPfad, pdfPfad);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn("BatchPdf", $"{file.Name}: {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { doc.Close(SaveChanges: false); } catch { }
                            try { Marshal.ReleaseComObject(doc); } catch { }
                        }
                        try { File.Delete(tempKopie); } catch { }
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
