using System;
using System.IO;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Verwaltet Pfade und Gültigkeitsprüfung des PDF-Vorschau-Caches.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal static class PdfCache
    {
        /// <summary>Basis-Verzeichnis für den PDF-Cache unter %APPDATA%.</summary>
        public static readonly string CacheBasis = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "StatikManager", "pdf-cache");

        /// <summary>Version des Caches. Erhöhen → alle gecachten PDFs werden neu generiert.</summary>
        public const string CacheVersion = "v5";

        /// <summary>Gibt den Cache-Pfad für das Basis-PDF einer Word-Datei zurück.</summary>
        public static string GetBasisPdfPfad(string quellPfad, string cacheDir)
        {
            var hash = ((uint)quellPfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            var name = Path.GetFileNameWithoutExtension(quellPfad);
            return Path.Combine(cacheDir, $"{name}_{hash}_{CacheVersion}.pdf");
        }

        /// <summary>Gibt den Cache-Pfad für ein mit Abdeckbändern versehenes PDF zurück.</summary>
        public static string GetCoveredPdfPfad(string quellPfad, string cacheDir,
                                                double kopfMm, double fussMm)
        {
            var hash   = ((uint)quellPfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            var name   = Path.GetFileNameWithoutExtension(quellPfad);
            var covKey = $"cov_k{(int)Math.Round(kopfMm)}_f{(int)Math.Round(fussMm)}";
            return Path.Combine(cacheDir, $"{name}_{hash}_{CacheVersion}_{covKey}.pdf");
        }

        /// <summary>Gibt den Cache-Pfad für einen Word-Klon einer PDF-Datei zurück.</summary>
        public static string GetWordKlonPfad(string pdfPfad, string cacheDir)
        {
            // Falls noch kein Projekt geladen → Temp-Verzeichnis
            if (string.IsNullOrEmpty(cacheDir))
                cacheDir = Path.Combine(Path.GetTempPath(), "StatikManager", "wortklone");

            var hash = ((uint)pdfPfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            var name = Path.GetFileNameWithoutExtension(pdfPfad);
            return Path.Combine(cacheDir, $"{name}_{hash}_klon.docx");
        }

        /// <summary>Prüft ob eine gecachte Datei noch gültig (aktueller als die Quelldatei) ist.</summary>
        public static bool CacheGültig(string pdfPfad, string quellPfad)
            => File.Exists(pdfPfad)
            && File.GetLastWriteTime(pdfPfad) >= File.GetLastWriteTime(quellPfad);

        /// <summary>Löscht alle Cache-Dateien für eine bestimmte Quelldatei.</summary>
        public static void LöscheCacheFürDatei(string quellPfad, string cacheDir)
        {
            if (string.IsNullOrEmpty(cacheDir)) return;
            try
            {
                var hash   = ((uint)quellPfad.ToLowerInvariant().GetHashCode()).ToString("X8");
                var name   = Path.GetFileNameWithoutExtension(quellPfad);
                var präfix = $"{name}_{hash}_{CacheVersion}";

                foreach (var f in Directory.GetFiles(cacheDir, präfix + "*.pdf"))
                    try { File.Delete(f); } catch { }  // Best-Effort: einzelne Datei gesperrt → überspringen
            }
            catch (Exception ex)
            {
                Logger.Warn("PdfCache.Löschen",
                    $"Cache-Bereinigung fehlgeschlagen für '{Path.GetFileName(quellPfad)}': {ex.Message}");
            }
        }
    }
}
