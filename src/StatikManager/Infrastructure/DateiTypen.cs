using System;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Zentrale Dateityp-Klassifikation für das StatikManager-Projekt.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal static class DateiTypen
    {
        public static bool IstWordDatei(string ext)
            => ext.Equals(".doc",  StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".docx", StringComparison.OrdinalIgnoreCase);

        public static bool IstPdfDatei(string ext)
            => ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase);

        public static bool IstBildDatei(string ext)
        {
            var e = ext.ToLowerInvariant();
            return e == ".jpg" || e == ".jpeg" || e == ".png"
                || e == ".gif" || e == ".bmp"  || e == ".tif" || e == ".tiff";
        }

        public static bool IstHtmlDatei(string ext)
        {
            var e = ext.ToLowerInvariant();
            return e == ".html" || e == ".htm";
        }

        public static bool IstJsonDatei(string ext)
            => ext.Equals(".json", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Dateitypen, die im Statik-Manager nie per Shell geöffnet werden dürfen.
        /// Beispiel: .axs würde AxisVM starten; .exe/.bat sind gefährlich.
        /// </summary>
        public static bool IstGesperrteExtension(string ext)
        {
            var e = ext.ToLowerInvariant();
            return e == ".axs"  // AxisVM-Projektdatei → würde AxisVM starten
                || e == ".exe"
                || e == ".bat"
                || e == ".cmd"
                || e == ".msi";
        }

        public static string DateiIcon(string ext)
        {
            var e = ext.ToLowerInvariant();
            if (e == ".doc"  || e == ".docx")                                return "📄";
            if (e == ".xls"  || e == ".xlsx"  || e == ".xlsm")               return "📊";
            if (e == ".pdf")                                                   return "📑";
            if (e == ".jpg"  || e == ".jpeg"  || e == ".png"
             || e == ".gif"  || e == ".bmp"   || e == ".tif"  || e == ".tiff") return "🖼️";
            if (e == ".json") return "{ }";
            return "📎";
        }
    }
}
