using System;
using System.IO;
using System.Text;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Zentraler Logger für StatikManager.
    /// Schreibt Einträge mit Zeitstempel und Level in eine Logdatei neben der EXE
    /// und gibt sie zusätzlich über System.Diagnostics.Debug aus.
    /// Wirft niemals selbst eine Exception.
    /// </summary>
    internal static class Logger
    {
        /// <summary>Pfad zur Logdatei neben der EXE.</summary>
        public static readonly string LogDatei = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "StatikManager.log");

        /// <summary>Informative Meldung über eine abgeschlossene Aktion.</summary>
        public static void Info(string kontext, string nachricht)
            => Schreibe("INFO ", kontext, nachricht);

        /// <summary>Warnung: unerwartete Situation, aber kein Fehler.</summary>
        public static void Warn(string kontext, string nachricht)
            => Schreibe("WARN ", kontext, nachricht);

        /// <summary>Fehler: Exception oder fehlgeschlagene Operation.</summary>
        public static void Fehler(string kontext, string details)
            => Schreibe("ERROR", kontext, details);

        /// <summary>Debug-Information für Diagnose (Ablauf, Zwischenstände).</summary>
        public static void Debug(string kontext, string nachricht)
            => Schreibe("DEBUG", kontext, nachricht);

        // ─────────────────────────────────────────────────────────────────────

        private static void Schreibe(string level, string kontext, string text)
        {
            try
            {
                string zeile =
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] [{kontext}]\n" +
                    text +
                    $"\n{new string('─', 80)}\n";

                File.AppendAllText(LogDatei, zeile, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine(zeile);
            }
            catch { /* Log-Fehler niemals propagieren */ }
        }
    }
}
