using System;
using System.IO;
using System.Threading;
using System.Windows.Threading;
using StatikManager.Core;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Orchestriert den Auto-Refresh der Word-Vorschau nach einer Dateiänderung.
    /// Empfängt DateiGeändertGemeldet() vom Panel, löscht den Cache, konvertiert
    /// .docx → PDF auf einem STA-Hintergrundthread und feuert Events zurück auf den UI-Thread.
    /// Bei Fehler: einmaliger Retry nach 3 Sekunden.
    /// </summary>
    internal sealed class WordAutoRefreshService : IDisposable
    {
        private readonly Dispatcher _dispatcher;

        private string? _pfad;
        private string  _cacheDir = "";

        private CancellationTokenSource? _cts;
        private DispatcherTimer?         _retryTimer;
        private bool                     _retryAusstehend;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung startet.</summary>
        public event Action? KonvertierungGestartet;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung erfolgreich war.
        /// Parameter: vollständiger Pfad zum erzeugten Basis-PDF.</summary>
        public event Action<string>? VorschauBereit;

        /// <summary>Ausgelöst auf dem UI-Thread wenn die Konvertierung fehlgeschlagen ist
        /// (nach Retry). Parameter: Fehlermeldung für die Statuszeile.</summary>
        public event Action<string>? KonvertierungFehlgeschlagen;

        public WordAutoRefreshService(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        /// <summary>
        /// Setzt die zu überwachende Datei und das Cache-Verzeichnis.
        /// Bricht eine laufende Konvertierung ab.
        /// </summary>
        public void Starte(string docxPfad, string cacheDir)
        {
            Stoppe();
            _pfad     = docxPfad;
            _cacheDir = cacheDir;
            _retryAusstehend = false;
        }

        /// <summary>
        /// Bricht laufende Konvertierung und ausstehenden Retry ab.
        /// </summary>
        public void Stoppe()
        {
            _cts?.Cancel();
            _cts = null;
            _retryTimer?.Stop();
            _retryTimer = null;
            _retryAusstehend = false;
            _pfad     = null;
        }

        /// <summary>
        /// Wird vom Panel aufgerufen wenn FileWatcherService eine Änderung meldet.
        /// Muss auf dem UI-Thread aufgerufen werden.
        /// </summary>
        public void DateiGeändertGemeldet()
        {
            if (_pfad == null) return;
            _retryTimer?.Stop();
            _retryTimer = null;
            _retryAusstehend = false;
            StarteKonvertierung(_pfad, _cacheDir, istRetry: false);
        }

        public void Dispose() => Stoppe();

        // ── Private ──────────────────────────────────────────────────────────

        private void StarteKonvertierung(string pfad, string cacheDir, bool istRetry)
        {
            // Laufenden Thread abbrechen
            _cts?.Cancel();
            var cts = new CancellationTokenSource();
            _cts = cts;
            var token = cts.Token;

            // Cache löschen damit WortDateiZuPdf neu konvertiert
            PdfCache.LöscheCacheFürDatei(pfad, cacheDir);

            var effektiverCacheDir = string.IsNullOrEmpty(cacheDir)
                ? Path.Combine(Path.GetTempPath(), "StatikManager", "preview")
                : cacheDir;

            var basisPdf = PdfCache.GetBasisPdfPfad(pfad, effektiverCacheDir);
            Directory.CreateDirectory(Path.GetDirectoryName(basisPdf)!);

            if (!istRetry)
            {
                AppZustand.Instanz.SetzeStatus(
                    "Vorschau wird aktualisiert: " + Path.GetFileName(pfad) + " …");
            }
            else
            {
                AppZustand.Instanz.SetzeStatus(
                    "Erneuter Versuch: " + Path.GetFileName(pfad) + " …");
            }

            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (token.IsCancellationRequested) return;
                KonvertierungGestartet?.Invoke();
            }));

            Logger.Info("WordAutoRefresh",
                $"{(istRetry ? "[Retry] " : "")}Starte Konvertierung: {Path.GetFileName(pfad)}");

            var t = new Thread(() =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;
                    WordInteropService.WortDateiZuPdf(pfad, basisPdf);

                    if (token.IsCancellationRequested) return;

                    Logger.Info("WordAutoRefresh",
                        $"Konvertierung erfolgreich: {Path.GetFileName(pfad)}");

                    var basisPdfFinal = basisPdf;
                    _dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested) return;
                        AppZustand.Instanz.SetzeStatus(
                            "Vorschau aktualisiert: " + Path.GetFileName(pfad));
                        VorschauBereit?.Invoke(basisPdfFinal);
                    }));
                }
                catch (Exception ex) when (!(ex is OperationCanceledException))
                {
                    Logger.Warn("WordAutoRefresh",
                        $"Konvertierung fehlgeschlagen ({Path.GetFileName(pfad)}): {ex.Message}");

                    _dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested) return;

                        if (!istRetry)
                        {
                            // Erster Fehler → Retry nach 3s
                            _retryAusstehend = true;
                            AppZustand.Instanz.SetzeStatus(
                                "Konvertierung fehlgeschlagen – erneuter Versuch in 3 s …",
                                StatusLevel.Warn);
                            _retryTimer = new DispatcherTimer
                                { Interval = TimeSpan.FromSeconds(3) };
                            _retryTimer.Tick += (_, _) =>
                            {
                                _retryTimer?.Stop();
                                _retryTimer = null;
                                if (_retryAusstehend && _pfad == pfad)
                                    StarteKonvertierung(pfad, cacheDir, istRetry: true);
                            };
                            _retryTimer.Start();
                        }
                        else
                        {
                            // Retry auch fehlgeschlagen → endgültig
                            _retryAusstehend = false;
                            var fehler = "Konvertierung fehlgeschlagen – Vorschau veraltet ("
                                         + Path.GetFileName(pfad) + ")";
                            AppZustand.Instanz.SetzeStatus(fehler, StatusLevel.Warn);
                            KonvertierungFehlgeschlagen?.Invoke(fehler);
                        }
                    }));
                }
            })
            { IsBackground = true, Name = "WordAutoRefresh" };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }
    }
}
