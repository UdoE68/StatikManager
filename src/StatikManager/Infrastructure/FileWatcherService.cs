using System;
using System.IO;
using System.Windows.Threading;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Kapselt einen FileSystemWatcher für eine einzelne überwachte Datei
    /// sowie einen Debounce-Timer (2 s), um mehrfache Speicher-Events zu bündeln.
    /// Feuert das Ereignis <see cref="DateiGeändert"/> nach dem Debounce auf dem UI-Thread.
    /// Der Dispatcher wird vom Aufrufer übergeben.
    /// Extrahiert aus DokumentePanel.xaml.cs.
    /// </summary>
    internal sealed class FileWatcherService : IDisposable
    {
        private FileSystemWatcher? _watcher;
        private DispatcherTimer?   _debounceTimer;
        private readonly Dispatcher _dispatcher;

        /// <summary>
        /// Wird nach einer Dateiänderung (nach 2-s-Debounce) auf dem UI-Thread ausgelöst.
        /// </summary>
        public event Action? DateiGeändert;

        public FileWatcherService(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        /// <summary>
        /// Startet die Überwachung der angegebenen Datei.
        /// Ein zuvor laufender Watcher wird automatisch gestoppt.
        /// </summary>
        public void Starte(string pfad)
        {
            Stoppe();

            var dir  = Path.GetDirectoryName(pfad);
            var name = Path.GetFileName(pfad);
            if (dir == null || name == null) return;

            // Gesamtes Verzeichnis beobachten (kein Dateinamen-Filter),
            // weil Word beim Speichern eine Temp-Datei erstellt und dann
            // zur Original-Datei umbenennt – dabei wäre ein reiner Changed-
            // Event auf den Originalnamen nicht zuverlässig.
            _watcher = new FileSystemWatcher(dir)
            {
                NotifyFilter          = NotifyFilters.LastWrite
                                      | NotifyFilters.FileName
                                      | NotifyFilters.Size,
                EnableRaisingEvents   = true,
                IncludeSubdirectories = false
            };

            // Nur Events für unsere Zieldatei weiterleiten
            FileSystemEventHandler filter = (s, e) =>
            {
                if (string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
                    OnFsWatcherEvent();
            };
            RenamedEventHandler filterRenamed = (s, e) =>
            {
                // e.Name = neuer Name → trifft zu, wenn Temp-Datei auf Original umbenannt wird
                if (string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
                    OnFsWatcherEvent();
            };

            _watcher.Changed += filter;
            _watcher.Created += filter;
            _watcher.Renamed += filterRenamed;
        }

        /// <summary>Stoppt und verwirft den aktiven Watcher und den Debounce-Timer.</summary>
        public void Stoppe()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }
            _debounceTimer?.Stop();
            _debounceTimer = null;
        }

        public void Dispose() => Stoppe();

        // Datei-Events kommen auf Background-Thread → auf UI-Thread wechseln, dann debounce starten.
        private void OnFsWatcherEvent()
        {
            _dispatcher.BeginInvoke(new Action(() =>
            {
                Logger.Debug("FileWatcher", "Dateiänderung erkannt – Debounce gestartet");
                // Debounce 2 s – Word schreibt in mehreren Schritten
                _debounceTimer?.Stop();
                _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(2000) };
                _debounceTimer.Tick += (s, ev) =>
                {
                    _debounceTimer?.Stop();
                    Logger.Info("FileWatcher", "Dateiänderung ausgelöst (nach Debounce)");
                    DateiGeändert?.Invoke();
                };
                _debounceTimer.Start();
            }));
        }
    }
}
