using System;
using System.IO;
using System.Windows.Threading;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Überwacht einen Projektordner (inkl. Unterordner) auf strukturelle Änderungen
    /// (neue Dateien/Ordner, gelöschte oder umbenannte Einträge).
    /// Feuert <see cref="OrdnerGeändert"/> nach einem 500-ms-Debounce auf dem UI-Thread.
    /// </summary>
    internal sealed class OrdnerWatcherService : IDisposable
    {
        private FileSystemWatcher? _watcher;
        private DispatcherTimer?   _debounceTimer;
        private readonly Dispatcher _dispatcher;

        /// <summary>
        /// Wird nach einer Ordner-/Dateiänderung (nach 500-ms-Debounce) auf dem UI-Thread ausgelöst.
        /// </summary>
        public event Action? OrdnerGeändert;

        public OrdnerWatcherService(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        /// <summary>
        /// Startet die Überwachung des angegebenen Ordners.
        /// Ein zuvor laufender Watcher wird automatisch gestoppt.
        /// </summary>
        public void Starte(string ordnerPfad)
        {
            Stoppe();
            if (!Directory.Exists(ordnerPfad)) return;

            _watcher = new FileSystemWatcher(ordnerPfad)
            {
                // Nur Datei-/Ordnernamen-Änderungen – kein LastWrite/Size.
                // LastWrite würde bei jedem Speichervorgang hunderte Events erzeugen
                // und den 500-ms-Debounce dauerhaft zurücksetzen.
                NotifyFilter          = NotifyFilters.FileName
                                      | NotifyFilters.DirectoryName,
                IncludeSubdirectories = true,
                EnableRaisingEvents   = true,
                InternalBufferSize    = 65536   // Standard 8 KB reicht bei großen Ordnern nicht
            };

            FileSystemEventHandler handler = (s, e) => OnFsEvent();
            RenamedEventHandler    renamed = (s, e) => OnFsEvent();

            _watcher.Created += handler;
            _watcher.Deleted += handler;
            _watcher.Renamed += renamed;
            // Changed bewusst NICHT abonniert – Changed feuert bei Dateiinhalt-Änderungen,
            // nicht bei strukturellen Änderungen (neue Dateien/Ordner).

            Logger.Debug("OrdnerWatcher", $"Überwachung gestartet: {ordnerPfad}");
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

        // FS-Events kommen auf Background-Thread → auf UI-Thread wechseln, dann Debounce starten.
        private void OnFsEvent()
        {
            _dispatcher.BeginInvoke(new Action(() =>
            {
                // Timer bei jedem Event zurücksetzen → erst nach 500 ms Ruhe feuern
                _debounceTimer?.Stop();
                _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                _debounceTimer.Tick += (s, ev) =>
                {
                    _debounceTimer?.Stop();
                    Logger.Debug("OrdnerWatcher", "Ordnerstruktur geändert – Dokumentenliste wird aktualisiert");
                    OrdnerGeändert?.Invoke();
                };
                _debounceTimer.Start();
            }));
        }
    }
}
