using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;

namespace StatikManager
{
    public partial class App : Application
    {
        // ── Build-Kennung (Zeitstempel wird beim Kompilieren eingebettet) ─────
        // Wenn dieser Text im Fehlerdialog erscheint, läuft definitiv dieser Code.
        internal static readonly string BuildZeit =
            new DateTime(2000, 1, 1)
            .AddDays(System.Reflection.Assembly.GetExecutingAssembly()
                .GetName().Version?.Build ?? 0)
            .ToString("yyyy-MM-dd");

        // Log-Pfad liegt im Logger – hier nur noch Weiterleitung
        internal static readonly string LogDatei = Infrastructure.Logger.LogDatei;

        // ── Einzelinstanz-Steuerung ───────────────────────────────────────────
        private const string MUTEX_NAME = "StatikManager_SingleInstance";
        private const string PIPE_NAME  = "StatikManagerPipe";
        private static System.Threading.Mutex _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            // ── 0. Einzelinstanz-Prüfung per Mutex ───────────────────────────
            bool neuInstanz;
            _mutex = new System.Threading.Mutex(true, MUTEX_NAME, out neuInstanz);

            if (!neuInstanz)
            {
                // Bereits eine Instanz aktiv – Pfad per Named Pipe übergeben
                if (e.Args.Length > 0)
                {
                    try
                    {
                        using (var client = new System.IO.Pipes.NamedPipeClientStream(
                            ".", PIPE_NAME, System.IO.Pipes.PipeDirection.Out))
                        {
                            client.Connect(1000);
                            using (var writer = new System.IO.StreamWriter(client))
                            {
                                writer.Write(e.Args[0]);
                                writer.Flush();
                            }
                        }
                    }
                    catch { }
                }
                Environment.Exit(0);
                return;
            }

            // Erste Instanz – Pipe-Server starten
            StarteNamedPipeServer();
            // ── 1. AppDomain – fängt auch Nicht-UI-Thread-Fehler ─────────────
            AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
            {
                var fehler = ex.ExceptionObject as Exception;
                var details = fehler != null
                    ? GetExceptionKette(fehler)
                    : ex.ExceptionObject?.ToString() ?? "(null)";
                LogFehler("AppDomain.UnhandledException", details);
                // Kein MessageBox – App ist ggf. schon in fataler Lage
            };

            // ── 2. WPF-UI-Thread ──────────────────────────────────────────────
            DispatcherUnhandledException += (s, ex) =>
            {
                ex.Handled = true;

                // Innerste Ursache auspacken:
                // TargetInvocationException ist nur der Dispatcher-Wrapper.
                var innerste = EntpackeTIE(ex.Exception);
                var kette    = GetExceptionKette(ex.Exception);
                LogFehler("DispatcherUnhandledException", kette);

                // Dialog zeigt sowohl innerste Ursache als auch vollständige Kette
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("═══ EIGENTLICHE URSACHE ═══");
                sb.AppendLine($"Typ:     {innerste.GetType().FullName}");
                sb.AppendLine($"Meldung: {innerste.Message}");
                if (innerste.StackTrace != null)
                {
                    var zeilen = innerste.StackTrace.Split('\n');
                    int zeige  = Math.Min(zeilen.Length, 6);
                    for (int i = 0; i < zeige; i++)
                        sb.AppendLine(zeilen[i].Trim());
                }
                sb.AppendLine();
                sb.AppendLine("═══ VOLLSTÄNDIGE KETTE ═══");
                sb.AppendLine(kette);
                sb.AppendLine();
                sb.AppendLine($"[Log: {LogDatei}]");

                MessageBox.Show(sb.ToString(),
                    "Unbehandelter Fehler – StatikManager",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            };

            // ── 3. Unbeobachtete Task-Exceptions ─────────────────────────────
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, ex) =>
            {
                ex.SetObserved();
                var kette = GetExceptionKette(ex.Exception);
                LogFehler("UnobservedTaskException", kette);
                // Nur loggen – kein Dialog für Background-Fehler
            };

            SetzeBrowserEmulationsModus();
            base.OnStartup(e);

            // ── Kommandozeilen-Pfad setzen (nach Fenster-Start) ──────────────
            if (e.Args.Length > 0 && System.IO.Directory.Exists(e.Args[0]))
            {
                Dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Loaded,
                    new Action(() =>
                    {
                        Core.AppZustand.Instanz.SetzeProjekt(e.Args[0]);
                    }));
            }
        }

        /// <summary>
        /// Startet den Named-Pipe-Server im Hintergrundthread.
        /// Wartet auf eingehende Pfade von weiteren Programminstanzen.
        /// </summary>
        private void StarteNamedPipeServer()
        {
            var pipeThread = new System.Threading.Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        using (var server = new System.IO.Pipes.NamedPipeServerStream(
                            PIPE_NAME, System.IO.Pipes.PipeDirection.In))
                        {
                            server.WaitForConnection();
                            using (var reader = new System.IO.StreamReader(server))
                            {
                                string neuerPfad = reader.ReadToEnd().Trim();
                                if (!string.IsNullOrEmpty(neuerPfad))
                                {
                                    Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        Core.AppZustand.Instanz.SetzeProjekt(neuerPfad);
                                    });
                                }
                            }
                        }
                    }
                    catch { }
                }
            });
            pipeThread.IsBackground = true;
            pipeThread.Start();
        }

        /// <summary>
        /// Entpackt TargetInvocationException-Ketten und gibt die innerste,
        /// eigentliche Ursache zurück. Unterstützt beliebig tiefe Verschachtelung.
        /// </summary>
        internal static Exception EntpackeTIE(Exception ex)
        {
            var current = ex;
            int guard   = 0;
            while (current is System.Reflection.TargetInvocationException
                   && current.InnerException != null
                   && guard++ < 20)
            {
                current = current.InnerException;
            }
            // Falls nach TIE-Entpackung noch weitere InnerExceptions vorhanden:
            while (current.InnerException != null
                   && guard++ < 20
                   && !(current is AggregateException))
            {
                // Nur wenn InnerException informativer ist (nicht gleicher Typ wie outer)
                if (current.InnerException.GetType() == current.GetType()
                    && current.InnerException.Message == current.Message)
                    break;
                current = current.InnerException;
            }
            return current;
        }

        /// <summary>
        /// Gibt die vollständige Exception-Kette als lesbaren Text zurück.
        /// Alle Ebenen werden rekursiv aufgebaut.
        /// </summary>
        internal static string GetExceptionKette(Exception? ex)
        {
            if (ex == null) return "(null)";
            var sb  = new StringBuilder();
            var cur = ex;
            int lv  = 0;
            while (cur != null && lv < 15)
            {
                string präfix = lv == 0 ? "" : new string(' ', lv * 2) + "↳ ";
                sb.AppendLine($"{präfix}[{cur.GetType().Name}] {cur.Message}");
                if (cur.StackTrace != null && lv <= 3)
                {
                    var z = cur.StackTrace.Split('\n');
                    for (int i = 0; i < Math.Min(z.Length, 5); i++)
                        sb.AppendLine("    " + z[i].Trim());
                }
                cur = cur.InnerException;
                lv++;
            }
            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Schreibt einen Fehler ins Log. Delegiert an Logger.Fehler.
        /// Bleibt als Kompatibilitäts-Einstiegspunkt für bestehende Aufrufer erhalten.
        /// </summary>
        internal static void LogFehler(string kontext, string details)
            => Infrastructure.Logger.Fehler(kontext, details);

        private static void SetzeBrowserEmulationsModus()
        {
            try
            {
                var exeName = Path.GetFileName(
                    Process.GetCurrentProcess().MainModule!.FileName);
                using var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION",
                    writable: true);
                key?.SetValue(exeName, 11001, RegistryValueKind.DWord);
            }
            catch { }
        }
    }
}
