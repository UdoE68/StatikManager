using System.Threading;
using System.Windows.Forms;

namespace StatikManager.Api.Infrastructure;

/// <summary>
/// Öffnet einen Windows-Ordnerdialog ausschließlich auf einem STA-Thread (kein Task.Run, kein async).
/// </summary>
internal static class WindowsFolderPicker
{
    private static readonly object VisualStylesLock = new();
    private static bool _visualStylesRegistered;

    private static void EnsureVisualStylesOnce()
    {
        lock (VisualStylesLock)
        {
            if (_visualStylesRegistered)
                return;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            _visualStylesRegistered = true;
        }
    }

    /// <summary>
    /// Zeigt den Dialog. Rückgabe: gewählter Pfad, oder <c>null</c> bei Abbruch.
    /// </summary>
    /// <exception cref="Exception">Dialog oder WinForms-Fehler.</exception>
    public static string? PickFolderOnStaThread()
    {
        string? auswahl = null;
        Exception? threadFehler = null;

        void DialogThreadProc()
        {
            try
            {
                EnsureVisualStylesOnce();

                using var dlg = new FolderBrowserDialog
                {
                    Description = "Projektordner auswählen",
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton = true,
                };

                var dr = dlg.ShowDialog();
                if (dr == DialogResult.OK)
                    auswahl = dlg.SelectedPath;
            }
            catch (Exception ex)
            {
                threadFehler = ex;
            }
        }

        var thread = new Thread(new ThreadStart(DialogThreadProc))
        {
            IsBackground = false,
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadFehler is not null)
            throw threadFehler;

        return auswahl;
    }
}
