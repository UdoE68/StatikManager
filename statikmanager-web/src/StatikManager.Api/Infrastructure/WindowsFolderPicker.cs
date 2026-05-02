using System.Threading;
using System.Windows.Forms;

namespace StatikManager.Api.Infrastructure;

/// <summary>
/// Öffnet einen Windows-Ordnerdialog auf einem dedizierten STA-Thread.
/// </summary>
internal static class WindowsFolderPicker
{
    /// <summary>Gibt den gewählten Pfad zurück oder null bei Abbruch/Fehler im Dialog.</summary>
    public static string? PickFolder()
    {
        string? selected = null;
        Exception? caught = null;

        void Run()
        {
            try
            {
                using var dlg = new FolderBrowserDialog
                {
                    Description = "Projektordner auswählen",
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton = true,
                };

                if (dlg.ShowDialog() == DialogResult.OK)
                    selected = dlg.SelectedPath;
            }
            catch (Exception ex)
            {
                caught = ex;
            }
        }

        var thread = new Thread(Run)
        {
            IsBackground = false,
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (caught is not null)
            throw caught;

        return selected;
    }
}
