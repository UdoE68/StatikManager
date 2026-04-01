using StatikManager.Core;
using System.IO;
using System.Windows;

namespace StatikManager.Modules.Dokumente
{
    public partial class ProjektLadenDialog : Window
    {
        /// <summary>Der vom Benutzer gewählte Pfad. Nur gültig wenn DialogResult == true.</summary>
        public string? GewähltPfad { get; private set; }

        public ProjektLadenDialog()
        {
            InitializeComponent();

            var standardPfad = Einstellungen.Instanz.StandardPfad;
            if (!string.IsNullOrEmpty(standardPfad) && Directory.Exists(standardPfad))
            {
                TxtStandardpfadAnzeige.Text      = standardPfad;
                TxtStandardpfadAnzeige.Foreground = null; // erbt vom TextBlock-Standard
                BtnStandardÖffnen.IsEnabled       = true;
                TxtPfad.Text                       = standardPfad;
            }
            else if (!string.IsNullOrEmpty(standardPfad))
            {
                TxtStandardpfadAnzeige.Text = standardPfad + "  (Ordner nicht gefunden)";
            }
            else
            {
                TxtStandardpfadAnzeige.Text = "Kein Standardpfad festgelegt  –  Datei › Standardpfad festlegen …";
            }
        }

        private void BtnStandardÖffnen_Click(object sender, RoutedEventArgs e)
        {
            GewähltPfad  = Einstellungen.Instanz.StandardPfad;
            DialogResult = true;
        }

        private void BtnDurchsuchen_Click(object sender, RoutedEventArgs e)
        {
            var startPfad = TxtPfad.Text.Trim();
            if (!Directory.Exists(startPfad))
                startPfad = Einstellungen.Instanz.StandardPfad ?? "";

            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description         = "Projektordner wählen",
                ShowNewFolderButton = false,
                SelectedPath        = startPfad
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                TxtPfad.Text = dialog.SelectedPath;
        }

        private void BtnÖffnen_Click(object sender, RoutedEventArgs e)
        {
            var pfad = TxtPfad.Text.Trim();
            if (!Directory.Exists(pfad))
            {
                MessageBox.Show("Der Ordner existiert nicht:\n" + pfad, "Ordner nicht gefunden",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            GewähltPfad  = pfad;
            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
            => DialogResult = false;
    }
}
