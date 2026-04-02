using StatikManager.Core;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Modules.EinstellungsDialog
{
    /// <summary>
    /// Einstellungs-Fenster mit Navigationsleiste und austauschbarem Inhaltsbereich.
    /// Änderungen werden erst beim Klick auf OK in Einstellungen.Instanz geschrieben.
    /// </summary>
    public partial class EinstellungenFenster : Window
    {
        // Arbeitskopie der Vorlagenliste (tiefe Kopie – Instanz bleibt bis OK unberührt)
        private readonly ObservableCollection<WordVorlage> _vorlagen;

        public EinstellungenFenster()
        {
            InitializeComponent();

            // Tiefe Kopie aller Vorlagen aus den aktuellen Einstellungen
            _vorlagen = new ObservableCollection<WordVorlage>(
                Einstellungen.Instanz.WordVorlagen.Select(v => new WordVorlage
                {
                    Name     = v.Name,
                    Pfad     = v.Pfad,
                    Standard = v.Standard,
                }));

            DgVorlagen.ItemsSource = _vorlagen;

            // "Löschen"-Button: nur aktiv wenn Zeile selektiert
            DgVorlagen.SelectionChanged += (_, _) =>
                BtnVorageLöschen.IsEnabled = DgVorlagen.SelectedItem != null;

            // Word-Vorlagen standardmäßig anzeigen (Index 1)
            LstNav.SelectedIndex = 1;
        }

        // ── Navigation ────────────────────────────────────────────────────────

        private void LstNav_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PanelWordVorlagen == null) return; // Schutz vor frühem Feuern

            var tag = (LstNav.SelectedItem as ListBoxItem)?.Tag as string;

            PanelWordVorlagen.Visibility =
                tag == "WordVorlagen" ? Visibility.Visible : Visibility.Collapsed;
            PanelPlatzhalter.Visibility =
                tag != "WordVorlagen" ? Visibility.Visible : Visibility.Collapsed;
        }

        // ── Aktionen: Word-Vorlagen ───────────────────────────────────────────

        private void BtnVorlageHinzufügen_Click(object sender, RoutedEventArgs e)
        {
            var neu = new WordVorlage { Name = "Neue Vorlage" };
            _vorlagen.Add(neu);
            DgVorlagen.SelectedItem = neu;
            DgVorlagen.ScrollIntoView(neu);
        }

        private void BtnVorageLöschen_Click(object sender, RoutedEventArgs e)
        {
            if (DgVorlagen.SelectedItem is not WordVorlage sel) return;

            var antwort = MessageBox.Show(
                $"Vorlage \"{sel.Name}\" wirklich löschen?",
                "Vorlage löschen",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (antwort == MessageBoxResult.Yes)
                _vorlagen.Remove(sel);
        }

        private void BtnDateiWählen_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn || btn.DataContext is not WordVorlage vorlage) return;

            var startDir = "";
            if (!string.IsNullOrWhiteSpace(vorlage.Pfad))
                startDir = System.IO.Path.GetDirectoryName(vorlage.Pfad) ?? "";

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title            = "Word-Vorlage auswählen",
                Filter           = "Word-Vorlagen|*.dotx;*.docx|Alle Dateien|*.*",
                InitialDirectory = startDir,
            };

            if (dlg.ShowDialog(this) == true)
                vorlage.Pfad = dlg.FileName;
        }

        /// <summary>
        /// Stellt sicher dass immer genau eine Vorlage als Standard markiert ist.
        /// Alle anderen werden beim Aktivieren automatisch deaktiviert.
        /// </summary>
        private void ChkStandard_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox chk || chk.DataContext is not WordVorlage gewählt) return;

            foreach (var v in _vorlagen)
                if (v != gewählt) v.Standard = false;
        }

        // ── OK / Abbrechen ────────────────────────────────────────────────────

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            // Laufende Zell-Bearbeitung abschließen (damit der letzte Tastendruck gespeichert ist)
            DgVorlagen.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);

            Einstellungen.Instanz.WordVorlagen = _vorlagen.ToList();
            Einstellungen.Instanz.Speichern();
            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
