using StatikManager.Core;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Modules.Dokumente
{
    public partial class ProjektVerwaltungDialog : Window
    {
        // ── Datenklasse für DataGrid-Zeilen ───────────────────────────────────

        private sealed class ProjektDialogItem : INotifyPropertyChanged
        {
            private bool   _sichtbar = true;
            private string _kurzname = "";

            public string Pfad { get; set; } = "";

            public bool Sichtbar
            {
                get => _sichtbar;
                set { _sichtbar = value; Notify(nameof(Sichtbar)); }
            }

            public string Kurzname
            {
                get => _kurzname;
                set { _kurzname = value ?? ""; Notify(nameof(Kurzname)); }
            }

            public event PropertyChangedEventHandler? PropertyChanged;
            private void Notify(string name) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // ── Felder ────────────────────────────────────────────────────────────

        private readonly ObservableCollection<ProjektDialogItem> _items = new();

        // ── Fabrik-Methode ────────────────────────────────────────────────────

        /// <summary>
        /// Zeigt den Dialog modal. Gibt true zurück wenn der Benutzer OK gedrückt hat
        /// (Einstellungen wurden bereits gespeichert).
        /// </summary>
        public static bool Zeigen(Window owner)
        {
            var dlg = new ProjektVerwaltungDialog { Owner = owner };
            return dlg.ShowDialog() == true;
        }

        // ── Konstruktor ───────────────────────────────────────────────────────

        public ProjektVerwaltungDialog()
        {
            InitializeComponent();
            DgProjekte.ItemsSource = _items;
            LadeEintraege();
        }

        // ── Datenladen ────────────────────────────────────────────────────────

        private void LadeEintraege()
        {
            _items.Clear();
            foreach (var e in Einstellungen.Instanz.ProjektEintraege)
            {
                _items.Add(new ProjektDialogItem
                {
                    Pfad     = e.Pfad,
                    Kurzname = e.Kurzname,
                    Sichtbar = e.Sichtbar
                });
            }
        }

        // ── Button-Status aktualisieren ───────────────────────────────────────

        private void DgProjekte_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int idx   = DgProjekte.SelectedIndex;
            int count = _items.Count;
            BtnEntfernen.IsEnabled  = idx >= 0;
            BtnNachOben.IsEnabled   = idx > 0;
            BtnNachUnten.IsEnabled  = idx >= 0 && idx < count - 1;
        }

        // ── Hinzufügen ────────────────────────────────────────────────────────

        private void BtnHinzufügen_Click(object sender, RoutedEventArgs e)
        {
            var pfad = OrdnerDialog.Zeigen(
                startPfad: Einstellungen.Instanz.StandardPfad ?? "",
                titel:     "Projektordner auswählen",
                besitzer:  this);

            if (string.IsNullOrWhiteSpace(pfad)) return;

            // Duplikat-Prüfung
            foreach (var item in _items)
            {
                if (string.Equals(item.Pfad, pfad, System.StringComparison.OrdinalIgnoreCase))
                {
                    DgProjekte.SelectedItem = item;
                    DgProjekte.ScrollIntoView(item);
                    return;
                }
            }

            var neuer = new ProjektDialogItem { Pfad = pfad, Sichtbar = true };
            _items.Add(neuer);
            DgProjekte.SelectedItem = neuer;
            DgProjekte.ScrollIntoView(neuer);
        }

        // ── Entfernen ─────────────────────────────────────────────────────────

        private void BtnEntfernen_Click(object sender, RoutedEventArgs e)
        {
            if (DgProjekte.SelectedItem is not ProjektDialogItem item) return;
            int idx = _items.IndexOf(item);
            _items.Remove(item);

            // Selektion auf Nachbareintrag setzen
            if (_items.Count > 0)
                DgProjekte.SelectedIndex = System.Math.Min(idx, _items.Count - 1);
        }

        // ── Reihenfolge ändern ────────────────────────────────────────────────

        private void BtnNachOben_Click(object sender, RoutedEventArgs e)
        {
            int idx = DgProjekte.SelectedIndex;
            if (idx <= 0) return;
            _items.Move(idx, idx - 1);
            DgProjekte.SelectedIndex = idx - 1;
        }

        private void BtnNachUnten_Click(object sender, RoutedEventArgs e)
        {
            int idx = DgProjekte.SelectedIndex;
            if (idx < 0 || idx >= _items.Count - 1) return;
            _items.Move(idx, idx + 1);
            DgProjekte.SelectedIndex = idx + 1;
        }

        // ── OK / Abbrechen ────────────────────────────────────────────────────

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            // Laufende Cell-Edits übernehmen
            DgProjekte.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);

            var neue = new System.Collections.Generic.List<ProjektEintrag>();
            foreach (var item in _items)
            {
                neue.Add(new ProjektEintrag
                {
                    Pfad     = item.Pfad,
                    Kurzname = item.Kurzname.Trim(),
                    Sichtbar = item.Sichtbar
                });
            }

            Einstellungen.Instanz.ProjektEintraege = neue;
            Einstellungen.Instanz.Speichern();

            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
