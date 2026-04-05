using StatikManager.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace StatikManager.Modules.Dokumente
{
    public partial class ProjektVerwaltungDialog : Window
    {
        // ── Datenklasse für die DataGrid-Zeilen ───────────────────────────────

        private sealed class ProjektDialogItem : INotifyPropertyChanged
        {
            private bool   _sichtbar = true;
            private string _kurzname = "";

            public string Pfad       { get; set; } = "";
            public string OrdnerName => Path.GetFileName(Pfad.TrimEnd(Path.DirectorySeparatorChar));

            public bool Sichtbar
            {
                get => _sichtbar;
                set { _sichtbar = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Sichtbar))); }
            }

            public string Kurzname
            {
                get => _kurzname;
                set { _kurzname = value ?? ""; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Kurzname))); }
            }

            public event PropertyChangedEventHandler? PropertyChanged;
        }

        // ── Felder ────────────────────────────────────────────────────────────

        private readonly ObservableCollection<ProjektDialogItem> _items = new();
        private string? _basisPfad;

        // ── Fabrik-Methode ────────────────────────────────────────────────────

        /// <summary>
        /// Zeigt den Dialog. Gibt true zurück wenn der Benutzer OK gedrückt hat
        /// (Einstellungen wurden gespeichert).
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

            _basisPfad = Einstellungen.Instanz.ProjektBasisPfad;
            ZeigeBasisPfad();
            Scannen();
        }

        // ── Basispfad ─────────────────────────────────────────────────────────

        private void ZeigeBasisPfad()
        {
            TxtBasisPfad.Text = string.IsNullOrWhiteSpace(_basisPfad)
                ? "(noch nicht festgelegt)"
                : _basisPfad;
        }

        private void BtnBasisPfadÄndern_Click(object sender, RoutedEventArgs e)
        {
            var neuerPfad = OrdnerDialog.Zeigen(
                startPfad: _basisPfad ?? "",
                titel:     "Projektbasis-Ordner wählen",
                besitzer:  this);

            if (string.IsNullOrWhiteSpace(neuerPfad)) return;

            _basisPfad = neuerPfad;
            ZeigeBasisPfad();
            Scannen();
        }

        private void BtnNeuScannen_Click(object sender, RoutedEventArgs e) => Scannen();

        // ── Scan-Logik ────────────────────────────────────────────────────────

        /// <summary>
        /// Liest alle ersten-Ebene-Unterordner des Basispfads ein und füllt die Liste.
        /// Bestehende Einträge (Sichtbar / Kurzname) werden beibehalten.
        /// </summary>
        private void Scannen()
        {
            _items.Clear();

            if (string.IsNullOrWhiteSpace(_basisPfad) || !Directory.Exists(_basisPfad))
                return;

            var bekannte = Einstellungen.Instanz.ProjektEintraege
                .ToDictionary(e => e.Pfad.ToLowerInvariant(), e => e);

            IEnumerable<string> unterordner;
            try
            {
                unterordner = Directory.GetDirectories(_basisPfad)
                    .OrderBy(p => Path.GetFileName(p), StringComparer.OrdinalIgnoreCase);
            }
            catch { unterordner = Enumerable.Empty<string>(); }

            foreach (var pfad in unterordner)
            {
                bekannte.TryGetValue(pfad.ToLowerInvariant(), out var bekannt);
                _items.Add(new ProjektDialogItem
                {
                    Pfad     = pfad,
                    Sichtbar = bekannt?.Sichtbar ?? true,
                    Kurzname = bekannt?.Kurzname ?? ""
                });
            }
        }

        // ── Alle an / aus ─────────────────────────────────────────────────────

        private void BtnAlleAn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.Sichtbar = true;
        }

        private void BtnAlleAus_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items) item.Sichtbar = false;
        }

        // ── OK / Abbrechen ────────────────────────────────────────────────────

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            // DataGrid-Änderungen übernehmen (laufende Cell-Edits committen)
            DgProjekte.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, exitEditingMode: true);

            Einstellungen.Instanz.ProjektBasisPfad = _basisPfad;

            // Alle gescannten Ordner (auch deaktivierte) speichern
            var neueEintraege = _items.Select(i => new ProjektEintrag
            {
                Pfad     = i.Pfad,
                Kurzname = i.Kurzname.Trim(),
                Sichtbar = i.Sichtbar
            }).ToList();

            // Einträge die nicht mehr im Basispfad liegen aber woanders waren: verwerfen
            Einstellungen.Instanz.ProjektEintraege = neueEintraege;
            Einstellungen.Instanz.Speichern();

            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
