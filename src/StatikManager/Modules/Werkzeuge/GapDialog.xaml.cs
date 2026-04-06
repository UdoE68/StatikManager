using System.Globalization;
using System.Windows;

namespace StatikManager.Modules.Werkzeuge
{
    public partial class GapDialog : Window
    {
        // ── Rückgabewerte ─────────────────────────────────────────────────────
        public bool     Bestätigt      { get; private set; }
        public GapModus GewählterModus { get; private set; }
        public double   EingabeGapMm   { get; private set; }

        // ── Konstruktor ───────────────────────────────────────────────────────
        /// <summary>
        /// Öffnet den Dialog. Optionale Parameter füllen ihn für "Bearbeiten"-Modus vor.
        /// </summary>
        public GapDialog(GapModus aktuellModus = GapModus.OriginalAbstand,
                         double   aktuellMm   = 0.0)
        {
            InitializeComponent();

            switch (aktuellModus)
            {
                case GapModus.KundenAbstand:
                    RbKunden.IsChecked = true;
                    TxtMm.Text = aktuellMm.ToString("F1", CultureInfo.CurrentCulture);
                    break;
                case GapModus.KeinAbstand:
                    RbKein.IsChecked = true;
                    break;
                default:
                    RbOriginal.IsChecked = true;
                    break;
            }

            AktualisiereTextBox();
        }

        // ── Events ────────────────────────────────────────────────────────────
        private void Rb_Checked(object sender, RoutedEventArgs e)
            => AktualisiereTextBox();

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (RbKunden.IsChecked == true)
            {
                // Komma und Punkt akzeptieren
                string raw = TxtMm.Text.Replace(",", ".");
                if (!double.TryParse(raw, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double mm) || mm < 0)
                {
                    MessageBox.Show("Bitte eine gültige Zahl ≥ 0 eingeben.",
                        "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Warning);
                    TxtMm.Focus();
                    return;
                }
                GewählterModus = GapModus.KundenAbstand;
                EingabeGapMm   = mm;
            }
            else if (RbKein.IsChecked == true)
            {
                GewählterModus = GapModus.KeinAbstand;
                EingabeGapMm   = 0.0;
            }
            else
            {
                GewählterModus = GapModus.OriginalAbstand;
                EingabeGapMm   = 0.0;
            }

            Bestätigt    = true;
            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            Bestätigt    = false;
            DialogResult = false;
        }

        // ── Hilfsmethoden ─────────────────────────────────────────────────────
        private void AktualisiereTextBox()
        {
            if (TxtMm == null) return;
            bool aktivieren = RbKunden.IsChecked == true;
            TxtMm.IsEnabled = aktivieren;
            if (aktivieren) TxtMm.Focus();
        }
    }
}
