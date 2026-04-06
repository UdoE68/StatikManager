using System.Windows;

namespace StatikManager.Modules.Werkzeuge
{
    /// <summary>
    /// Modaler Dialog zur Auswahl des Lückenabstands beim Block-Löschen.
    /// Vollständige Implementierung folgt in Task 3.
    /// </summary>
    public partial class GapDialog : Window
    {
        public GapDialog()
        {
            InitializeComponent();
        }

        private void Rb_Checked(object sender, RoutedEventArgs e)
        {
            // Implementierung folgt in Task 3
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            // Implementierung folgt in Task 3
            DialogResult = true;
        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            // Implementierung folgt in Task 3
            DialogResult = false;
        }
    }
}
