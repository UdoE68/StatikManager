using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;

namespace StatikManager.Core
{
    /// <summary>
    /// Beschreibt eine Word-Vorlage (.dotx / .docx) für den Export.
    /// Serialisierbar via XmlSerializer; implementiert INotifyPropertyChanged
    /// für direkte Databinding-Unterstützung in der Einstellungs-UI.
    /// </summary>
    public sealed class WordVorlage : INotifyPropertyChanged
    {
        private string _name     = "";
        private string _pfad     = "";
        private bool   _standard;

        /// <summary>Anzeigename der Vorlage.</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        /// <summary>Vollständiger Pfad zur .dotx- / .docx-Datei.</summary>
        public string Pfad
        {
            get => _pfad;
            set
            {
                _pfad = value;
                OnPropertyChanged(nameof(Pfad));
                OnPropertyChanged(nameof(PfadGültig));
            }
        }

        /// <summary>Gibt an ob dies die aktive Standardvorlage ist.</summary>
        public bool Standard
        {
            get => _standard;
            set { _standard = value; OnPropertyChanged(nameof(Standard)); }
        }

        /// <summary>
        /// true wenn Pfad leer (noch nicht gesetzt) ODER die Datei tatsächlich existiert.
        /// false = Pfad angegeben, Datei aber nicht gefunden → rote Markierung in der UI.
        /// </summary>
        [XmlIgnore]
        public bool PfadGültig =>
            string.IsNullOrWhiteSpace(_pfad) || File.Exists(_pfad);

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
