using System.Collections.Generic;

namespace StatikManager.Core
{
    /// <summary>
    /// Verwaltet alle registrierten Module.
    /// Neue Funktionalität wird durch Registrieren eines weiteren IModul hinzugefügt –
    /// ohne Änderungen an MainWindow oder anderen Modulen.
    /// </summary>
    public class ModulManager
    {
        private readonly List<IModul> _module = new();

        public IReadOnlyList<IModul> Module => _module;

        /// <summary>Registriert ein Modul. Reihenfolge bestimmt die Anzeige.</summary>
        public void Registrieren(IModul modul) => _module.Add(modul);

        /// <summary>Sucht ein Modul anhand seiner Id.</summary>
        public IModul? FindeModul(string id) => _module.Find(m => m.Id == id);

        /// <summary>Ruft Bereinigen() auf allen Modulen auf.</summary>
        public void AllesBereinigen()
        {
            foreach (var m in _module)
                m.Bereinigen();
        }
    }
}
