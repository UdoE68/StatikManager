using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Core
{
    /// <summary>
    /// Schnittstelle für alle Programmmodule.
    /// Ein neues Modul implementiert diese Schnittstelle, registriert sich im ModulManager
    /// und wird automatisch in Menü, Werkzeugleiste und Inhaltsbereich integriert.
    /// </summary>
    public interface IModul
    {
        /// <summary>Eindeutige Kennung des Moduls (z. B. "dokumente").</summary>
        string Id { get; }

        /// <summary>Anzeigename des Moduls.</summary>
        string Name { get; }

        /// <summary>
        /// Erstellt das Hauptpanel, das im Inhaltsbereich des Fensters angezeigt wird.
        /// Wird genau einmal aufgerufen.
        /// </summary>
        UIElement ErstellePanel();

        /// <summary>
        /// Gibt einen Menüeintrag zurück, der ins Hauptmenü (zwischen "Bearbeiten" und "Hilfe")
        /// eingefügt wird. Null, wenn das Modul kein eigenes Menü benötigt.
        /// </summary>
        MenuItem? ErzeugeMenüEintrag();

        /// <summary>
        /// Gibt Elemente zurück, die in die Werkzeugleiste eingefügt werden.
        /// Trennlinien werden als Border mit Width=1 übergeben.
        /// </summary>
        IEnumerable<FrameworkElement> ErzeugeWerkzeugleistenEinträge();

        /// <summary>
        /// Wird beim Schließen der Anwendung aufgerufen.
        /// Ressourcen (z. B. temporäre Dateien) können hier freigegeben werden.
        /// </summary>
        void Bereinigen();
    }
}
