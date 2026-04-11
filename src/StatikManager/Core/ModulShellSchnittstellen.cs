namespace StatikManager.Core
{
    /// <summary>Wird beim Start aufgerufen, um die gespeicherte Sitzung wiederherzustellen.</summary>
    public interface ISitzungsWiederherstellung
    {
        void SitzungWiederherstellen(SitzungsZustand sitzung);
    }

    /// <summary>Liefert den aktuellen Sitzungszustand zum Speichern beim Beenden.</summary>
    public interface ISitzungsPersistenz
    {
        SitzungsZustand SitzungSpeichern();
    }

    /// <summary>„Projekt laden …“ aus dem Datei-Menü der Shell.</summary>
    public interface IProjektLadenAusShell
    {
        void ProjektLaden();
    }

    /// <summary>Rückfrage vor Schließen des Hauptfensters (z. B. ungespeicherte PDF-Änderungen).</summary>
    public interface IHauptfensterSchliessenPruefung
    {
        /// <returns>false, wenn der Benutzer das Schließen abgebrochen hat</returns>
        bool DarfHauptfensterSchliessen();
    }

    /// <summary>Wird aufgerufen, wenn das Modul-Panel zur Anzeige gewechselt wurde.</summary>
    public interface IBeiModulAnzeige
    {
        /// <param name="kontextDateiPfad">z. B. beim Wechsel zu Bildschnitt die zu ladende Datei</param>
        void BeiAnzeige(string? kontextDateiPfad);
    }
}
