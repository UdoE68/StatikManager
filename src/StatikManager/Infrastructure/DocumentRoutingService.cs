namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Legt fest, welcher Vorschau-Typ für eine gegebene Datei verwendet wird.
    /// Extrahiert aus DokumentePanel.xaml.cs (LadeVorschau-Routing).
    /// </summary>
    internal enum VorschauTyp
    {
        SchnittEditor,
        WordVorschau,
        Browser,
        KeinVorschau
    }

    /// <summary>
    /// Bestimmt anhand der Dateiendung den passenden Vorschau-Typ.
    /// Die eigentlichen UI-Aufrufe verbleiben im DokumentePanel.
    /// </summary>
    internal static class DocumentRoutingService
    {
        public static VorschauTyp ErmittleVorschauTyp(string dateipfad)
        {
            var ext = System.IO.Path.GetExtension(dateipfad).ToLowerInvariant();
            VorschauTyp typ;
            if      (DateiTypen.IstPdfDatei(ext))  typ = VorschauTyp.SchnittEditor;
            else if (DateiTypen.IstWordDatei(ext))  typ = VorschauTyp.WordVorschau;
            else if (DateiTypen.IstBildDatei(ext))  typ = VorschauTyp.Browser;
            else                                     typ = VorschauTyp.KeinVorschau;

            Logger.Debug("Routing", $"{ext} → {typ} ({System.IO.Path.GetFileName(dateipfad)})");
            return typ;
        }
    }
}
