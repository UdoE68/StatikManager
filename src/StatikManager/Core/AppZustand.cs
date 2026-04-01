using System;

namespace StatikManager.Core
{
    public enum StatusLevel { Info, Warn, Error }

    /// <summary>
    /// Gemeinsamer Anwendungszustand. Module und Shell kommunizieren über Events,
    /// ohne direkte Referenzen aufeinander zu haben.
    /// </summary>
    public sealed class AppZustand
    {
        public static AppZustand Instanz { get; } = new();
        private AppZustand() { }

        // ── Statusleiste ─────────────────────────────────────────────────────

        private string _statusText = "Bereit";
        public string StatusText => _statusText;
        public StatusLevel StatusLevel { get; private set; } = StatusLevel.Info;

        public event Action<string, StatusLevel>? StatusGeändert;

        public void SetzeStatus(string text, StatusLevel level = StatusLevel.Info)
        {
            _statusText = text;
            StatusLevel = level;
            StatusGeändert?.Invoke(text, level);
        }

        // ── Ladezustand ──────────────────────────────────────────────────────────
        // Signalisiert, ob gerade ein Dokument gerendert wird.
        // PdfSchnittEditor setzt true/false; DokumentePanel reagiert via Event.

        public bool IsLaden { get; private set; }
        public event Action<bool>? LadeZustandGeändert;

        public void SetzeLaden(bool aktiv)
        {
            IsLaden = aktiv;
            LadeZustandGeändert?.Invoke(aktiv);
        }

        // ── Fortschritt ──────────────────────────────────────────────────────
        // gesamt = 0 → Fortschrittsleiste ausblenden

        public event Action<int, int>? FortschrittGeändert;

        public void SetzeProgress(int aktuell, int gesamt)
            => FortschrittGeändert?.Invoke(aktuell, gesamt);

        public void ResetProgress()
            => FortschrittGeändert?.Invoke(0, 0);

        // ── PDF/Vorschau-Render-Schutz ───────────────────────────────────────
        // Serialisiert alle pdfium- und Word-COM-Zugriffe aus Hintergrund-Threads.
        // Beide UserControls (DokumentePanel + PdfSchnittEditor) verwenden diese
        // gemeinsame Semaphore, damit niemals zwei Threads gleichzeitig auf die
        // nativen Bibliotheken zugreifen.
        public static readonly System.Threading.SemaphoreSlim RenderSem = new(1, 1);

        // ── Aktuelles Projekt ─────────────────────────────────────────────────

        private string? _projektPfad;
        public string? ProjektPfad => _projektPfad;

        public event Action<string?>? ProjektGeändert;

        public void SetzeProjekt(string? pfad)
        {
            _projektPfad = pfad;
            ProjektGeändert?.Invoke(pfad);
        }
    }
}
