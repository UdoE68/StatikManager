using System;

namespace StatikManager.Core
{
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

        public event Action<string>? StatusGeändert;

        public void SetzeStatus(string text)
        {
            _statusText = text;
            StatusGeändert?.Invoke(text);
        }

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
