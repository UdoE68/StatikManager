using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace StatikManager.Core
{
    public enum AnsichtModus { Baum, Liste }

    /// <summary>
    /// Persistente Anwendungseinstellungen.
    /// Werden in %APPDATA%\StatikManager\einstellungen.xml gespeichert.
    /// Zugriff über Einstellungen.Instanz (Singleton, lazy geladen).
    /// </summary>
    [XmlRoot("Einstellungen")]
    public sealed class Einstellungen
    {
        // ── Singleton ─────────────────────────────────────────────────────────

        private static Einstellungen? _instanz;
        public static Einstellungen Instanz => _instanz ??= Laden();

        // ── Eigenschaften ─────────────────────────────────────────────────────

        /// <summary>Standardpfad der beim Projekt-öffnen-Dialog vorausgewählt ist.</summary>
        public string? StandardPfad { get; set; }

        /// <summary>Anzeigemodus der Dokumentenliste (Baumstruktur oder flache Liste).</summary>
        public AnsichtModus DokumentAnsicht { get; set; } = AnsichtModus.Baum;

        /// <summary>Liste der konfigurierten Word-Vorlagen für den Export.</summary>
        [XmlArray("WordVorlagen")]
        [XmlArrayItem("Vorlage")]
        public List<WordVorlage> WordVorlagen { get; set; } = new List<WordVorlage>();

        /// <summary>Zuletzt verwendete Projektpfade (keine Dateilöschung beim Entfernen).</summary>
        [XmlArray("ProjektPfade")]
        [XmlArrayItem("Pfad")]
        public List<string> ProjektPfade { get; set; } = new List<string>();

        // ── Persistenz ────────────────────────────────────────────────────────

        private static readonly string DateiPfad = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "StatikManager", "einstellungen.xml");

        private static Einstellungen Laden()
        {
            try
            {
                if (!File.Exists(DateiPfad)) return new Einstellungen();
                using var fs = File.OpenRead(DateiPfad);
                var xs = new XmlSerializer(typeof(Einstellungen));
                return (Einstellungen?)xs.Deserialize(fs) ?? new Einstellungen();
            }
            catch { return new Einstellungen(); }
        }

        public void Speichern()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(DateiPfad)!);
                using var fs = File.Create(DateiPfad);
                var xs = new XmlSerializer(typeof(Einstellungen));
                xs.Serialize(fs, this);
            }
            catch { }
        }
    }
}
