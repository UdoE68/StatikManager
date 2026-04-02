using System;
using System.IO;
using System.Xml.Serialization;

namespace StatikManager.Core
{
    /// <summary>
    /// Persistenter Sitzungszustand.
    /// Wird beim Schließen gespeichert und beim nächsten Start wiederhergestellt.
    /// Gespeichert unter %APPDATA%\StatikManager\sitzung.xml.
    /// </summary>
    [XmlRoot("SitzungsZustand")]
    public sealed class SitzungsZustand
    {
        // ── Projekt & Datei ───────────────────────────────────────────────────

        /// <summary>Zuletzt geöffneter Projektordner.</summary>
        public string? ProjektPfad { get; set; }

        /// <summary>Zuletzt ausgewählte Datei.</summary>
        public string? AktiveDatei { get; set; }

        // ── PdfSchnittEditor – Darstellung ────────────────────────────────────

        public double ZoomFaktor       { get; set; } = 1.0;
        public bool   LayoutHorizontal { get; set; }
        public double ScrollH          { get; set; }
        public double ScrollV          { get; set; }

        // ── Crop-Anwendungsmodus ──────────────────────────────────────────────

        /// <summary>0 = NurDiese, 1 = Alle, 2 = Ausgewählt, 3 = AlsStandard</summary>
        public int   CropModus        { get; set; }
        public int[] CropAuswahlSeiten { get; set; } = Array.Empty<int>();

        // ── Default-Crop ──────────────────────────────────────────────────────

        public bool   DefaultCropGesetzt { get; set; }
        public double DefaultCropLinks   { get; set; }
        public double DefaultCropRechts  { get; set; }
        public double DefaultCropOben    { get; set; }
        public double DefaultCropUnten   { get; set; }

        // ── Per-Seite Crop-Arrays ─────────────────────────────────────────────

        public double[] CropLinks  { get; set; } = Array.Empty<double>();
        public double[] CropRechts { get; set; } = Array.Empty<double>();
        public double[] CropOben   { get; set; } = Array.Empty<double>();
        public double[] CropUnten  { get; set; } = Array.Empty<double>();

        // ── Persistenz ────────────────────────────────────────────────────────

        private static readonly string DateiPfad = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "StatikManager", "sitzung.xml");

        /// <summary>
        /// Lädt die gespeicherte Sitzung. Gibt eine leere Sitzung zurück wenn keine existiert
        /// oder die Datei nicht gelesen werden kann.
        /// </summary>
        public static SitzungsZustand Laden()
        {
            try
            {
                if (!File.Exists(DateiPfad)) return new SitzungsZustand();
                using var fs = File.OpenRead(DateiPfad);
                var xs = new XmlSerializer(typeof(SitzungsZustand));
                return (SitzungsZustand?)xs.Deserialize(fs) ?? new SitzungsZustand();
            }
            catch { return new SitzungsZustand(); }
        }

        /// <summary>
        /// Speichert den Sitzungszustand. Fehler werden stillschweigend ignoriert.
        /// </summary>
        public static void Speichern(SitzungsZustand sitzung)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(DateiPfad)!);
                using var fs = File.Create(DateiPfad);
                var xs = new XmlSerializer(typeof(SitzungsZustand));
                xs.Serialize(fs, sitzung);
            }
            catch { }
        }
    }
}
