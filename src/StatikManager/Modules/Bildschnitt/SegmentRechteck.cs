using System.Drawing;

namespace StatikManager.Modules.Bildschnitt
{
    /// <summary>
    /// Ein rechteckiges Bildsegment, das aus der Schnittlinien-Berechnung entsteht.
    /// Koordinaten beziehen sich auf Pixel des Quellbildes.
    /// </summary>
    public class SegmentRechteck
    {
        /// <summary>Zeile im Segmentraster (0-basiert, von oben).</summary>
        public int Zeile  { get; }
        /// <summary>Spalte im Segmentraster (0-basiert, von links).</summary>
        public int Spalte { get; }
        /// <summary>Bildbereich in Quellbild-Pixeln.</summary>
        public Rectangle Bereich { get; }
        /// <summary>true = wird beim Export nach Word berücksichtigt.</summary>
        public bool Aktiv { get; set; }

        public SegmentRechteck(int zeile, int spalte, Rectangle bereich)
        {
            Zeile   = zeile;
            Spalte  = spalte;
            Bereich = bereich;
            Aktiv   = true;
        }
    }
}
