using System.Drawing;

namespace StatikManager.Modules.Bildschnitt
{
    /// <summary>
    /// Beschreibt einen rechteckigen Ausschnitt (Crop-Bereich) innerhalb eines Bildes.
    /// Koordinaten beziehen sich auf Pixel des Quellbildes (0-basiert).
    /// </summary>
    public class RandRechteck
    {
        public int Links  { get; set; }
        public int Oben   { get; set; }
        public int Rechts { get; set; }
        public int Unten  { get; set; }

        public int Breite => Rechts - Links;
        public int Höhe   => Unten  - Oben;

        public RandRechteck() { }

        public RandRechteck(int links, int oben, int rechts, int unten)
        {
            Links = links;
            Oben  = oben;
            Rechts = rechts;
            Unten  = unten;
        }

        /// <summary>Erzeugt eine unabhängige Kopie.</summary>
        public RandRechteck Clone() =>
            new RandRechteck(Links, Oben, Rechts, Unten);

        /// <summary>Konvertiert zu System.Drawing.Rectangle.</summary>
        public Rectangle AlsDrawingRect() =>
            new Rectangle(Links, Oben, Breite, Höhe);

        /// <summary>Klemmt alle Werte auf gültige Bildbounds.</summary>
        public void Klemmen(int bildBreite, int bildHöhe)
        {
            Links  = Klemme(Links,  0, bildBreite - 1);
            Oben   = Klemme(Oben,   0, bildHöhe   - 1);
            Rechts = Klemme(Rechts, Links  + 1, bildBreite);
            Unten  = Klemme(Unten,  Oben   + 1, bildHöhe);
        }

        private static int Klemme(int v, int min, int max) =>
            v < min ? min : v > max ? max : v;
    }
}
