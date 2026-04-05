using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace StatikManager.Modules.Bildschnitt
{
    /// <summary>
    /// Berechnet aus einem Crop-Bereich und einer Menge von Schnittlinien
    /// ein rechteckiges Raster von SegmentRechteck-Objekten.
    ///
    /// Logik:
    ///   – Horizontale Linien erzeugen Zeilen-Grenzen (Y-Achse).
    ///   – Vertikale Linien erzeugen Spalten-Grenzen (X-Achse).
    ///   – Linien außerhalb des Crop-Bereichs werden ignoriert.
    ///   – Reihenfolge: von oben-links nach unten-rechts.
    /// </summary>
    public static class SegmentierungsEngine
    {
        public static List<SegmentRechteck> Segmentiere(
            RandRechteck           cropBereich,
            IEnumerable<SchnittLinie> linien)
        {
            var yGrenzen = new List<int> { cropBereich.Oben };
            var xGrenzen = new List<int> { cropBereich.Links };

            foreach (var linie in linien)
            {
                if (linie.Richtung == SchnittRichtung.Horizontal)
                {
                    if (linie.Position > cropBereich.Oben && linie.Position < cropBereich.Unten)
                        yGrenzen.Add(linie.Position);
                }
                else
                {
                    if (linie.Position > cropBereich.Links && linie.Position < cropBereich.Rechts)
                        xGrenzen.Add(linie.Position);
                }
            }

            yGrenzen.Add(cropBereich.Unten);
            xGrenzen.Add(cropBereich.Rechts);
            yGrenzen.Sort();
            xGrenzen.Sort();

            var segmente = new List<SegmentRechteck>();
            for (int zeile = 0; zeile < yGrenzen.Count - 1; zeile++)
            {
                for (int spalte = 0; spalte < xGrenzen.Count - 1; spalte++)
                {
                    int x = xGrenzen[spalte];
                    int y = yGrenzen[zeile];
                    int w = xGrenzen[spalte + 1] - x;
                    int h = yGrenzen[zeile + 1]  - y;
                    if (w > 0 && h > 0)
                        segmente.Add(new SegmentRechteck(zeile, spalte,
                            new Rectangle(x, y, w, h)));
                }
            }
            return segmente;
        }
    }
}
