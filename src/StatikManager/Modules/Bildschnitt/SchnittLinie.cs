namespace StatikManager.Modules.Bildschnitt
{
    public enum SchnittRichtung { Horizontal, Vertikal }

    /// <summary>
    /// Eine horizontale oder vertikale Schnittlinie auf dem Bild.
    /// Position ist Y (Horizontal) bzw. X (Vertikal) in Bildpixeln.
    /// </summary>
    public class SchnittLinie
    {
        public SchnittRichtung Richtung  { get; }
        public int             Position  { get; set; }   // veränderbar per Drag

        public SchnittLinie(SchnittRichtung richtung, int position)
        {
            Richtung = richtung;
            Position = position;
        }
    }
}
