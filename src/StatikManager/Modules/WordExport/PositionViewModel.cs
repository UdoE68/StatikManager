// src/StatikManager/Modules/WordExport/PositionViewModel.cs
using System.Collections.Generic;

namespace StatikManager.Modules.WordExport
{
    /// <summary>Repräsentiert eine AxisVM-Position (Pos_01_Fundament/).</summary>
    internal sealed class PositionViewModel
    {
        public string Id         { get; set; } = "";
        public string Name       { get; set; } = "";
        public string OrdnerPfad { get; set; } = "";

        public List<AusschnittViewModel> Ausschnitte { get; set; } = new();
    }

    /// <summary>Repräsentiert einen einzelnen Ausschnitt (PNG + Metadaten).</summary>
    internal sealed class AusschnittViewModel
    {
        public int    Nr          { get; set; }
        public string Ueberschrift { get; set; } = "";
        public string PngPfad     { get; set; } = "";
        public string Massstab    { get; set; } = "";

        /// <summary>Anzeigename in der Liste: "001 – Ansicht vorne  [1:100]"</summary>
        public string Anzeigename =>
            string.IsNullOrEmpty(Massstab)
                ? $"{Nr:000}  –  {Ueberschrift}"
                : $"{Nr:000}  –  {Ueberschrift}   [{Massstab}]";
    }
}
