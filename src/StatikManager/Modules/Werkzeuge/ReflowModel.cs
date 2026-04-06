using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatikManager.Modules.Werkzeuge
{
    // ─────────────────────────────────────────────────────────────────────────
    //  Reflow-Datenmodell für den PdfSchnittEditor
    //
    //  Designprinzip:
    //    - ContentBlock  = atomare Inhaltseinheit (Bitmapstreifen oder Leerzeile)
    //    - PlacedBlock   = ContentBlock auf einer Ausgabe-Seite platziert
    //                      (kann kleiner als Block sein wenn am Seitenumbruch geteilt)
    //    - OutputPage    = eine Ausgabe-Seite mit ihren platzierten Blöcken
    //    - ReflowResult  = vollständiges Layout (Liste von OutputPages)
    //    - ReflowEngine  = pure static Reflow-Funktion ohne UI-Abhängigkeit
    //
    //  Beide Modelle (altes Fraktions-Modell + dieses Block-Modell) koexistieren
    //  während der Migration. Schritt 4+ tauscht dann ZeicheCanvas / SpeicherInStream.
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Atomarer Inhaltsblock: ein vertikaler Streifen aus einer Quell-PDF-Seite
    /// oder eine eingefügte Leerzeile (FracOben == FracUnten, ExtraHeightPx > 0).
    /// </summary>
    public class ContentBlock
    {
        /// <summary>Eindeutige ID innerhalb der Session (für Debugging / Undo).</summary>
        public int BlockId { get; set; }

        /// <summary>
        /// Index in _seitenBilder (Quell-Bitmap).
        /// Sonderwert -1 = reine Leerzeile (kein Bitmap-Inhalt, wird beim Rendern ignoriert).
        /// </summary>
        public int SourcePageIdx { get; set; } = -1;

        /// <summary>Obere Grenze im Quell-Bitmap (0.0 = Seitenanfang).</summary>
        public double FracOben { get; set; }

        /// <summary>Untere Grenze im Quell-Bitmap (1.0 = Seitenende).</summary>
        public double FracUnten { get; set; }

        /// <summary>Wenn true, wird dieser Block beim Reflow übersprungen.</summary>
        public bool IsDeleted { get; set; }

        /// <summary>Wie groß die Lücke nach dem Löschen dargestellt wird. Nur relevant wenn IsDeleted == true.</summary>
        public GapModus GapArt { get; set; } = GapModus.OriginalAbstand;

        /// <summary>Lückengröße in mm. Nur relevant für GapModus.KundenAbstand. Muss >= 0 sein.</summary>
        public double GapMm { get; set; } = 0.0;

        /// <summary>
        /// Zusätzliche Leerzeilen-Höhe in Pixeln.
        /// Für normale Blöcke immer 0.
        /// Für eingefügte Leerzeilen: ExtraHeightPx > 0, FracOben == FracUnten.
        /// </summary>
        public double ExtraHeightPx { get; set; }

        /// <summary>True wenn dieser Block eine reine Leerzeile ohne Bitmap-Inhalt ist.</summary>
        public bool IsLeerzeile => FracOben >= FracUnten && ExtraHeightPx > 0.0;

        /// <summary>
        /// Gesamthöhe dieses Blocks in Pixeln.
        /// <paramref name="sourcePageHeightPx"/> = PixelHeight der Quell-Bitmap dieser Seite.
        /// </summary>
        public double ContentHeightPx(double sourcePageHeightPx)
            => sourcePageHeightPx * Math.Max(0.0, FracUnten - FracOben) + ExtraHeightPx;

        public override string ToString()
            => IsLeerzeile
                ? $"[Block {BlockId}] Leerzeile {ExtraHeightPx:F0}px (SourcePageIdx=-1)"
                : $"[Block {BlockId}] Seite {SourcePageIdx} Frac {FracOben:F3}–{FracUnten:F3}" +
                  (IsDeleted ? $" GELÖSCHT [{GapArt}/{GapMm:F1}mm]" : "");
    }

    /// <summary>
    /// Ein ContentBlock – oder ein Teil davon – platziert auf einer OutputPage.
    /// Bei Seitenumbruch mitten im Block wird er in zwei PlacedBlocks aufgeteilt;
    /// SrcFracOben/Unten geben dann den tatsächlich gerenderten Ausschnitt an.
    /// </summary>
    public class PlacedBlock
    {
        /// <summary>Referenz auf den ContentBlock aus dem der Streifen stammt.</summary>
        public ContentBlock Block { get; set; }

        /// <summary>Y-Position des Blocks auf dieser Seite (Pixel, ab Seitenanfang).</summary>
        public double YOffset { get; set; }

        /// <summary>Gerenderte Höhe auf dieser Seite (Pixel). Kann < ContentHeightPx bei Split.</summary>
        public double HeightPx { get; set; }

        /// <summary>
        /// Crop-Fraktion im Quell-Bitmap – obere Grenze.
        /// Entspricht Block.FracOben, außer der Block wurde oben abgetrennt.
        /// Für Leerzeilen irrelevant (kein Bitmap-Inhalt).
        /// </summary>
        public double SrcFracOben { get; set; }

        /// <summary>
        /// Crop-Fraktion im Quell-Bitmap – untere Grenze.
        /// Entspricht Block.FracUnten, außer der Block wurde unten abgetrennt.
        /// Für Leerzeilen irrelevant.
        /// </summary>
        public double SrcFracUnten { get; set; }

        public override string ToString()
            => $"  Y={YOffset:F0} H={HeightPx:F0}px  Frac={SrcFracOben:F3}–{SrcFracUnten:F3}  {Block}";
    }

    /// <summary>
    /// Eine Ausgabe-Seite nach dem Reflow-Lauf.
    /// MaxHeightPx ist die Kapazität (= Originalseitenhöhe).
    /// FilledHeightPx ist tatsächlich belegte Höhe.
    /// </summary>
    public class OutputPage
    {
        /// <summary>Maximale nutzbare Höhe (Pixel) — entspricht der Original-Seitenhöhe.</summary>
        public double MaxHeightPx { get; set; }

        /// <summary>Seitenbreite (Pixel).</summary>
        public double WidthPx { get; set; }

        /// <summary>Index der Quell-PDF-Seite, aus der diese OutputPage entstammt. -1 = unbekannt.</summary>
        public int SourcePageIdx { get; set; } = -1;

        /// <summary>True wenn diese Seite durch Überlauf der vorigen Quellseite entstanden ist.</summary>
        public bool IsOverflowPage { get; set; }

        /// <summary>Platzierte Blöcke in ihrer vertikalen Reihenfolge.</summary>
        public List<PlacedBlock> Blocks { get; } = new List<PlacedBlock>();

        /// <summary>Tatsächlich belegte Höhe (Ende des letzten Blocks).</summary>
        public double FilledHeightPx
            => Blocks.Count > 0
                ? Blocks[Blocks.Count - 1].YOffset + Blocks[Blocks.Count - 1].HeightPx
                : 0.0;
    }

    /// <summary>
    /// Vollständiges Layout-Ergebnis: alle Ausgabe-Seiten in Reihenfolge.
    /// Wird bei jeder Änderung am Inhaltsmodell neu berechnet.
    /// </summary>
    public class ReflowResult
    {
        public List<OutputPage> Pages { get; } = new List<OutputPage>();
    }

    /// <summary>
    /// Pure Reflow-Engine — keine WPF-UI-Abhängigkeit, keine Seiteneffekte.
    /// Nimmt flache Blockliste + Seitenhöhen-Array und berechnet daraus das Layout.
    /// </summary>
    public static class ReflowEngine
    {
        /// <summary>Mindesthöhe in Pixeln, damit ein Block-Stummel noch auf eine Seite passt.
        /// Ist die verbleibende Seitenhöhe kleiner, wird der Block komplett auf die nächste Seite geschoben.</summary>
        public const double MinSplitHeightPx = 8.0;

        /// <summary>
        /// Berechnet das vollständige Seitenlayout für die gegebene Blockliste.
        /// </summary>
        /// <param name="blöcke">Geordnete ContentBlock-Liste (inkl. gelöschter — werden gefiltert).</param>
        /// <param name="sourcePageHeightsPx">
        ///   PixelHeight der Quell-Bitmaps, indiziert wie ContentBlock.SourcePageIdx.
        ///   Wird für ContentHeightPx-Berechnung benötigt.
        /// </param>
        /// <param name="pageMaxHeightPx">Maximale Seitenhöhe der Ausgabe (Pixel).</param>
        /// <param name="pageWidthPx">Seitenbreite der Ausgabe (Pixel).</param>
        /// <returns>ReflowResult mit einer oder mehr OutputPages.</returns>
        public static ReflowResult RunReflow(
            IReadOnlyList<ContentBlock> blöcke,
            IReadOnlyList<double>       sourcePageHeightsPx,
            double                     pageMaxHeightPx,
            double                     pageWidthPx)
        {
            if (blöcke        == null) throw new ArgumentNullException(nameof(blöcke));
            if (sourcePageHeightsPx == null) throw new ArgumentNullException(nameof(sourcePageHeightsPx));
            if (pageMaxHeightPx <= 0) throw new ArgumentOutOfRangeException(nameof(pageMaxHeightPx));

            var result = new ReflowResult();

            // Nur nicht-gelöschte Blöcke mit messbarer Höhe
            var sichtbar = blöcke.Where(b => !b.IsDeleted).ToList();
            if (sichtbar.Count == 0)
                return result;

            var aktuelleSeite = NeueSeite(pageMaxHeightPx, pageWidthPx);
            double currentY   = 0.0;

            foreach (var block in sichtbar)
            {
                double sourcePH = (block.SourcePageIdx >= 0 && block.SourcePageIdx < sourcePageHeightsPx.Count)
                    ? sourcePageHeightsPx[block.SourcePageIdx]
                    : pageMaxHeightPx;  // -1 = Leerzeile, kein Quell-Bitmap

                double blockH = block.ContentHeightPx(sourcePH);
                if (blockH <= 0.0) continue;

                // Verbleibende Fraktionen für den aktuellen Durchlauf
                double restFracOben  = block.FracOben;
                double restFracUnten = block.FracUnten;
                double restH         = blockH;

                // Schleife: Block kann durch Seitenumbrüche in mehrere PlacedBlocks geteilt werden
                while (restH > 0.0)
                {
                    double verfügbar = pageMaxHeightPx - currentY;

                    // Seite voll → abschließen, neue starten
                    if (verfügbar < MinSplitHeightPx)
                    {
                        result.Pages.Add(aktuelleSeite);
                        aktuelleSeite = NeueSeite(pageMaxHeightPx, pageWidthPx);
                        currentY  = 0.0;
                        verfügbar = pageMaxHeightPx;
                    }

                    if (restH <= verfügbar)
                    {
                        // Block passt vollständig auf die aktuelle Seite
                        aktuelleSeite.Blocks.Add(new PlacedBlock
                        {
                            Block        = block,
                            YOffset      = currentY,
                            HeightPx     = restH,
                            SrcFracOben  = restFracOben,
                            SrcFracUnten = restFracUnten
                        });
                        currentY += restH;
                        restH = 0.0;
                    }
                    else
                    {
                        // Block überläuft — teilen
                        double splitH = verfügbar;

                        double splitFracUnten;
                        if (block.IsLeerzeile)
                        {
                            // Leerzeile hat kein Bitmap — SrcFrac bleibt konstant (wird beim Rendern ignoriert)
                            splitFracUnten = restFracOben;
                        }
                        else
                        {
                            // Fraktion des Split-Punkts im Quell-Bitmap
                            splitFracUnten = restFracOben + (splitH / sourcePH);
                            splitFracUnten = Math.Min(splitFracUnten, restFracUnten);
                        }

                        aktuelleSeite.Blocks.Add(new PlacedBlock
                        {
                            Block        = block,
                            YOffset      = currentY,
                            HeightPx     = splitH,
                            SrcFracOben  = restFracOben,
                            SrcFracUnten = splitFracUnten
                        });

                        result.Pages.Add(aktuelleSeite);
                        aktuelleSeite = NeueSeite(pageMaxHeightPx, pageWidthPx);
                        currentY  = 0.0;

                        // Rest-Block für nächste Iteration
                        restFracOben = block.IsLeerzeile ? restFracOben : splitFracUnten;
                        restH       -= splitH;
                    }
                }
            }

            // Letzte Seite nur hinzufügen wenn sie Inhalt hat
            if (aktuelleSeite.Blocks.Count > 0)
                result.Pages.Add(aktuelleSeite);

            return result;
        }

        private static OutputPage NeueSeite(double maxH, double w)
            => new OutputPage { MaxHeightPx = maxH, WidthPx = w };

        // ── Debug-Ausgabe ────────────────────────────────────────────────────

        /// <summary>
        /// Erzeugt eine lesbare Beschreibung des ReflowResult für Debug-Zwecke.
        /// Wird über System.Diagnostics.Debug.WriteLine ausgegeben.
        /// </summary>
        public static string DebugBeschreibung(ReflowResult r, IReadOnlyList<ContentBlock> blöcke)
        {
            if (r == null) return "(null)";
            var sb = new StringBuilder();
            sb.AppendLine($"=== ReflowResult: {r.Pages.Count} Seite(n), " +
                          $"{blöcke.Count(b => !b.IsDeleted)} aktive Blöcke ===");
            for (int pi = 0; pi < r.Pages.Count; pi++)
            {
                var page = r.Pages[pi];
                sb.AppendLine($"  Seite {pi + 1}: {page.Blocks.Count} Block(e), " +
                              $"gefüllt {page.FilledHeightPx:F0}/{page.MaxHeightPx:F0} px");
                foreach (var pb in page.Blocks)
                    sb.AppendLine("    " + pb.ToString());
            }
            return sb.ToString();
        }
    }
}
