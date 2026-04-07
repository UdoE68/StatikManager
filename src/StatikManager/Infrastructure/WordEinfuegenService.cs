// src/StatikManager/Infrastructure/WordEinfuegenService.cs
using System;
using System.IO;
using System.Runtime.InteropServices;
using StatikManager.Core;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Infrastructure
{
    /// <summary>
    /// Interaktiver Word-Service: verbindet sich mit dem laufenden Word,
    /// öffnet/erstellt Dokumente und fügt Inhalte an der Cursorposition ein.
    /// Alle Methoden müssen auf dem UI-Thread aufgerufen werden (STA).
    /// </summary>
    internal static class WordEinfuegenService
    {
        private static Word.Application? _wordApp;

        // ── Verbindung ─────────────────────────────────────────────────────

        /// <summary>
        /// Gibt true zurück wenn Word läuft und ein Dokument geöffnet ist.
        /// </summary>
        public static bool IstWordBereit()
        {
            try
            {
                var app = HoleWordApp();
                return app != null && app.Documents.Count > 0;
            }
            catch { return false; }
        }

        /// <summary>
        /// Gibt den Pfad des aktiven Word-Dokuments zurück, oder null.
        /// </summary>
        public static string? GetAktiveDokumentPfad()
        {
            try
            {
                var app = HoleWordApp();
                if (app == null || app.Documents.Count == 0) return null;
                var doc = app.ActiveDocument;
                return string.IsNullOrEmpty(doc.FullName) ? null : doc.FullName;
            }
            catch { return null; }
        }

        // ── Dokument öffnen / erstellen ────────────────────────────────────

        /// <summary>
        /// Öffnet eine bestehende .docx-Datei in Word (sichtbar).
        /// </summary>
        public static void OeffneDokument(string pfad)
        {
            if (!File.Exists(pfad))
                throw new FileNotFoundException("Word-Datei nicht gefunden: " + pfad);

            var app = HoleOderStarteWord();
            app.Documents.Open(
                FileName: pfad,
                ReadOnly: false,
                AddToRecentFiles: true,
                Visible: true);
            app.Visible = true;
            app.Activate();
        }

        /// <summary>
        /// Erstellt ein neues Dokument, optional auf Basis einer .dotx-Vorlage.
        /// </summary>
        public static void ErstelleDokument(string? vorlagePfad)
        {
            var app = HoleOderStarteWord();
            if (!string.IsNullOrEmpty(vorlagePfad) && File.Exists(vorlagePfad))
                app.Documents.Add(Template: vorlagePfad, Visible: true);
            else
                app.Documents.Add(Visible: true);
            app.Visible = true;
            app.Activate();
        }

        // ── Einfügen ──────────────────────────────────────────────────────

        /// <summary>
        /// Fügt ein PNG-Bild mit optionaler Beschriftungszeile an der aktuellen
        /// Cursor-Position im aktiven Word-Dokument ein.
        /// </summary>
        public static void EinfuegenAnCursor(
            string pngPfad,
            string ueberschrift,
            string massstab,
            BildbreiteOption bildbreiteOption,
            bool mitUeberschrift,
            bool mitMassstab)
        {
            if (!File.Exists(pngPfad))
                throw new FileNotFoundException($"PNG nicht gefunden: {pngPfad}");

            var app = HoleWordApp()
                ?? throw new InvalidOperationException("Word ist nicht geöffnet.");
            if (app.Documents.Count == 0)
                throw new InvalidOperationException("Kein Word-Dokument geöffnet.");

            Word.Range cursor = null;
            Word.InlineShape shape = null;
            try
            {
                cursor = app.Selection.Range;

                shape = cursor.InlineShapes.AddPicture(
                    FileName: pngPfad,
                    LinkToFile: false,
                    SaveWithDocument: true);

                float breitePixel = BerechneBildbreite(app, bildbreiteOption);
                if (breitePixel > 0)
                {
                    // LockAspectRatio ist MsoTriState aus Microsoft.Office.Core (office-Assembly).
                    // Da EmbedInteropTypes=true, Zugriff via dynamic um Namespace-Referenz zu vermeiden.
                    // msoTrue = -1
                    dynamic dynShape = shape;
                    dynShape.LockAspectRatio = -1;
                    shape.Width = breitePixel;
                }

                if (mitUeberschrift || mitMassstab)
                {
                    shape.Range.InsertParagraphAfter();

                    Word.Range nachBild = shape.Range;
                    nachBild.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    nachBild.MoveEnd(Word.WdUnits.wdParagraph, 1);
                    nachBild.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    string beschriftung = BaueBeschriftung(ueberschrift, massstab, mitUeberschrift, mitMassstab);
                    nachBild.InsertAfter(beschriftung);
                }

                app.Selection.EndOf(Word.WdUnits.wdParagraph);
            }
            finally
            {
                if (shape != null) try { Marshal.ReleaseComObject(shape); } catch { }
                if (cursor != null) try { Marshal.ReleaseComObject(cursor); } catch { }
            }
        }

        // ── Hilfsmethoden ─────────────────────────────────────────────────

        private static Word.Application? HoleWordApp()
        {
            try
            {
                if (_wordApp != null)
                {
                    _ = _wordApp.Version;
                    return _wordApp;
                }
            }
            catch
            {
                _wordApp = null;
            }

            try
            {
                _wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                return _wordApp;
            }
            catch
            {
                return null;
            }
        }

        private static Word.Application HoleOderStarteWord()
        {
            var app = HoleWordApp();
            if (app != null) return app;

            _wordApp = new Word.Application { Visible = true };
            return _wordApp;
        }

        private static float BerechneBildbreite(Word.Application app, BildbreiteOption option)
        {
            const float CmZuPunkt = 28.3465f;
            switch (option)
            {
                case BildbreiteOption.Seitenbreite:
                {
                    var doc = app.ActiveDocument;
                    float seitenBreite = (float)doc.PageSetup.PageWidth;
                    float randLinks    = (float)doc.PageSetup.LeftMargin;
                    float randRechts   = (float)doc.PageSetup.RightMargin;
                    return seitenBreite - randLinks - randRechts;
                }
                case BildbreiteOption.HalbeSeitenbreite:
                {
                    var doc = app.ActiveDocument;
                    float seitenBreite = (float)doc.PageSetup.PageWidth;
                    float randLinks    = (float)doc.PageSetup.LeftMargin;
                    float randRechts   = (float)doc.PageSetup.RightMargin;
                    return (seitenBreite - randLinks - randRechts) / 2f;
                }
                case BildbreiteOption.Manuell_14cm:
                    return 14f * CmZuPunkt;
                case BildbreiteOption.Manuell_10cm:
                    return 10f * CmZuPunkt;
                default:
                    return 0;
            }
        }

        private static string BaueBeschriftung(string ueberschrift, string massstab,
            bool mitUeberschrift, bool mitMassstab)
        {
            if (mitUeberschrift && mitMassstab && !string.IsNullOrEmpty(massstab))
                return $"{ueberschrift}   |   Maßstab {massstab}";
            if (mitUeberschrift)
                return ueberschrift;
            if (mitMassstab && !string.IsNullOrEmpty(massstab))
                return $"Maßstab {massstab}";
            return "";
        }
    }

}
