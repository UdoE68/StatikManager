using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf.IO;
using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using IO = System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Modules.Werkzeuge
{
    /// <summary>
    /// PDF-Vorschau mit Beschnittrahmen und Word-Export.
    /// Zoom per Strg+Mausrad, verschiebbarer Beschnittrahmen, präziser Word-Export.
    /// </summary>
    public partial class PdfSchnittEditor : UserControl
    {
        // ── Render-Konstanten ─────────────────────────────────────────────────
        private const int    RenderBreite  = 900;
        private const double ExportDpi     = 150.0;  // DPI für Export-Render
        private const double SeitenAbstand = 30;
        private const int    SeiteX        = 20;

        // Zoom-Grenzen
        private const double ZoomMin  = 0.20;
        private const double ZoomMax  = 4.00;
        private const double ZoomStep = 0.15;

        public PdfSchnittEditor() => InitializeComponent();

        // ── Zustand ───────────────────────────────────────────────────────────
        private string?            _pdfPfad;
        private List<BitmapSource> _seitenBilder  = new();
        private double[]           _seitenYStart   = Array.Empty<double>();
        private double[]           _seitenHöhe     = Array.Empty<double>();

        // Horizontaler Layout-Modus
        private bool     _layoutHorizontal;
        private double[] _seitenXStart = Array.Empty<double>();

        private double _zoomFaktor = 1.0;

        // Smooth-Zoom
        private double _zielZoom    = 1.0;
        private Point? _zoomAnker;      // Canvas-Koordinaten des Maus-Ankerpunkts
        private Point? _zoomAnkerView;  // Viewport-Koordinaten des Maus-Ankerpunkts
        private System.Windows.Threading.DispatcherTimer? _zoomTimer;

        // Crop-Ränder als Bruchteil der Seitenbreite/-höhe (0 = kein Rand, 0.5 = halbe Seite) – pro Seite
        private double[] _cropLinks  = Array.Empty<double>();
        private double[] _cropRechts = Array.Empty<double>();
        private double[] _cropOben   = Array.Empty<double>();
        private double[] _cropUnten  = Array.Empty<double>();

        // Default-Crop: wird beim Laden neuer PDFs angewendet, sofern gesetzt; null = kein Standard
        private (double Links, double Rechts, double Oben, double Unten)? _defaultCrop;

        // Einheitlicher Anwendungsmodus für Drag, Dialog und Auto-Rand
        private enum CropAnwendungsModus { NurDiese = 0, Alle = 1, Ausgewählt = 2, AlsStandard = 3 }
        private CropAnwendungsModus _cropModus = CropAnwendungsModus.NurDiese;

        // ── Gruppen-System ────────────────────────────────────────────────────
        private sealed class CropGruppe
        {
            public int       Id     { get; set; }
            public string    Name   { get; set; } = "";
            public List<int> Seiten { get; set; } = new List<int>();
        }
        private static readonly Color[] GruppenFarben =
        {
            Color.FromRgb(  0, 120, 215),  // Blau
            Color.FromRgb( 16, 124,  16),  // Grün
            Color.FromRgb(202,  80,  16),  // Orange
            Color.FromRgb(136,  23, 152),  // Lila
            Color.FromRgb(196,  43,  28),  // Rot
            Color.FromRgb(  0, 153, 153),  // Türkis
        };
        private List<CropGruppe> _gruppen     = new List<CropGruppe>();
        private int              _aktGruppeId = 1;
        private int              _nächsteId   = 2;
        private bool             _modusWahlLäuft;   // Re-Entranz-Schutz für CmbCropModus_SelectionChanged
        private bool             _gruppeWahlLäuft;  // Re-Entranz-Schutz für CmbGruppe_SelectionChanged

        // Bearbeitungsmodus für Seitenzuweisung
        private bool      _bearbeitungsModus    = false;
        private List<int> _tempSeiten           = new List<int>();  // Arbeitsauswahl während Bearbeitung
        private List<int> _tempSeitenOriginal   = new List<int>();  // Snapshot für Abbrechen
        private bool      _bearbeitungsEndet;   // Re-Entranz-Schutz für BtnAuswahlmodus_Unchecked

        // Sicherheitsabstand für automatische Rand-Erkennung (Vorschau-Pixel, min 0, max 50)
        private double _cropSicherheitMm  = 2.0;   // Sicherheitsabstand fuer Auto-Rand (mm)
        private double _pxPerMm           = 4.0;   // Pixel pro mm (beim PDF-Laden berechnet)
        private bool   _autoRandAktiv;              // true wenn Auto-Rand zuletzt angewendet

        // Drag-State für Crop-Linien (Tag-Wert der gezogenen Linie)
        private string? _gezogeneCropSeite;

        // Export-Sperre: verhindert Doppelstart
        private volatile bool _exportLäuft;

        // Abbruch laufender Lade-Aufträge
        private CancellationTokenSource? _ladeCts;
        // Abbruch laufender Auto-Rand-Berechnung
        private CancellationTokenSource? _autoRandCts;
        // Generationszähler: verhindert, dass veraltete Dispatcher-Callbacks UI überschreiben
        private volatile int _ladeGeneration;

        // Sitzungszustand der nach dem nächsten PDF-Laden angewendet wird (null = kein Restore)
        private Core.SitzungsZustand? _pendingSitzung;

        // ── Fehler-Hilfsmethoden ──────────────────────────────────────────────

        private static void LogException(Exception ex, string kontext)
            => App.LogFehler(kontext, App.GetExceptionKette(ex));

        private bool SafeExecute(Action aktion, string kontext)
        {
            try { aktion(); return true; }
            catch (Exception ex) { LogException(ex, kontext); return false; }
        }

        // ── Öffentliche API ───────────────────────────────────────────────────

        /// <summary>
        /// Bereitet die Wiederherstellung eines Sitzungszustands vor.
        /// Muss VOR LadePdf() aufgerufen werden; Crop/Scroll werden nach dem Laden angewendet.
        /// </summary>
        public void SitzungVorbereiten(Core.SitzungsZustand sitzung)
        {
            _pendingSitzung    = sitzung;
            _layoutHorizontal  = sitzung.LayoutHorizontal;
            _cropModus         = (CropAnwendungsModus)Math.Max(0, Math.Min(3, sitzung.CropModus));
            _defaultCrop       = sitzung.DefaultCropGesetzt
                ? (sitzung.DefaultCropLinks, sitzung.DefaultCropRechts,
                   sitzung.DefaultCropOben,  sitzung.DefaultCropUnten)
                : (null as (double, double, double, double)?);
            // ComboBox synchronisieren
            _modusWahlLäuft = true;
            CmbCropModus.SelectedIndex = (int)_cropModus;
            _modusWahlLäuft = false;
        }

        /// <summary>
        /// Liest den aktuellen Zustand des Editors für die Sitzungsspeicherung.
        /// </summary>
        public Core.SitzungsZustand SitzungSpeichern()
        {
            var s = new Core.SitzungsZustand
            {
                ZoomFaktor         = _zoomFaktor,
                LayoutHorizontal   = _layoutHorizontal,
                CropModus          = (int)_cropModus,
                AktiveGruppeId     = _aktGruppeId,
                CropGruppen        = _gruppen.Select(g => new Core.SitzungsZustand.GruppeSitzung
                                     {
                                         Id     = g.Id,
                                         Name   = g.Name,
                                         Seiten = g.Seiten.ToArray()
                                     }).ToArray(),
                DefaultCropGesetzt = _defaultCrop.HasValue,
                CropLinks          = (double[])_cropLinks.Clone(),
                CropRechts         = (double[])_cropRechts.Clone(),
                CropOben           = (double[])_cropOben.Clone(),
                CropUnten          = (double[])_cropUnten.Clone(),
                ScrollH            = ScrollView.HorizontalOffset,
                ScrollV            = ScrollView.VerticalOffset,
            };
            if (_defaultCrop.HasValue)
            {
                s.DefaultCropLinks  = _defaultCrop.Value.Links;
                s.DefaultCropRechts = _defaultCrop.Value.Rechts;
                s.DefaultCropOben   = _defaultCrop.Value.Oben;
                s.DefaultCropUnten  = _defaultCrop.Value.Unten;
            }
            return s;
        }

        public void LadePdf(string pfad)
        {
            if (string.IsNullOrEmpty(pfad) || !IO.File.Exists(pfad))
            {
                TxtInfo.Text = "Datei nicht gefunden.";
                return;
            }

            var oldCts = _ladeCts;
            _ladeCts = null;
            oldCts?.Cancel();
            oldCts?.Dispose();
            _autoRandCts?.Cancel();
            _ladeGeneration++;
            int myGen = _ladeGeneration;
            var cts = new CancellationTokenSource();
            _ladeCts = cts;
            var token = cts.Token;

            _pdfPfad = pfad;
            PdfCanvas.Children.Clear();
            _seitenBilder.Clear();
            _seitenYStart = Array.Empty<double>();
            _seitenHöhe   = Array.Empty<double>();
            _cropLinks = _cropRechts = _cropOben = _cropUnten = Array.Empty<double>();
            TxtInfo.Text              = "Lade PDF …";
            BtnExport.IsEnabled       = false;
            BtnAuswahlmodus.IsEnabled = false;

            LogException(new Exception($"[DEBUG] Starte PDF-Laden: {IO.Path.GetFileName(pfad)}"), "LadePdf");

            string pfadKopie = pfad;

            var ladeThread = new Thread(() =>
            {
                List<BitmapSource>? bilder = null;
                double[]? yStart = null, höhe = null;
                string? fehler = null;

                // Serialisierung: kein paralleler pdfium-Zugriff mit Word-Vorschau-Threads.
                try { AppZustand.RenderSem.Wait(token); }
                catch (OperationCanceledException) { return; }
                try
                {
                try
                {
                    if (token.IsCancellationRequested) return;
                    bilder = PdfRenderer.RenderiereAlleSeiten(pfadKopie, RenderBreite, token: token);
                    if (token.IsCancellationRequested) return;
                    if (bilder.Count == 0) { fehler = "PDF enthält keine lesbaren Seiten."; }
                    else
                    {
                        BerechneLayoutStatic(bilder, SeitenAbstand, out yStart, out höhe);
                    }
                }
                catch (OperationCanceledException) { return; }
                catch (Exception ex) { LogException(ex, "LadePdf"); fehler = App.GetExceptionKette(ex); }
                }
                finally { AppZustand.RenderSem.Release(); }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                    try
                    {
                        if (fehler != null)
                        {
                            TxtInfo.Text = "Ladefehler. Siehe Log.";
                            MessageBox.Show(fehler, "PDF-Schnitt-Editor",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        _seitenBilder = bilder!;
                        _seitenYStart = yStart!;
                        _seitenHöhe   = höhe!;
                        InitCropArrays(_seitenBilder.Count);

                        // Sitzungs-Crop wiederherstellen (überschreibt Default-Crop wenn Länge passt)
                        var ps = _pendingSitzung;
                        _pendingSitzung = null;
                        if (ps?.CropLinks?.Length == _seitenBilder.Count)
                        {
                            _cropLinks  = (double[])ps.CropLinks.Clone();
                            _cropRechts = (double[])ps.CropRechts.Clone();
                            _cropOben   = (double[])ps.CropOben.Clone();
                            _cropUnten  = (double[])ps.CropUnten.Clone();
                        }
                        // Gruppen aus Sitzung wiederherstellen
                        if (ps?.CropGruppen?.Length > 0)
                        {
                            int n = _seitenBilder.Count;
                            var gl = new List<CropGruppe>();
                            foreach (var gs in ps.CropGruppen)
                            {
                                gl.Add(new CropGruppe
                                {
                                    Id     = gs.Id,
                                    Name   = string.IsNullOrWhiteSpace(gs.Name) ? $"Gruppe {gs.Id}" : gs.Name,
                                    Seiten = gs.Seiten?.Where(s => s >= 0 && s < n).ToList() ?? new List<int>()
                                });
                            }
                            if (gl.Count > 0)
                            {
                                // Gruppe 0 muss immer existieren
                                if (!gl.Any(g => g.Id == 0))
                                    gl.Insert(0, new CropGruppe { Id = 0, Name = "Gruppe 0" });

                                // Seiten ohne Gruppe → Gruppe 0
                                var gruppe0 = gl.First(g => g.Id == 0);
                                var zugewiesen = new HashSet<int>(gl.SelectMany(g => g.Seiten));
                                for (int si = 0; si < n; si++)
                                    if (!zugewiesen.Contains(si)) gruppe0.Seiten.Add(si);

                                _gruppen     = gl;
                                _aktGruppeId = ps.AktiveGruppeId;
                                if (!_gruppen.Any(g => g.Id == _aktGruppeId)) _aktGruppeId = 0;
                                _nächsteId   = _gruppen.Max(g => g.Id) + 1;
                            }
                        }
                        double pendingZoom    = ps != null ? Math.Max(ZoomMin, Math.Min(ZoomMax, ps.ZoomFaktor)) : _zoomFaktor;
                        double pendingScrollH = ps?.ScrollH ?? 0;
                        double pendingScrollV = ps?.ScrollV ?? 0;

                        ZeicheCanvas();
                        // Pixel-pro-mm fuer Sicherheitsabstand-Konvertierung
                        if (_pdfPfad != null && bilder!.Count > 0)
                        {
                            var (wPts, _) = HolePdfSeitenGrösse(_pdfPfad);
                            _pxPerMm = wPts > 0 ? bilder![0].PixelWidth / (wPts / 72.0 * 25.4) : 4.0;
                        }
                        BtnExport.IsEnabled      = true;
                        BtnAuswahlmodus.IsEnabled = true;
                        TxtInfo.Text = $"{bilder!.Count} Seite(n) geladen";

                        // Zoom und Scroll aus Sitzung anwenden
                        if (ps != null)
                        {
                            WendeZoomAnSofort(pendingZoom);
                            _zielZoom = pendingZoom;
                            if (pendingScrollH > 0 || pendingScrollV > 0)
                            {
                                double sH = pendingScrollH, sV = pendingScrollV;
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    ScrollView.ScrollToHorizontalOffset(sH);
                                    ScrollView.ScrollToVerticalOffset(sV);
                                }), System.Windows.Threading.DispatcherPriority.Loaded);
                            }
                        }
                    }
                    catch (Exception ex) { LogException(ex, "LadePdf/UI"); }
                    finally { AppZustand.Instanz.SetzeLaden(false); }
                }));
            }) { IsBackground = true, Name = "PdfSchnittLaden" };
            AppZustand.Instanz.SetzeLaden(true);
            ladeThread.Start();
        }

        public void ExportierenNachWord()
        {
            // Doppelstart verhindern: kein paralleler Word-Vorgang erlaubt
            if (_exportLäuft)
            {
                MessageBox.Show(
                    "Solange Word geöffnet wird oder ein Word-Vorgang aktiv ist,\n" +
                    "kann kein weiteres Word-Dokument erstellt werden.",
                    "Word beschäftigt",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (_seitenBilder.Count == 0 || _pdfPfad == null) return;

            string zielPfad = IO.Path.Combine(
                IO.Path.GetDirectoryName(_pdfPfad)!,
                IO.Path.GetFileNameWithoutExtension(_pdfPfad) + ".docx");

            // Überschreiben-Abfrage VOR dem Sperren
            if (IO.File.Exists(zielPfad))
            {
                var antwort = MessageBox.Show(
                    $"Die Datei existiert bereits:\n{IO.Path.GetFileName(zielPfad)}\n\nMöchten Sie sie überschreiben?",
                    "Datei überschreiben?",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (antwort != MessageBoxResult.Yes) return;
            }

            // Ab hier gesperrt
            _exportLäuft = true;
            BtnExport.IsEnabled = false;
            TxtInfo.Text = "Export gestartet …";

            var yStartK  = (double[])_seitenYStart.Clone();
            var höheK    = (double[])_seitenHöhe.Clone();
            string pdfK  = _pdfPfad;
            var cropLK = (double[])_cropLinks.Clone();
            var cropRK = (double[])_cropRechts.Clone();
            var cropOK = (double[])_cropOben.Clone();
            var cropUK = (double[])_cropUnten.Clone();

            var thread = new Thread(() =>
                ExportThreadWorker(zielPfad, pdfK, yStartK, höheK, cropLK, cropRK, cropOK, cropUK))
            { IsBackground = true, Name = "PdfSchnittExport" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        // ── Layout ────────────────────────────────────────────────────────────

        private static void BerechneLayoutStatic(
            List<BitmapSource> bilder, double abstand,
            out double[] yStart, out double[] höhe)
        {
            yStart = new double[bilder.Count];
            höhe   = new double[bilder.Count];
            double y = 0;
            for (int i = 0; i < bilder.Count; i++)
            {
                yStart[i] = y;
                höhe[i]   = Math.Max(1, bilder[i].PixelHeight);
                y += höhe[i] + abstand;
            }
        }
        private static void BerechneLayoutHorizontalStatic(
            List<BitmapSource> bilder, double abstand,
            out double[] xStart)
        {
            xStart = new double[bilder.Count];
            double x = SeiteX;
            for (int i = 0; i < bilder.Count; i++)
            {
                xStart[i] = x;
                x += Math.Max(1, bilder[i].PixelWidth) + abstand;
            }
        }

        // ── Crop-Linien ───────────────────────────────────────────────────────

        // Initialisiert alle Crop-Arrays auf 0, Länge = Seitenanzahl.
        // Wenn _defaultCrop gesetzt ist, wird er auf alle Seiten angewendet.
        private void InitCropArrays(int n)
        {
            _cropLinks  = new double[n];
            _cropRechts = new double[n];
            _cropOben   = new double[n];
            _cropUnten  = new double[n];
            if (_defaultCrop.HasValue)
            {
                var d = _defaultCrop.Value;
                for (int i = 0; i < n; i++)
                {
                    _cropLinks[i]  = d.Links;
                    _cropRechts[i] = d.Rechts;
                    _cropOben[i]   = d.Oben;
                    _cropUnten[i]  = d.Unten;
                }
            }
            // Gruppen zurücksetzen: Gruppe 0 (unverlöschbar) enthält alle Seiten
            _gruppen = new List<CropGruppe>
            {
                new CropGruppe { Id = 0, Name = "Gruppe 0", Seiten = Enumerable.Range(0, n).ToList() }
            };
            _aktGruppeId = 0;
            _nächsteId   = 1;
        }

        // Gibt den Index der Seite zurück, die aktuell am stärksten sichtbar ist.
        private int AktiveSeiteIndex()
        {
            if (_seitenBilder.Count == 0) return -1;
            if (_seitenBilder.Count == 1) return 0;
            double zoom = Math.Max(0.001, _zoomFaktor);
            double vpT  = ScrollView.VerticalOffset   / zoom;
            double vpB  = (ScrollView.VerticalOffset  + ScrollView.ViewportHeight) / zoom;
            double vpL  = ScrollView.HorizontalOffset / zoom;
            double vpR  = (ScrollView.HorizontalOffset + ScrollView.ViewportWidth)  / zoom;
            int    best = 0;
            double bestOvl = -1;
            for (int i = 0; i < _seitenBilder.Count; i++)
            {
                double ovl;
                if (_layoutHorizontal && i < _seitenXStart.Length)
                {
                    double pH = i < _seitenHöhe.Length ? _seitenHöhe[i] : 0;
                    ovl = Math.Max(0, Math.Min(vpR, _seitenXStart[i] + _seitenBilder[i].PixelWidth) - Math.Max(vpL, _seitenXStart[i]))
                        * Math.Max(0, Math.Min(vpB, SeiteX + pH) - Math.Max(vpT, SeiteX));
                }
                else if (!_layoutHorizontal && i < _seitenYStart.Length)
                {
                    ovl = Math.Max(0, Math.Min(vpB, _seitenYStart[i] + _seitenHöhe[i]) - Math.Max(vpT, _seitenYStart[i]));
                }
                else ovl = 0;
                if (ovl > bestOvl) { bestOvl = ovl; best = i; }
            }
            return best;
        }

        // ── Gruppen-Hilfsmethoden ─────────────────────────────────────────────

        private CropGruppe? AktiveGruppe() => _gruppen.FirstOrDefault(g => g.Id == _aktGruppeId);

        private CropGruppe? GruppeVonSeite(int idx) => _gruppen.FirstOrDefault(g => g.Seiten.Contains(idx));

        private Color FarbeVonGruppe(CropGruppe g)
        {
            int pos = _gruppen.IndexOf(g);
            return GruppenFarben[(pos < 0 ? 0 : pos) % GruppenFarben.Length];
        }

        private void WeiseSeiteZu(int seitenIdx, CropGruppe ziel)
        {
            foreach (var g in _gruppen) g.Seiten.Remove(seitenIdx);
            if (!ziel.Seiten.Contains(seitenIdx)) ziel.Seiten.Add(seitenIdx);
        }

        // Synchronisiert CmbGruppe mit der aktuellen Gruppenliste.
        private void AktualisiereGruppenComboBox()
        {
            if (CmbGruppe == null) return;
            _gruppeWahlLäuft = true;
            try
            {
                CmbGruppe.Items.Clear();
                foreach (var g in _gruppen)
                    CmbGruppe.Items.Add(g.Name);
                int aktIdx = _gruppen.FindIndex(g => g.Id == _aktGruppeId);
                CmbGruppe.SelectedIndex = aktIdx >= 0 ? aktIdx : 0;
                if (aktIdx >= 0) CmbGruppe.Text = _gruppen[aktIdx].Name;
                // Gruppe 0 ist unverlöschbar; außerdem muss mindestens 1 weitere Gruppe existieren
                int aktId = _gruppen.Count > 0 && aktIdx >= 0 ? _gruppen[aktIdx].Id : -1;
                BtnGruppeLöschen.IsEnabled = _gruppen.Count > 1 && aktId != 0;
            }
            finally { _gruppeWahlLäuft = false; }
        }

        // ── Bearbeitungsmodus ─────────────────────────────────────────────────

        private void StarteBearbeitungsModus()
        {
            var aktGruppe = AktiveGruppe();
            _tempSeiten         = aktGruppe?.Seiten.ToList() ?? new List<int>();
            _tempSeitenOriginal = _tempSeiten.ToList();
            _bearbeitungsModus  = true;
            BorderBearbeitungsModus.Visibility = Visibility.Visible;
            if (TxtSeitenBereich != null) TxtSeitenBereich.Clear();
            AktualisiereTempInfo();
            AktualisiereAuswahlAnzeige();
        }

        private void BeendeBearbeitungsModus(bool übernehmen)
        {
            if (übernehmen)
            {
                var aktGruppe = AktiveGruppe();
                if (aktGruppe != null)
                {
                    // Seiten die aus der Gruppe entfernt werden sollen
                    var entfernt = aktGruppe.Seiten.Except(_tempSeiten).ToList();
                    // Seiten die neu zur Gruppe hinzukommen
                    var hinzugekommen = _tempSeiten.Except(aktGruppe.Seiten).ToList();

                    // Hinzugekommene Seiten: WeiseSeiteZu entfernt sie aus alter Gruppe
                    foreach (int s in hinzugekommen)
                        WeiseSeiteZu(s, aktGruppe);

                    // Entfernte Seiten → Gruppe 0 (Fallback); wenn aktive Gruppe IS Gruppe 0 → bleiben dort
                    if (entfernt.Count > 0)
                    {
                        var gruppe0 = _gruppen.FirstOrDefault(g => g.Id == 0);
                        var ziel    = aktGruppe.Id != 0 ? gruppe0 : null; // bei Gruppe-0-Bearbeitung: kein Fallback
                        foreach (int s in entfernt)
                        {
                            aktGruppe.Seiten.Remove(s);
                            if (ziel != null && !ziel.Seiten.Contains(s))
                                ziel.Seiten.Add(s);
                            else if (ziel == null)
                                aktGruppe.Seiten.Add(s); // aktive Gruppe ist Gruppe 0 → Seite bleibt
                        }
                    }
                }
            }

            _tempSeiten.Clear();
            _tempSeitenOriginal.Clear();
            _bearbeitungsModus = false;
            BorderBearbeitungsModus.Visibility = Visibility.Collapsed;

            // BtnAuswahlmodus zurücksetzen ohne Unchecked-Handler auszulösen
            _bearbeitungsEndet = true;
            try { if (BtnAuswahlmodus?.IsChecked == true) BtnAuswahlmodus.IsChecked = false; }
            finally { _bearbeitungsEndet = false; }

            AktualisiereAuswahlAnzeige();
        }

        private void AktualisiereTempInfo()
        {
            if (TxtBearbeitungInfo == null) return;
            int n = _seitenBilder.Count;
            TxtBearbeitungInfo.Text = n > 0
                ? $"{_tempSeiten.Count} von {n} Seiten ausgewählt"
                : "";
        }

        // Parst Seitenbereiche wie "1-5,8,10-12" (1-basiert) → 0-basierte Indizes.
        private List<int> ParseSeitenBereich(string eingabe)
        {
            var seiten = new SortedSet<int>();
            int n = _seitenBilder.Count;
            if (n == 0 || string.IsNullOrWhiteSpace(eingabe)) return new List<int>();
            foreach (var teil in eingabe.Split(','))
            {
                var t = teil.Trim();
                if (t.Contains('-'))
                {
                    var parts = t.Split(new[] { '-' }, 2);
                    if (parts.Length == 2
                        && int.TryParse(parts[0].Trim(), out int von)
                        && int.TryParse(parts[1].Trim(), out int bis))
                    {
                        for (int i = Math.Max(1, von); i <= Math.Min(n, bis); i++)
                            seiten.Add(i - 1);
                    }
                }
                else if (int.TryParse(t, out int s) && s >= 1 && s <= n)
                {
                    seiten.Add(s - 1);
                }
            }
            return seiten.ToList();
        }

        // Liefert die Seiten-Menge, auf die Crop-Ops im Modus "Ausgewählt" wirken.
        // Im Bearbeitungsmodus: _tempSeiten (Arbeitsauswahl), sonst: aktive Gruppe.
        private IEnumerable<int> AktuelleZielSeiten()
        {
            if (_cropModus != CropAnwendungsModus.Ausgewählt) return Enumerable.Empty<int>();
            if (_bearbeitungsModus) return _tempSeiten;
            return AktiveGruppe()?.Seiten ?? (IEnumerable<int>)Enumerable.Empty<int>();
        }

        // ── Canvas zeichnen ───────────────────────────────────────────────────

        private void ZeicheCanvas()
        {
            try
            {
                PdfCanvas.Children.Clear();
                if (_seitenBilder.Count == 0) return;
                if (_seitenYStart.Length != _seitenBilder.Count ||
                    _seitenHöhe.Length   != _seitenBilder.Count) return;

                if (_layoutHorizontal)
                {
                    BerechneLayoutHorizontalStatic(_seitenBilder, SeitenAbstand, out _seitenXStart);
                    int lastH = _seitenXStart.Length - 1;
                    double gesamtW = _seitenXStart[lastH] + _seitenBilder[lastH].PixelWidth + SeitenAbstand;
                    double maxH    = _seitenBilder.Max(b => (double)b.PixelHeight);
                    PdfCanvas.Width  = Math.Max(gesamtW, 1);
                    PdfCanvas.Height = Math.Max(maxH + SeiteX * 2, 1);
                }
                else
                {
                    int    last    = _seitenYStart.Length - 1;
                    double gesamtH = _seitenYStart[last] + _seitenHöhe[last] + SeitenAbstand;
                    double maxBmpW = _seitenBilder.Max(b => (double)b.PixelWidth);
                    PdfCanvas.Width  = maxBmpW + SeiteX * 2;
                    PdfCanvas.Height = Math.Max(gesamtH, 1);
                }

                for (int i = 0; i < _seitenBilder.Count; i++)
                    SafeExecute(() => ZeicheSeite(i), $"ZeicheSeite[{i}]");

                ZeicheCropLinien();
                AktualisiereAuswahlAnzeige();
                AktualisiereGruppenComboBox();
            }
            catch (Exception ex) { LogException(ex, "ZeicheCanvas"); }
        }

        // Aktualisiert Opacity und Rahmen-Farbe aller Seiten-Borders im Canvas.
        // Im Modus "Ausgewählt": ausgewählte Seiten = blauer Rahmen, nicht gewählte = Opacity 0.6.
        // Alle anderen Modi: einheitlich Opacity 1.0, grauer Rahmen.
        private void AktualisiereAuswahlAnzeige()
        {
            if (PdfCanvas == null) return;
            bool auswahlModus = _cropModus == CropAnwendungsModus.Ausgewählt;
            foreach (UIElement child in PdfCanvas.Children)
            {
                if (child is not Border b) continue;
                if (b.Tag is not string tag || !tag.StartsWith("SEITE_")) continue;
                if (!int.TryParse(tag.Substring(6), out int idx)) continue;

                if (_bearbeitungsModus)
                {
                    bool inTemp = _tempSeiten.Contains(idx);
                    var aktGruppe = AktiveGruppe();
                    Color farbe   = aktGruppe != null ? FarbeVonGruppe(aktGruppe) : GruppenFarben[0];
                    b.Opacity         = inTemp ? 1.0 : 0.40;
                    b.BorderBrush     = new SolidColorBrush(inTemp ? farbe : Color.FromRgb(160, 160, 160));
                    b.BorderThickness = new Thickness(inTemp ? 3 : 1);
                }
                else if (auswahlModus)
                {
                    var gruppe = GruppeVonSeite(idx);
                    bool istAktiv = gruppe != null && gruppe.Id == _aktGruppeId;
                    Color farbe   = gruppe != null ? FarbeVonGruppe(gruppe) : GruppenFarben[0];
                    b.Opacity         = istAktiv ? 1.0 : 0.75;
                    b.BorderBrush     = new SolidColorBrush(farbe);
                    b.BorderThickness = new Thickness(istAktiv ? 3 : 1);
                }
                else
                {
                    b.Opacity         = 1.0;
                    b.BorderBrush     = new SolidColorBrush(Color.FromRgb(160, 160, 160));
                    b.BorderThickness = new Thickness(2);
                }
            }
        }

        private void ZeicheSeite(int i)
        {
            if (_seitenBilder[i] == null) return;
            // Tatsächliche Bitmap-Abmessungen verwenden – kein Strecken auf RenderBreite
            double bmpW = _seitenBilder[i].PixelWidth;
            double bmpH = Math.Max(1, _seitenHöhe[i]); // = PixelHeight

            DropShadowEffect? shadow = null;
            try
            {
                shadow = new DropShadowEffect
                {
                    BlurRadius  = 20, ShadowDepth = 7,
                    Direction   = 280, Color       = Colors.Black, Opacity = 0.85
                };
            }
            catch { /* ohne Schatten weiterzeichnen */ }

            var blatt = new Border
            {
                Tag                 = $"SEITE_{i}",
                Width               = bmpW,
                Height              = bmpH,
                Background          = Brushes.White,
                BorderBrush         = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                BorderThickness     = new Thickness(2),
                Child               = new Image
                {
                    Source  = _seitenBilder[i],
                    Width   = bmpW,
                    Height  = bmpH,
                    Stretch = Stretch.Uniform,   // Seitenverhältnis bleibt erhalten
                    SnapsToDevicePixels = true
                },
                Effect              = shadow,
                SnapsToDevicePixels = true
            };

            // Seitenzuweisung per Klick: NUR im aktiven Bearbeitungsmodus
            int seitenIdx = i; // capture für Lambda
            blatt.MouseLeftButtonDown += (_, ev) =>
            {
                // Außerhalb Bearbeitungsmodus: keine Gruppenänderung, normale Navigation
                if (!_bearbeitungsModus) return;
                if (_cropModus != CropAnwendungsModus.Ausgewählt) return;

                // Toggle in _tempSeiten — permanente Gruppenstruktur bleibt unberührt
                if (_tempSeiten.Contains(seitenIdx))
                {
                    _tempSeiten.Remove(seitenIdx);
                }
                else
                {
                    // Warnung nur wenn Seite einer echten (nicht-0) anderen Gruppe zugeordnet ist.
                    // Gruppe 0 ist der unzugewiesene Standard-Fallback → keine Bestätigung nötig.
                    var alteGruppe = GruppeVonSeite(seitenIdx);
                    var aktGruppe  = AktiveGruppe();
                    if (alteGruppe != null && aktGruppe != null
                        && alteGruppe.Id != aktGruppe.Id
                        && alteGruppe.Id != 0)
                    {
                        var antwort = MessageBox.Show(
                            $"Seite {seitenIdx + 1} gehört bereits zu \"{alteGruppe.Name}\".\n\n" +
                            $"Soll sie aus dieser Gruppe entfernt und der aktuellen Gruppe \"{aktGruppe.Name}\" zugeordnet werden?",
                            "Gruppe wechseln",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);
                        if (antwort != MessageBoxResult.Yes) { ev.Handled = true; return; }
                    }
                    _tempSeiten.Add(seitenIdx);
                }
                AktualisiereTempInfo();
                AktualisiereAuswahlAnzeige();
                ev.Handled = true;
            };

            double setX = _layoutHorizontal && i < _seitenXStart.Length ? _seitenXStart[i] : SeiteX;
            double setY = _layoutHorizontal ? SeiteX : _seitenYStart[i];
            Canvas.SetLeft(blatt, setX);
            Canvas.SetTop(blatt,  setY);
            PdfCanvas.Children.Add(blatt);
        }
        // Crop-Linien-Tags: visuelle Linie = "CROP_XXX", Hit-Zone = "CROP_XXX_HIT"
        private static bool IstCropLinie(Line l)
            => l.Tag is string s && s.StartsWith("CROP_");

        private static string BasisTag(string tag)
            => tag.EndsWith("_HIT") ? tag.Substring(0, tag.Length - 4) : tag;

        private void ZeicheCropLinien()
        {
            if (_seitenBilder.Count == 0) return;
            bool sichtbar = BtnRandAnzeigen.IsChecked == true;

            if (_layoutHorizontal)
            {
                if (_seitenXStart.Length != _seitenBilder.Count) return;
                for (int i = 0; i < _seitenBilder.Count; i++)
                {
                    double cL = i < _cropLinks.Length  ? _cropLinks[i]  : 0;
                    double cR = i < _cropRechts.Length ? _cropRechts[i] : 0;
                    double cO = i < _cropOben.Length   ? _cropOben[i]   : 0;
                    double cU = i < _cropUnten.Length  ? _cropUnten[i]  : 0;
                    double pageW = _seitenBilder[i].PixelWidth;
                    double pageH = _seitenHöhe[i];
                    MacheCropLinie(_seitenXStart[i] + cL * pageW, SeiteX,
                                   _seitenXStart[i] + cL * pageW, SeiteX + pageH,        $"CROP_LINKS_{i}",  sichtbar);
                    MacheCropLinie(_seitenXStart[i] + (1.0 - cR) * pageW, SeiteX,
                                   _seitenXStart[i] + (1.0 - cR) * pageW, SeiteX + pageH, $"CROP_RECHTS_{i}", sichtbar);
                    MacheCropLinie(_seitenXStart[i], SeiteX + cO * pageH,
                                   _seitenXStart[i] + pageW, SeiteX + cO * pageH,        $"CROP_OBEN_{i}",  sichtbar);
                    MacheCropLinie(_seitenXStart[i], SeiteX + (1.0 - cU) * pageH,
                                   _seitenXStart[i] + pageW, SeiteX + (1.0 - cU) * pageH, $"CROP_UNTEN_{i}", sichtbar);
                }
            }
            else
            {
                if (_seitenYStart.Length == 0) return;
                for (int i = 0; i < _seitenBilder.Count; i++)
                {
                    if (i >= _seitenYStart.Length || i >= _seitenHöhe.Length) break;
                    double cL = i < _cropLinks.Length  ? _cropLinks[i]  : 0;
                    double cR = i < _cropRechts.Length ? _cropRechts[i] : 0;
                    double cO = i < _cropOben.Length   ? _cropOben[i]   : 0;
                    double cU = i < _cropUnten.Length  ? _cropUnten[i]  : 0;
                    double pageW = _seitenBilder[i].PixelWidth;
                    double pT    = _seitenYStart[i];
                    double pH    = _seitenHöhe[i];
                    MacheCropLinie(SeiteX + cL * pageW, pT,
                                   SeiteX + cL * pageW, pT + pH,            $"CROP_LINKS_{i}",  sichtbar);
                    MacheCropLinie(SeiteX + (1.0 - cR) * pageW, pT,
                                   SeiteX + (1.0 - cR) * pageW, pT + pH,    $"CROP_RECHTS_{i}", sichtbar);
                    MacheCropLinie(SeiteX, pT + cO * pH,
                                   SeiteX + pageW, pT + cO * pH,             $"CROP_OBEN_{i}",  sichtbar);
                    MacheCropLinie(SeiteX, pT + (1.0 - cU) * pH,
                                   SeiteX + pageW, pT + (1.0 - cU) * pH,    $"CROP_UNTEN_{i}", sichtbar);
                }
            }
        }

        // Fügt zwei übereinanderliegende Linien zum Canvas hinzu:
        //   visLine  – dünne grüne Linie (Anzeige, kein HitTest)
        //   hitLine  – breite transparente Zone (HitTest, nimmt Mausereignisse)
        // So bleibt die Anzeige sauber und die Greiffläche ist groß genug.
        private void MacheCropLinie(
            double x1, double y1, double x2, double y2, string tag, bool sichtbar)
        {
            double dickeVis = 2.0 / _zoomFaktor;
            double dickeHit = Math.Max(14.0, 14.0 / _zoomFaktor); // ≥14 px Greifbereich
            bool istVertikal = (tag == "CROP_LINKS" || tag == "CROP_RECHTS"
                             || tag.StartsWith("CROP_LINKS_") || tag.StartsWith("CROP_RECHTS_"));
            var cursor     = istVertikal ? Cursors.SizeWE : Cursors.SizeNS;
            var visibility = sichtbar ? Visibility.Visible : Visibility.Collapsed;

            // Visuelle (dünne) Linie – nur Anzeige
            var visLine = new Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke           = Brushes.LimeGreen,
                StrokeThickness  = dickeVis,
                Tag              = tag,
                IsHitTestVisible = false,
                Visibility       = visibility
            };
            Panel.SetZIndex(visLine, 100);
            PdfCanvas.Children.Add(visLine);

            // Transparente Hit-Zone – breit, leicht zu greifen
            var hitLine = new Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke           = Brushes.Transparent,
                StrokeThickness  = dickeHit,
                Tag              = tag + "_HIT",
                Cursor           = cursor,
                IsHitTestVisible = true,
                Visibility       = visibility
            };
            Panel.SetZIndex(hitLine, 101);

            // Hover: visuelle Linie hervorheben
            hitLine.MouseEnter += (_, __) => visLine.Stroke = Brushes.Yellow;
            hitLine.MouseLeave += (_, __) => visLine.Stroke = Brushes.LimeGreen;

            hitLine.MouseLeftButtonDown += CropLinie_MouseDown;

            var menu     = new ContextMenu();
            var itemEdit = new MenuItem { Header = "Rand bearbeiten …" };
            itemEdit.Click += (_, __) => BtnRandBearbeiten_Click(this, new RoutedEventArgs());
            menu.Items.Add(itemEdit);
            hitLine.ContextMenu = menu;

            PdfCanvas.Children.Add(hitLine);
        }

        private void AktualisiereCropLinien()
        {
            if (_seitenBilder.Count == 0) return;
            foreach (var l in PdfCanvas.Children.OfType<Line>())
            {
                string tag = BasisTag(l.Tag?.ToString() ?? "");

                if (tag.StartsWith("CROP_LINKS_") &&
                    int.TryParse(tag.Substring(11), out int iL) && iL < _seitenBilder.Count)
                {
                    double cL    = iL < _cropLinks.Length ? _cropLinks[iL] : 0;
                    double pageW = _seitenBilder[iL].PixelWidth;
                    if (_layoutHorizontal && iL < _seitenXStart.Length)
                    {
                        l.X1 = l.X2 = _seitenXStart[iL] + cL * pageW;
                        l.Y1 = SeiteX; l.Y2 = SeiteX + _seitenHöhe[iL];
                    }
                    else if (!_layoutHorizontal && iL < _seitenYStart.Length)
                    {
                        l.X1 = l.X2 = SeiteX + cL * pageW;
                        l.Y1 = _seitenYStart[iL]; l.Y2 = _seitenYStart[iL] + _seitenHöhe[iL];
                    }
                }
                else if (tag.StartsWith("CROP_RECHTS_") &&
                         int.TryParse(tag.Substring(12), out int iR) && iR < _seitenBilder.Count)
                {
                    double cR    = iR < _cropRechts.Length ? _cropRechts[iR] : 0;
                    double pageW = _seitenBilder[iR].PixelWidth;
                    if (_layoutHorizontal && iR < _seitenXStart.Length)
                    {
                        l.X1 = l.X2 = _seitenXStart[iR] + (1.0 - cR) * pageW;
                        l.Y1 = SeiteX; l.Y2 = SeiteX + _seitenHöhe[iR];
                    }
                    else if (!_layoutHorizontal && iR < _seitenYStart.Length)
                    {
                        l.X1 = l.X2 = SeiteX + (1.0 - cR) * pageW;
                        l.Y1 = _seitenYStart[iR]; l.Y2 = _seitenYStart[iR] + _seitenHöhe[iR];
                    }
                }
                else if (tag.StartsWith("CROP_OBEN_") &&
                         int.TryParse(tag.Substring(10), out int iO) && iO < _seitenHöhe.Length)
                {
                    double cO = iO < _cropOben.Length ? _cropOben[iO] : 0;
                    if (_layoutHorizontal && iO < _seitenXStart.Length)
                    {
                        l.Y1 = l.Y2 = SeiteX + cO * _seitenHöhe[iO];
                        l.X1 = _seitenXStart[iO]; l.X2 = _seitenXStart[iO] + _seitenBilder[iO].PixelWidth;
                    }
                    else if (!_layoutHorizontal && iO < _seitenYStart.Length)
                    {
                        l.Y1 = l.Y2 = _seitenYStart[iO] + cO * _seitenHöhe[iO];
                        l.X1 = SeiteX; l.X2 = SeiteX + _seitenBilder[iO].PixelWidth;
                    }
                }
                else if (tag.StartsWith("CROP_UNTEN_") &&
                         int.TryParse(tag.Substring(11), out int iU) && iU < _seitenHöhe.Length)
                {
                    double cU = iU < _cropUnten.Length ? _cropUnten[iU] : 0;
                    if (_layoutHorizontal && iU < _seitenXStart.Length)
                    {
                        l.Y1 = l.Y2 = SeiteX + (1.0 - cU) * _seitenHöhe[iU];
                        l.X1 = _seitenXStart[iU]; l.X2 = _seitenXStart[iU] + _seitenBilder[iU].PixelWidth;
                    }
                    else if (!_layoutHorizontal && iU < _seitenYStart.Length)
                    {
                        l.Y1 = l.Y2 = _seitenYStart[iU] + (1.0 - cU) * _seitenHöhe[iU];
                        l.X1 = SeiteX; l.X2 = SeiteX + _seitenBilder[iU].PixelWidth;
                    }
                }
            }
        }

        // ── Automatische Rand-Erkennung ───────────────────────────────────────

        private static (double Links, double Rechts, double Oben, double Unten)
            ErkenneCropRänderVonBitmap(BitmapSource bmp, int sicherheitsPx = 4)
        {
            int w = bmp.PixelWidth, h = bmp.PixelHeight;
            if (w <= 0 || h <= 0) return (0, 0, 0, 0);
            int stride = w * 4;
            var pixels = new byte[(long)h * stride];
            try { bmp.CopyPixels(pixels, stride, 0); }
            catch { return (0, 0, 0, 0); }

            // Dichte-basierte Erkennung: Pixel deutlich dunkler als Weiß = Inhalt.
            // Schwelle 245 (statt 240) damit hellgraue Hintergründe ignoriert werden.
            // Dichte 0.15 % (statt 0.3 %) damit auch leichter Inhalt erkannt wird.
            // Mindestens 2 aufeinanderfolgende Zeilen/Spalten mit Inhalt → robuster
            // gegen einzelne Rauschpixel am Blattrand.
            const byte   BgSchwelle     = 245;
            const double DichteSchwelle = 0.0015;
            const int    MinKonsekutiv  = 2;

            bool ZeileHatInhalt(int y)
            {
                int off = y * stride, dunkel = 0;
                for (int x = 0; x < w; x++)
                {
                    int idx = off + x * 4;
                    if (pixels[idx] < BgSchwelle || pixels[idx+1] < BgSchwelle || pixels[idx+2] < BgSchwelle)
                        dunkel++;
                }
                return (double)dunkel / w >= DichteSchwelle;
            }

            bool SpalteHatInhalt(int x)
            {
                int dunkel = 0;
                for (int y = 0; y < h; y++)
                {
                    int idx = y * stride + x * 4;
                    if (pixels[idx] < BgSchwelle || pixels[idx+1] < BgSchwelle || pixels[idx+2] < BgSchwelle)
                        dunkel++;
                }
                return (double)dunkel / h >= DichteSchwelle;
            }

            // Erste Position suchen wo MinKonsekutiv aufeinanderfolgende Zeilen/Spalten
            // Inhalt haben → verhindert, dass einzelne Rauschpixel den Rand zu früh setzen
            int FindeErste(int von, int bis, int schritt, System.Func<int, bool> hatInhalt)
            {
                int konsek = 0, erster = -1;
                for (int p = von; p != bis; p += schritt)
                {
                    if (hatInhalt(p)) { if (konsek == 0) erster = p; konsek++; if (konsek >= MinKonsekutiv) return erster; }
                    else              { konsek = 0; erster = -1; }
                }
                return -1;
            }

            int oben   = FindeErste(0,   h,    1,  ZeileHatInhalt);
            if (oben < 0) return (0, 0, 0, 0);
            int unten  = FindeErste(h-1, -1,  -1,  ZeileHatInhalt);
            int links  = FindeErste(0,   w,    1,  SpalteHatInhalt);
            int rechts = FindeErste(w-1, -1,  -1,  SpalteHatInhalt);

            if (links < 0 || rechts < 0 || unten < 0) return (0, 0, 0, 0);

            // Sicherheitsabstand
            oben   = Math.Max(0,   oben  - sicherheitsPx);
            unten  = Math.Min(h-1, unten + sicherheitsPx);
            links  = Math.Max(0,   links - sicherheitsPx);
            rechts = Math.Min(w-1, rechts + sicherheitsPx);

            return (
                (double)links        / w,
                (double)(w-1-rechts) / w,
                (double)oben         / h,
                (double)(h-1-unten)  / h
            );
        }

        // ── Crop-Linien Drag & Drop ───────────────────────────────────────────

        private void CropLinie_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is not Line l || !IstCropLinie(l)) return;
                // Basis-Tag speichern (ohne "_HIT") damit MouseMove-Handler ihn vergleichen kann
                _gezogeneCropSeite = BasisTag(l.Tag?.ToString() ?? "");
                // Capture auf Canvas – robuster als auf der schmalen Linie selbst
                PdfCanvas.CaptureMouse();
                PdfCanvas.MouseMove         += CropCanvas_MouseMove;
                PdfCanvas.MouseLeftButtonUp += CropCanvas_MouseUp;
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropLinie_MouseDown"); }
        }

        private void CropCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (_gezogeneCropSeite == null || _seitenBilder.Count == 0) return;
                Point pos = e.GetPosition(PdfCanvas);

                string gs = _gezogeneCropSeite;
                if (gs.StartsWith("CROP_LINKS_") &&
                    int.TryParse(gs.Substring(11), out int iL) && iL < _seitenBilder.Count)
                {
                    double cR    = iL < _cropRechts.Length ? _cropRechts[iL] : 0;
                    double pageW = _seitenBilder[iL].PixelWidth;
                    double xBase = _layoutHorizontal && iL < _seitenXStart.Length ? _seitenXStart[iL] : SeiteX;
                    if (iL < _cropLinks.Length)
                        _cropLinks[iL] = Math.Max(0, Math.Min(0.49 - cR, (pos.X - xBase) / pageW));
                }
                else if (gs.StartsWith("CROP_RECHTS_") &&
                         int.TryParse(gs.Substring(12), out int iR) && iR < _seitenBilder.Count)
                {
                    double cL    = iR < _cropLinks.Length ? _cropLinks[iR] : 0;
                    double pageW = _seitenBilder[iR].PixelWidth;
                    double xBase = _layoutHorizontal && iR < _seitenXStart.Length ? _seitenXStart[iR] : SeiteX;
                    if (iR < _cropRechts.Length)
                        _cropRechts[iR] = Math.Max(0, Math.Min(0.49 - cL, (xBase + pageW - pos.X) / pageW));
                }
                else if (gs.StartsWith("CROP_OBEN_") &&
                         int.TryParse(gs.Substring(10), out int iO) &&
                         iO < _seitenHöhe.Length && _seitenHöhe[iO] > 0)
                {
                    double cU    = iO < _cropUnten.Length ? _cropUnten[iO] : 0;
                    double yBase = _layoutHorizontal ? SeiteX : (_seitenYStart.Length > iO ? _seitenYStart[iO] : 0);
                    if (iO < _cropOben.Length)
                        _cropOben[iO] = Math.Max(0, Math.Min(0.49 - cU, (pos.Y - yBase) / _seitenHöhe[iO]));
                }
                else if (gs.StartsWith("CROP_UNTEN_") &&
                         int.TryParse(gs.Substring(11), out int iU) &&
                         iU < _seitenHöhe.Length && _seitenHöhe[iU] > 0)
                {
                    double cO    = iU < _cropOben.Length ? _cropOben[iU] : 0;
                    double yBase = _layoutHorizontal ? SeiteX : (_seitenYStart.Length > iU ? _seitenYStart[iU] : 0);
                    double bot   = yBase + _seitenHöhe[iU];
                    if (iU < _cropUnten.Length)
                        _cropUnten[iU] = Math.Max(0, Math.Min(0.49 - cO, (bot - pos.Y) / _seitenHöhe[iU]));
                }

                AktualisiereCropLinien();
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropCanvas_MouseMove"); }
        }

        private void CropCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (_gezogeneCropSeite == null) return;
                PdfCanvas.ReleaseMouseCapture();
                PdfCanvas.MouseMove         -= CropCanvas_MouseMove;
                PdfCanvas.MouseLeftButtonUp -= CropCanvas_MouseUp;

                // Modus-Propagierung: finalen Wert der gezogenen Seite auf Zielseiten übertragen
                int dragIdx = HoleSeitenIndexAusTag(_gezogeneCropSeite);
                if (dragIdx >= 0) WendeCropModusAnNachDrag(dragIdx);

                _gezogeneCropSeite = null;
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "CropCanvas_MouseUp"); _gezogeneCropSeite = null; }
        }

        // ── Toolbar-Handler: Rand ─────────────────────────────────────────────

        private void CmbCropModus_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_modusWahlLäuft) return;
            // Null-Guard: Handler kann während InitializeComponent() feuern,
            // bevor BtnAuswahlmodus via Connect() gesetzt wurde.
            if (BtnAuswahlmodus == null) return;
            int idx = CmbCropModus?.SelectedIndex ?? -1;
            if (idx < 0) return;

            _cropModus = (CropAnwendungsModus)Math.Min(idx, 3);

            AktualisiereAuswahlAnzeige();
        }

        private void BtnAuswahlmodus_Checked(object sender, RoutedEventArgs e)
        {
            if (_seitenBilder.Count == 0) { BtnAuswahlmodus.IsChecked = false; return; }
            StarteBearbeitungsModus();
        }

        private void BtnAuswahlmodus_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_bearbeitungsEndet) return; // ausgelöst durch BeendeBearbeitungsModus → ignorieren
            if (_bearbeitungsModus) BeendeBearbeitungsModus(false); // User deaktiviert Toggle = Abbrechen
        }

        // ── Gruppen-Handler ───────────────────────────────────────────────────

        private void CmbGruppe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_gruppeWahlLäuft) return;
            int selIdx = CmbGruppe?.SelectedIndex ?? -1;
            if (selIdx < 0 || selIdx >= _gruppen.Count) return;
            _aktGruppeId = _gruppen[selIdx].Id;
            AktualisiereAuswahlAnzeige();
        }

        private void CmbGruppe_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Return) return;
            BenennGruppeUm();
            e.Handled = true;
        }

        private void CmbGruppe_LostFocus(object sender, RoutedEventArgs e)
        {
            BenennGruppeUm();
        }

        private void BenennGruppeUm()
        {
            if (_gruppeWahlLäuft || CmbGruppe == null) return;
            var aktGruppe = AktiveGruppe();
            if (aktGruppe == null) return;
            string neuerName = (CmbGruppe.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(neuerName) || neuerName == aktGruppe.Name) return;
            aktGruppe.Name = neuerName;
            AktualisiereGruppenComboBox();
        }

        private void BtnGruppeNeu_Click(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                int id   = _nächsteId++;
                int pos  = _gruppen.Count;
                var name = $"Gruppe {pos + 1}";
                _gruppen.Add(new CropGruppe { Id = id, Name = name });
                _aktGruppeId = id;
                AktualisiereGruppenComboBox();
                AktualisiereAuswahlAnzeige();
            }, "BtnGruppeNeu_Click");
        }

        private void BtnGruppeLöschen_Click(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                var aktGruppe = AktiveGruppe();
                if (aktGruppe == null) return;
                if (aktGruppe.Id == 0) return; // Gruppe 0 ist unverlöschbar
                if (_gruppen.Count <= 1) return;

                // Seiten der gelöschten Gruppe → Gruppe 0
                var gruppe0 = _gruppen.FirstOrDefault(g => g.Id == 0);
                if (gruppe0 != null)
                    foreach (int s in aktGruppe.Seiten)
                        if (!gruppe0.Seiten.Contains(s)) gruppe0.Seiten.Add(s);

                _gruppen.Remove(aktGruppe);
                _aktGruppeId = 0; // nach Löschen immer zu Gruppe 0 zurück
                AktualisiereGruppenComboBox();
                AktualisiereAuswahlAnzeige();
            }, "BtnGruppeLöschen_Click");
        }

        // ── Bearbeitungs-Handler ──────────────────────────────────────────────

        private void BtnBearbeitungOk_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => BeendeBearbeitungsModus(true), "BtnBearbeitungOk_Click");

        private void BtnBearbeitungAbbrechen_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => BeendeBearbeitungsModus(false), "BtnBearbeitungAbbrechen_Click");

        private void BtnAlleSeiten_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                _tempSeiten = Enumerable.Range(0, _seitenBilder.Count).ToList();
                AktualisiereTempInfo();
                AktualisiereAuswahlAnzeige();
            }, "BtnAlleSeiten_Click");

        private void BtnKeineSeiten_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                _tempSeiten.Clear();
                AktualisiereTempInfo();
                AktualisiereAuswahlAnzeige();
            }, "BtnKeineSeiten_Click");

        private void BtnBereichHinzufügen_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                var neu = ParseSeitenBereich(TxtSeitenBereich?.Text ?? "");
                foreach (int s in neu)
                    if (!_tempSeiten.Contains(s)) _tempSeiten.Add(s);
                if (TxtSeitenBereich != null) TxtSeitenBereich.Clear();
                AktualisiereTempInfo();
                AktualisiereAuswahlAnzeige();
            }, "BtnBereichHinzufügen_Click");

        private void TxtSeitenBereich_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                BtnBereichHinzufügen_Click(sender, new RoutedEventArgs());
                e.Handled = true;
            }
        }

        private void BtnRandAnzeigen_Checked(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                foreach (var l in PdfCanvas.Children.OfType<Line>().Where(IstCropLinie))
                    l.Visibility = Visibility.Visible;
            }, "BtnRandAnzeigen_Checked");

        private void BtnRandAnzeigen_Unchecked(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                foreach (var l in PdfCanvas.Children.OfType<Line>().Where(IstCropLinie))
                    l.Visibility = Visibility.Collapsed;
            }, "BtnRandAnzeigen_Unchecked");

        private void BtnRandAuto_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => { _autoRandAktiv = true; AktualisiereAutoRand(); }, "BtnRandAuto_Click");

        private void AktualisiereAutoRand()
        {
            if (_seitenBilder.Count == 0) return;

            _autoRandCts?.Cancel();
            var cts = new CancellationTokenSource();
            _autoRandCts = cts;
            var token = cts.Token;

            TxtInfo.Text = "Erkenne Rand …";
            var bilder = _seitenBilder.ToList();
            int sicherheitPx = Math.Max(0, (int)Math.Round(_cropSicherheitMm * _pxPerMm));
            int n = bilder.Count;

            AppZustand.Instanz.SetzeProgress(0, n);
            AppZustand.Instanz.SetzeStatus($"Auto-Rand: 0 von {n} Seiten");

            Task.Run(() =>
            {
                var perL = new double[n]; var perR = new double[n];
                var perO = new double[n]; var perU = new double[n];
                try
                {
                    for (int i = 0; i < n; i++)
                    {
                        if (token.IsCancellationRequested)
                        {
                            Dispatcher.BeginInvoke(new Action(() => AppZustand.Instanz.ResetProgress()));
                            return;
                        }
                        var (l, r, o, u) = ErkenneCropRänderVonBitmap(bilder[i], sicherheitPx);
                        perL[i] = l; perR[i] = r; perO[i] = o; perU[i] = u;
                        int seite = i + 1;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            AppZustand.Instanz.SetzeProgress(seite, n);
                            AppZustand.Instanz.SetzeStatus($"Auto-Rand: Seite {seite} von {n}");
                        }));
                    }
                }
                catch (OperationCanceledException) { return; }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (token.IsCancellationRequested) { AppZustand.Instanz.ResetProgress(); return; }

                    // Anwendung gemäß aktuellem Modus
                    IEnumerable<int> zielIdxs;
                    int aktIdx = AktiveSeiteIndex();
                    switch (_cropModus)
                    {
                        case CropAnwendungsModus.AlsStandard:
                            if (aktIdx >= 0 && aktIdx < n)
                                _defaultCrop = (perL[aktIdx], perR[aktIdx], perO[aktIdx], perU[aktIdx]);
                            TxtInfo.Text = "Auto-Rand als Standard gespeichert.";
                            AppZustand.Instanz.ResetProgress();
                            return;
                        case CropAnwendungsModus.NurDiese:
                            zielIdxs = aktIdx >= 0 ? new[] { aktIdx } : Enumerable.Empty<int>();
                            break;
                        case CropAnwendungsModus.Ausgewählt:
                            var zielMenge = AktuelleZielSeiten().ToList();
                            zielIdxs = zielMenge.Count > 0
                                ? zielMenge
                                : (aktIdx >= 0 ? new[] { aktIdx } : Enumerable.Empty<int>());
                            break;
                        default: // Alle
                            zielIdxs = Enumerable.Range(0, n);
                            break;
                    }
                    foreach (int i in zielIdxs)
                    {
                        if (i < n) { _cropLinks[i] = perL[i]; _cropRechts[i] = perR[i];
                                     _cropOben[i]  = perO[i]; _cropUnten[i]  = perU[i]; }
                    }
                    AktualisiereCropLinien();
                    if (BtnRandAnzeigen.IsChecked != true) BtnRandAnzeigen.IsChecked = true;
                    int anz = zielIdxs.Count();
                    TxtInfo.Text = anz == 1 && aktIdx >= 0
                        ? $"Rand · L:{perL[aktIdx]*100:F1}%  O:{perO[aktIdx]*100:F1}%  R:{perR[aktIdx]*100:F1}%  U:{perU[aktIdx]*100:F1}%"
                        : $"Auto-Rand für {anz} Seite(n) erkannt.";
                    AppZustand.Instanz.ResetProgress();
                }));
            });
        }

        private void BtnRandBearbeiten_Click(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0)
                { MessageBox.Show("Kein PDF geladen.", "Rand bearbeiten"); return; }

                int aktSeite = AktiveSeiteIndex();
                if (aktSeite < 0) return;

                double refH    = aktSeite < _seitenHöhe.Length ? _seitenHöhe[aktSeite] : 1000;
                double refW    = _seitenBilder[aktSeite].PixelWidth;
                double pxPerMm = _pxPerMm > 0 ? _pxPerMm : 4.0;

                // Originalwerte der aktiven Seite für Cancel-Restore sichern
                double origOben   = aktSeite < _cropOben.Length   ? _cropOben[aktSeite]   : 0;
                double origUnten  = aktSeite < _cropUnten.Length  ? _cropUnten[aktSeite]  : 0;
                double origLinks  = aktSeite < _cropLinks.Length  ? _cropLinks[aktSeite]  : 0;
                double origRechts = aktSeite < _cropRechts.Length ? _cropRechts[aktSeite] : 0;

                // Beschnittrahmen einblenden damit Live-Aktualisierung sichtbar ist
                if (BtnRandAnzeigen.IsChecked != true) BtnRandAnzeigen.IsChecked = true;

                // Crop-Fractions -> Pixel -> mm (aktive Seite)
                double obenMm   = Math.Round(origOben   * refH / pxPerMm, 1);
                double untenMm  = Math.Round(origUnten  * refH / pxPerMm, 1);
                double linksMm  = Math.Round(origLinks  * refW / pxPerMm, 1);
                double rechtsMm = Math.Round(origRechts * refW / pxPerMm, 1);

                TextBox MkTxt(double v) => new TextBox
                {
                    Text = v.ToString("F1"), Width = 72,
                    TextAlignment = TextAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center
                };
                TextBlock MkLbl(string s) => new TextBlock
                {
                    Text = s, VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 6, 0), Width = 60
                };

                var txtOben   = MkTxt(obenMm);
                var txtUnten  = MkTxt(untenMm);
                var txtLinks  = MkTxt(linksMm);
                var txtRechts = MkTxt(rechtsMm);

                var sp = new StackPanel { Margin = new Thickness(14) };
                sp.Children.Add(new TextBlock
                {
                    Text = "Rand in mm (0 = kein Rand):",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 8)
                });

                void AddRow(string label, TextBox tb)
                {
                    var row = new DockPanel { Margin = new Thickness(0, 3, 0, 3) };
                    var lbl = MkLbl(label); DockPanel.SetDock(lbl, Dock.Left); row.Children.Add(lbl);
                    var mm  = new TextBlock { Text = " mm", VerticalAlignment = VerticalAlignment.Center };
                    DockPanel.SetDock(mm, Dock.Right); row.Children.Add(mm);
                    row.Children.Add(tb);
                    sp.Children.Add(row);
                }

                AddRow("Oben:",   txtOben);
                AddRow("Unten:",  txtUnten);
                AddRow("Links:",  txtLinks);
                AddRow("Rechts:", txtRechts);

                // Anwenden-auf-Modus
                sp.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 6) });
                sp.Children.Add(new TextBlock
                {
                    Text = "Anwenden auf:",
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 4)
                });
                var rbNurDiese = new RadioButton
                    { Content = $"Nur diese Seite (Seite {aktSeite + 1})", IsChecked = true, Margin = new Thickness(0, 2, 0, 2) };
                var rbAlle     = new RadioButton
                    { Content = "Alle Seiten", Margin = new Thickness(0, 2, 0, 2) };
                var rbAuswahl  = new RadioButton
                    { Content = "Ausgew\u00e4hlte Seiten \u2026", Margin = new Thickness(0, 2, 0, 2) };
                var rbStandard = new RadioButton
                    { Content = "Als Standard speichern", Margin = new Thickness(0, 2, 0, 2) };
                sp.Children.Add(rbNurDiese);
                if (_seitenBilder.Count > 1) { sp.Children.Add(rbAlle); sp.Children.Add(rbAuswahl); }
                sp.Children.Add(rbStandard);

                // Live-Aktualisierung: Rahmen sofort bei jeder Eingabe verschieben
                bool TryMm(string t, out double v)
                    => double.TryParse(t.Trim().Replace(",", "."),
                           System.Globalization.NumberStyles.Any,
                           System.Globalization.CultureInfo.InvariantCulture, out v) && v >= 0;

                void AktualisiereLive(object s, TextChangedEventArgs ev)
                {
                    if (TryMm(txtOben.Text,   out double lO) && aktSeite < _cropOben.Length)   _cropOben[aktSeite]   = Math.Min(lO * pxPerMm / refH, 0.49);
                    if (TryMm(txtUnten.Text,  out double lU) && aktSeite < _cropUnten.Length)  _cropUnten[aktSeite]  = Math.Min(lU * pxPerMm / refH, 0.49);
                    if (TryMm(txtLinks.Text,  out double lL) && aktSeite < _cropLinks.Length)  _cropLinks[aktSeite]  = Math.Min(lL * pxPerMm / refW, 0.49);
                    if (TryMm(txtRechts.Text, out double lR) && aktSeite < _cropRechts.Length) _cropRechts[aktSeite] = Math.Min(lR * pxPerMm / refW, 0.49);
                    AktualisiereCropLinien();
                }
                txtOben.TextChanged   += AktualisiereLive;
                txtUnten.TextChanged  += AktualisiereLive;
                txtLinks.TextChanged  += AktualisiereLive;
                txtRechts.TextChanged += AktualisiereLive;

                bool ok = false;
                Window? dlg = null;
                var btnOk = new Button
                    { Content = "OK", Width = 70, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
                var btnCancel = new Button { Content = "Abbrechen", Width = 80, IsCancel = true };
                btnOk.Click     += (_, __) => { ok = true; dlg!.Close(); };
                btnCancel.Click += (_, __) =>
                {
                    if (aktSeite < _cropOben.Length)   _cropOben[aktSeite]   = origOben;
                    if (aktSeite < _cropUnten.Length)  _cropUnten[aktSeite]  = origUnten;
                    if (aktSeite < _cropLinks.Length)  _cropLinks[aktSeite]  = origLinks;
                    if (aktSeite < _cropRechts.Length) _cropRechts[aktSeite] = origRechts;
                    AktualisiereCropLinien();
                    dlg!.Close();
                };

                var btnRow = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 12, 0, 0)
                };
                btnRow.Children.Add(btnOk);
                btnRow.Children.Add(btnCancel);
                sp.Children.Add(btnRow);

                dlg = new Window
                {
                    Title = $"Rand bearbeiten – Seite {aktSeite + 1}", Content = sp,
                    SizeToContent = SizeToContent.WidthAndHeight,
                    MinWidth = 260,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = Window.GetWindow(this),
                    ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
                };
                dlg.ShowDialog();
                if (!ok) return;

                if (!double.TryParse(txtOben.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nO) || nO < 0) return;
                if (!double.TryParse(txtUnten.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nU) || nU < 0) return;
                if (!double.TryParse(txtLinks.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nL) || nL < 0) return;
                if (!double.TryParse(txtRechts.Text.Trim().Replace(",", "."),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double nR) || nR < 0) return;

                double newObenPx   = nO * pxPerMm;
                double newUntenPx  = nU * pxPerMm;
                double newLinksPx  = nL * pxPerMm;
                double newRechtsPx = nR * pxPerMm;

                if (newObenPx + newUntenPx >= refH || newLinksPx + newRechtsPx >= refW)
                {
                    MessageBox.Show("Rand-Summe überschreitet die Seitengröße.",
                        "Ungültige Eingabe", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Modus auswerten
                if (rbStandard.IsChecked == true)
                {
                    // Als Standard speichern – Live-Preview auf aktSeite rückgängig machen
                    _defaultCrop = (newLinksPx / refW, newRechtsPx / refW, newObenPx / refH, newUntenPx / refH);
                    if (aktSeite < _cropOben.Length)   _cropOben[aktSeite]   = origOben;
                    if (aktSeite < _cropUnten.Length)  _cropUnten[aktSeite]  = origUnten;
                    if (aktSeite < _cropLinks.Length)  _cropLinks[aktSeite]  = origLinks;
                    if (aktSeite < _cropRechts.Length) _cropRechts[aktSeite] = origRechts;
                    AktualisiereCropLinien();
                    TxtInfo.Text = "Standard-Rand gespeichert (wird beim n\u00e4chsten PDF-Laden angewendet).";
                    return;
                }

                List<int> zielSeiten;
                if (rbAlle.IsChecked == true)
                {
                    zielSeiten = Enumerable.Range(0, _seitenBilder.Count).ToList();
                }
                else if (rbAuswahl.IsChecked == true)
                {
                    // Im Bearbeitungsmodus: _tempSeiten; sonst aktive Gruppe. Fallback: aktive Seite.
                    var zielDialog = AktuelleZielSeiten().ToList();
                    zielSeiten = zielDialog.Count > 0
                        ? zielDialog
                        : new List<int> { aktSeite };
                }
                else
                {
                    zielSeiten = new List<int> { aktSeite };
                }

                // Werte auf Zielseiten anwenden (mm → Bruchteil der jeweiligen Seitengröße)
                foreach (int idx in zielSeiten)
                {
                    double iRefH = idx < _seitenHöhe.Length  ? _seitenHöhe[idx]               : refH;
                    double iRefW = idx < _seitenBilder.Count ? _seitenBilder[idx].PixelWidth   : refW;
                    if (idx < _cropOben.Length)   _cropOben[idx]   = Math.Min(newObenPx  / iRefH, 0.49);
                    if (idx < _cropUnten.Length)  _cropUnten[idx]  = Math.Min(newUntenPx / iRefH, 0.49);
                    if (idx < _cropLinks.Length)  _cropLinks[idx]  = Math.Min(newLinksPx / iRefW, 0.49);
                    if (idx < _cropRechts.Length) _cropRechts[idx] = Math.Min(newRechtsPx / iRefW, 0.49);
                }
                _autoRandAktiv = false;

                AktualisiereCropLinien();
                if (BtnRandAnzeigen.IsChecked != true) BtnRandAnzeigen.IsChecked = true;
            }, "BtnRandBearbeiten_Click");
        }

        // ── Seitenauswahl-Dialog ──────────────────────────────────────────────

        // Zeigt einen Dialog mit Checkboxen für alle Seiten.
        // Gibt die Indizes der gewählten Seiten zurück; leere Liste = Abbruch.
        private List<int> WähleSeiten(int aktSeite, Window? owner = null)
        {
            int n        = _seitenBilder.Count;
            var selected = new bool[n];
            if (aktSeite >= 0 && aktSeite < n) selected[aktSeite] = true;
            int aktIdx   = Math.Max(0, Math.Min(n - 1, aktSeite));

            // ── Alle UI-Controls vorab deklarieren (Reihenfolge wichtig für lokale Funktionen) ──

            // Rechte Vorschau
            var imgVorschau = new Image
            {
                Stretch             = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin              = new Thickness(6)
            };

            // Navigations-Zeile
            var btnZurück = new Button { Content = "\u25c4  Zur\u00fcck", Width = 90, Height = 26, Margin = new Thickness(0, 0, 8, 0) };
            var btnWeiter = new Button { Content = "Weiter  \u25ba", Width = 90, Height = 26, Margin = new Thickness(8, 0, 0, 0) };
            var txtSeite  = new TextBlock
            {
                VerticalAlignment   = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                FontSize            = 12,
                MinWidth            = 100,
                TextAlignment       = TextAlignment.Center
            };

            // Toggle "Diese Seite auswählen"
            var cbDiese = new CheckBox
            {
                Content  = "Diese Seite ausw\u00e4hlen",
                FontSize = 12,
                Margin   = new Thickness(8, 2, 8, 6)
            };

            // Header-Anzahl
            var txtAnzahl = new TextBlock
            {
                FontSize            = 11,
                Foreground          = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin              = new Thickness(0, 0, 6, 0)
            };

            // Linke Seitenliste
            var listSp  = new StackPanel();
            var listCbs = new CheckBox[n];
            var listBtn = new Border[n];

            // Footer
            var btnAlle   = new Button { Content = "Alle",      Width = 58, Height = 24, Margin = new Thickness(0, 0, 6, 0) };
            var btnKeine  = new Button { Content = "Keine",     Width = 58, Height = 24 };
            bool ok       = false;
            Window? dlg   = null;
            var btnOk     = new Button { Content = "OK",        Width = 70, Height = 26, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var btnCancel = new Button { Content = "Abbrechen", Width = 80, Height = 26, IsCancel  = true };
            btnOk.Click     += (_, __) => { ok = true; dlg!.Close(); };
            btnCancel.Click += (_, __) => dlg!.Close();

            // ── Interne Aktualisierungsfunktionen ─────────────────────────────

            int AnzahlGewählt() => selected.Count(v => v);

            void AktualisiereAnzahl()
                => txtAnzahl.Text = $"{AnzahlGewählt()} von {n} ausgew\u00e4hlt";

            void AktualisiereListenMarkierung()
            {
                for (int i = 0; i < n; i++)
                {
                    listCbs[i].IsChecked   = selected[i];
                    bool istAktuell        = (i == aktIdx);
                    listBtn[i].BorderThickness = istAktuell ? new Thickness(2) : new Thickness(0);
                    listBtn[i].BorderBrush     = istAktuell
                        ? new SolidColorBrush(Color.FromRgb(0, 120, 215))
                        : Brushes.Transparent;
                    if (listBtn[i].Child is StackPanel rowSp)
                    {
                        var lbl = rowSp.Children.OfType<TextBlock>().FirstOrDefault();
                        if (lbl != null)
                            lbl.FontWeight = istAktuell ? FontWeights.SemiBold : FontWeights.Normal;
                    }
                }
                listBtn[aktIdx].BringIntoView();
            }

            void ZeigeSeite(int idx)
            {
                aktIdx              = Math.Max(0, Math.Min(n - 1, idx));
                imgVorschau.Source  = _seitenBilder[aktIdx];
                txtSeite.Text       = $"Seite {aktIdx + 1} von {n}";
                cbDiese.IsChecked   = selected[aktIdx];
                btnZurück.IsEnabled = aktIdx > 0;
                btnWeiter.IsEnabled = aktIdx < n - 1;
                AktualisiereListenMarkierung();
            }

            // ── Linke Seitenliste befüllen ────────────────────────────────────

            for (int i = 0; i < n; i++)
            {
                int idx = i; // capture für Lambda

                listCbs[i] = new CheckBox
                {
                    IsChecked         = selected[i],
                    IsHitTestVisible  = false,          // rein visuell
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin            = new Thickness(0, 0, 6, 0)
                };

                var lbl = new TextBlock
                {
                    Text              = $"Seite {i + 1}",
                    VerticalAlignment = VerticalAlignment.Center,
                    FontSize          = 12
                };

                var rowSp = new StackPanel { Orientation = Orientation.Horizontal };
                rowSp.Children.Add(listCbs[i]);
                rowSp.Children.Add(lbl);

                listBtn[i] = new Border
                {
                    Child      = rowSp,
                    Padding    = new Thickness(8, 4, 8, 4),
                    Cursor     = Cursors.Hand,
                    Background = Brushes.Transparent
                };

                // Hover-Effekt
                listBtn[i].MouseEnter += (s, _) =>
                    ((Border)s!).Background = new SolidColorBrush(Color.FromRgb(220, 232, 248));
                listBtn[i].MouseLeave += (s, _) =>
                    ((Border)s!).Background = Brushes.Transparent;

                // Klick → nur navigieren, keine Auswahländerung
                listBtn[i].MouseLeftButtonUp += (_, __) => ZeigeSeite(idx);

                listSp.Children.Add(listBtn[i]);
            }

            // ── Event-Handler ─────────────────────────────────────────────────

            btnZurück.Click += (_, __) => ZeigeSeite(aktIdx - 1);
            btnWeiter.Click += (_, __) => ZeigeSeite(aktIdx + 1);

            btnZurück.PreviewKeyDown += (_, ev) =>
            {
                if (ev.Key == Key.Left || ev.Key == Key.Up) { ZeigeSeite(aktIdx - 1); ev.Handled = true; }
            };
            btnWeiter.PreviewKeyDown += (_, ev) =>
            {
                if (ev.Key == Key.Right || ev.Key == Key.Down) { ZeigeSeite(aktIdx + 1); ev.Handled = true; }
            };

            cbDiese.Checked   += (_, __) => { selected[aktIdx] = true;  AktualisiereListenMarkierung(); AktualisiereAnzahl(); };
            cbDiese.Unchecked += (_, __) => { selected[aktIdx] = false; AktualisiereListenMarkierung(); AktualisiereAnzahl(); };

            btnAlle.Click  += (_, __) => { for (int i = 0; i < n; i++) selected[i] = true;  ZeigeSeite(aktIdx); AktualisiereAnzahl(); };
            btnKeine.Click += (_, __) => { for (int i = 0; i < n; i++) selected[i] = false; ZeigeSeite(aktIdx); AktualisiereAnzahl(); };

            // ── Layout zusammenbauen ──────────────────────────────────────────

            var listScroll = new ScrollViewer
            {
                Content                       = listSp,
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                MinWidth                      = 130,
                MaxWidth                      = 160
            };

            var vorschauBorder = new Border
            {
                Child           = imgVorschau,
                Background      = new SolidColorBrush(Color.FromRgb(74, 74, 74)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                BorderThickness = new Thickness(1),
                Margin          = new Thickness(6, 0, 6, 0)
            };

            var navRow = new DockPanel { Margin = new Thickness(6, 4, 6, 4) };
            DockPanel.SetDock(btnZurück, Dock.Left);
            DockPanel.SetDock(btnWeiter, Dock.Right);
            navRow.Children.Add(btnZurück);
            navRow.Children.Add(btnWeiter);
            navRow.Children.Add(txtSeite);

            var headerRow = new DockPanel { Margin = new Thickness(6, 6, 6, 4) };
            var headerLbl = new TextBlock
            {
                Text              = "Seiten ausw\u00e4hlen",
                FontWeight        = FontWeights.SemiBold,
                FontSize          = 13,
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(txtAnzahl, Dock.Right);
            headerRow.Children.Add(txtAnzahl);
            headerRow.Children.Add(headerLbl);

            var hauptGrid = new Grid();
            hauptGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            hauptGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetColumn(listScroll,     0);
            Grid.SetColumn(vorschauBorder, 1);
            hauptGrid.Children.Add(listScroll);
            hauptGrid.Children.Add(vorschauBorder);

            var footerLinksRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(6, 0, 0, 0) };
            footerLinksRow.Children.Add(btnAlle);
            footerLinksRow.Children.Add(btnKeine);

            var footerRechtsRow = new StackPanel
            {
                Orientation         = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin              = new Thickness(0, 0, 6, 0)
            };
            footerRechtsRow.Children.Add(btnOk);
            footerRechtsRow.Children.Add(btnCancel);

            var footerRow = new DockPanel { Margin = new Thickness(0, 4, 0, 6) };
            DockPanel.SetDock(footerRechtsRow, Dock.Right);
            footerRow.Children.Add(footerRechtsRow);
            footerRow.Children.Add(footerLinksRow);

            var rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            void SetRow(UIElement el, int row) { Grid.SetRow(el, row); rootGrid.Children.Add(el); }
            SetRow(headerRow,                  0);
            SetRow(new Separator(),            1);
            SetRow(hauptGrid,                  2);
            SetRow(new Separator(),            3);
            SetRow(navRow,                     4);
            SetRow(cbDiese,                    5);
            SetRow(new Separator(),            6);
            SetRow(footerRow,                  7);

            dlg = new Window
            {
                Title                 = "Seiten ausw\u00e4hlen",
                Content               = rootGrid,
                Width                 = 680,
                Height                = 520,
                MinWidth              = 500,
                MinHeight             = 380,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner                 = owner,
                ResizeMode            = ResizeMode.CanResize,
                ShowInTaskbar         = false
            };

            AktualisiereAnzahl();
            ZeigeSeite(aktIdx);
            dlg.ShowDialog();

            if (!ok) return new List<int>();
            return Enumerable.Range(0, n).Where(i => selected[i]).ToList();
        }

        // ── Crop-Modus Propagierung ───────────────────────────────────────────

        // Extrahiert Seiten-Index aus Tag-String, z.B. "CROP_LINKS_2" → 2.
        private static int HoleSeitenIndexAusTag(string tag)
        {
            int lastUs = tag.LastIndexOf('_');
            return lastUs >= 0 && int.TryParse(tag.Substring(lastUs + 1), out int idx) ? idx : -1;
        }

        // Kopiert die Crop-Werte der Quellseite auf alle Zielseiten.
        // Skalierung: gleiche physische mm-Menge, auf die jeweilige Seitengröße der Zielseite umgerechnet.
        private void KopiereCropAufSeiten(
            int quellIdx, double qL, double qR, double qO, double qU, IEnumerable<int> ziele)
        {
            double qRefW = quellIdx < _seitenBilder.Count ? _seitenBilder[quellIdx].PixelWidth : 1;
            double qRefH = quellIdx < _seitenHöhe.Length  ? _seitenHöhe[quellIdx]              : 1;
            foreach (int idx in ziele)
            {
                if (idx == quellIdx) continue;
                double iRefW = idx < _seitenBilder.Count ? _seitenBilder[idx].PixelWidth : qRefW;
                double iRefH = idx < _seitenHöhe.Length  ? _seitenHöhe[idx]              : qRefH;
                if (idx < _cropLinks.Length)  _cropLinks[idx]  = Math.Min(qL * qRefW / iRefW, 0.49);
                if (idx < _cropRechts.Length) _cropRechts[idx] = Math.Min(qR * qRefW / iRefW, 0.49);
                if (idx < _cropOben.Length)   _cropOben[idx]   = Math.Min(qO * qRefH / iRefH, 0.49);
                if (idx < _cropUnten.Length)  _cropUnten[idx]  = Math.Min(qU * qRefH / iRefH, 0.49);
            }
            AktualisiereCropLinien();
        }

        // Wendet _cropModus nach Abschluss eines Drag-Vorgangs an.
        // Drag-Preview hat bereits quellIdx aktualisiert; diese Methode verteilt den Wert.
        private void WendeCropModusAnNachDrag(int quellIdx)
        {
            if (quellIdx < 0 || quellIdx >= _seitenBilder.Count) return;
            double qL = quellIdx < _cropLinks.Length  ? _cropLinks[quellIdx]  : 0;
            double qR = quellIdx < _cropRechts.Length ? _cropRechts[quellIdx] : 0;
            double qO = quellIdx < _cropOben.Length   ? _cropOben[quellIdx]   : 0;
            double qU = quellIdx < _cropUnten.Length  ? _cropUnten[quellIdx]  : 0;

            switch (_cropModus)
            {
                case CropAnwendungsModus.NurDiese:
                    // bereits durch MouseMove gesetzt – nichts weiter zu tun
                    break;
                case CropAnwendungsModus.AlsStandard:
                    // Wert bleibt auf Quellseite sichtbar; zusätzlich als Standard speichern
                    _defaultCrop = (qL, qR, qO, qU);
                    TxtInfo.Text = "Standard-Rand gespeichert.";
                    break;
                case CropAnwendungsModus.Alle:
                    KopiereCropAufSeiten(quellIdx, qL, qR, qO, qU,
                        Enumerable.Range(0, _seitenBilder.Count));
                    break;
                case CropAnwendungsModus.Ausgewählt:
                    var zielDrag = AktuelleZielSeiten().ToList();
                    if (zielDrag.Count > 0)
                        KopiereCropAufSeiten(quellIdx, qL, qR, qO, qU, zielDrag);
                    break;
            }
        }

        // ── Zoom ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Strg + Mausrad → zentriert auf Mausposition zoomen.
        /// Normales Mausrad → Scrollen (Standard-Verhalten des ScrollViewer).
        /// </summary>
        private void ScrollView_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                if (_layoutHorizontal)
                {
                    ScrollView.ScrollToHorizontalOffset(ScrollView.HorizontalOffset - e.Delta / 3.0);
                    e.Handled = true;
                }
                return;
            }
            e.Handled = true;

            try
            {
                // Mausposition vor dem Zoom erfassen (Anker für Scroll-Korrektur im Timer)
                Point mouseInCanvas = e.GetPosition(PdfCanvas);
                Point mouseInView   = e.GetPosition(ScrollView);

                // Nächste Stufe auf Basis des aktuellen Zielzoom berechnen
                double faktor    = e.Delta > 0 ? 1.0 + ZoomStep : 1.0 / (1.0 + ZoomStep);
                double neuerZoom = Math.Max(ZoomMin, Math.Min(ZoomMax, _zielZoom * faktor));
                if (Math.Abs(neuerZoom - _zielZoom) < 0.001) return;

                // Smooth-Animation starten; Anker wird im Timer pro Frame zur Scroll-Korrektur genutzt
                SetzeZoom(neuerZoom, mouseInCanvas, mouseInView);
            }
            catch (Exception ex) { LogException(ex, "Zoom/MouseWheel"); }
        }

        private void BtnZoomMinus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(Math.Max(ZoomMin, _zielZoom / (1.0 + ZoomStep))), "ZoomMinus");

        private void BtnZoomPlus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(Math.Min(ZoomMax, _zielZoom * (1.0 + ZoomStep))), "ZoomPlus");

        private void BtnZoomReset_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() => SetzeZoom(1.0), "ZoomReset");

        // Setzt den Zielzoom und startet die Smooth-Animation.
        // ankerCanvas/ankerView: Mausposition vor dem Zoom (für Scroll-Korrektur pro Frame).
        private void SetzeZoom(double ziel, Point? ankerCanvas = null, Point? ankerView = null)
        {
            _zielZoom = Math.Max(ZoomMin, Math.Min(ZoomMax, ziel));
            if (ankerCanvas.HasValue) _zoomAnker     = ankerCanvas;
            if (ankerView.HasValue)   _zoomAnkerView = ankerView;

            if (_zoomTimer == null)
            {
                _zoomTimer = new System.Windows.Threading.DispatcherTimer
                    { Interval = TimeSpan.FromMilliseconds(16) };
                _zoomTimer.Tick += ZoomTimer_Tick;
            }
            if (!_zoomTimer.IsEnabled) _zoomTimer.Start();
        }

        // Timer-Callback: ein Animations-Frame pro 16 ms.
        private void ZoomTimer_Tick(object? sender, EventArgs e)
        {
            double diff     = _zielZoom - _zoomFaktor;
            bool   fertig   = Math.Abs(diff) < 0.001;
            double neuerZoom = fertig ? _zielZoom : _zoomFaktor + diff * 0.35;

            WendeZoomAnSofort(neuerZoom);

            // Scroll-Korrektur: Ankerpunkt unter dem Mauszeiger stabil halten
            if (_zoomAnker.HasValue && _zoomAnkerView.HasValue)
            {
                ScrollView.ScrollToHorizontalOffset(_zoomAnker.Value.X * neuerZoom - _zoomAnkerView.Value.X);
                ScrollView.ScrollToVerticalOffset  (_zoomAnker.Value.Y * neuerZoom - _zoomAnkerView.Value.Y);
            }

            if (fertig)
            {
                _zoomTimer!.Stop();
                _zoomAnker     = null;
                _zoomAnkerView = null;
            }
        }

        // Wendet den Zoom sofort visuell an (ScaleTransform + Liniendicken).
        private void WendeZoomAnSofort(double zoom)
        {
            _zoomFaktor   = zoom;
            CanvasZoom.ScaleX = zoom;
            CanvasZoom.ScaleY = zoom;
            TxtZoom.Text  = $"{zoom * 100:0}%";

            // Liniendicken anpassen: visuelle Linien dünn, Hit-Zonen breit
            foreach (var l in PdfCanvas.Children.OfType<Line>())
            {
                string tag = l.Tag?.ToString() ?? "";
                if (tag.EndsWith("_HIT"))
                    l.StrokeThickness = Math.Max(14.0, 14.0 / zoom);
                else if (tag.StartsWith("CROP_"))
                    l.StrokeThickness = 2.0 / zoom;
                l.StrokeDashArray = null;
            }
        }

        private void BtnZoomFit_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                double sichtbarB = ScrollView.ActualWidth;
                double sichtbarH = ScrollView.ActualHeight;
                double ersteSeiteBreite = _seitenBilder.Count > 0 ? _seitenBilder[0].PixelWidth : RenderBreite;
                double ersteSeiteHöhe   = _seitenHöhe.Length   > 0 ? _seitenHöhe[0]             : RenderBreite * 2;
                double canvBreite = ersteSeiteBreite + 2.0 * SeiteX;
                double canvHöhe   = ersteSeiteHöhe   + 2.0 * SeiteX;
                if (canvBreite <= 0 || canvHöhe <= 0 || sichtbarB <= 0 || sichtbarH <= 0) return;
                double zoomB = sichtbarB / canvBreite;
                double zoomH = sichtbarH / canvHöhe;
                SetzeZoom(Math.Max(ZoomMin, Math.Min(ZoomMax, Math.Min(zoomB, zoomH))));
            }, "BtnZoomFit_Click");
        private void BtnExport_Click(object sender, RoutedEventArgs e)
            => SafeExecute(ExportierenNachWord, "BtnExport_Click");

        private void BtnSicherheitMinus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_cropSicherheitMm <= 0) return;
                _cropSicherheitMm = Math.Round(Math.Max(0.0, _cropSicherheitMm - 1.0), 1);
                TxtSicherheit.Text = _cropSicherheitMm.ToString("F0");
                TxtInfo.Text = $"Sicherheitsabstand: {_cropSicherheitMm} mm";
                if (_autoRandAktiv) AktualisiereAutoRand();
            }, "BtnSicherheitMinus_Click");

        private void BtnSicherheitPlus_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_cropSicherheitMm >= 20) return;
                _cropSicherheitMm = Math.Round(Math.Min(20.0, _cropSicherheitMm + 1.0), 1);
                TxtSicherheit.Text = _cropSicherheitMm.ToString("F0");
                TxtInfo.Text = $"Sicherheitsabstand: {_cropSicherheitMm} mm";
                if (_autoRandAktiv) AktualisiereAutoRand();
            }, "BtnSicherheitPlus_Click");
        // ── Word-Export ───────────────────────────────────────────────────────

        private void ExportThreadWorker(
            string zielPfad, string pdfPfad,
            double[] seitenYStart, double[] seitenHöhe,
            double[] cropLinks, double[] cropRechts, double[] cropOben, double[] cropUnten)
        {
            var alleTemp = new List<string>();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                // ── 1. PDF-Seitengröße und effektive Maße nach Beschnitt ──────────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Lese PDF-Maße …"));
                var (nativeW_pts, nativeH_pts) = HolePdfSeitenGrösse(pdfPfad);
                int exportPixW = Math.Max(1, (int)Math.Round(nativeW_pts / 72.0 * ExportDpi));
                int exportPixH = Math.Max(1, (int)Math.Round(nativeH_pts / 72.0 * ExportDpi));

                // Für Word-Seitenformat wird Seite 0 als Referenz verwendet
                double cL0 = cropLinks.Length  > 0 ? cropLinks[0]  : 0;
                double cR0 = cropRechts.Length > 0 ? cropRechts[0] : 0;
                double cO0 = cropOben.Length   > 0 ? cropOben[0]   : 0;
                double cU0 = cropUnten.Length  > 0 ? cropUnten[0]  : 0;
                double nativeW_eff = nativeW_pts * Math.Max(0.01, 1.0 - cL0 - cR0);
                double nativeH_eff = nativeH_pts * Math.Max(0.01, 1.0 - cO0 - cU0);

                App.LogFehler("Export/Maße",
                    $"PDF: {nativeW_pts:F1}×{nativeH_pts:F1} pt | " +
                    $"Render: {exportPixW}×{exportPixH} px | DPI={ExportDpi} | " +
                    $"Eff: {nativeW_eff:F1}×{nativeH_eff:F1} pt | " +
                    $"Crop[0] L:{cL0:P1} R:{cR0:P1} O:{cO0:P1} U:{cU0:P1}");

                // ── 2. Hochauflösend rendern ───────────────────────────────────────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Rendere in hoher Auflösung …"));
                long t1 = sw.ElapsedMilliseconds;
                List<BitmapSource> hochRes;
                try { hochRes = PdfRenderer.RenderiereAlleSeiten(pdfPfad, exportPixW, exportPixH); }
                catch (Exception ex) { LogException(ex, "Export/Render"); hochRes = new List<BitmapSource>(); }
                App.LogFehler("Export/Timing", $"Rendern: {sw.ElapsedMilliseconds - t1} ms ({hochRes.Count} Seiten, {exportPixW}×{exportPixH} px)");

                if (hochRes.Count == 0)
                {
                    Dispatcher.BeginInvoke(new Action(() => {
                        BtnExport.IsEnabled = true;
                        TxtInfo.Text = "Export abgebrochen: Render fehlgeschlagen.";
                    }));
                    return;
                }

                // ── 3. Seitenanzahl prüfen ────────────────────────────────────────
                if (hochRes.Count == 0)
                {
                    Dispatcher.BeginInvoke(new Action(() => {
                        BtnExport.IsEnabled = true;
                        TxtInfo.Text = "Export abgebrochen: Keine Seiten gerendert.";
                    }));
                    return;
                }

                // ── 4. Word öffnen, Schreibbereich lesen, Skalierungsdialog ─────
                Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Word wird geöffnet …"));
                Word.Application? wordApp = null;
                Word.Document?    wordDoc = null;
                int seiteNr = 0;
                try
                {
                    wordApp = new Word.Application { Visible = false };
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wordDoc = wordApp.Documents.Add();
                    NormiereAbsatzstil(wordDoc);

                    // Seitenformat in Word passend zur PDF-Seite einstellen
                    SetzeWordSeitenFormat(wordDoc, nativeW_pts, nativeH_pts);

                    var (availW_pt, availH_pt) = HoleSchreibbereich(wordDoc);
                    App.LogFehler("Export/Schreibbereich",
                        $"Vorlage: {availW_pt:F1} × {availH_pt:F1} pt | " +
                        $"Druckbereich: {nativeW_eff:F1} × {nativeH_eff:F1} pt");

                    double scaleW      = nativeW_eff > 0 ? availW_pt / nativeW_eff : 1.0;
                    double scaleH      = nativeH_eff > 0 ? availH_pt / nativeH_eff : 1.0;
                    double globalScale = Math.Min(1.0, Math.Min(scaleW, scaleH));
                    bool   doScale     = true;

                    App.LogFehler("Export/Skalierung-Detail",
                        $"nativeW_eff={nativeW_eff:F2} pt | nativeH_eff={nativeH_eff:F2} pt | " +
                        $"availW_pt={availW_pt:F2} pt | availH_pt={availH_pt:F2} pt | " +
                        $"scaleW={scaleW:F4} ({scaleW*100:F1}%) | scaleH={scaleH:F4} ({scaleH*100:F1}%) | " +
                        $"globalScale={globalScale:F4} ({globalScale*100:F1}%)");

                    if (globalScale < 0.99)
                    {
                        int prozent = (int)Math.Round(globalScale * 100.0);
                        string engpassAchse = scaleW <= scaleH
                            ? $"Breite: PDF {nativeW_eff:F0} pt > Schreibbereich {availW_pt:F0} pt"
                            : $"Höhe: PDF {nativeH_eff:F0} pt > Schreibbereich {availH_pt:F0} pt";
                        Dispatcher.Invoke(new Action(() =>
                        {
                            var antwort = MessageBox.Show(
                                $"Der Druckbereich des PDFs ({nativeW_eff:F0} × {nativeH_eff:F0} pt)\n" +
                                $"ist größer als der Schreibbereich der Word-Vorlage ({availW_pt:F0} × {availH_pt:F0} pt).\n" +
                                $"  Engpass: {engpassAchse}\n\n" +
                                $"Hinweis: Die Word-Seitenränder reduzieren den Schreibbereich.\n" +
                                $"Proportionale Verkleinerung auf {prozent} % passt den Inhalt vollständig ein.\n\n" +
                                $"Möchten Sie alle Seiten auf {prozent} % verkleinern?\n" +
                                $"(Nein = Überstand rechts/unten abschneiden, Originalgröße bleibt erhalten)",
                                "Proportionale Verkleinerung",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Question);
                            doScale = (antwort == MessageBoxResult.Yes);
                        }));
                    }
                    else
                    {
                        globalScale = 1.0;
                    }

                    App.LogFehler("Export/Skalierung",
                        $"globalScale={globalScale:F4} ({(int)Math.Round(globalScale * 100)} %) | doScale={doScale}");

                    int availW_px = Math.Max(1, (int)Math.Round(availW_pt / 72.0 * ExportDpi));
                    int availH_px = Math.Max(1, (int)Math.Round(availH_pt / 72.0 * ExportDpi));
                    int gesamtGruppen = hochRes.Count;

                    // ── 5. Pro Original-PDF-Seite genau ein Bild → eine Word-Seite ──
                    long t4 = sw.ElapsedMilliseconds;
                    for (int gi = 0; gi < hochRes.Count; gi++)
                    {
                        seiteNr++;
                        int sNr = seiteNr, gGes = gesamtGruppen;
                        Dispatcher.BeginInvoke(new Action(() =>
                            TxtInfo.Text = $"Erstelle Seite {sNr} von {gGes} …"));

                        BitmapSource? segBmp = null;
                        try
                        {
                            segBmp = RendereBlockHochRes(hochRes, seitenYStart, seitenHöhe,
                                    seitenYStart[gi], seitenYStart[gi] + seitenHöhe[gi],
                                    cropLinks, cropRechts, cropOben, cropUnten);
                        }
                        catch (Exception ex) { LogException(ex, $"Export/Block[{gi}]"); }
                        if (segBmp == null) continue;

                        BitmapSource finalBmp;
                        float finalW_pt, finalH_pt;
                        double segW_pt = segBmp.PixelWidth  / ExportDpi * 72.0;
                        double segH_pt = segBmp.PixelHeight / ExportDpi * 72.0;

                        if (doScale)
                        {
                            finalBmp  = segBmp;
                            finalW_pt = (float)Math.Min(availW_pt, segW_pt * globalScale);
                            finalH_pt = (float)Math.Min(availH_pt, segH_pt * globalScale);
                        }
                        else
                        {
                            int clipW = Math.Min(segBmp.PixelWidth,  availW_px);
                            int clipH = Math.Min(segBmp.PixelHeight, availH_px);
                            BitmapSource clipped = segBmp;
                            try
                            {
                                var cb = new CroppedBitmap(segBmp, new Int32Rect(0, 0, clipW, clipH));
                                cb.Freeze();
                                clipped = cb;
                            }
                            catch (Exception ex) { LogException(ex, $"Export/Clip[{gi}]"); }
                            finalBmp  = clipped;
                            finalW_pt = (float)(finalBmp.PixelWidth  / ExportDpi * 72.0);
                            finalH_pt = (float)(finalBmp.PixelHeight / ExportDpi * 72.0);
                        }

                        App.LogFehler("Export/Seite",
                            $"[{gi}] {segBmp.PixelWidth}x{segBmp.PixelHeight} px " +
                            $"Shape: {finalW_pt:F1}x{finalH_pt:F1} pt | doScale={doScale}");

                        string png = IO.Path.Combine(IO.Path.GetTempPath(), $"sm_{Guid.NewGuid():N}.png");
                        try
                        {
                            var enc = new PngBitmapEncoder();
                            enc.Frames.Add(BitmapFrame.Create(finalBmp));
                            using var fs = IO.File.Create(png);
                            enc.Save(fs);
                            alleTemp.Add(png);
                        }
                        catch (Exception ex) { LogException(ex, $"Export/PNG[{gi}]"); continue; }

                        bool seitenumbruch = (seiteNr > 1);
                        try { EinfügenSegment(wordDoc, png, finalW_pt, finalH_pt,
                                neuerAbsatz: false, seitenumbruchVorher: seitenumbruch); }
                        catch (Exception ex) { LogException(ex, $"Export/Word[{gi}]"); }
                    }
                    App.LogFehler("Export/Timing",
                        $"Bilder+Einfügen: {sw.ElapsedMilliseconds - t4} ms ({seiteNr} Seiten)");

                    BereinigeDokumentAbsätze(wordDoc);

                    // Leerseite am Ende verhindern: Word-Dokumente haben immer einen
                    // abschließenden leeren Absatz. Wenn der letzte Inhalt die Seite
                    // komplett füllt, landet dieser Absatz auf einer neuen (leeren) Seite.
                    // Lösung: letzten leeren Absatz auf 1 pt Schriftgröße setzen.
                    try
                    {
                        int absNr = wordDoc.Paragraphs.Count;
                        if (absNr > 0)
                        {
                            var letzter = wordDoc.Paragraphs[absNr];
                            var txt     = letzter.Range.Text;
                            // Absatz gilt als leer wenn er nur Absatzmarke enthält
                            if (txt == "\r" || (txt != null && txt.Replace("\r","").Replace("\n","").Trim().Length == 0))
                            {
                                letzter.Range.Font.Size    = 1f;
                                letzter.Format.SpaceBefore = 0f;
                                letzter.Format.SpaceAfter  = 0f;
                                App.LogFehler("Export/LetzterAbsatz", "Auf 1pt minimiert (war leer)");
                            }
                        }
                    }
                    catch (Exception ex) { LogException(ex, "Export/LetzterAbsatz"); }

                    // Datei vorher löschen → kein Word-Überschreiben-Dialog
                    if (IO.File.Exists(zielPfad))
                        try { IO.File.Delete(zielPfad); } catch { }

                    Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Speichere Word-Dokument …"));
                    wordDoc.SaveAs2(zielPfad, Word.WdSaveFormat.wdFormatXMLDocument);

                    string dateiName = IO.Path.GetFileName(zielPfad);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = $"Exportiert: {seiteNr} Seite(n) → {dateiName}";
                        if (MessageBox.Show(
                            $"Word-Dokument erstellt:\n{zielPfad}\n\nJetzt öffnen?",
                            "Export fertig", MessageBoxButton.YesNo,
                            MessageBoxImage.Information) == MessageBoxResult.Yes)
                        {
                            try { System.Diagnostics.Process.Start(
                                new System.Diagnostics.ProcessStartInfo(zielPfad)
                                { UseShellExecute = true }); }
                            catch { }
                        }
                    }));
                }
                catch (Exception ex)
                {
                    LogException(ex, "Export/Word");
                    var inner = App.EntpackeTIE(ex);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        TxtInfo.Text = "Export fehlgeschlagen. Siehe Log.";
                        MessageBox.Show(
                            $"Word-Export Fehler:\n[{inner.GetType().Name}] {inner.Message}\n\n[Log: {App.LogDatei}]",
                            "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                    }));
                }
                finally
                {
                    try { wordDoc?.Close(false); }  catch { }
                    try { wordApp?.Quit(); }          catch { }
                    if (wordDoc != null) try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc); } catch { }
                    if (wordApp != null) try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
            finally
            {
                App.LogFehler("Export/Timing", $"Gesamt: {sw.ElapsedMilliseconds} ms");
                foreach (var f in alleTemp) try { IO.File.Delete(f); } catch { }

                // Export-Sperre immer zurücksetzen (auch im Fehlerfall)
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _exportLäuft = false;
                    BtnExport.IsEnabled = true;
                }));
            }
        }
        /// <summary>
        /// Liest die nativen Seitenmaße der ersten PDF-Seite in Points (72 pt = 1 inch).
        /// Wird für die exakte 300-DPI-Render-Berechnung verwendet.
        /// Fallback: A4 (595 × 842 pt).
        /// </summary>
        private static (double widthPts, double heightPts) HolePdfSeitenGrösse(string pfad)
        {
            try
            {
                using var doc = PdfReader.Open(pfad, PdfDocumentOpenMode.ReadOnly);
                if (doc.PageCount > 0)
                {
                    var page = doc.Pages[0];
                    double w = page.Width.Point;
                    double h = page.Height.Point;
                    if (w > 0 && h > 0) return (w, h);
                }
            }
            catch (Exception ex) { LogException(ex, "HolePdfSeitenGrösse"); }
            return (595.0, 842.0); // A4-Fallback
        }

        /// <summary>
        /// Liest den tatsächlich verfügbaren Schreibbereich des Word-Dokuments
        /// aus den vorhandenen Seiteneinstellungen (Ränder bleiben unverändert).
        /// Rückgabe in Points (72 pt = 1 inch).
        /// Fallback: A4 mit 2,5 cm Rand ≈ 451 × 694 pt.
        /// </summary>
        private static (double BreiteP, double HöheP) HoleSchreibbereich(Word.Document doc)
        {
            try
            {
                var ps = doc.PageSetup;
                double w = ps.PageWidth  - ps.LeftMargin - ps.RightMargin;
                double h = ps.PageHeight - ps.TopMargin  - ps.BottomMargin;
                return (Math.Max(10.0, w), Math.Max(10.0, h));
            }
            catch (Exception ex)
            {
                LogException(ex, "HoleSchreibbereich");
                return (451.0, 694.0); // A4, 2,5 cm Rand
            }
        }

        /// <summary>
        /// Setzt das Word-Seitenformat (Größe, Ausrichtung, Ränder) passend zur PDF-Seite.
        /// Ränder werden auf 0 gesetzt, damit das Bild die volle Seite füllt.
        /// </summary>
        private static void SetzeWordSeitenFormat(Word.Document doc, double pdfW_pts, double pdfH_pts)
        {
            try
            {
                var ps = doc.PageSetup;

                // Orientation ZUERST setzen: Word begrenzt PageWidth im Hochformat auf
                // ≤ aktuelle PageHeight. Bei Querformat-PDFs (z.B. A3: pdfW=1190 pt)
                // würde PageWidth ohne diesen Schritt auf ~842 pt gekürzt → falsches Format.
                // Nach dem Orientation-Swap werden die exakten PDF-Maße gesetzt.
                ps.Orientation = pdfW_pts > pdfH_pts
                    ? Word.WdOrientation.wdOrientLandscape
                    : Word.WdOrientation.wdOrientPortrait;

                ps.PageWidth  = (float)pdfW_pts;
                ps.PageHeight = (float)pdfH_pts;

                // Minimale Ränder (1 pt ≈ 0,35 mm) statt 0 – verhindert COM-Ausnahmen
                // bei Druckertreibern die keine 0-Ränder akzeptieren
                const float RandPt = 1f;
                ps.TopMargin    = RandPt;
                ps.BottomMargin = RandPt;
                ps.LeftMargin   = RandPt;
                ps.RightMargin  = RandPt;

                App.LogFehler("Export/SeitenFormat",
                    $"Word-Seite: {pdfW_pts:F1} × {pdfH_pts:F1} pt " +
                    $"({(pdfW_pts > pdfH_pts ? "Quer" : "Hoch")})");
            }
            catch (Exception ex) { LogException(ex, "SetzeWordSeitenFormat"); }
        }

        /// <summary>
        /// Setzt den Normal-Stil des Dokuments auf 0 Abstand / einfachen Zeilenabstand.
        /// </summary>
        private static void NormiereAbsatzstil(Word.Document doc)
        {
            try
            {
                var normal = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                var pf     = normal.ParagraphFormat;
                pf.SpaceBefore         = 0f;
                pf.SpaceAfter          = 0f;
                pf.SpaceBeforeAuto     = 0;  // False: nicht automatisch
                pf.SpaceAfterAuto      = 0;
                pf.LineSpacingRule     = Word.WdLineSpacing.wdLineSpaceSingle;
                pf.LeftIndent          = 0f;
                pf.RightIndent         = 0f;
                pf.FirstLineIndent     = 0f;
            }
            catch (Exception ex) { LogException(ex, "NormiereAbsatzstil"); }
        }

        /// <summary>
        /// Fügt ein Bildsegment in das Word-Dokument ein.
        ///
        /// Geometrie:  pts = px / ExportDpi * 72  (× skalierung)
        ///
        /// seitenumbruchVorher = true:
        ///   Vor diesem Segment wird ein neuer Absatz eingefügt und
        ///   PageBreakBefore = wdTrue gesetzt, damit die neue PDF-Seite
        ///   sauber auf einer neuen Word-Seite beginnt.
        ///
        /// neuerAbsatz = true:
        ///   Folgesegment innerhalb derselben PDF-Seite — neuer Absatz,
        ///   kein Seitenumbruch.
        /// </summary>
        private static void EinfügenSegment(
            Word.Document doc,
            string bildPfad,
            float imgW_pt, float imgH_pt,
            bool neuerAbsatz,
            bool seitenumbruchVorher = false)
        {
            if (seitenumbruchVorher || neuerAbsatz)
            {
                Word.Range rNew = doc.Content;
                rNew.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                rNew.InsertParagraphAfter();
            }

            Word.Range r = doc.Content;
            r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            SetzeAbsatzFormat(r.ParagraphFormat);
            try { r.ParagraphFormat.PageBreakBefore = seitenumbruchVorher ? -1 : 0; } catch { }

            var shape = r.InlineShapes.AddPicture(bildPfad, false, true);
            shape.Width  = imgW_pt;
            shape.Height = imgH_pt;
        }

        private static void SetzeAbsatzFormat(Word.ParagraphFormat pf)
        {
            try
            {
                pf.SpaceBefore      = 0f;
                pf.SpaceAfter       = 0f;
                pf.SpaceBeforeAuto  = 0;
                pf.SpaceAfterAuto   = 0;
                pf.LineSpacingRule  = Word.WdLineSpacing.wdLineSpaceSingle;
                pf.LeftIndent       = 0f;
                pf.RightIndent      = 0f;
                pf.FirstLineIndent  = 0f;
            }
            catch { /* ParagraphFormat-Fehler nicht propagieren */ }
        }

        /// <summary>
        /// Letzte Sicherung: alle Absätze im fertigen Dokument nochmals bereinigen.
        /// Verhindert, dass Word Abstände nachträglich einfügt.
        /// </summary>
        private static void BereinigeDokumentAbsätze(Word.Document doc)
        {
            try
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                    SetzeAbsatzFormat(para.Format);
            }
            catch (Exception ex) { LogException(ex, "BereinigeDokumentAbsätze"); }
        }

        // ── Hochauflösung: Block zusammensetzen ──────────────────────────────

        private static BitmapSource? RendereBlockHochRes(
            List<BitmapSource> hochRes,
            double[] seitenYStart, double[] seitenHöhe,
            double yStart, double yEnd,
            double[]? cropLinks = null, double[]? cropRechts = null,
            double[]? cropOben = null, double[]? cropUnten = null)
        {
            if (hochRes == null || hochRes.Count == 0) return null;
            if (seitenYStart == null || seitenHöhe  == null) return null;
            if (yEnd <= yStart) return null;

            var teile = new List<(BitmapSource Bmp, int H)>();
            int maxW  = 0;
            int n     = Math.Min(hochRes.Count, Math.Min(seitenYStart.Length, seitenHöhe.Length));

            for (int i = 0; i < n; i++)
            {
                try
                {
                    if (hochRes[i] == null || seitenHöhe[i] <= 0) continue;

                    double sTop    = seitenYStart[i];
                    double sBottom = sTop + seitenHöhe[i];

                    // Per-Seite Crop-Werte (null-safe)
                    double cL = cropLinks?.Length  > i ? cropLinks![i]  : 0;
                    double cR = cropRechts?.Length > i ? cropRechts![i] : 0;
                    double cO = cropOben?.Length   > i ? cropOben![i]   : 0;
                    double cU = cropUnten?.Length  > i ? cropUnten![i]  : 0;

                    // Effektive Seitengrenzen nach Beschnitt (in Vorschau-Koordinaten)
                    double effTop    = sTop    + cO * seitenHöhe[i];
                    double effBottom = sBottom - cU * seitenHöhe[i];

                    double übTop = Math.Max(yStart, effTop);
                    double übBot = Math.Min(yEnd,   effBottom);
                    if (übBot <= übTop) continue;

                    double scY = hochRes[i].PixelHeight / seitenHöhe[i];
                    int pixT   = (int)Math.Round((übTop - sTop) * scY);  // ab Seiten-Top (inkl. Crop-Offset)
                    int pixH   = (int)Math.Round((übBot - übTop) * scY);
                    int w      = hochRes[i].PixelWidth;

                    // Horizontaler Beschnitt in Export-Pixeln
                    int cropL_px = (int)Math.Round(cL * w);
                    int cropR_px = (int)Math.Round(cR * w);
                    int cropW    = Math.Max(1, w - cropL_px - cropR_px);
                    if (cropL_px + cropW > w) cropW = w - cropL_px;
                    if (cropW <= 0) continue;

                    pixT = Math.Max(0, pixT);
                    int maxH = hochRes[i].PixelHeight - pixT;
                    if (maxH <= 0) continue;
                    pixH = Math.Min(maxH, Math.Max(1, pixH));
                    if (pixT + pixH > hochRes[i].PixelHeight) continue;

                    var crop = new CroppedBitmap(hochRes[i], new Int32Rect(cropL_px, pixT, cropW, pixH));
                    crop.Freeze();
                    teile.Add((crop, pixH));
                    if (cropW > maxW) maxW = cropW;
                }
                catch (Exception ex) { LogException(ex, $"RendereBlockHochRes/Seite[{i}]"); }
            }

            if (teile.Count == 0 || maxW == 0) return null;

            int totalH = teile.Sum(t => t.H);
            if (totalH <= 0) return null;

            try
            {
                var visual = new DrawingVisual();
                using (var ctx = visual.RenderOpen())
                {
                    double y = 0;
                    foreach (var (bmp, h) in teile)
                    {
                        ctx.DrawImage(bmp, new Rect(0, y, maxW, h));
                        y += h;
                    }
                }
                // 96 DPI: 1 WPF-Unit = 1 Pixel, damit Koordinaten direkt als Pixel gelten.
                var rtb = new RenderTargetBitmap(
                    maxW, totalH, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(visual);
                rtb.Freeze();
                return rtb;
            }
            catch (Exception ex) { LogException(ex, "RendereBlockHochRes/Render"); return null; }
        }

        // ── Layout-Modus wechseln ─────────────────────────────────────────────

        private void BtnLayoutWechsel_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0) return;
                _layoutHorizontal = !_layoutHorizontal;

                ZeicheCanvas();
            }, "BtnLayoutWechsel_Click");
    }
}
