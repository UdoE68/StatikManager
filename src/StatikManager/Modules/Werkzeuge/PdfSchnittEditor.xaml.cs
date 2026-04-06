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

        public PdfSchnittEditor()
        {
            InitializeComponent();
            InitAutoSave();
        }

        // ── Zustand ───────────────────────────────────────────────────────────
        private string?            _pdfPfad;
        private byte[]?            _pdfBytes;      // PDF als byte[] – verhindert Datei-Sperr-Probleme beim AutoSpeichern
        private bool               _hatUngespeicherteÄnderungen = false;
        // Gesetzt bei jeder Änderung; NICHT vom AutoSave gelöscht – nur bei neuem LadePdf oder
        // explizitem Speichern/Verwerfen. Steuert den "Änderungen speichern?"-Dialog beim Dateiwechsel.
        private bool               _hatSitzungsÄnderungen = false;
        private System.Windows.Threading.DispatcherTimer? _autoSaveTimer;
        private volatile bool      _autoSaveLäuft;
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

            // Gruppeneigene Randwerte (Bruchteil 0..0.49, relativ zu Seitenmaßen)
            public double CropLinks  { get; set; }
            public double CropRechts { get; set; }
            public double CropOben   { get; set; }
            public double CropUnten  { get; set; }
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

        // ── Scheren-Werkzeug ──────────────────────────────────────────────────
        private bool   _scherenModus = false;
        // (SeitenIdx, YFraction 0..1) – multiple Schnitte pro Seite möglich
        private readonly List<(int Seite, double YFraction)> _scherenschnitte
            = new List<(int, double)>();
        private Line?  _scherenVorschauLinie;

        // ── Seitenwechsel-Werkzeug ────────────────────────────────────────────
        private bool _seitenwechselModus = false;
        private Line? _seitenwechselVorschauLinie;

        // Schnittlinie verschieben (Verbesserung 4)
        private int    _gezogenesSchnittIdx  = -1;
        private bool   _schnittDragAktiv;
        private double _schnittDragOrigFrac;

        // ── Teil-Auswahl & Löschung ───────────────────────────────────────────
        // Ausgewählte (SeitenIdx, TeilIdx) – TeilIdx=0 ist der erste Teil von oben
        // NICHT persistiert: rein temporärer UI-Zustand pro Sitzung (kein Einfluss auf Export/Speichern).
        private readonly HashSet<(int Seite, int Teil)> _ausgewählteParts
            = new HashSet<(int, int)>();
        // Gelöschte (SeitenIdx, TeilIdx) – beim Export übersprungen
        private readonly HashSet<(int Seite, int Teil)> _gelöschteParts
            = new HashSet<(int, int)>();
        // Komposit-Bilder: Schlüssel = SeitenIdx, Wert = zusammengeschobenes Bitmap
        private readonly Dictionary<int, BitmapSource> _kompositBilder = new Dictionary<int, BitmapSource>();
        // Unified undo stack
        private readonly Stack<ScherenZustand> _undoStack = new Stack<ScherenZustand>();

        // Feature 3: Gelöschte ganze Seiten
        private readonly HashSet<int> _gelöschteSeiten = new HashSet<int>();

        // Features 4 & 5: Seitenreihenfolge (null = Identität [0,1,...,n-1])
        private List<int> _seitenReihenfolge;
        // Quellinformation für extern eingefügte Seiten: BitmapIdx → (Pfad, OrigSeitenIdx)
        private readonly Dictionary<int, (string Pfad, int Idx)> _eingefügteSeitenInfo
            = new Dictionary<int, (string, int)>();

        // Drag-State für Seiten-Drag & Drop (Feature 5)
        private int    _dragQuellIdx    = -1;
        private Point  _dragStartPunkt;
        private bool   _dragAktiv;
        private Border _dragGhost;

        private int _markierteSeitenIdx = -1; // Für Seite-einfügen vor/nach Auswahl

        private sealed class ScherenZustand
        {
            public List<(int Seite, double YFraction)> Schnitte { get; set; } = new List<(int, double)>();
            public HashSet<(int Seite, int Teil)> Gelöscht { get; set; } = new HashSet<(int, int)>();
            public Dictionary<int, BitmapSource> KompositBilder { get; set; } = new Dictionary<int, BitmapSource>();
            public HashSet<int> GelöschteSeiten  { get; set; } = new HashSet<int>();
            public List<int>    SeitenReihenfolge { get; set; }
        }

        // Abbruch laufender Lade-Aufträge
        private CancellationTokenSource? _ladeCts;
        // Abbruch laufender Auto-Rand-Berechnung
        private CancellationTokenSource? _autoRandCts;
        // Generationszähler: verhindert, dass veraltete Dispatcher-Callbacks UI überschreiben
        private volatile int _ladeGeneration;

        // Sitzungszustand der nach dem nächsten PDF-Laden angewendet wird (null = kein Restore)
        private Core.SitzungsZustand? _pendingSitzung;

        // ── Reflow-Primärdaten ────────────────────────────────────────────────
        /// <summary>
        /// Führendes Inhaltsmodell für den Reflow-Renderer.
        /// Wird nach dem Laden einer PDF (nach LadeSchnittState) aus dem Altmodell konvertiert.
        /// null solange noch keine PDF geladen wurde.
        /// </summary>
        private List<ContentBlock> _contentBlocks;

        /// <summary>Monoton steigender Zähler für BlockIds — wird nach jeder Konvertierung und nach Undo-Sync neu gesetzt.</summary>
        private int _nextBlockId = 0;

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
                                         Id         = g.Id,
                                         Name       = g.Name,
                                         Seiten     = g.Seiten.ToArray(),
                                         CropLinks  = g.CropLinks,
                                         CropRechts = g.CropRechts,
                                         CropOben   = g.CropOben,
                                         CropUnten  = g.CropUnten,
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

            // Speicher-Dialog: ungespeicherte Änderungen prüfen
            if (!FrageObSpeichern()) return;

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

            // Bearbeitungs- und Scheren-Modus vor dem Laden zurücksetzen (Bug 4: Ansicht-Reste beim Wechsel)
            if (_scherenModus) BeendeScherenModus();
            if (_seitenwechselModus) BeendeSeitenwechselModus();
            if (_bearbeitungsModus) BeendeBearbeitungsModus(false);

            _pdfPfad  = pfad;
            _pdfBytes = null;
            _hatUngespeicherteÄnderungen = false;
            _hatSitzungsÄnderungen = false;
            PdfCanvas.Children.Clear();
            _seitenBilder.Clear();
            _seitenYStart = Array.Empty<double>();
            _seitenHöhe   = Array.Empty<double>();
            _cropLinks = _cropRechts = _cropOben = _cropUnten = Array.Empty<double>();
            _scherenschnitte.Clear();
            _scherenVorschauLinie = null;
            _seitenwechselVorschauLinie = null;
            _ausgewählteParts.Clear();
            _gelöschteParts.Clear();
            _kompositBilder.Clear();
            _undoStack.Clear();
            _gelöschteSeiten.Clear();
            _markierteSeitenIdx = -1;
            _seitenReihenfolge = null;
            _eingefügteSeitenInfo.Clear();
            _dragAktiv = false;
            _dragQuellIdx = -1;
            _dragGhost = null;
            _gezogenesSchnittIdx = -1;
            _schnittDragAktiv    = false;
            TxtInfo.Text              = "Lade PDF …";
            BtnExport.IsEnabled       = false;
            BtnAuswahlmodus.IsEnabled = false;
            BtnSchereToggle.IsEnabled         = false;
            BtnSeitenwechsel.IsEnabled        = false;
            BtnSchnittZurücksetzen.IsEnabled  = false;
            BtnTeileExportieren.IsEnabled     = false;
            BtnTeilLöschen.IsEnabled          = false;

            System.Diagnostics.Debug.WriteLine($"[LadePdf] Starte PDF-Laden: {IO.Path.GetFileName(pfad)}");

            // Lade-Entscheidung: Original vs. _bearbeitet.pdf
            // Wenn _bearbeitet.pdf existiert → immer diese laden. Sie enthält den physisch
            //   korrekten Zustand (Komposit-Bitmaps, zusammengeschobene Teile etc.).
            //   LadeSchnittState() lädt danach nur _scherenschnitte für Interaktivität,
            //   aber NICHT _gelöschteParts/_gelöschteSeiten (schon eingearbeitet in _bearbeitet.pdf).
            // Wenn kein _bearbeitet.pdf aber JSON vorhanden → Original laden + vollen State anwenden.
            string bearbeitetPfad    = BearbeitetPfadFür(pfad);
            bool   bearbeitetVorhanden = IO.File.Exists(bearbeitetPfad);
            bool   jsonVorhanden     = IO.File.Exists(pfad + ".edit.json");
            string pfadKopie         = bearbeitetVorhanden ? bearbeitetPfad : pfad;
            if (pfadKopie != pfad)
                System.Diagnostics.Debug.WriteLine($"[LADEN] Bearbeitete Version geladen: {IO.Path.GetFileName(bearbeitetPfad)}");
            else if (jsonVorhanden)
                System.Diagnostics.Debug.WriteLine($"[LADEN] Kein _bearbeitet.pdf → Original + JSON-State: {IO.Path.GetFileName(pfad)}");

            var ladeThread = new Thread(() =>
            {
                List<BitmapSource>? bilder = null;
                double[]? yStart = null, höhe = null;
                string? fehler = null;
                byte[]? pdfBytesLoaded = null;

                // Serialisierung: kein paralleler pdfium-Zugriff mit Word-Vorschau-Threads.
                try { AppZustand.RenderSem.Wait(token); }
                catch (OperationCanceledException) { return; }
                try
                {
                try
                {
                    if (token.IsCancellationRequested) return;
                    // PDF in byte[] laden – pdfium hält damit keine Datei-Handle offen
                    pdfBytesLoaded = IO.File.ReadAllBytes(pfadKopie);
                    bilder = PdfRenderer.RenderiereAlleSeiten(pdfBytesLoaded, RenderBreite, token: token);
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
                        _pdfBytes     = pdfBytesLoaded;   // byte[]-Puffer für spätere Operationen
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
                                    Id         = gs.Id,
                                    Name       = string.IsNullOrWhiteSpace(gs.Name) ? $"Gruppe {gs.Id}" : gs.Name,
                                    Seiten     = gs.Seiten?.Where(s => s >= 0 && s < n).ToList() ?? new List<int>(),
                                    CropLinks  = gs.CropLinks,
                                    CropRechts = gs.CropRechts,
                                    CropOben   = gs.CropOben,
                                    CropUnten  = gs.CropUnten,
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

                        // SchnittState aus JSON wiederherstellen — NACH Seitenaufbau, VOR ZeicheCanvas
                        // Wenn _bearbeitet.pdf geladen wurde: nur Schnittlinien laden (für Interaktivität),
                        // aber NICHT _gelöschteParts/_gelöschteSeiten — diese sind bereits physisch
                        // in der _bearbeitet.pdf eingearbeitet (Komposit-Bitmaps, Lücken geschlossen etc.).
                        LadeSchnittState(nurSchnittlinien: bearbeitetVorhanden);

                        // Reflow-Primärdaten initialisieren — NACH LadeSchnittState, damit Schnitte + gelöschte
                        // Teile bereits bekannt sind und korrekt in ContentBlocks übertragen werden.
                        _contentBlocks = KonvertiereAltesModellZuBlöcken();
                        _nextBlockId   = _contentBlocks.Count > 0 ? _contentBlocks.Max(b => b.BlockId) + 1 : 0;

                        ZeicheCanvas();
                        // Pixel-pro-mm fuer Sicherheitsabstand-Konvertierung — via _pdfBytes (kein Datei-Handle)
                        if (bilder!.Count > 0)
                        {
                            var (wPts, _) = _pdfBytes != null
                                ? HolePdfSeitenGrösse(_pdfBytes)
                                : (_pdfPfad != null ? HolePdfSeitenGrösse(_pdfPfad) : (595.0, 842.0));
                            _pxPerMm = wPts > 0 ? bilder![0].PixelWidth / (wPts / 72.0 * 25.4) : 4.0;
                        }
                        BtnExport.IsEnabled               = true;
                        BtnAuswahlmodus.IsEnabled         = true;
                        BtnSchereToggle.IsEnabled         = true;
                        BtnSeitenwechsel.IsEnabled        = true;
                        // SchnittZurücksetzen nur aktiv wenn Schnitte vorhanden (auch nach JSON-Restore)
                        BtnSchnittZurücksetzen.IsEnabled  = _scherenschnitte.Count > 0;
                        string ladeInfo = $"{bilder!.Count} Seite(n) geladen";
                        if (_scherenschnitte.Count > 0 || _gelöschteParts.Count > 0 || _gelöschteSeiten.Count > 0)
                            ladeInfo += $" – {_scherenschnitte.Count} Schnitt(e), {_gelöschteSeiten.Count} gelöschte Seite(n) wiederhergestellt";
                        TxtInfo.Text = ladeInfo;

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

            string? vorlagePfad = WähleVorlage();

            // ── DIAGNOSE (testweise) ─────────────────────────────────────────────
            if (vorlagePfad != null)
            {
                TxtInfo.Text = $"Vorlage erkannt: {IO.Path.GetFileName(vorlagePfad)}";
                MessageBox.Show("Vorlage erkannt:\n" + vorlagePfad,
                    "DIAGNOSE – Vorlage", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                TxtInfo.Text = "Keine Vorlage konfiguriert – Standard-Export";
                MessageBox.Show(
                    "Keine gültige Vorlage gefunden.\n\n" +
                    "Konfigurierte Vorlagen in den Einstellungen:\n" +
                    string.Join("\n", Core.Einstellungen.Instanz.WordVorlagen
                        .Select(v => $"  {v.Name}: {v.Pfad} (existiert={IO.File.Exists(v.Pfad ?? "")})")),
                    "DIAGNOSE – Keine Vorlage", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            // ── DIAGNOSE ENDE ────────────────────────────────────────────────────

            var yStartK  = (double[])_seitenYStart.Clone();
            var höheK    = (double[])_seitenHöhe.Clone();
            string pdfK  = _pdfPfad;
            var cropLK = (double[])_cropLinks.Clone();
            var cropRK = (double[])_cropRechts.Clone();
            var cropOK = (double[])_cropOben.Clone();
            var cropUK = (double[])_cropUnten.Clone();

            var thread = new Thread(() =>
                ExportThreadWorker(zielPfad, pdfK, yStartK, höheK, cropLK, cropRK, cropOK, cropUK, vorlagePfad))
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
            // Gruppe und Crop-Modus während Bearbeitung sperren
            // (Löschen bleibt verfügbar: Gruppe 0 nicht löschbar, alle anderen schon)
            CmbGruppe.IsEnabled        = false;
            BtnGruppeNeu.IsEnabled     = false;
            BtnGruppeLöschen.IsEnabled = AktiveGruppe()?.Id != 0;
            CmbCropModus.IsEnabled     = false;
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

            // Gruppe und Crop-Modus wieder freigeben
            CmbGruppe.IsEnabled        = true;
            BtnGruppeNeu.IsEnabled     = true;
            BtnGruppeLöschen.IsEnabled = _gruppen.Count > 1 && AktiveGruppe()?.Id != 0;
            CmbCropModus.IsEnabled     = true;

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
            if (_reflowDebugModus) { ZeicheCanvasReflow(); return; }
            try
            {
                PdfCanvas.Children.Clear();
                if (_seitenBilder.Count == 0) return;

                // Effektive Bitmaps: Komposit wenn vorhanden, sonst Original
                var effBilder = Enumerable.Range(0, _seitenBilder.Count)
                    .Select(i => (BitmapSource)(_kompositBilder.TryGetValue(i, out var k) ? k : _seitenBilder[i]))
                    .ToList();

                // Reihenfolge: _seitenReihenfolge falls aktiv, sonst Identität
                var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
                // Sichtbare Seiten (nicht gelöscht), in Anzeigereihenfolge
                var sichtbar = reihenfolge.Where(i => i < _seitenBilder.Count && !_gelöschteSeiten.Contains(i)).ToList();

                // Layout-Arrays initialisieren
                _seitenYStart = new double[_seitenBilder.Count];
                _seitenHöhe   = new double[_seitenBilder.Count];
                _seitenXStart = new double[_seitenBilder.Count];
                for (int i = 0; i < _seitenBilder.Count; i++) { _seitenYStart[i] = -9999; _seitenHöhe[i] = 0; }

                if (sichtbar.Count == 0) return;

                var sichtbarBmps = sichtbar.Select(i => effBilder[i]).ToList();

                if (_layoutHorizontal)
                {
                    BerechneLayoutHorizontalStatic(sichtbarBmps, SeitenAbstand, out var ordXStart);
                    BerechneLayoutStatic(sichtbarBmps, SeitenAbstand, out var ordYStart, out var ordHöhe);
                    for (int di = 0; di < sichtbar.Count; di++)
                    {
                        int oi = sichtbar[di];
                        _seitenYStart[oi] = ordYStart[di];
                        _seitenHöhe[oi]   = ordHöhe[di];
                        _seitenXStart[oi] = ordXStart[di];
                    }
                    int lastOi = sichtbar[sichtbar.Count - 1];
                    double gesamtW = _seitenXStart[lastOi] + effBilder[lastOi].PixelWidth + SeitenAbstand;
                    double maxH    = sichtbarBmps.Max(b => (double)b.PixelHeight);
                    PdfCanvas.Width  = Math.Max(gesamtW, 1);
                    PdfCanvas.Height = Math.Max(maxH + SeiteX * 2, 1);
                }
                else
                {
                    BerechneLayoutStatic(sichtbarBmps, SeitenAbstand, out var ordYStart, out var ordHöhe);
                    for (int di = 0; di < sichtbar.Count; di++)
                    {
                        int oi = sichtbar[di];
                        _seitenYStart[oi] = ordYStart[di];
                        _seitenHöhe[oi]   = ordHöhe[di];
                    }
                    int lastOi = sichtbar[sichtbar.Count - 1];
                    double gesamtH = _seitenYStart[lastOi] + _seitenHöhe[lastOi] + SeitenAbstand;
                    double maxBmpW = sichtbarBmps.Max(b => (double)b.PixelWidth);
                    PdfCanvas.Width  = maxBmpW + SeiteX * 2;
                    PdfCanvas.Height = Math.Max(gesamtH, 1);
                }

                foreach (int i in sichtbar)
                    SafeExecute(() => ZeicheSeite(i), $"ZeicheSeite[{i}]");

                ZeicheCropLinien();
                _scherenVorschauLinie      = null;  // Canvas.Clear() hat Linie entfernt; Pointer resetten
                _seitenwechselVorschauLinie = null;  // Canvas.Clear() hat Linie entfernt; Pointer resetten
                AktualisiereSchnitteLinien();
                AktualisiereAuswahlAnzeige();
                // Markierte Seite mit blauem Rahmen hervorheben
                if (_markierteSeitenIdx >= 0)
                {
                    var seiteEl = PdfCanvas.Children.OfType<Border>()
                        .FirstOrDefault(b => b.Tag?.ToString() == $"SEITE_{_markierteSeitenIdx}");
                    if (seiteEl != null)
                    {
                        seiteEl.BorderBrush     = new SolidColorBrush(Color.FromRgb(0, 100, 200));
                        seiteEl.BorderThickness = new Thickness(3);
                    }
                }
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
            System.Diagnostics.Debug.WriteLine($"[RENDER] Seite {i}, Blocks: {_contentBlocks?.Count}");
            // Wenn diese Seite physisch geschnitten wurde → Blöcke einzeln zeichnen
            if (_contentBlocks != null)
            {
                var blöckeDeserSeite = _contentBlocks
                    .Where(b => b.SourcePageIdx == i)
                    .ToList();
                if (blöckeDeserSeite.Count > 1)
                {
                    ZeicheSeiteAlsBlöcke(i, blöckeDeserSeite);
                    return;
                }
            }

            var displayBmp = _kompositBilder.TryGetValue(i, out var kb) ? kb : _seitenBilder[i];
            if (displayBmp == null) return;
            // Tatsächliche Bitmap-Abmessungen verwenden – kein Strecken auf RenderBreite
            double bmpW = displayBmp.PixelWidth;
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
                    Source  = displayBmp,
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

            // Context menu: Seite löschen / Reihenfolge
            var seitenMenu = new ContextMenu();
            var itemLöschSeite = new MenuItem { Header = "\U0001F5D1  Seite löschen" };
            int capturedIdx = seitenIdx;
            itemLöschSeite.Click += (_, __) => LöscheSeiteMitBestätigung(capturedIdx);
            seitenMenu.Items.Add(itemLöschSeite);
            if (_seitenReihenfolge != null || _seitenBilder.Count > 1)
            {
                seitenMenu.Items.Add(new Separator());
                var itemNachOben = new MenuItem { Header = "\u2191  Nach oben verschieben" };
                itemNachOben.Click += (_, __) => VerschiebeSeite(capturedIdx, -1);
                var itemNachUnten = new MenuItem { Header = "\u2193  Nach unten verschieben" };
                itemNachUnten.Click += (_, __) => VerschiebeSeite(capturedIdx, +1);
                seitenMenu.Items.Add(itemNachOben);
                seitenMenu.Items.Add(itemNachUnten);
                seitenMenu.Items.Add(new Separator());
                var itemEinfügen = new MenuItem { Header = "\u2795  Nach dieser Seite einfügen" };
                itemEinfügen.Click += (_, __) => FügeSeiteEinNach(capturedIdx);
                seitenMenu.Items.Add(itemEinfügen);
                seitenMenu.Items.Add(new Separator());
                var itemEinfügenVor = new MenuItem { Header = "\u2795  Vor dieser Seite einfügen" };
                itemEinfügenVor.Click += (_, __) =>
                {
                    var reihenfolge2 = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
                    int pos = reihenfolge2.IndexOf(capturedIdx);
                    // FügeSeiteEinDialog(p) ruft intern FügeInReihenfolgeEin(neuIdx, p+1) auf.
                    // Für "vor Seite an pos": FügeInReihenfolgeEin(neuIdx, pos) → übergib pos-1
                    FügeSeiteEinDialog(Math.Max(0, pos - 1));
                };
                seitenMenu.Items.Add(itemEinfügenVor);
            }
            blatt.ContextMenu = seitenMenu;

            // Drag & Drop: Seiten umsortieren (Feature 5)
            blatt.PreviewMouseLeftButtonDown += (_, ev) =>
            {
                if (_scherenModus || _seitenwechselModus || _bearbeitungsModus) return;
                // Drag-Vorbereitung nur setzen, wenn kein Schnittlinie-Drag aktiv ist
                if (_schnittDragAktiv) return;
                _dragQuellIdx    = capturedIdx;
                _dragStartPunkt  = ev.GetPosition(PdfCanvas);
                _dragAktiv       = false;
            };
            blatt.MouseMove += (s, ev) =>
            {
                if (_dragQuellIdx != capturedIdx || ev.LeftButton != MouseButtonState.Pressed
                    || _schnittDragAktiv) return;
                var pos = ev.GetPosition(PdfCanvas);
                double dx = pos.X - _dragStartPunkt.X;
                double dy = pos.Y - _dragStartPunkt.Y;
                if (!_dragAktiv)
                {
                    if (Math.Abs(dx) + Math.Abs(dy) < 12) return;
                    _dragAktiv = true;
                    ((UIElement)s).CaptureMouse();
                    StartDragGhost(capturedIdx, blatt.Width, blatt.Height, pos);
                }
                else { MoveDragGhost(pos); }
                ev.Handled = true;
            };
            blatt.MouseLeftButtonUp += (s, ev) =>
            {
                if (_dragAktiv && _dragQuellIdx == capturedIdx)
                {
                    ((UIElement)s).ReleaseMouseCapture();
                    EndDrag(ev.GetPosition(PdfCanvas));
                    ev.Handled = true;
                }
                _dragAktiv = false;
                _dragQuellIdx = -1;
            };
            blatt.LostMouseCapture += (_, __) =>
            {
                if (_dragGhost != null) { PdfCanvas.Children.Remove(_dragGhost); _dragGhost = null; }
                _dragAktiv = false;
            };

            double setX = _layoutHorizontal && i < _seitenXStart.Length ? _seitenXStart[i] : SeiteX;
            double setY = _layoutHorizontal ? SeiteX : _seitenYStart[i];
            Canvas.SetLeft(blatt, setX);
            Canvas.SetTop(blatt,  setY);
            PdfCanvas.Children.Add(blatt);
        }
        /// <summary>
        /// Zeichnet eine physisch geschnittene Seite als unabhängige ContentBlock-Elemente.
        /// Jeder Block wird als eigenes CroppedBitmap auf dem Canvas platziert.
        /// Wird nur aufgerufen wenn _contentBlocks > 1 Block für diese Seite enthält.
        /// </summary>
        private void ZeicheSeiteAlsBlöcke(int seitenIdx, List<ContentBlock> blöcke)
        {
            System.Diagnostics.Debug.WriteLine($"[RENDER] BLOCK-PFAD für Seite {seitenIdx}, Anzahl Blöcke: {blöcke.Count}");
            // Grundregel (MIGRATION-01): immer _seitenBilder — Fraktionen beziehen sich auf Original.
            // Ausnahme: _kompositBilder wenn vorhanden UND gleiche Pixelhöhe wie Original
            // (= von SchiebeTeileZusammen padded, Fraktionen im Komposit-Raum).
            var origBmp = _seitenBilder[seitenIdx];
            var sourceBmp = (_kompositBilder.TryGetValue(seitenIdx, out var kb)
                             && kb != null
                             && origBmp != null
                             && kb.PixelHeight == origBmp.PixelHeight)
                ? kb : origBmp;
            if (sourceBmp == null) return;

            int    bmpPixelW = sourceBmp.PixelWidth;
            int    bmpPixelH = sourceBmp.PixelHeight;
            double pageH     = _seitenHöhe[seitenIdx];
            double setX      = _layoutHorizontal && seitenIdx < _seitenXStart.Length
                                   ? _seitenXStart[seitenIdx] : SeiteX;
            double yBase     = _layoutHorizontal ? SeiteX : _seitenYStart[seitenIdx];

            bool   ersterBlock = true;
            double currentY    = yBase;   // gestapeltes Layout: kein frac-basierter Offset
            int    teilIdx     = 0;       // 0-basierter Index pro Seite — kompatibel mit _ausgewählteParts/(si,t)
            foreach (var block in blöcke)
            {
                int capturedTeilIdx = teilIdx;
                teilIdx++;  // VOR continue — damit gelöschte Blöcke den Index trotzdem verbrauchen

                double fracO = Math.Max(0.0, Math.Min(1.0, block.FracOben));
                double fracU = Math.Max(fracO, Math.Min(1.0, block.FracUnten));
                if (fracU <= fracO) continue;

                double originalDisplayH = (fracU - fracO) * pageH;

                // Gelöschter Block → Lücken-Platzhalter mit konfigurierter Höhe
                if (block.IsDeleted)
                {
                    double gapH = BerechneGapHöhe(block, sourceBmp.DpiY, originalDisplayH);
                    if (gapH > 0)
                    {
                        double blockY      = currentY;
                        currentY          += gapH;
                        int capturedBlockId = block.BlockId;

                        var placeholder = new Border
                        {
                            Width           = bmpPixelW,
                            Height          = gapH,
                            Background      = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8)),
                            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0)),
                            BorderThickness = new Thickness(1),
                            Child           = new TextBlock
                            {
                                Text                = block.GapArt == GapModus.KundenAbstand
                                                        ? $"↕  {block.GapMm:F1} mm"
                                                        : "↕  Originalgröße",
                                Foreground          = new SolidColorBrush(Color.FromRgb(0xA0, 0xA0, 0xA0)),
                                FontStyle           = FontStyles.Italic,
                                FontSize            = 11,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment   = VerticalAlignment.Center
                            }
                        };

                        // Rechtsklick → Abstand nachbearbeiten
                        var cm = new ContextMenu();
                        var mi = new MenuItem { Header = "↕  Abstand bearbeiten …" };
                        mi.Click += (_, __) => BearbeiteBlockGap(capturedBlockId);
                        cm.Items.Add(mi);
                        placeholder.ContextMenu = cm;

                        Canvas.SetLeft(placeholder, setX);
                        Canvas.SetTop(placeholder,  blockY);
                        PdfCanvas.Children.Add(placeholder);
                    }
                    // gapH == 0 (KeinAbstand): kein Platzhalter, currentY nicht erhöhen
                    continue;
                }

                double blockDisplayH = Math.Max(1, originalDisplayH);
                double blockY2       = currentY;
                currentY            += blockDisplayH;

                int pixelY = (int)Math.Round(fracO * bmpPixelH);
                int pixelH = (int)Math.Round((fracU - fracO) * bmpPixelH);
                pixelY = Math.Max(0, Math.Min(pixelY, bmpPixelH - 1));
                pixelH = Math.Max(1, Math.Min(pixelH, bmpPixelH - pixelY));
                if (pixelH <= 0) continue;

                BitmapSource croppedBmp;
                try
                {
                    croppedBmp = new CroppedBitmap(
                        sourceBmp,
                        new Int32Rect(0, pixelY, bmpPixelW, pixelH));
                }
                catch (Exception ex)
                {
                    LogException(ex, $"ZeicheSeiteAlsBlöcke CroppedBitmap si={seitenIdx} B{block.BlockId}");
                    continue;
                }

                DropShadowEffect? shadow = null;
                try
                {
                    shadow = new DropShadowEffect
                    {
                        BlurRadius = 20, ShadowDepth = 7,
                        Direction  = 280, Color = Colors.Black, Opacity = 0.85
                    };
                }
                catch { /* ohne Schatten weiterzeichnen */ }

                // Erster Block bekommt Tag SEITE_{i} für Kompatibilität mit Highlight-Code
                string tag = ersterBlock
                    ? $"SEITE_{seitenIdx}"
                    : $"SEITE_{seitenIdx}_BLK_{block.BlockId}";
                ersterBlock = false;

                bool isSelected = _ausgewählteParts.Contains((seitenIdx, capturedTeilIdx));

                var blatt = new Border
                {
                    Tag                 = tag,
                    Width               = bmpPixelW,
                    Height              = blockDisplayH,
                    Background          = Brushes.White,
                    BorderBrush         = isSelected
                                            ? new SolidColorBrush(Color.FromRgb(0, 100, 220))
                                            : new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                    BorderThickness     = new Thickness(isSelected ? 3 : 2),
                    Child               = new Image
                    {
                        Source              = croppedBmp,
                        Width               = bmpPixelW,
                        Height              = blockDisplayH,
                        Stretch             = Stretch.Uniform,
                        SnapsToDevicePixels = true
                    },
                    Effect              = shadow,
                    SnapsToDevicePixels = true
                };

                // Klick → Block markieren (kompatibel mit LöscheAusgewählteParts)
                blatt.MouseLeftButtonDown += (_, ev) =>
                {
                    if (_scherenModus) return;   // Schnitt-Modus: Event an Canvas durchlassen (Schere_MouseDown)
                    bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
                    if (!ctrl) _ausgewählteParts.Clear();
                    if (_ausgewählteParts.Contains((seitenIdx, capturedTeilIdx)))
                        _ausgewählteParts.Remove((seitenIdx, capturedTeilIdx));
                    else
                        _ausgewählteParts.Add((seitenIdx, capturedTeilIdx));
                    BtnTeilLöschen.IsEnabled = _ausgewählteParts.Count > 0;
                    TxtInfo.Text = _ausgewählteParts.Count > 0
                        ? $"{_ausgewählteParts.Count} Teil(e) markiert – Entf oder Rechtsklick zum Löschen"
                        : "";
                    ZeicheCanvas();
                    ev.Handled = true;
                };

                // Rechtsklick-Menü → Teil löschen
                var blockMenu = new ContextMenu();
                var itemLösch = new MenuItem { Header = "\u2715  Teil löschen" };
                itemLösch.Click += (_, __) =>
                {
                    if (!_ausgewählteParts.Contains((seitenIdx, capturedTeilIdx)))
                    {
                        _ausgewählteParts.Clear();
                        _ausgewählteParts.Add((seitenIdx, capturedTeilIdx));
                    }
                    LöscheAusgewählteParts();
                };
                blockMenu.Items.Add(itemLösch);
                blatt.ContextMenu = blockMenu;

                Canvas.SetLeft(blatt, setX);
                Canvas.SetTop(blatt,  blockY2);
                PdfCanvas.Children.Add(blatt);
            }
        }

        /// <summary>
        /// Berechnet die Anzeige-Höhe der Lücke in Pixeln für einen gelöschten Block.
        /// </summary>
        /// <param name="block">Der gelöschte ContentBlock.</param>
        /// <param name="sourceDpiY">DPI-Y der Quell-Bitmap (für mm→px Umrechnung).</param>
        /// <param name="originalDisplayH">Originale Anzeige-Höhe des Blocks in Pixeln (für GapModus.OriginalAbstand).</param>
        private static double BerechneGapHöhe(ContentBlock block, double sourceDpiY, double originalDisplayH)
        {
            switch (block.GapArt)
            {
                case GapModus.OriginalAbstand:
                    return originalDisplayH;
                case GapModus.KundenAbstand:
                    double dpi = sourceDpiY > 0 ? sourceDpiY : 96.0;
                    return block.GapMm * dpi / 25.4;
                case GapModus.KeinAbstand:
                default:
                    return 0.0;
            }
        }

        private void BearbeiteBlockGap(int blockId) { /* Task 6 */ }

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
                if (_reflowDebugModus) return;      // im Debug-Modus keine Crop-Interaktion
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

                MarkiereAlsGeändert();
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
            var gruppe = _gruppen[selIdx];
            _aktGruppeId = gruppe.Id;

            // Wenn Gruppe eigene Randwerte hat, sofort auf alle Seiten der Gruppe anwenden
            if (_cropLinks.Length > 0
                && (gruppe.CropLinks != 0 || gruppe.CropRechts != 0
                    || gruppe.CropOben != 0 || gruppe.CropUnten != 0))
            {
                foreach (int s in gruppe.Seiten)
                {
                    if (s < _cropLinks.Length)  _cropLinks[s]  = gruppe.CropLinks;
                    if (s < _cropRechts.Length) _cropRechts[s] = gruppe.CropRechts;
                    if (s < _cropOben.Length)   _cropOben[s]   = gruppe.CropOben;
                    if (s < _cropUnten.Length)  _cropUnten[s]  = gruppe.CropUnten;
                }
                AktualisiereCropLinien();
            }

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
                int pos  = _gruppen.Count;   // Gruppe 0 ist immer vorhanden; pos=1 → "Gruppe 1"
                var name = $"Gruppe {pos}";
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
                // Defensiv: _aktGruppeId aus ComboBox-Auswahl synchronisieren
                int selIdx = CmbGruppe?.SelectedIndex ?? -1;
                if (selIdx >= 0 && selIdx < _gruppen.Count)
                    _aktGruppeId = _gruppen[selIdx].Id;

                var aktGruppe = AktiveGruppe();
                if (aktGruppe == null) return;
                if (aktGruppe.Id == 0) return; // Gruppe 0 ist unverlöschbar

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

                // Aktive Gruppe ermitteln – bei Modus "Ausgewählt" Gruppenrandwerte bevorzugen
                var dialogGruppe = (_cropModus == CropAnwendungsModus.Ausgewählt) ? AktiveGruppe() : null;
                bool hatGruppenCrop = dialogGruppe != null
                    && (dialogGruppe.CropLinks != 0 || dialogGruppe.CropRechts != 0
                        || dialogGruppe.CropOben != 0 || dialogGruppe.CropUnten != 0);

                // Originalwerte: aus aktiver Gruppe (wenn vorhanden), sonst aus aktiver Seite
                double origOben   = hatGruppenCrop ? dialogGruppe!.CropOben   : (aktSeite < _cropOben.Length   ? _cropOben[aktSeite]   : 0);
                double origUnten  = hatGruppenCrop ? dialogGruppe!.CropUnten  : (aktSeite < _cropUnten.Length  ? _cropUnten[aktSeite]  : 0);
                double origLinks  = hatGruppenCrop ? dialogGruppe!.CropLinks  : (aktSeite < _cropLinks.Length  ? _cropLinks[aktSeite]  : 0);
                double origRechts = hatGruppenCrop ? dialogGruppe!.CropRechts : (aktSeite < _cropRechts.Length ? _cropRechts[aktSeite] : 0);

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

                // Bei "Ausgewählte Seiten": Werte auch in der Gruppe speichern (Bruchteil der Referenzseite)
                if (rbAuswahl.IsChecked == true && dialogGruppe != null)
                {
                    dialogGruppe.CropLinks  = Math.Min(newLinksPx  / refW, 0.49);
                    dialogGruppe.CropRechts = Math.Min(newRechtsPx / refW, 0.49);
                    dialogGruppe.CropOben   = Math.Min(newObenPx   / refH, 0.49);
                    dialogGruppe.CropUnten  = Math.Min(newUntenPx  / refH, 0.49);
                }

                MarkiereAlsGeändert();
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
                {
                    l.StrokeThickness = Math.Max(14.0, 14.0 / zoom);
                    l.StrokeDashArray = null;
                }
                else if (tag.StartsWith("CROP_"))
                {
                    l.StrokeThickness = 2.0 / zoom;
                    l.StrokeDashArray = null;
                }
                else if (tag.StartsWith("SCHERE_"))
                {
                    l.StrokeThickness = 2.0 / zoom;
                    // StrokeDashArray bleibt erhalten
                }
                else if (tag == "SEITENWECHSEL_PREVIEW")
                {
                    l.StrokeThickness = 2.0 / zoom;
                    // StrokeDashArray bleibt erhalten
                }
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
            double[] cropLinks, double[] cropRechts, double[] cropOben, double[] cropUnten,
            string? vorlagePfad = null)
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

                    Word.Range? einfügePunkt    = null;
                    bool        textmarkeGefunden = false;
                    if (vorlagePfad != null)
                    {
                        // ── Diagnose: Neuer Vorlagen-Pfad aktiv ────────────────
                        Dispatcher.Invoke(new Action(() => MessageBox.Show(
                            $"Neuer Vorlagen-Exportpfad aktiv.\n\nVorlage:\n{vorlagePfad}",
                            "Export-Diagnose", MessageBoxButton.OK, MessageBoxImage.Information)));

                        wordDoc = ÖffneVorlage(wordApp, vorlagePfad);
                        NormiereAbsatzstil(wordDoc);

                        // Bookmark-Status VOR dem Löschen prüfen
                        textmarkeGefunden = wordDoc.Bookmarks.Exists("BILD_BEREICH");
                        einfügePunkt = HoleEinfügePunkt(wordDoc);

                        // Status-Anzeige: Textmarke gefunden oder Fallback aktiv
                        string bmInfo = textmarkeGefunden
                            ? "Textmarke BILD_BEREICH gefunden"
                            : "Keine Textmarke – automatischer Seitenbereich verwendet";
                        Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = bmInfo));
                        App.LogFehler("Export/Textmarke", bmInfo);
                    }
                    else
                    {
                        wordDoc = wordApp.Documents.Add();
                        NormiereAbsatzstil(wordDoc);
                        SetzeWordSeitenFormat(wordDoc, nativeW_pts, nativeH_pts);
                    }

                    var (availW_pt, availH_pt) = HoleSchreibbereich(wordDoc);

                    // Für Template: nutzbaren Bereich aus Tabellenstruktur ermitteln
                    double? zellenB = null, zellenH = null;
                    if (vorlagePfad != null && einfügePunkt != null)
                    {
                        (zellenB, zellenH) = HoleZellenMasse(einfügePunkt);
                        if (zellenB.HasValue)
                        {
                            App.LogFehler("Export/Schreibbereich", $"Tabellenbreite: {zellenB.Value:F1} pt");
                            availW_pt = zellenB.Value;
                        }
                        if (zellenH.HasValue)
                        {
                            App.LogFehler("Export/Schreibbereich", $"Tabellenhöhe: {zellenH.Value:F1} pt");
                            availH_pt = zellenH.Value;
                        }
                    }
                    App.LogFehler("Export/Schreibbereich",
                        $"Vorlage: {availW_pt:F1} × {availH_pt:F1} pt | " +
                        $"Druckbereich: {nativeW_eff:F1} × {nativeH_eff:F1} pt");

                    // ── DIAGNOSE: Zielbereich-Breakdown als MessageBox ───────────
                    {
                        var ps2 = wordDoc!.PageSetup;
                        double effUntererRand2    = Math.Max(ps2.BottomMargin, ps2.FooterDistance);
                        double footerOberkante2   = ps2.PageHeight - ps2.FooterDistance;
                        double satzspiegelH2      = ps2.PageHeight - ps2.TopMargin - effUntererRand2;
                        string quelleBotM         = ps2.FooterDistance > ps2.BottomMargin
                                                    ? $"FooterDistance ({ps2.FooterDistance:F1} pt)"
                                                    : $"BottomMargin ({ps2.BottomMargin:F1} pt)";
                        string quelleB = zellenB.HasValue ? "Tabellenbreite" : "Satzspiegel";
                        string quelleH = zellenH.HasValue ? "Zeilenhöhe fix" : "Satzspiegel";
                        Dispatcher.Invoke(new Action(() =>
                            MessageBox.Show(
                                $"── Seitenmaße ─────────────────────────\n" +
                                $"Seitenhöhe:         {ps2.PageHeight:F1} pt\n" +
                                $"TopMargin:          {ps2.TopMargin:F1} pt\n" +
                                $"BottomMargin:       {ps2.BottomMargin:F1} pt\n" +
                                $"HeaderDistance:     {ps2.HeaderDistance:F1} pt\n" +
                                $"FooterDistance:     {ps2.FooterDistance:F1} pt\n\n" +
                                $"── Footer-Analyse ──────────────────────\n" +
                                $"Footer-Oberkante (PageH−FooterDist): {footerOberkante2:F1} pt von oben\n" +
                                $"Eff. unterer Rand:  {effUntererRand2:F1} pt  (Quelle: {quelleBotM})\n" +
                                $"Satzspiegel-Höhe:   {satzspiegelH2:F1} pt\n\n" +
                                $"── Finaler Zielbereich ─────────────────\n" +
                                $"Breite: {availW_pt:F0} pt  (Quelle: {quelleB})\n" +
                                $"Höhe:   {availH_pt:F0} pt  (Quelle: {quelleH})\n\n" +
                                $"── Bild ────────────────────────────────\n" +
                                $"Breite: {nativeW_eff:F0} pt\n" +
                                $"Höhe:   {nativeH_eff:F0} pt",
                                "DIAGNOSE – Zielbereich",
                                MessageBoxButton.OK, MessageBoxImage.Information)));
                    }

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
                        // ── DIAGNOSE (testweise): Bild passt nicht ──────────────
                        Dispatcher.Invoke(new Action(() =>
                            MessageBox.Show(
                                $"Bild passt nicht.\nErforderliche Skalierung: {prozent} %\nSkalierungsdialog folgt.",
                                "DIAGNOSE – Bild zu groß", MessageBoxButton.OK, MessageBoxImage.Information)));

                        SkalierungWahl wahl = SkalierungWahl.Abbrechen;
                        Dispatcher.Invoke(new Action(() =>
                        {
                            var dlg = new SkalierungDialog(
                                prozent,
                                bildB_pt: nativeW_eff, bildH_pt: nativeH_eff,
                                zielB_pt: availW_pt,   zielH_pt: availH_pt)
                            {
                                Owner = Application.Current?.MainWindow
                            };
                            dlg.ShowDialog();
                            wahl = dlg.Wahl;
                        }));

                        if (wahl == SkalierungWahl.Abbrechen)
                        {
                            Dispatcher.BeginInvoke(new Action(() => TxtInfo.Text = "Export abgebrochen."));
                            return;
                        }
                        doScale = (wahl == SkalierungWahl.Verkleinern);
                    }
                    else
                    {
                        globalScale = 1.0;
                        // ── Diagnose: Bild passt ────────────────────────────────
                        Dispatcher.BeginInvoke(new Action(() =>
                            TxtInfo.Text =
                                $"Bild passt vollständig  •  Zielbereich: {availW_pt:F0} × {availH_pt:F0} pt  •  Bild: {nativeW_eff:F0} × {nativeH_eff:F0} pt"));
                    }

                    App.LogFehler("Export/Skalierung",
                        $"globalScale={globalScale:F4} ({(int)Math.Round(globalScale * 100)} %) | doScale={doScale}");

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
                            // Originalgröße: linksbündig, rechter/unterer Überstand bleibt
                            finalBmp  = segBmp;
                            finalW_pt = (float)segW_pt;
                            finalH_pt = (float)segH_pt;
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
                        Word.Range? thisInsertAt = (gi == 0) ? einfügePunkt : null;
                        try { EinfügenSegment(wordDoc, png, finalW_pt, finalH_pt,
                                neuerAbsatz: false, seitenumbruchVorher: seitenumbruch,
                                insertBei: thisInsertAt); }
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
                byte[] bytes = IO.File.ReadAllBytes(pfad);
                return HolePdfSeitenGrösse(bytes);
            }
            catch (Exception ex) { LogException(ex, "HolePdfSeitenGrösse(pfad)"); }
            return (595.0, 842.0); // A4-Fallback
        }

        private static (double widthPts, double heightPts) HolePdfSeitenGrösse(byte[] bytes)
        {
            try
            {
                using var ms  = new IO.MemoryStream(bytes, writable: false);
                using var doc = PdfReader.Open(ms, PdfDocumentOpenMode.ReadOnly);
                if (doc.PageCount > 0)
                {
                    var page = doc.Pages[0];
                    double w = page.Width.Point;
                    double h = page.Height.Point;
                    if (w > 0 && h > 0) return (w, h);
                }
            }
            catch (Exception ex) { LogException(ex, "HolePdfSeitenGrösse(bytes)"); }
            return (595.0, 842.0); // A4-Fallback
        }

        /// <summary>
        /// Liest den tatsächlich verfügbaren Schreibbereich des Word-Dokuments.
        ///
        /// Breite = PageWidth − LeftMargin − RightMargin
        ///
        /// Höhe: untere Nutzgrenze = PageHeight − max(BottomMargin, FooterDistance)
        ///   • BottomMargin  = Abstand Seitenende → Ende Körperbereich
        ///   • FooterDistance = Abstand Seitenende → Oberkante Fußzeile
        ///   → Wenn FooterDistance > BottomMargin, ragt die Fußzeile in den Body:
        ///     FooterDistance wird dann als harte Untergrenze verwendet.
        ///   → Wenn FooterDistance ≤ BottomMargin, liegt Fußzeile innerhalb des Rands:
        ///     BottomMargin bleibt maßgeblich (bestehende Formel).
        ///
        /// Rückgabe in Points (72 pt = 1 inch).
        /// Fallback: A4 mit 2,5 cm Rand ≈ 451 × 694 pt.
        /// </summary>
        private static (double BreiteP, double HöheP) HoleSchreibbereich(Word.Document doc)
        {
            try
            {
                var ps = doc.PageSetup;
                double w = ps.PageWidth - ps.LeftMargin - ps.RightMargin;

                // Untere Nutzgrenze: Fußzeilen-Oberkante als harte Grenze wenn sie über
                // die BottomMargin hinausragt (FooterDistance = Abstand Seitenende → Footer-Oberkante)
                double effUntererRand = Math.Max(ps.BottomMargin, ps.FooterDistance);
                double footerOberkanteVonOben = ps.PageHeight - ps.FooterDistance;
                double h = ps.PageHeight - ps.TopMargin - effUntererRand;

                App.LogFehler("HoleSchreibbereich",
                    $"PageH={ps.PageHeight:F1} | TopM={ps.TopMargin:F1} | " +
                    $"BotM={ps.BottomMargin:F1} | FooterDist={ps.FooterDistance:F1} | " +
                    $"FooterOberkanteVonOben={footerOberkanteVonOben:F1} | " +
                    $"EffUntererRand={effUntererRand:F1} (={( ps.FooterDistance > ps.BottomMargin ? "FooterDist" : "BottomMargin")}) | " +
                    $"W={w:F1} | H={h:F1}");

                return (Math.Max(10.0, w), Math.Max(10.0, h));
            }
            catch (Exception ex)
            {
                LogException(ex, "HoleSchreibbereich");
                return (451.0, 694.0); // A4, 2,5 cm Rand
            }
        }

        // ── Vorlagen-Hilfsmethoden ────────────────────────────────────────────

        /// <summary>
        /// Wählt die zu verwendende Word-Vorlage aus den Einstellungen.
        /// Gibt null zurück wenn keine Vorlage konfiguriert ist (Legacy-Modus).
        /// </summary>
        private string? WähleVorlage()
        {
            var vorlagen = Einstellungen.Instanz.WordVorlagen
                .Where(v => !string.IsNullOrWhiteSpace(v.Pfad) && IO.File.Exists(v.Pfad))
                .ToList();

            if (vorlagen.Count == 0) return null;
            if (vorlagen.Count == 1) return vorlagen[0].Pfad;

            var standard = vorlagen.FirstOrDefault(v => v.Standard);
            return standard?.Pfad ?? vorlagen[0].Pfad;
        }

        /// <summary>
        /// Öffnet ein neues Dokument auf Basis der angegebenen Word-Vorlage.
        /// </summary>
        private static Word.Document ÖffneVorlage(Word.Application app, string vorlagePfad)
        {
            return app.Documents.Add(Template: vorlagePfad);
        }

        /// <summary>
        /// Sucht die Textmarke "BILD_BEREICH" und gibt deren Range zurück (Inhalt wird gelöscht).
        /// Fallback (keine Textmarke): Anfang des Dokuments = oben links im Satzspiegel.
        /// Der Export bricht bei fehlender Textmarke NICHT ab.
        /// </summary>
        private static Word.Range HoleEinfügePunkt(Word.Document doc)
        {
            try
            {
                if (doc.Bookmarks.Exists("BILD_BEREICH"))
                {
                    var bm = doc.Bookmarks["BILD_BEREICH"];
                    var r  = bm.Range;
                    r.Delete();
                    App.LogFehler("Export/Textmarke", "Textmarke BILD_BEREICH gefunden und geleert");
                    return r;
                }
                App.LogFehler("Export/Textmarke",
                    "Textmarke BILD_BEREICH nicht gefunden – Fallback: Anfang des Satzspiegels");
            }
            catch (Exception ex) { LogException(ex, "HoleEinfügePunkt"); }

            // Kein Bookmark → oben links im Satzspiegel (Anfang des Body-Inhalts)
            var fallback = doc.Content;
            fallback.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            return fallback;
        }

        /// <summary>
        /// Gibt Breite und Höhe der Tabellenzelle zurück, in der sich die Range befindet.
        /// Gibt (null, null) zurück wenn die Range nicht in einer Tabelle liegt.
        ///
        /// Höhe: nur bei explizit gesetzter fester/Mindesthöhe (HeightRule ≠ Auto).
        /// Bei Auto-Höhe wird null zurückgegeben → Fallback auf HoleSchreibbereich.
        /// (wdVerticalPositionRelativeToPage ist bei Visible=false unzuverlässig.)
        ///
        /// Rückgabe in Points.
        /// </summary>
        private static (double? Breite, double? Höhe) HoleZellenMasse(Word.Range r)
        {
            try
            {
                var inTable = r.Information[Word.WdInformation.wdWithInTable];
                bool istInTabelle = inTable is bool b ? b : (inTable is int iv && iv != 0);
                if (istInTabelle)
                {
                    var cell  = r.Cells[1];
                    double?   breite = cell.Width > 10 ? cell.Width : (double?)null;
                    double?   höhe   = null;
                    try
                    {
                        var row = cell.Row;
                        if (row.HeightRule != Word.WdRowHeightRule.wdRowHeightAuto && row.Height > 10)
                        {
                            // Feste oder Mindesthöhe → direkt verwenden
                            höhe = row.Height;
                            App.LogFehler("Export/ZellenHöhe",
                                $"Feste Zeilenhöhe: {row.Height:F1} pt (Rule={row.HeightRule})");
                        }
                        else
                        {
                            // Auto-Höhe → null, Fallback auf Satzspiegel-Höhe (HoleSchreibbereich)
                            App.LogFehler("Export/ZellenHöhe",
                                "Auto-Zeilenhöhe → Höhe aus Satzspiegel (HoleSchreibbereich)");
                        }
                    }
                    catch { }
                    return (breite, höhe);
                }
            }
            catch { }
            return (null, null);
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
            bool seitenumbruchVorher = false,
            Word.Range? insertBei = null)
        {
            if (seitenumbruchVorher || neuerAbsatz)
            {
                Word.Range rNew = doc.Content;
                rNew.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                rNew.InsertParagraphAfter();
            }

            Word.Range r;
            if (insertBei != null)
            {
                r = insertBei;
            }
            else
            {
                r = doc.Content;
                r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

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

        // ── Scheren-Werkzeug ──────────────────────────────────────────────────

        private void BtnSchereToggle_Checked(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0) { BtnSchereToggle.IsChecked = false; return; }
                // Seitenwechsel-Modus deaktivieren wenn aktiv
                if (_seitenwechselModus) BeendeSeitenwechselModus();
                _scherenModus = true;
                _ausgewählteParts.Clear();
                ScrollView.Cursor                = Cursors.Cross;
                PdfCanvas.MouseMove              += Schere_MouseMove;
                PdfCanvas.MouseLeftButtonDown    += Schere_MouseDown;
                AktualisiereTeilOverlays();   // overlays auf IsHitTestVisible=false setzen
                ScrollView.Focus();
                TxtInfo.Text = "✂ Scheren aktiv – Klick = Schnitt | Strg+Z = rückgängig | Esc = beenden";
            }, "BtnSchereToggle_Checked");
        }

        private void BtnSchereToggle_Unchecked(object sender, RoutedEventArgs e)
            => SafeExecute(BeendeScherenModus, "BtnSchereToggle_Unchecked");

        private void BeendeScherenModus()
        {
            _scherenModus = false;
            ScrollView.Cursor                 = null;
            PdfCanvas.MouseMove              -= Schere_MouseMove;
            PdfCanvas.MouseLeftButtonDown    -= Schere_MouseDown;

            if (_scherenVorschauLinie != null)
            {
                PdfCanvas.Children.Remove(_scherenVorschauLinie);
                _scherenVorschauLinie = null;
            }

            // Toggle-Button zurücksetzen ohne erneutes Unchecked-Ereignis
            if (BtnSchereToggle.IsChecked == true) BtnSchereToggle.IsChecked = false;

            AktualisiereTeilOverlays();   // overlays wieder auf IsHitTestVisible=true
            int anzahl = _scherenschnitte.Count;
            TxtInfo.Text = anzahl > 0 ? $"{anzahl} Schnitt(e) – Klick zum Markieren" : "";
        }

        // ── Seitenwechsel-Werkzeug ────────────────────────────────────────────

        private void BtnSeitenwechsel_Checked(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0) { BtnSeitenwechsel.IsChecked = false; return; }
                // Scheren-Modus deaktivieren wenn aktiv
                if (_scherenModus) BeendeScherenModus();
                _seitenwechselModus = true;
                ScrollView.Cursor             = Cursors.Cross;
                PdfCanvas.MouseMove           += Seitenwechsel_MouseMove;
                PdfCanvas.MouseLeftButtonDown += Seitenwechsel_MouseDown;
                ScrollView.Focus();
                TxtInfo.Text = "\u23CE Seitenwechsel aktiv – Klick = Seite teilen | Esc = beenden";
            }, "BtnSeitenwechsel_Checked");
        }

        private void BtnSeitenwechsel_Unchecked(object sender, RoutedEventArgs e)
            => SafeExecute(BeendeSeitenwechselModus, "BtnSeitenwechsel_Unchecked");

        private void BeendeSeitenwechselModus()
        {
            _seitenwechselModus = false;
            ScrollView.Cursor             = null;
            PdfCanvas.MouseMove           -= Seitenwechsel_MouseMove;
            PdfCanvas.MouseLeftButtonDown -= Seitenwechsel_MouseDown;

            if (_seitenwechselVorschauLinie != null)
            {
                PdfCanvas.Children.Remove(_seitenwechselVorschauLinie);
                _seitenwechselVorschauLinie = null;
            }

            // Toggle-Button zurücksetzen ohne erneutes Unchecked-Ereignis
            if (BtnSeitenwechsel.IsChecked == true) BtnSeitenwechsel.IsChecked = false;

            TxtInfo.Text = "";
        }

        private void Seitenwechsel_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (!_seitenwechselModus || _reflowDebugModus) return;   // kein Seitenwechsel-Feedback im Debug-Modus
                Point pos = e.GetPosition(PdfCanvas);
                int si = HoleSeitenIndexBeiPos(pos.Y, pos.X);

                if (si < 0)
                {
                    if (_seitenwechselVorschauLinie != null)
                        _seitenwechselVorschauLinie.Visibility = Visibility.Collapsed;
                    return;
                }

                double pageW = _seitenBilder[si].PixelWidth;
                double setX  = _layoutHorizontal && si < _seitenXStart.Length ? _seitenXStart[si] : SeiteX;

                if (_seitenwechselVorschauLinie == null)
                {
                    _seitenwechselVorschauLinie = new Line
                    {
                        Stroke           = new SolidColorBrush(Color.FromRgb(0, 100, 204)),
                        StrokeThickness  = 2.0 / _zoomFaktor,
                        StrokeDashArray  = new DoubleCollection(new[] { 8.0, 4.0 }),
                        IsHitTestVisible = false,
                        Tag              = "SEITENWECHSEL_PREVIEW"
                    };
                    Panel.SetZIndex(_seitenwechselVorschauLinie, 200);
                    PdfCanvas.Children.Add(_seitenwechselVorschauLinie);
                }

                _seitenwechselVorschauLinie.X1             = setX;
                _seitenwechselVorschauLinie.X2             = setX + pageW;
                _seitenwechselVorschauLinie.Y1             = pos.Y;
                _seitenwechselVorschauLinie.Y2             = pos.Y;
                _seitenwechselVorschauLinie.StrokeThickness = 2.0 / _zoomFaktor;
                _seitenwechselVorschauLinie.Visibility     = Visibility.Visible;
            }
            catch (Exception ex) { LogException(ex, "Seitenwechsel_MouseMove"); }
        }

        private void Seitenwechsel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (!_seitenwechselModus || e.ChangedButton != MouseButton.Left || _reflowDebugModus) return;   // kein Seitenwechsel im Debug-Modus
                Point pos = e.GetPosition(PdfCanvas);
                int si = HoleSeitenIndexBeiPos(pos.Y, pos.X);
                if (si < 0) return;

                double pageH = _seitenHöhe[si];
                if (pageH <= 0) return;

                double yBase  = _layoutHorizontal ? SeiteX : (si < _seitenYStart.Length ? _seitenYStart[si] : 0);
                double yFrac  = Math.Max(0.01, Math.Min(0.99, (pos.Y - yBase) / pageH));

                // Seitenwechsel sofort ausführen und Modus beenden
                BeendeSeitenwechselModus();
                SetzeSeitenwechsel(si, yFrac);
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "Seitenwechsel_MouseDown"); }
        }

        private void SetzeSeitenwechsel(int si, double yFraction)
        {
            if (si >= _seitenBilder.Count) return;

            // Undo-Snapshot
            _undoStack.Push(SpeichereZustand());

            // Quell-Bitmap
            var sourceBmp = _kompositBilder.TryGetValue(si, out var kb) && kb != null
                ? kb : _seitenBilder[si];
            int origH   = _seitenBilder[si].PixelHeight;
            int sourceH = sourceBmp.PixelHeight;
            int sourceW = sourceBmp.PixelWidth;

            int cutY = (int)Math.Round(yFraction * sourceH);
            cutY = Math.Max(1, Math.Min(cutY, sourceH - 1));

            // Oberer Teil: auf origH padden
            var topVisual = new DrawingVisual();
            using (var ctx = topVisual.RenderOpen())
            {
                ctx.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, origH));
                if (cutY > 0)
                {
                    var topCrop = new CroppedBitmap(sourceBmp, new Int32Rect(0, 0, sourceW, cutY));
                    ctx.DrawImage(topCrop, new Rect(0, 0, sourceW, cutY));
                }
            }
            var topRtb = new RenderTargetBitmap(sourceW, origH, 96, 96, PixelFormats.Pbgra32);
            topRtb.Render(topVisual);
            topRtb.Freeze();
            _kompositBilder[si] = topRtb;

            // Unterer Teil: auf origH padden (oben anfangend)
            int belowH = sourceH - cutY;
            var botVisual = new DrawingVisual();
            using (var ctx = botVisual.RenderOpen())
            {
                ctx.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, origH));
                if (belowH > 0)
                {
                    var botCrop = new CroppedBitmap(sourceBmp, new Int32Rect(0, cutY, sourceW, belowH));
                    ctx.DrawImage(botCrop, new Rect(0, 0, sourceW, Math.Min(belowH, origH)));
                }
            }
            var botRtb = new RenderTargetBitmap(sourceW, origH, 96, 96, PixelFormats.Pbgra32);
            botRtb.Render(botVisual);
            botRtb.Freeze();

            // Neue Seite einfügen
            EnsureReihenfolge();
            int anzeigePos = _seitenReihenfolge.IndexOf(si);
            int neuIdx = _seitenBilder.Count;
            _seitenBilder.Add(botRtb);
            InitCropEintrag(neuIdx);
            _seitenReihenfolge.Insert(anzeigePos >= 0 ? anzeigePos + 1 : _seitenReihenfolge.Count, neuIdx);
            _kompositBilder[neuIdx] = botRtb;

            // Schnittlinien der alten Seite entfernen (Seite wurde geteilt)
            _scherenschnitte.RemoveAll(s => s.Seite == si);
            _gelöschteParts.RemoveWhere(p => p.Seite == si);

            ZeicheCanvas();
            AktualisiereSchnitteLinien();
            TxtInfo.Text = $"Seitenwechsel gesetzt – Seite {si + 1} geteilt in 2 – Strg+Z zum R\u00fckg\u00e4ngigmachen";
            MarkiereAlsGeändert();
        }

        private void Schere_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (!_scherenModus || _reflowDebugModus) return;   // kein Schere-Feedback im Debug-Modus
                Point pos = e.GetPosition(PdfCanvas);
                int si = HoleSeitenIndexBeiPos(pos.Y, pos.X);

                if (si < 0)
                {
                    if (_scherenVorschauLinie != null)
                        _scherenVorschauLinie.Visibility = Visibility.Collapsed;
                    return;
                }

                double pageW = _seitenBilder[si].PixelWidth;
                double setX  = _layoutHorizontal && si < _seitenXStart.Length ? _seitenXStart[si] : SeiteX;

                if (_scherenVorschauLinie == null)
                {
                    _scherenVorschauLinie = new Line
                    {
                        Stroke           = Brushes.Red,
                        StrokeThickness  = 2.0 / _zoomFaktor,
                        StrokeDashArray  = new DoubleCollection(new[] { 8.0, 4.0 }),
                        IsHitTestVisible = false,
                        Tag              = "SCHERE_PREVIEW"
                    };
                    Panel.SetZIndex(_scherenVorschauLinie, 200);
                    PdfCanvas.Children.Add(_scherenVorschauLinie);
                }

                _scherenVorschauLinie.X1             = setX;
                _scherenVorschauLinie.X2             = setX + pageW;
                _scherenVorschauLinie.Y1             = pos.Y;
                _scherenVorschauLinie.Y2             = pos.Y;
                _scherenVorschauLinie.StrokeThickness = 2.0 / _zoomFaktor;
                _scherenVorschauLinie.Visibility     = Visibility.Visible;
            }
            catch (Exception ex) { LogException(ex, "Schere_MouseMove"); }
        }

        private void Schere_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (!_scherenModus || e.ChangedButton != MouseButton.Left || _reflowDebugModus) return;   // kein Schnitt im Debug-Modus
                Point pos = e.GetPosition(PdfCanvas);
                int si = HoleSeitenIndexBeiPos(pos.Y, pos.X);
                if (si < 0) return;

                double pageH = _seitenHöhe[si];
                if (pageH <= 0) return;

                double yBase = _layoutHorizontal ? SeiteX : (si < _seitenYStart.Length ? _seitenYStart[si] : 0);
                double yFrac = Math.Max(0.01, Math.Min(0.99, (pos.Y - yBase) / pageH));

                _scherenschnitte.Add((si, yFrac));

                // Neues Modell: ContentBlock an yFrac physisch aufteilen
                if (_contentBlocks != null)
                {
                    SplitContentBlockBeiSchnitt(si, yFrac);
                    System.Diagnostics.Debug.WriteLine($"[SPLIT] Blocks jetzt: {_contentBlocks.Count}");
                }

                MarkiereAlsGeändert();
                if (_contentBlocks != null) ZeicheCanvas();
                else AktualisiereSchnitteLinien();
                TxtInfo.Text = $"✂ {_scherenschnitte.Count} Schnitt(e) – Strg+Z rückgängig";
                e.Handled = true;
            }
            catch (Exception ex) { LogException(ex, "Schere_MouseDown"); }
        }

        /// <summary>Gibt den Seiten-Index zurück, der die Canvas-Position enthält; -1 wenn keine.</summary>
        private int HoleSeitenIndexBeiPos(double canvasY, double canvasX)
        {
            if (_seitenBilder.Count == 0) return -1;
            if (_layoutHorizontal)
            {
                for (int i = 0; i < _seitenBilder.Count; i++)
                {
                    if (i >= _seitenXStart.Length) break;
                    double x0 = _seitenXStart[i];
                    double x1 = x0 + _seitenBilder[i].PixelWidth;
                    if (canvasX >= x0 && canvasX <= x1 && canvasY >= SeiteX && canvasY <= SeiteX + _seitenHöhe[i])
                        return i;
                }
            }
            else
            {
                for (int i = 0; i < _seitenBilder.Count; i++)
                {
                    if (i >= _seitenYStart.Length || i >= _seitenHöhe.Length) break;
                    double y0 = _seitenYStart[i];
                    double y1 = y0 + _seitenHöhe[i];
                    double x0 = SeiteX;
                    double x1 = x0 + _seitenBilder[i].PixelWidth;
                    if (canvasY >= y0 && canvasY <= y1 && canvasX >= x0 && canvasX <= x1)
                        return i;
                }
            }
            return -1;
        }

        private void AktualisiereSchnitteLinien()
        {
            if (_reflowDebugModus) { ZeicheCanvas(); return; }   // Reflow-Canvas neu zeichnen, keine Alt-Schnittlinien

            // Alte feste Schnittlinien entfernen (nicht Vorschau)
            var alte = PdfCanvas.Children.OfType<Line>()
                .Where(l => { var t = l.Tag?.ToString() ?? ""; return t.StartsWith("SCHERE_") && t != "SCHERE_PREVIEW"; })
                .ToList();
            foreach (var l in alte) PdfCanvas.Children.Remove(l);

            for (int k = 0; k < _scherenschnitte.Count; k++)
            {
                var (si, yFrac) = _scherenschnitte[k];
                if (si >= _seitenBilder.Count || si >= _seitenHöhe.Length) continue;

                double pageW = _seitenBilder[si].PixelWidth;
                double setX  = _layoutHorizontal && si < _seitenXStart.Length ? _seitenXStart[si] : SeiteX;
                double yBase = _layoutHorizontal ? SeiteX : (si < _seitenYStart.Length ? _seitenYStart[si] : 0);
                double canY  = yBase + yFrac * _seitenHöhe[si];

                var linie = new Line
                {
                    X1               = setX,
                    Y1               = canY,
                    X2               = setX + pageW,
                    Y2               = canY,
                    Stroke           = Brushes.Red,
                    StrokeThickness  = 2.0 / _zoomFaktor,
                    StrokeDashArray  = new DoubleCollection(new[] { 8.0, 4.0 }),
                    IsHitTestVisible = false,
                    Tag              = $"SCHERE_{k}"
                };
                Panel.SetZIndex(linie, 150);
                PdfCanvas.Children.Add(linie);
            }

            bool hatSchnitte = _scherenschnitte.Count > 0;
            BtnSchnittZurücksetzen.IsEnabled = hatSchnitte;
            BtnTeileExportieren.IsEnabled    = (_scherenschnitte.Count > 0 || _gelöschteParts.Count > 0 || _kompositBilder.Count > 0) && _pdfPfad != null;

            AktualisiereTeilOverlays();
        }

        private void BtnSchnittZurücksetzen_Click(object sender, RoutedEventArgs e)
        {
            SafeExecute(() =>
            {
                _scherenschnitte.Clear();
                MarkiereAlsGeändert();
                AktualisiereSchnitteLinien();
                TxtInfo.Text = "Alle Schnitte zurückgesetzt.";
            }, "BtnSchnittZurücksetzen_Click");
        }

        private void BtnTeileExportieren_Click(object sender, RoutedEventArgs e)
            => SafeExecute(ExportierteTeile, "BtnTeileExportieren_Click");

        private void ExportierteTeile()
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0) return;
            if (_scherenschnitte.Count == 0 && _gelöschteParts.Count == 0 && _kompositBilder.Count == 0) return;

            var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description       = "Ordner für PDF-Teile wählen",
                SelectedPath      = IO.Path.GetDirectoryName(_pdfPfad) ?? "",
                ShowNewFolderButton = true
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            string ordner = dlg.SelectedPath;
            string basis  = IO.Path.GetFileNameWithoutExtension(_pdfPfad);

            try
            {
                var seitenMitSchnitten = _scherenschnitte
                    .Select(s => s.Seite).Distinct().OrderBy(x => x).ToList();

                int exportiert = 0;

                foreach (int si in seitenMitSchnitten)
                {
                    if (si >= _seitenBilder.Count) continue;

                    var cuts = _scherenschnitte
                        .Where(s => s.Seite == si)
                        .Select(s => s.YFraction)
                        .OrderBy(f => f)
                        .ToList();

                    var grenzen = new List<double> { 0.0 };
                    grenzen.AddRange(cuts);
                    grenzen.Add(1.0);

                    // PDF-Seitengröße (Punkte, bottom-up) — via _pdfBytes, kein Datei-Handle
                    (double pageWPts, double pageHPts) = _pdfBytes != null
                        ? HolePdfSeitenGrösse(_pdfBytes)
                        : (595.28, 841.89); // A4 Fallback wenn _pdfBytes nicht geladen

                    for (int t = 0; t < grenzen.Count - 1; t++)
                    {
                        // Gelöschte Teile überspringen
                        if (_gelöschteParts.Contains((si, t))) continue;

                        double fracOben  = grenzen[t];
                        double fracUnten = grenzen[t + 1];

                        // PDF-Koordinaten: Y-Achse bottom-up
                        double pdfY1 = (1.0 - fracUnten) * pageHPts;  // untere Kante des Teils
                        double pdfY2 = (1.0 - fracOben)  * pageHPts;  // obere Kante des Teils

                        string suffix    = seitenMitSchnitten.Count > 1
                            ? $"_s{si + 1}_t{t + 1}" : $"_teil{t + 1}";
                        string zielDatei = IO.Path.Combine(ordner, basis + suffix + ".pdf");

                        // MemoryStream statt Dateipfad – verhindert Datei-Sperr-Konflikt mit AutoSpeichern
                        if (_pdfBytes == null)
                        {
                            System.Diagnostics.Debug.WriteLine("[SAVE] WARNUNG: _pdfBytes ist null, Seite wird übersprungen");
                            continue;
                        }
                        using var pdfIn = PdfReader.Open(new IO.MemoryStream(_pdfBytes), PdfDocumentOpenMode.Import);
                        var pdfOut      = new PdfSharp.Pdf.PdfDocument();
                        var seite       = pdfOut.AddPage(pdfIn.Pages[si]);
                        seite.CropBox   = new PdfSharp.Pdf.PdfRectangle(
                            new PdfSharp.Drawing.XPoint(0,        pdfY1),
                            new PdfSharp.Drawing.XPoint(pageWPts, pdfY2));
                        pdfOut.Save(zielDatei);
                        exportiert++;
                    }
                }

                // Komposit-Seiten exportieren (zusammengeschobene Seiten ohne Schnitte)
                foreach (var kv in _kompositBilder)
                {
                    int si = kv.Key;
                    if (si >= _seitenBilder.Count) continue;
                    // Seiten die bereits oben exportiert wurden überspringen
                    if (_scherenschnitte.Any(s => s.Seite == si)) continue;

                    var kompBmp = kv.Value;
                    if (kompBmp == null) continue;

                    (double pageWPts, double pageHPts) = _pdfBytes != null
                        ? HolePdfSeitenGrösse(_pdfBytes)
                        : (595.28, 841.89); // A4 Fallback wenn _pdfBytes nicht geladen
                    double scaleH  = (double)kompBmp.PixelHeight / Math.Max(1, _seitenBilder[si].PixelHeight);
                    double newHPts = Math.Max(1, pageHPts * scaleH);

                    string suffix    = _kompositBilder.Count > 1 ? $"_s{si + 1}_zs" : "_zusammengeschoben";
                    string zielDatei = IO.Path.Combine(ordner, basis + suffix + ".pdf");

                    using var ms = new IO.MemoryStream();
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(kompBmp));
                    enc.Save(ms);
                    ms.Position = 0;
                    using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);

                    var pdfOut = new PdfSharp.Pdf.PdfDocument();
                    var seite  = pdfOut.AddPage();
                    seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(pageWPts);
                    seite.Height = PdfSharp.Drawing.XUnit.FromPoint(newHPts);
                    using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                    gfx.DrawImage(xImg, 0, 0, pageWPts, newHPts);
                    pdfOut.Save(zielDatei);
                    exportiert++;
                }

                TxtInfo.Text = $"✂ {exportiert} PDF-Teil(e) exportiert";
                AppZustand.Instanz.SetzeStatus($"Scheren-Export: {exportiert} Dateien → {IO.Path.GetFileName(ordner)}");
            }
            catch (Exception ex)
            {
                LogException(ex, "ExportierteTeile");
                MessageBox.Show($"Export fehlgeschlagen:\n{App.GetExceptionKette(ex)}",
                    "Export-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Teil-Auswahl & Löschung ───────────────────────────────────────────

        /// <summary>Gibt sortierte Schnitt-Y-Fraktionen für eine Seite zurück.</summary>
        private List<double> GetSchnitteVonSeite(int si)
            => _scherenschnitte.Where(s => s.Seite == si)
                               .Select(s => s.YFraction)
                               .OrderBy(f => f)
                               .ToList();

        /// <summary>Gibt die Teil-Grenzen (FracOben, FracUnten) für eine Seite zurück.</summary>
        private List<(double Oben, double Unten)> GetTeilGrenzen(int si)
        {
            var cuts    = GetSchnitteVonSeite(si);
            var grenzen = new List<double> { 0.0 };
            grenzen.AddRange(cuts);
            grenzen.Add(1.0);
            var result = new List<(double, double)>();
            for (int t = 0; t < grenzen.Count - 1; t++)
                result.Add((grenzen[t], grenzen[t + 1]));
            return result;
        }

        /// <summary>Gibt den Teil-Index zurück, in dem die canvas-Y-Koordinate auf Seite si liegt.</summary>
        private int GetTeilBeiCanvasY(int si, double canvasY)
        {
            if (si < 0 || si >= _seitenHöhe.Length || _seitenHöhe[si] <= 0) return 0;
            double yBase = _layoutHorizontal ? SeiteX : (si < _seitenYStart.Length ? _seitenYStart[si] : 0);
            double yFrac = Math.Max(0, Math.Min(1, (canvasY - yBase) / _seitenHöhe[si]));
            var grenzen  = GetTeilGrenzen(si);
            for (int t = 0; t < grenzen.Count; t++)
                if (yFrac < grenzen[t].Unten) return t;
            return grenzen.Count - 1;
        }

        /// <summary>Zeichnet klickbare Teil-Overlays auf den Canvas (Auswahl + Lösch-Markierung).</summary>
        private void AktualisiereTeilOverlays()
        {
            if (_reflowDebugModus) return;   // Reflow-Canvas bleibt frei von Altmodell-Overlays
            if (_contentBlocks != null) return;   // Block-Renderpfad: keine Altmodell-Overlays

            // Alte Overlays entfernen
            var alte = PdfCanvas.Children.OfType<Border>()
                .Where(b => (b.Tag?.ToString() ?? "").StartsWith("TEIL_"))
                .ToList();
            foreach (var b in alte) PdfCanvas.Children.Remove(b);

            if (_seitenBilder.Count == 0) return;

            for (int si = 0; si < _seitenBilder.Count; si++)
            {
                if (_gelöschteSeiten.Contains(si)) continue;
                if (_seitenYStart.Length > si && _seitenYStart[si] < -100) continue; // off-screen (deleted or not visible)
                if (si >= _seitenHöhe.Length) continue;
                if (!_layoutHorizontal && si >= _seitenYStart.Length) continue;

                double pageW = _seitenBilder[si].PixelWidth;
                double setX  = _layoutHorizontal && si < _seitenXStart.Length ? _seitenXStart[si] : SeiteX;
                double yBase = _layoutHorizontal ? SeiteX : _seitenYStart[si];
                double pH    = _seitenHöhe[si];

                var teilGrenzen = GetTeilGrenzen(si);

                for (int t = 0; t < teilGrenzen.Count; t++)
                {
                    var (fracOben, fracUnten) = teilGrenzen[t];
                    double y1 = yBase + fracOben  * pH;
                    double h  = Math.Max(1, (fracUnten - fracOben) * pH);

                    bool ausgewählt = _ausgewählteParts.Contains((si, t));
                    bool gelöscht   = _gelöschteParts.Contains((si, t));

                    var overlay = new Border
                    {
                        Tag             = $"TEIL_{si}_{t}",
                        Width           = pageW,
                        Height          = h,
                        Background      = gelöscht
                            ? new SolidColorBrush(Color.FromArgb(150, 80, 80, 80))
                            : ausgewählt
                                ? new SolidColorBrush(Color.FromArgb(55, 30, 100, 220))
                                : Brushes.Transparent,
                        BorderBrush     = ausgewählt
                            ? new SolidColorBrush(Color.FromRgb(30, 100, 220))
                            : gelöscht
                                ? new SolidColorBrush(Color.FromRgb(120, 40, 40))
                                : Brushes.Transparent,
                        BorderThickness = new Thickness(ausgewählt || gelöscht ? 2.0 / _zoomFaktor : 0),
                        IsHitTestVisible = !_scherenModus,
                        Cursor           = _scherenModus ? null : Cursors.Hand,
                        ClipToBounds     = true
                    };

                    if (gelöscht)
                    {
                        overlay.Child = new TextBlock
                        {
                            Text                = "✕  entfernt",
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment   = VerticalAlignment.Center,
                            Foreground          = Brushes.White,
                            FontSize            = Math.Max(10, Math.Min(18, h / 4)),
                            FontWeight          = FontWeights.Bold,
                            IsHitTestVisible    = false
                        };
                    }

                    int finalSi = si, finalT = t;
                    double capturedFracOben = fracOben, capturedFracUnten = fracUnten;
                    overlay.MouseLeftButtonDown += (_, ev) =>
                    {
                        if (_scherenModus) return;
                        System.Diagnostics.Debug.WriteLine($"[BUG1-KLICK] Klick auf Overlay si={finalSi} t={finalT} fracOben={capturedFracOben:F3} fracUnten={capturedFracUnten:F3}");
                        bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
                        if (!ctrl) _ausgewählteParts.Clear();
                        if (_ausgewählteParts.Contains((finalSi, finalT)))
                            _ausgewählteParts.Remove((finalSi, finalT));
                        else
                            _ausgewählteParts.Add((finalSi, finalT));
                        System.Diagnostics.Debug.WriteLine($"[BUG1-KLICK] _ausgewählteParts nach Klick: {string.Join(",", _ausgewählteParts.Select(p => $"si={p.Seite},t={p.Teil}"))}");
                        _markierteSeitenIdx = finalSi; // Letzte angeklickte Seite merken
                        AktualisiereTeilOverlays();
                        BtnTeilLöschen.IsEnabled = _ausgewählteParts.Count > 0;
                        int n = _ausgewählteParts.Count;
                        TxtInfo.Text = n > 0
                            ? $"{n} Teil(e) markiert – Entf oder Rechtsklick zum Löschen"
                            : "";
                        ScrollView.Focus();
                        ev.Handled = true;
                    };

                    // Schnittlinie über diesem Teil draggbar machen (wenn nicht im Scheren-Modus)
                    if (t > 0 && !_scherenModus)
                    {
                        // Index der Schnittlinie über diesem Teil bestimmen
                        double cutFrac = fracOben; // fracOben dieses Teils = Position der Schnittlinie darüber
                        int capturedCutIdx = -1;
                        for (int ci2 = 0; ci2 < _scherenschnitte.Count; ci2++)
                        {
                            if (_scherenschnitte[ci2].Seite == si && Math.Abs(_scherenschnitte[ci2].YFraction - cutFrac) < 0.001)
                            {
                                capturedCutIdx = ci2; break;
                            }
                        }

                        if (capturedCutIdx >= 0)
                        {
                            double dragSchwelle = Math.Max(7.0, 7.0 / _zoomFaktor); // px innerhalb des Overlays
                            int capturedDragSi = si;

                            // Cursor-Wechsel: oben im Overlay = SizeNS, sonst Hand
                            overlay.MouseMove += (_, ev2) =>
                            {
                                if (_schnittDragAktiv && _gezogenesSchnittIdx == capturedCutIdx)
                                {
                                    // Im Drag: Schnittposition aktualisieren
                                    var pos2 = ev2.GetPosition(PdfCanvas);
                                    int dsi2 = _scherenschnitte[capturedCutIdx].Seite;
                                    if (dsi2 < _seitenHöhe.Length && _seitenHöhe[dsi2] > 0)
                                    {
                                        double yBase3 = _layoutHorizontal ? SeiteX : (dsi2 < _seitenYStart.Length ? _seitenYStart[dsi2] : 0);
                                        double newFrac = Math.Max(0.01, Math.Min(0.99, (pos2.Y - yBase3) / _seitenHöhe[dsi2]));
                                        var andereFracs2 = _scherenschnitte
                                            .Where((s3, ix) => ix != capturedCutIdx && s3.Seite == dsi2)
                                            .Select(s3 => s3.YFraction).OrderBy(f => f).ToList();
                                        int myPos2 = _scherenschnitte
                                            .Select((s3, ix) => (s3, ix))
                                            .Where(x2 => x2.ix != capturedCutIdx && x2.s3.Seite == dsi2)
                                            .OrderBy(x2 => x2.s3.YFraction)
                                            .TakeWhile(x2 => x2.s3.YFraction < _schnittDragOrigFrac)
                                            .Count();
                                        double minF2 = myPos2 > 0 ? andereFracs2[myPos2 - 1] + 0.01 : 0.01;
                                        double maxF2 = myPos2 < andereFracs2.Count ? andereFracs2[myPos2] - 0.01 : 0.99;
                                        newFrac = Math.Max(minF2, Math.Min(maxF2, newFrac));
                                        var sOld2 = _scherenschnitte[capturedCutIdx];
                                        _scherenschnitte[capturedCutIdx] = (sOld2.Seite, newFrac);
                                        AktualisiereSchnitteLinien();
                                    }
                                    ev2.Handled = true;
                                }
                                else
                                {
                                    var pos2 = ev2.GetPosition(overlay);
                                    overlay.Cursor = pos2.Y < dragSchwelle ? Cursors.SizeNS : Cursors.Hand;
                                }
                            };

                            // Drag starten wenn nahe der Oberkante (= Schnittlinie) gedrückt wird
                            overlay.PreviewMouseLeftButtonDown += (_, ev2) =>
                            {
                                if (_scherenModus || capturedCutIdx < 0) return;
                                var pos2 = ev2.GetPosition(overlay);
                                if (pos2.Y > dragSchwelle) return; // Nicht nahe genug an der Schnittlinie
                                _undoStack.Push(SpeichereZustand());
                                _gezogenesSchnittIdx = capturedCutIdx;
                                _schnittDragAktiv    = true;
                                _schnittDragOrigFrac = _scherenschnitte[capturedCutIdx].YFraction;
                                overlay.CaptureMouse();
                                ev2.Handled = true; // WICHTIG: Verhindert Teil-Auswahl und PreviewMouseLeftButtonDown vom SEITE_-Border
                            };

                            overlay.MouseLeftButtonUp += (s2, ev2) =>
                            {
                                if (_schnittDragAktiv && _gezogenesSchnittIdx == capturedCutIdx)
                                {
                                    ((UIElement)s2).ReleaseMouseCapture();
                                    bool hatBewegt = Math.Abs(_scherenschnitte[capturedCutIdx].YFraction - _schnittDragOrigFrac) > 0.002;
                                    _schnittDragAktiv    = false;
                                    _gezogenesSchnittIdx = -1;
                                    if (hatBewegt)
                                    {
                                        AktualisiereSchnitteLinien();
                                        TxtInfo.Text = "Schnittlinie verschoben – Strg+Z zum Rückgängigmachen";
                                        MarkiereAlsGeändert();
                                    }
                                    else
                                    {
                                        // Kein echter Drag — als Klick behandeln: Segment selektieren
                                        if (_undoStack.Count > 0) _undoStack.Pop(); // Drag-Start-Undo rückgängig
                                        _ausgewählteParts.Clear();
                                        _ausgewählteParts.Add((capturedDragSi, finalT));
                                        AktualisiereTeilOverlays();
                                        BtnTeilLöschen.IsEnabled = true;
                                        TxtInfo.Text = "1 Teil(e) markiert – Entf oder Rechtsklick zum Löschen";
                                    }
                                    ev2.Handled = true;
                                }
                            };

                            overlay.LostMouseCapture += (_, __) =>
                            {
                                if (_gezogenesSchnittIdx == capturedCutIdx)
                                {
                                    _schnittDragAktiv    = false;
                                    _gezogenesSchnittIdx = -1;
                                }
                            };
                        }
                    }

                    // Kontext-Menü (Rechtsklick)
                    var menu       = new ContextMenu();

                    // Verbesserung 1: "Seite löschen" als erstes Item (Fix ZIndex-Problem)
                    var itemLöschSeite = new MenuItem { Header = "\U0001F5D1  Seite löschen" };
                    int capSi = finalSi;
                    itemLöschSeite.Click += (_, __) => LöscheSeiteMitBestätigung(capSi);
                    menu.Items.Insert(0, itemLöschSeite);
                    menu.Items.Insert(1, new Separator());

                    var itemLösch  = new MenuItem { Header = "✕  Teil löschen" };
                    itemLösch.Click += (_, __) =>
                    {
                        if (!_ausgewählteParts.Contains((finalSi, finalT)))
                        {
                            _ausgewählteParts.Clear();
                            _ausgewählteParts.Add((finalSi, finalT));
                        }
                        LöscheAusgewählteParts();
                    };
                    var itemAbwählen = new MenuItem { Header = "Auswahl aufheben" };
                    itemAbwählen.Click += (_, __) =>
                    {
                        _ausgewählteParts.Clear();
                        AktualisiereTeilOverlays();
                        BtnTeilLöschen.IsEnabled = false;
                        TxtInfo.Text = "";
                    };
                    menu.Items.Add(itemLösch);

                    if (gelöscht)
                    {
                        var itemLücke = new MenuItem { Header = "\u2B06  Lücke schließen" };
                        int capSi2 = finalSi;
                        int capT2  = finalT;
                        itemLücke.Click += (_, __) => SchliesseLücke(capSi2, capT2);
                        menu.Items.Insert(menu.Items.Count, itemLücke); // nach "Teil löschen"
                    }

                    menu.Items.Add(new Separator());

                    // Leerzeile einfügen
                    var itemLeerzeile = new MenuItem { Header = "\u2195  Leerzeile einfügen" };
                    int capLSi = finalSi, capLT = finalT;
                    itemLeerzeile.Click += (_, __) =>
                    {
                        bool oberhalb = ZeigeBinaryDialog(
                            "Leerzeile einfügen",
                            "Wo soll die Leerzeile eingefügt werden?",
                            "Oberhalb",
                            "Unterhalb");
                        FügeLeerzeileEin(capLSi, capLT, oberhalb);
                    };
                    menu.Items.Add(itemLeerzeile);

                    menu.Items.Add(new Separator());
                    menu.Items.Add(itemAbwählen);

                    // Verbesserung 2: "Teile verschmelzen" wenn ≥2 benachbarte Teile ausgewählt
                    if (_ausgewählteParts.Count >= 2)
                    {
                        var seiten = _ausgewählteParts.Select(p => p.Seite).Distinct().ToList();
                        if (seiten.Count == 1)
                        {
                            var sortierteTeilIndizes = _ausgewählteParts
                                .Where(p => p.Seite == seiten[0])
                                .Select(p => p.Teil)
                                .OrderBy(t2 => t2)
                                .ToList();
                            bool lückenlos = sortierteTeilIndizes.Last() - sortierteTeilIndizes.First()
                                             == sortierteTeilIndizes.Count - 1;
                            if (lückenlos)
                            {
                                menu.Items.Add(new Separator());
                                var itemVerschmelzen = new MenuItem { Header = "\u2702  Teile verschmelzen" };
                                itemVerschmelzen.Click += (_, __) => VerschmelzeAusgewählteParts();
                                menu.Items.Add(itemVerschmelzen);
                            }
                        }
                    }

                    overlay.ContextMenu = menu;

                    Canvas.SetLeft(overlay, setX);
                    Canvas.SetTop(overlay,  y1);
                    Panel.SetZIndex(overlay, 50);
                    PdfCanvas.Children.Add(overlay);
                }
            }

            BtnTeilLöschen.IsEnabled = _ausgewählteParts.Count > 0;
        }

        private void LöscheAusgewählteParts()
        {
            if (_ausgewählteParts.Count == 0) return;

            System.Diagnostics.Debug.WriteLine($"[LOESCHEN] _ausgewählteParts: {string.Join(",", _ausgewählteParts.Select(p => $"si={p.Seite},t={p.Teil}"))}");

            // GapDialog nur zeigen wenn Schnitte vorhanden (sonst kein sinnvoller Lücken-Kontext)
            bool hatSchnitte = _ausgewählteParts.Any(p => GetTeilGrenzen(p.Seite).Count > 1);

            GapModus gewählterModus = GapModus.OriginalAbstand;
            double   eingabeMm     = 0.0;

            if (hatSchnitte)
            {
                var dlg = new GapDialog { Owner = Window.GetWindow(this) };
                if (dlg.ShowDialog() != true) return;
                gewählterModus = dlg.GewählterModus;
                eingabeMm      = dlg.EingabeGapMm;
            }

            // Aktuellen Zustand für Undo sichern
            _undoStack.Push(SpeichereZustand());

            foreach (var p in _ausgewählteParts)
            {
                _gelöschteParts.Add(p);

                if (_contentBlocks != null)
                {
                    SetzeContentBlockGelöscht(p.Seite, p.Teil, true);
                    SetzeContentBlockGapInfo(p.Seite, p.Teil, gewählterModus, eingabeMm);
                }
            }

            _ausgewählteParts.Clear();
            if (_contentBlocks != null) ZeicheCanvas();
            else AktualisiereSchnitteLinien();
            BtnTeilLöschen.IsEnabled = false;
            TxtInfo.Text = $"{_gelöschteParts.Count} Teil(e) entfernt – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        private ScherenZustand SpeichereZustand()
            => new ScherenZustand
            {
                Schnitte          = new List<(int, double)>(_scherenschnitte),
                Gelöscht          = new HashSet<(int, int)>(_gelöschteParts),
                KompositBilder    = new Dictionary<int, BitmapSource>(_kompositBilder),
                GelöschteSeiten   = new HashSet<int>(_gelöschteSeiten),
                SeitenReihenfolge = _seitenReihenfolge != null ? new List<int>(_seitenReihenfolge) : null,
            };

        /// <summary>Zeigt einen Dialog mit zwei benutzerdefinierten Schaltflächen.
        /// Gibt true zurück wenn der erste Button geklickt wurde.</summary>
        private static bool ZeigeBinaryDialog(string titel, string nachricht, string btnErster, string btnZweiter)
        {
            var win = new Window
            {
                Title                 = titel,
                SizeToContent         = SizeToContent.WidthAndHeight,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode            = ResizeMode.NoResize,
                MinWidth              = 320
            };
            bool ersterGeklickt = false;
            var sp = new StackPanel { Margin = new Thickness(20) };
            sp.Children.Add(new TextBlock
            {
                Text         = nachricht,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth     = 400,
                Margin       = new Thickness(0, 0, 0, 16)
            });
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var btn1 = new Button { Content = btnErster, Padding = new Thickness(16, 6, 16, 6),
                                    Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var btn2 = new Button { Content = btnZweiter, Padding = new Thickness(16, 6, 16, 6), IsCancel = true };
            btn1.Click += (_, __) => { ersterGeklickt = true; win.Close(); };
            btn2.Click += (_, __) => win.Close();
            buttons.Children.Add(btn1);
            buttons.Children.Add(btn2);
            sp.Children.Add(buttons);
            win.Content = sp;
            if (Application.Current?.MainWindow is Window mw) win.Owner = mw;
            win.ShowDialog();
            return ersterGeklickt;
        }

        /// <summary>Verbesserung 2: Verschmelzt alle ausgewählten benachbarten Teile.</summary>
        private void VerschmelzeAusgewählteParts()
        {
            if (_ausgewählteParts.Count < 2) return;
            var seiten = _ausgewählteParts.Select(p => p.Seite).Distinct().ToList();
            if (seiten.Count != 1) return;
            int si = seiten[0];

            var sortierteTeilIndizes = _ausgewählteParts
                .Where(p => p.Seite == si)
                .Select(p => p.Teil)
                .OrderBy(t => t)
                .ToList();

            _undoStack.Push(SpeichereZustand());

            var alleGrenzen = GetTeilGrenzen(si);

            // Schnittlinien zwischen den ausgewählten benachbarten Teilen entfernen
            for (int k = 0; k < sortierteTeilIndizes.Count - 1; k++)
            {
                int t = sortierteTeilIndizes[k];
                if (t < alleGrenzen.Count)
                {
                    double schnittFrac = alleGrenzen[t].Unten;
                    _scherenschnitte.RemoveAll(s => s.Seite == si
                                               && Math.Abs(s.YFraction - schnittFrac) < 0.001);
                }
            }

            _ausgewählteParts.Clear();
            AktualisiereSchnitteLinien();
            TxtInfo.Text = "Teile verschmolzen – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        /// <summary>Verbesserung 6: Schließt die Lücke eines als gelöscht markierten Parts.</summary>
        private void SchliesseLücke(int si, int t)
        {
            if (!_gelöschteParts.Contains((si, t))) return;
            _undoStack.Push(SpeichereZustand());
            var toClose = new HashSet<(int Seite, int Teil)> { (si, t) };
            SchiebeTeileZusammen(toClose, true); // verschmelzen=true schließt Lücke ohne Dialog
        }

        private void SchiebeTeileZusammen(HashSet<(int Seite, int Teil)> gelöschteParts, bool verschmelzen)
        {
            System.Diagnostics.Debug.WriteLine($"[BUG1-SCHIEBEN] gelöschteParts: {string.Join(",", gelöschteParts.Select(p => $"si={p.Seite},t={p.Teil}"))}");
            var betroffeneSeiten = gelöschteParts.Select(p => p.Seite).Distinct().ToList();

            foreach (int si in betroffeneSeiten)
            {
                if (si >= _seitenBilder.Count) continue;

                var alleGrenzen = GetTeilGrenzen(si);
                var sichtbarTeilIndizes = Enumerable.Range(0, alleGrenzen.Count)
                    .Where(t => !gelöschteParts.Contains((si, t)))
                    .ToList();
                System.Diagnostics.Debug.WriteLine($"[BUG1-SCHIEBEN] si={si} alleGrenzen=[{string.Join(", ", alleGrenzen.Select(g => $"{g.Oben:F3}-{g.Unten:F3}"))}] sichtbar=[{string.Join(",", sichtbarTeilIndizes)}]");

                if (sichtbarTeilIndizes.Count == 0)
                {
                    // Alle Teile gelöscht: Komposit-Bild weglassen (Seite behält Original)
                    _kompositBilder.Remove(si);
                }
                else if (sichtbarTeilIndizes.Count == alleGrenzen.Count)
                {
                    // Nichts gelöscht auf dieser Seite
                    continue;
                }
                else
                {
                    var kompositBmp = ErzeugeKompositBild(si, sichtbarTeilIndizes);
                    if (kompositBmp != null)
                        _kompositBilder[si] = kompositBmp;
                }

                // Schnitte für diese Seite entfernen
                _scherenschnitte.RemoveAll(s => s.Seite == si);

                // Verschmelzen=Nein: neue Schnittlinien an den Grenzen der sichtbaren Teile
                // so dass die Teile im Komposit weiterhin einzeln auswählbar sind
                if (!verschmelzen && sichtbarTeilIndizes.Count > 1)
                {
                    double akkFrac = 0;
                    for (int k = 0; k < sichtbarTeilIndizes.Count - 1; k++)
                    {
                        int t = sichtbarTeilIndizes[k];
                        var (o, u) = alleGrenzen[t];
                        akkFrac += (u - o);
                        _scherenschnitte.Add((si, akkFrac));
                    }
                }

                // Gelöschte Parts für diese Seite entfernen (sind jetzt physisch weg)
                _gelöschteParts.RemoveWhere(p => p.Seite == si);
            }

            _ausgewählteParts.Clear();

            // _contentBlocks nach Altmodell-Änderung synchronisieren (wie im Undo-Pfad)
            if (_contentBlocks != null)
            {
                _contentBlocks = KonvertiereAltesModellZuBlöcken();
                _nextBlockId   = _contentBlocks.Count > 0 ? _contentBlocks.Max(b => b.BlockId) + 1 : 0;
            }

            // Canvas neu aufbauen mit aktualisierten Composite-Bitmaps
            ZeicheCanvas();
            BtnTeilLöschen.IsEnabled = false;
            string modusText = verschmelzen ? "verschmolzen" : "zusammengeschoben (getrennte Teile)";
            TxtInfo.Text = $"Teile {modusText} – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        private BitmapSource? ErzeugeKompositBild(int si, List<int> sichtbareTeilIndizes)
        {
            try
            {
                if (sichtbareTeilIndizes.Count == 0) return null;

                // Wichtig: Wenn bereits ein Komposit-Bild existiert, dessen Fraktionen
                // (in _scherenschnitte) im Komposit-Raum liegen – daher das Komposit als Quelle verwenden.
                var sourceBmp   = _kompositBilder.TryGetValue(si, out var vorhandenesKomposit) && vorhandenesKomposit != null
                    ? vorhandenesKomposit
                    : _seitenBilder[si];
                var alleGrenzen = GetTeilGrenzen(si);
                int pW = sourceBmp.PixelWidth;

                var teile = new List<CroppedBitmap>();
                foreach (int t in sichtbareTeilIndizes)
                {
                    if (t >= alleGrenzen.Count) continue;
                    var (oben, unten) = alleGrenzen[t];
                    int y0 = (int)Math.Round(oben  * sourceBmp.PixelHeight);
                    int h  = (int)Math.Round(unten * sourceBmp.PixelHeight) - y0;
                    h = Math.Max(1, Math.Min(h, sourceBmp.PixelHeight - y0));
                    if (y0 < 0 || y0 >= sourceBmp.PixelHeight || h <= 0) continue;
                    teile.Add(new CroppedBitmap(sourceBmp, new Int32Rect(0, y0, pW, h)));
                }

                if (teile.Count == 0) return null;
                // Verbesserung 3: Kein Frühabbruch bei 1 Teil – immer durch Komposit-Erstellung
                // damit das Seitenformat (origH) immer beibehalten wird.

                int totalH = teile.Sum(t2 => t2.PixelHeight);
                if (totalH <= 0) return null;

                var visual = new DrawingVisual();
                using (var ctx = visual.RenderOpen())
                {
                    double y = 0;
                    ctx.DrawRectangle(Brushes.White, null, new Rect(0, 0, pW, totalH)); // Hintergrund weiß
                    foreach (var t2 in teile)
                    {
                        ctx.DrawImage(t2, new Rect(0, y, pW, t2.PixelHeight));
                        y += t2.PixelHeight;
                    }
                }
                var rtb = new RenderTargetBitmap(pW, totalH, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(visual);

                // Seitenformat beibehalten: auf Original-Höhe auffüllen (weißer Bereich unten)
                int origH = sourceBmp.PixelHeight;
                if (totalH < origH)
                {
                    var padVisual = new DrawingVisual();
                    using (var padCtx = padVisual.RenderOpen())
                    {
                        padCtx.DrawRectangle(Brushes.White, null, new Rect(0, 0, pW, origH));
                        padCtx.DrawImage(rtb, new Rect(0, 0, pW, totalH));
                    }
                    var padRtb = new RenderTargetBitmap(pW, origH, 96, 96, PixelFormats.Pbgra32);
                    padRtb.Render(padVisual);
                    padRtb.Freeze();
                    return padRtb;
                }

                rtb.Freeze();
                return rtb;
            }
            catch (Exception ex) { LogException(ex, "ErzeugeKompositBild"); return null; }
        }

        /// <summary>Fügt einen weißen Leerstreifen (30 px) ober- oder unterhalb des Segments t ein.</summary>
        private void FügeLeerzeileEin(int si, int t, bool oberhalb)
        {
            if (si >= _seitenBilder.Count) return;

            // Quell-Bitmap: Komposit wenn vorhanden, sonst Original
            var sourceBmp = _kompositBilder.TryGetValue(si, out var kb) && kb != null
                ? kb : _seitenBilder[si];

            var grenzen = GetTeilGrenzen(si);
            if (t >= grenzen.Count) return;

            var (fracOben, fracUnten) = grenzen[t];
            int origH   = _seitenBilder[si].PixelHeight;  // unveränderliche Originalhöhe (Seitenformat)
            int sourceW = sourceBmp.PixelWidth;
            const int stripH = 30;

            // Tatsächliche Inhaltshöhe bestimmen:
            // Das sourceBmp kann auf origH gepaddet sein (weißer Leerraum unten).
            // Die echte Inhaltshöhe ist die Summe aller sichtbaren Teile in Pixeln.
            // So wissen wir wieviel Platz noch auf der Seite frei ist.
            int inhaltH;
            {
                var schnitte = GetSchnitteVonSeite(si);
                if (schnitte.Count == 0)
                {
                    // Keine Schnitte: Inhalt geht von 0 bis letzter sichtbarer Fraktion
                    // Bei unverändertem Original: Inhalt = origH
                    // Bei zusammengeschobenem Komposit: Inhalt = totalH der sichtbaren Teile
                    // Wir ermitteln dies aus dem letzten sichtbaren Schnitt-Abschluss.
                    // Einfachste Annäherung: wenn Komposit kleiner als origH war (kein Padding nötig),
                    // dann ist der Inhalt totalH. Da ErzeugeKompositBild auf origH padded,
                    // können wir den echten Inhalt über die früheren Teil-Grenzen ermitteln.
                    // Für den Fall ohne Schnitte: nehmen wir origH als Inhalt (konservativ).
                    inhaltH = origH;
                }
                else
                {
                    // Mit Schnitten: sichtbare Teile aufsummieren
                    var alleGrenzen = GetTeilGrenzen(si);
                    double totalFrac = alleGrenzen
                        .Where((g, idx) => !_gelöschteParts.Contains((si, idx)))
                        .Sum(g => g.Unten - g.Oben);
                    inhaltH = (int)Math.Round(totalFrac * sourceBmp.PixelHeight);
                    inhaltH = Math.Max(1, Math.Min(inhaltH, origH));
                }
            }

            // Einfügeposition in Pixeln (im Quell-Bitmap-Koordinatensystem)
            int insertY = oberhalb
                ? (int)Math.Round(fracOben  * sourceBmp.PixelHeight)
                : (int)Math.Round(fracUnten * sourceBmp.PixelHeight);
            insertY = Math.Max(0, Math.Min(insertY, sourceBmp.PixelHeight));

            // Undo-Snapshot VOR der Änderung
            _undoStack.Push(SpeichereZustand());

            // Gesamtes neues Bitmap aufbauen (Quelle + Leerstreifen eingefügt)
            int sourceH = sourceBmp.PixelHeight;
            int newH    = inhaltH + stripH;   // tatsächlicher Inhalt + Streifen

            var allVisual = new DrawingVisual();
            using (var ctx = allVisual.RenderOpen())
            {
                ctx.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, newH));
                // Bereich ÜBER der Einfügeposition
                if (insertY > 0)
                {
                    var topCrop = new CroppedBitmap(sourceBmp, new Int32Rect(0, 0, sourceW, Math.Min(insertY, sourceH)));
                    ctx.DrawImage(topCrop, new Rect(0, 0, sourceW, Math.Min(insertY, sourceH)));
                }
                // Weißer Streifen
                ctx.DrawRectangle(Brushes.White, null, new Rect(0, insertY, sourceW, stripH));
                // Bereich UNTER der Einfügeposition (bis max inhaltH)
                int below = Math.Min(inhaltH - insertY, sourceH - insertY);
                if (below > 0)
                {
                    var botCrop = new CroppedBitmap(sourceBmp, new Int32Rect(0, insertY, sourceW, below));
                    ctx.DrawImage(botCrop, new Rect(0, insertY + stripH, sourceW, below));
                }
            }
            var fullRtb = new RenderTargetBitmap(sourceW, newH, 96, 96, PixelFormats.Pbgra32);
            fullRtb.Render(allVisual);
            fullRtb.Freeze();

            // Schnittlinien für Seite si verschieben: alle unterhalb insertY um stripH nach unten.
            // Wichtig: Fraktionen müssen auf origH normiert sein, weil das Komposit-Bitmap
            // immer origH Pixel hoch ist (ErzeugeKompositBild paddet auf origH).
            // Nicht auf newH normieren (newH < origH bei kein Überlauf)!
            double insertFrac = sourceH > 0 ? (double)insertY / sourceH : 0;
            for (int i = 0; i < _scherenschnitte.Count; i++)
            {
                var (cSi, cFrac) = _scherenschnitte[i];
                if (cSi == si && cFrac >= insertFrac)
                {
                    double oldYPx = cFrac * sourceH;
                    double newYPx = oldYPx + stripH;
                    // Normieren auf origH (die tatsächliche Bitmap-Höhe nach Padding)
                    _scherenschnitte[i] = (cSi, Math.Min(newYPx / origH, 1.0));
                }
            }

            // Seitenformat-Invariante: Bitmap darf origH nicht überschreiten
            if (newH <= origH)
            {
                // Leerzeile passt noch auf die Seite → auf origH padden (kein Überlauf, KEINE neue Seite)
                var padVisual = new DrawingVisual();
                using (var padCtx = padVisual.RenderOpen())
                {
                    padCtx.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, origH));
                    padCtx.DrawImage(fullRtb, new Rect(0, 0, sourceW, newH));
                }
                var paddedRtb = new RenderTargetBitmap(sourceW, origH, 96, 96, PixelFormats.Pbgra32);
                paddedRtb.Render(padVisual);
                paddedRtb.Freeze();
                _kompositBilder[si] = paddedRtb;
                System.Diagnostics.Debug.WriteLine($"[LEERZEILE] Kein Überlauf: inhaltH={inhaltH}+{stripH}={newH} <= origH={origH}");
            }
            else
            {
                // Überlauf: oberen Teil (origH Pixel) auf Seite si belassen, Rest auf neue Seite
                var topV = new DrawingVisual();
                using (var c = topV.RenderOpen())
                {
                    c.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, origH));
                    c.DrawImage(new CroppedBitmap(fullRtb, new Int32Rect(0, 0, sourceW, origH)),
                                new Rect(0, 0, sourceW, origH));
                }
                var topRtb = new RenderTargetBitmap(sourceW, origH, 96, 96, PixelFormats.Pbgra32);
                topRtb.Render(topV);
                topRtb.Freeze();
                _kompositBilder[si] = topRtb;

                // Überlauf-Teil (newH - origH Pixel), max eine Seite
                int overH = Math.Min(newH - origH, origH);
                var botV = new DrawingVisual();
                using (var c = botV.RenderOpen())
                {
                    c.DrawRectangle(Brushes.White, null, new Rect(0, 0, sourceW, origH));
                    c.DrawImage(new CroppedBitmap(fullRtb, new Int32Rect(0, origH, sourceW, overH)),
                                new Rect(0, 0, sourceW, overH));
                }
                var botRtb = new RenderTargetBitmap(sourceW, origH, 96, 96, PixelFormats.Pbgra32);
                botRtb.Render(botV);
                botRtb.Freeze();

                // Schnittlinien die auf die neue Seite überlaufen entfernen
                // (Fraktionen > 1.0 sind durch Math.Min bereits auf 1.0 geklemmmt → diese entfernen)
                _scherenschnitte.RemoveAll(s => s.Seite == si && s.YFraction >= 1.0);

                // Neue Seite nach si einfügen
                EnsureReihenfolge();
                int anzeigePos = _seitenReihenfolge.IndexOf(si);
                int neuIdx = _seitenBilder.Count;
                _seitenBilder.Add(botRtb);      // als neues _seitenBilder-Element (ORIGINAL = Überlauf)
                InitCropEintrag(neuIdx);        // CropArrays erweitern
                _seitenReihenfolge.Insert(anzeigePos >= 0 ? anzeigePos + 1 : _seitenReihenfolge.Count, neuIdx);
                _kompositBilder[neuIdx] = botRtb;

                System.Diagnostics.Debug.WriteLine($"[LEERZEILE] Überlauf: inhaltH={inhaltH}+{stripH}={newH} > origH={origH}, {overH}px auf neue Seite {neuIdx}");
            }

            // ── Reflow-Zweig: Leerzeilen-ContentBlock in _contentBlocks einfügen ──
            // Der Altpfad (Bitmap-Manipulation oben) bleibt vollständig erhalten.
            // Dieser Block hält _contentBlocks synchron, damit ZeicheCanvasReflow()
            // korrekte Ergebnisse liefert.
            if (_contentBlocks != null)
            {
                int insertIdx = FindReflowEinfügeIndex(si, t, oberhalb);
                int newId = _contentBlocks.Count > 0
                    ? _contentBlocks.Max(b => b.BlockId) + 1
                    : 0;
                _contentBlocks.Insert(insertIdx, new ContentBlock
                {
                    BlockId       = newId,
                    SourcePageIdx = -1,
                    FracOben      = 0.0,
                    FracUnten     = 0.0,
                    ExtraHeightPx = stripH
                });
                System.Diagnostics.Debug.WriteLine(
                    $"[REFLOW] Leerzeile B{newId} ({stripH}px) in _contentBlocks[{insertIdx}] eingefügt " +
                    $"(si={si}, t={t}, oberhalb={oberhalb})");
            }

            ZeicheCanvas();
            AktualisiereSchnitteLinien();
            TxtInfo.Text = $"Leerzeile {(oberhalb ? "ober" : "unter")}halb eingefügt – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        private void BtnTeilLöschen_Click(object sender, RoutedEventArgs e)
            => SafeExecute(LöscheAusgewählteParts, "BtnTeilLöschen_Click");

        private void BtnSeiteEinfügen_Click(object sender, RoutedEventArgs e)
            => SafeExecute(() =>
            {
                if (_seitenBilder.Count == 0) return;
                int nachSeitePos;
                if (_markierteSeitenIdx >= 0 && _markierteSeitenIdx < _seitenBilder.Count)
                {
                    // Anzeige-Position der markierten Seite in der Reihenfolge ermitteln
                    var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
                    int anzeigePos = reihenfolge.IndexOf(_markierteSeitenIdx);
                    if (anzeigePos < 0) anzeigePos = reihenfolge.Count;

                    bool vorher = ZeigeBinaryDialog(
                        "Wo einfügen?",
                        $"Neue Seite relativ zu Seite {_markierteSeitenIdx + 1} einfügen:",
                        "Vor dieser Seite",
                        "Nach dieser Seite");
                    // FügeSeiteEinDialog ruft intern FügeInReihenfolgeEin(neuIdx, nachSeitenIdx + 1) auf.
                    // Daher: vorher → anzeigePos - 1 (damit +1 → anzeigePos), nachher → anzeigePos (damit +1 → anzeigePos+1)
                    nachSeitePos = vorher ? Math.Max(0, anzeigePos - 1) : anzeigePos;
                }
                else
                {
                    // Keine Seite markiert: ans Ende anhängen
                    nachSeitePos = _seitenBilder.Count;
                }
                FügeSeiteEinDialog(nachSeitePos);
            }, "BtnSeiteEinfügen_Click");

        private void BtnPdfSpeichern_Click(object sender, RoutedEventArgs e)
            => SafeExecute(SpeicherePdfMitÄnderungen, "BtnPdfSpeichern_Click");

        private void ScrollView_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Z && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
                {
                    if (_undoStack.Count > 0)
                    {
                        var z = _undoStack.Pop();
                        _scherenschnitte.Clear(); _scherenschnitte.AddRange(z.Schnitte);
                        _gelöschteParts.Clear();  foreach (var p in z.Gelöscht) _gelöschteParts.Add(p);
                        _kompositBilder.Clear();  foreach (var kv in z.KompositBilder) _kompositBilder[kv.Key] = kv.Value;
                        _gelöschteSeiten.Clear(); foreach (var s in z.GelöschteSeiten) _gelöschteSeiten.Add(s);
                        _seitenReihenfolge = z.SeitenReihenfolge != null ? new List<int>(z.SeitenReihenfolge) : null;
                        _ausgewählteParts.Clear();
                        // Neumodell nach Altpfad-Undo resynchronisieren (Leerzeilen gehen verloren — Stufe 1)
                        // BEKANNTE EINSCHRÄNKUNG: GapArt/GapMm gehen bei Undo verloren, da ScherenZustand
                        // diese Werte nicht sichert. Nach Undo wird OriginalAbstand verwendet.
                        // Siehe Plan: 2026-04-06-loeschdialog-gap-optionen.md, Abschnitt "Bekannte Einschränkungen".
                        if (_contentBlocks != null)
                        {
                            _contentBlocks = KonvertiereAltesModellZuBlöcken();
                            _nextBlockId   = _contentBlocks.Count > 0 ? _contentBlocks.Max(b => b.BlockId) + 1 : 0;
                        }
                        ZeicheCanvas();
                        TxtInfo.Text = $"Rückgängig ({_undoStack.Count} weitere Schritte)";
                        MarkiereAlsGeändert();
                        e.Handled = true;
                    }
                    else if (_scherenModus && _scherenschnitte.Count > 0)
                    {
                        _scherenschnitte.RemoveAt(_scherenschnitte.Count - 1);
                        // Neumodell nach einfachem Schnitt-Undo resynchronisieren (Leerzeilen gehen verloren — Stufe 1)
                        if (_contentBlocks != null)
                        {
                            _contentBlocks = KonvertiereAltesModellZuBlöcken();
                            _nextBlockId   = _contentBlocks.Count > 0 ? _contentBlocks.Max(b => b.BlockId) + 1 : 0;
                        }
                        AktualisiereSchnitteLinien();
                        TxtInfo.Text = _scherenschnitte.Count > 0
                            ? $"✂ {_scherenschnitte.Count} Schnitt(e) – Strg+Z rückgängig"
                            : "✂ Alle Schnitte rückgängig gemacht";
                        e.Handled = true;
                    }
                }
                else if (e.Key == Key.R && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
                {
                    _reflowDebugModus = !_reflowDebugModus;
                    ResetMausZustandBeimModuswechsel();
                    ZeicheCanvas();
                    e.Handled = true;
                }
                else if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
                {
                    SpeichereÄnderungen(); e.Handled = true;
                }
                else if (e.Key == Key.Delete && !_scherenModus)
                {
                    if (_ausgewählteParts.Count > 0) { LöscheAusgewählteParts(); e.Handled = true; }
                }
                else if (e.Key == Key.Escape)
                {
                    if (_scherenModus) { BeendeScherenModus(); e.Handled = true; }
                    else if (_seitenwechselModus) { BeendeSeitenwechselModus(); e.Handled = true; }
                    else if (_ausgewählteParts.Count > 0)
                    {
                        _ausgewählteParts.Clear();
                        AktualisiereTeilOverlays();
                        TxtInfo.Text = "";
                        e.Handled = true;
                    }
                }
            }
            catch (Exception ex) { LogException(ex, "ScrollView_PreviewKeyDown"); }
        }
        // ── Feature 3: Ganze Seite löschen ───────────────────────────────────────

        private void LöscheSeiteMitBestätigung(int seitenIdx)
        {
            var result = MessageBox.Show(
                $"Seite {seitenIdx + 1} löschen?\n\nDie Seite wird visuell entfernt.\nStrg+Z macht dies rückgängig.",
                "Seite löschen",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes) return;
            _undoStack.Push(SpeichereZustand());
            _gelöschteSeiten.Add(seitenIdx);
            // Schnitte, Parts und Komposit der gelöschten Seite aufräumen
            _scherenschnitte.RemoveAll(s => s.Seite == seitenIdx);
            _gelöschteParts.RemoveWhere(p => p.Seite == seitenIdx);
            _kompositBilder.Remove(seitenIdx);
            ZeicheCanvas();
            TxtInfo.Text = $"Seite {seitenIdx + 1} gelöscht – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        // ── Feature 4: Seite einfügen ─────────────────────────────────────────────

        private void FügeSeiteEinNach(int nachSeitenIdx)
            => SafeExecute(() => FügeSeiteEinDialog(nachSeitenIdx), "FügeSeiteEinNach");

        private void FügeSeiteEinDialog(int nachSeitenIdx)
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0) return;

            // Art der Seite wählen
            var artResult = MessageBox.Show(
                "Leere Seite einfügen?\n\nJa = leere weiße Seite\nNein = Seite aus anderer PDF auswählen",
                "Seitenart", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (artResult == MessageBoxResult.Cancel) return;

            _undoStack.Push(SpeichereZustand());

            if (artResult == MessageBoxResult.Yes)
            {
                // Leere Seite: weißes Bitmap mit Maßen der vorherigen Seite
                int refIdx = Math.Max(0, Math.Min(nachSeitenIdx, _seitenBilder.Count - 1));
                var refBmp = _seitenBilder[refIdx];
                var leereBmp = ErzeugeLeereBitmap(refBmp.PixelWidth, refBmp.PixelHeight);
                int neuIdx = _seitenBilder.Count;
                _seitenBilder.Add(leereBmp);
                InitCropEintrag(neuIdx);
                FügeInReihenfolgeEin(neuIdx, nachSeitenIdx + 1);
                TxtInfo.Text = $"Leere Seite nach Position {nachSeitenIdx + 1} eingefügt – Strg+Z zum Rückgängigmachen";
            }
            else
            {
                // Aus anderer PDF
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title  = "PDF auswählen",
                    Filter = "PDF-Dateien|*.pdf"
                };
                if (dlg.ShowDialog() != true) { if (_undoStack.Count > 0) _undoStack.Pop(); return; }
                string quellPfad = dlg.FileName;

                // Seitenzahl der Quell-PDF ermitteln
                int quellSeitenAnzahl;
                try
                {
                    using var quellDoc = PdfReader.Open(quellPfad, PdfDocumentOpenMode.Import);
                    quellSeitenAnzahl = quellDoc.PageCount;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Öffnen der PDF:\n{App.GetExceptionKette(ex)}", "Fehler",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    if (_undoStack.Count > 0) _undoStack.Pop();
                    return;
                }
                if (quellSeitenAnzahl == 0) { if (_undoStack.Count > 0) _undoStack.Pop(); return; }

                // Bei mehreren Seiten: einfaches Input-Fenster
                int quellSeite = 1;
                if (quellSeitenAnzahl > 1)
                {
                    var inputDlg = ErzeugeEinfacheEingabe(
                        $"Welche Seite der Quell-PDF einfügen? (1–{quellSeitenAnzahl})",
                        "Seite aus PDF", "1");
                    if (inputDlg == null) { if (_undoStack.Count > 0) _undoStack.Pop(); return; }
                    if (!int.TryParse(inputDlg, out quellSeite) || quellSeite < 1 || quellSeite > quellSeitenAnzahl)
                    {
                        MessageBox.Show("Ungültige Seitenzahl.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Warning);
                        if (_undoStack.Count > 0) _undoStack.Pop();
                        return;
                    }
                }

                // Bitmap rendern
                var bitmap = RendereExternSeite(quellPfad, quellSeite - 1);
                if (bitmap == null)
                {
                    MessageBox.Show("Seite konnte nicht gerendert werden.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                    if (_undoStack.Count > 0) _undoStack.Pop();
                    return;
                }
                int neuIdx = _seitenBilder.Count;
                _seitenBilder.Add(bitmap);
                InitCropEintrag(neuIdx);
                _eingefügteSeitenInfo[neuIdx] = (quellPfad, quellSeite - 1);
                FügeInReihenfolgeEin(neuIdx, nachSeitenIdx + 1);
                TxtInfo.Text = $"Seite aus '{IO.Path.GetFileName(quellPfad)}' eingefügt – Strg+Z zum Rückgängigmachen";
            }

            ZeicheCanvas();
            MarkiereAlsGeändert();
        }

        private string ErzeugeEinfacheEingabe(string nachricht, string titel, string standard)
        {
            string result = null;
            bool ok = false;
            var txtBox = new TextBox { Text = standard, Width = 120, Height = 24, VerticalContentAlignment = VerticalAlignment.Center };
            var sp = new StackPanel { Margin = new Thickness(12) };
            sp.Children.Add(new TextBlock { Text = nachricht, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
            sp.Children.Add(txtBox);
            var btnOk     = new Button { Content = "OK", Width = 70, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var btnCancel = new Button { Content = "Abbrechen", Width = 80, IsCancel = true };
            var btnRow    = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            btnRow.Children.Add(btnOk);
            btnRow.Children.Add(btnCancel);
            sp.Children.Add(btnRow);
            Window dlg = null;
            btnOk.Click     += (_, __) => { ok = true; result = txtBox.Text; dlg.Close(); };
            btnCancel.Click += (_, __) => dlg.Close();
            dlg = new Window
            {
                Title = titel, Content = sp,
                SizeToContent = SizeToContent.WidthAndHeight,
                MinWidth = 280,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = Window.GetWindow(this),
                ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
            };
            txtBox.SelectAll();
            txtBox.Focus();
            dlg.ShowDialog();
            return ok ? result : null;
        }

        private BitmapSource ErzeugeLeereBitmap(int breite, int höhe)
        {
            var visual = new DrawingVisual();
            using (var ctx = visual.RenderOpen())
                ctx.DrawRectangle(Brushes.White, null, new Rect(0, 0, breite, höhe));
            var rtb = new RenderTargetBitmap(breite, höhe, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(visual);
            rtb.Freeze();
            return rtb;
        }

        private BitmapSource RendereExternSeite(string pfad, int seitenIdx)
        {
            try
            {
                AppZustand.RenderSem.Wait();
                try
                {
                    using var lib = DocLib.Instance;
                    using var doc = lib.GetDocReader(pfad, new PageDimensions(RenderBreite, RenderBreite * 2));
                    using var page = doc.GetPageReader(seitenIdx);
                    int w = page.GetPageWidth(), h = page.GetPageHeight();
                    if (w <= 0 || h <= 0) return null;
                    var bytes = page.GetImage();
                    var wb = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
                    wb.WritePixels(new Int32Rect(0, 0, w, h), bytes, w * 4, 0);
                    wb.Freeze();
                    return wb;
                }
                finally { AppZustand.RenderSem.Release(); }
            }
            catch { return null; }
        }

        private void InitCropEintrag(int idx)
        {
            // Crop-Arrays um einen Eintrag erweitern
            Array.Resize(ref _cropLinks,  idx + 1);
            Array.Resize(ref _cropRechts, idx + 1);
            Array.Resize(ref _cropOben,   idx + 1);
            Array.Resize(ref _cropUnten,  idx + 1);
            if (_defaultCrop.HasValue)
            {
                _cropLinks[idx]  = _defaultCrop.Value.Links;
                _cropRechts[idx] = _defaultCrop.Value.Rechts;
                _cropOben[idx]   = _defaultCrop.Value.Oben;
                _cropUnten[idx]  = _defaultCrop.Value.Unten;
            }
        }

        private void FügeInReihenfolgeEin(int bitmapIdx, int nachAnzeigePos)
        {
            // _seitenReihenfolge aktivieren falls nötig
            if (_seitenReihenfolge == null)
                _seitenReihenfolge = Enumerable.Range(0, _seitenBilder.Count - 1).ToList(); // ohne neuen Eintrag

            // Einfügeposition bestimmen
            int einfügePos = Math.Min(nachAnzeigePos, _seitenReihenfolge.Count);
            _seitenReihenfolge.Insert(einfügePos, bitmapIdx);
        }

        // ── Feature 5: Drag & Drop (Seiten umsortieren) ───────────────────────────

        private void VerschiebeSeite(int seitenIdx, int richtung)
        {
            EnsureReihenfolge();
            int pos = _seitenReihenfolge.IndexOf(seitenIdx);
            if (pos < 0) return;
            int ziel = pos + richtung;
            if (ziel < 0 || ziel >= _seitenReihenfolge.Count) return;
            _undoStack.Push(SpeichereZustand());
            _seitenReihenfolge.RemoveAt(pos);
            _seitenReihenfolge.Insert(ziel, seitenIdx);
            ZeicheCanvas();
            TxtInfo.Text = "Seite verschoben – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        private void EnsureReihenfolge()
        {
            if (_seitenReihenfolge == null)
                _seitenReihenfolge = Enumerable.Range(0, _seitenBilder.Count).ToList();
        }

        private void StartDragGhost(int seitenIdx, double breite, double höhe, Point pos)
        {
            _dragGhost = new Border
            {
                Tag             = "DRAGGHOST",
                Width           = Math.Min(breite * _zoomFaktor, 120),
                Height          = Math.Min(höhe * _zoomFaktor, 160),
                Background      = new SolidColorBrush(Color.FromArgb(120, 50, 100, 200)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(50, 100, 200)),
                BorderThickness = new Thickness(2),
                IsHitTestVisible = false,
                Child = new TextBlock
                {
                    Text = $"Seite {seitenIdx + 1}",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                    Foreground          = Brushes.White,
                    FontWeight          = FontWeights.Bold
                }
            };
            Canvas.SetLeft(_dragGhost, pos.X - _dragGhost.Width / 2);
            Canvas.SetTop(_dragGhost,  pos.Y - _dragGhost.Height / 2);
            Panel.SetZIndex(_dragGhost, 1000);
            PdfCanvas.Children.Add(_dragGhost);
        }

        private void MoveDragGhost(Point pos)
        {
            if (_dragGhost == null) return;
            Canvas.SetLeft(_dragGhost, pos.X - _dragGhost.Width / 2);
            Canvas.SetTop(_dragGhost,  pos.Y - _dragGhost.Height / 2);
        }

        private void EndDrag(Point dropPos)
        {
            if (_dragGhost != null) { PdfCanvas.Children.Remove(_dragGhost); _dragGhost = null; }
            if (_dragQuellIdx < 0) return;

            // Zielposition bestimmen: welche sichtbare Seite liegt am nächsten?
            var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
            var sichtbar    = reihenfolge.Where(i => !_gelöschteSeiten.Contains(i) && i < _seitenBilder.Count).ToList();

            int zielPos = sichtbar.Count; // Default: ans Ende
            for (int di = 0; di < sichtbar.Count; di++)
            {
                int oi = sichtbar[di];
                if (_seitenYStart[oi] < -100) continue;
                double mitte = _seitenYStart[oi] + _seitenHöhe[oi] / 2;
                if (dropPos.Y < mitte) { zielPos = di; break; }
            }

            EnsureReihenfolge();
            int quellPos = _seitenReihenfolge.IndexOf(_dragQuellIdx);
            if (quellPos < 0 || quellPos == zielPos) return;

            _undoStack.Push(SpeichereZustand());
            _seitenReihenfolge.RemoveAt(quellPos);
            int insertPos = quellPos < zielPos ? zielPos - 1 : zielPos;
            insertPos = Math.Max(0, Math.Min(insertPos, _seitenReihenfolge.Count));
            _seitenReihenfolge.Insert(insertPos, _dragQuellIdx);
            ZeicheCanvas();
            TxtInfo.Text = $"Seite {_dragQuellIdx + 1} verschoben – Strg+Z zum Rückgängigmachen";
            MarkiereAlsGeändert();
        }

        // ── AutoSave-Infrastruktur ────────────────────────────────────────────────

        /// <summary>Initialisiert den Debounce-Timer für AutoSpeichern (1 Sekunde nach letzter Änderung).</summary>
        private void InitAutoSave()
        {
            _autoSaveTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _autoSaveTimer.Tick += (_, __) =>
            {
                _autoSaveTimer.Stop();
                if (_hatUngespeicherteÄnderungen && _pdfPfad != null && !_autoSaveLäuft)
                {
                    _autoSaveLäuft = true;
                    try { AutoSpeichern(); }
                    finally { _autoSaveLäuft = false; }
                }
            };
        }

        /// <summary>Setzt das Dirty-Flag und startet (oder resettet) den AutoSave-Debounce-Timer.</summary>
        private void MarkiereAlsGeändert()
        {
            _hatUngespeicherteÄnderungen = true;
            _hatSitzungsÄnderungen = true;
            _autoSaveTimer?.Stop();
            _autoSaveTimer?.Start();
        }

        /// <summary>Speichert die bearbeitete PDF ohne Dialog als _bearbeitet.pdf (Auto-Save).</summary>
        private void AutoSpeichern()
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[AUTOSAVE] Übersprungen: kein Pfad oder keine Bilder");
                return;
            }
            System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] Starte → {_pdfPfad}");

            try
            {
                // Schritt 1: PDF komplett im Speicher zusammenbauen — KEIN Dateizugriff
                byte[] neueBytes;
                using (var ms = new IO.MemoryStream())
                {
                    SpeicherInStream(ms);
                    neueBytes = ms.ToArray();
                }

                if (neueBytes.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[AUTOSAVE] FEHLER: SpeicherInStream lieferte 0 Bytes");
                    return;
                }

                // Schritt 2: Auf Festplatte schreiben mit Retry bei Sperre
                Exception letzterFehler = null;
                string bearbeitetPfad = BearbeitetPfadFür(_pdfPfad);
                for (int versuch = 1; versuch <= 3; versuch++)
                {
                    try
                    {
                        IO.File.WriteAllBytes(bearbeitetPfad, neueBytes);
                        _pdfBytes = neueBytes;
                        // Schritt 3: SchnittState-JSON nach PDF-Schreiben speichern
                        try { SpeichereSchnittState(); }
                        catch (Exception exJson)
                        {
                            System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] JSON-Fehler (nicht kritisch): {exJson.Message}");
                        }
                        _hatUngespeicherteÄnderungen = false;
                        System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] OK: {bearbeitetPfad} ({neueBytes.Length} Bytes, Versuch {versuch})");
                        return;
                    }
                    catch (IO.IOException ex)
                    {
                        letzterFehler = ex;
                        System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] Versuch {versuch} gesperrt: {ex.Message}");
                        Thread.Sleep(200);
                    }
                }

                // Fallback: unter anderem Namen speichern
                string fallback = IO.Path.Combine(
                    IO.Path.GetDirectoryName(_pdfPfad),
                    IO.Path.GetFileNameWithoutExtension(_pdfPfad) + "_autosave.pdf");
                IO.File.WriteAllBytes(fallback, neueBytes);
                _pdfBytes = neueBytes;
                // SchnittState-JSON auch beim Fallback schreiben
                try { SpeichereSchnittState(); }
                catch (Exception exJson)
                {
                    System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] Fallback JSON-Fehler (nicht kritisch): {exJson.Message}");
                }
                _hatUngespeicherteÄnderungen = false;
                System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] Fallback gespeichert: {fallback}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTOSAVE] FEHLER: {ex.Message}\n{ex.StackTrace}");
                LogException(ex, "AutoSpeichern");
                // KEINE MessageBox — blockiert UI nicht
            }
        }

        /// <summary>Gibt den Pfad zur bearbeiteten PDF-Version zurück.</summary>
        private static string BearbeitetPfadFür(string originalPfad)
        {
            string ordner = IO.Path.GetDirectoryName(originalPfad)!;
            string name   = IO.Path.GetFileNameWithoutExtension(originalPfad);
            return IO.Path.Combine(ordner, name + "_bearbeitet.pdf");
        }

        /// <summary>
        /// Fragt ob ungespeicherte Änderungen gespeichert werden sollen.
        /// Gibt true zurück wenn weitergegangen werden soll (Ja oder Nein),
        /// false wenn beim aktuellen Dokument geblieben werden soll (Abbrechen).
        /// </summary>
        public bool FrageObSpeichern()
        {
            // _hatSitzungsÄnderungen bleibt auch nach AutoSave true – erfasst alle Änderungen
            // dieser Sitzung gegenüber dem Original, unabhängig davon ob AutoSave bereits lief.
            if (!_hatSitzungsÄnderungen) return true;

            var antwort = MessageBox.Show(
                "Die aktuelle PDF wurde in dieser Sitzung bearbeitet.\n\nÄnderungen in '_bearbeitet.pdf' speichern?",
                "Änderungen speichern?",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (antwort == MessageBoxResult.Cancel) return false;

            if (antwort == MessageBoxResult.Yes)
            {
                SpeichereÄnderungen(); // setzt _hatSitzungsÄnderungen = false via SpeichereGesamtzustand
            }
            else // Nein = Verwerfen
            {
                // AutoSave hat evtl. bereits eine _bearbeitet.pdf erstellt → löschen
                if (_pdfPfad != null)
                {
                    string bearbeitetPfad = BearbeitetPfadFür(_pdfPfad);
                    if (System.IO.File.Exists(bearbeitetPfad))
                    {
                        try { System.IO.File.Delete(bearbeitetPfad); }
                        catch { /* Löschen nicht kritisch */ }
                    }
                    // SchnittState-JSON ebenfalls entfernen
                    string jsonPfad = _pdfPfad + ".edit.json";
                    if (System.IO.File.Exists(jsonPfad))
                    {
                        try { System.IO.File.Delete(jsonPfad); }
                        catch { }
                    }
                }
                _hatSitzungsÄnderungen = false;
            }
            return true;
        }

        /// <summary>
        /// Speichert die Änderungen als _bearbeitet.pdf + SchnittState-JSON.
        /// Delegiert an SpeichereGesamtzustand(). Original wird NIE angefasst.
        /// </summary>
        private void SpeichereÄnderungen()
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0) return;

            if (!SpeichereGesamtzustand())
            {
                MessageBox.Show("Speichern fehlgeschlagen. Siehe Log.", "Speichern fehlgeschlagen",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string bearbeitetPfad = BearbeitetPfadFür(_pdfPfad);
            AppZustand.Instanz.SetzeStatus("Gespeichert: " + IO.Path.GetFileName(bearbeitetPfad));
            System.Diagnostics.Debug.WriteLine($"[SAVE] OK → {bearbeitetPfad}");
        }

        /// <summary>
        /// Speichert Schnittlinien und Lösch-Zustand als JSON neben dem Original.
        /// Dateipfad: _pdfPfad + ".edit.json" — immer relativ zum ORIGINAL (nicht _bearbeitet.pdf).
        /// Wird beim AutoSpeichern, expliziten Speichern und SpeichereGesamtzustand aufgerufen.
        /// </summary>
        private void SpeichereSchnittState()
        {
            if (_pdfPfad == null) return;
            try
            {
                string jsonPfad = _pdfPfad + ".edit.json";

                // Meta für Konsistenzprüfung beim Laden
                var fi = new IO.FileInfo(_pdfPfad);
                long   groesse   = fi.Exists ? fi.Length : 0;
                string geaendert = fi.Exists
                    ? fi.LastWriteTimeUtc.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                    : "";
                string dateiName = IO.Path.GetFileName(_pdfPfad);

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("{");

                sb.AppendLine($"  \"meta\": {{\"datei\":\"{EscapeJsonString(dateiName)}\",\"groesse\":{groesse},\"letztGeaendert\":\"{geaendert}\"}},");

                // Schnittlinien
                sb.Append("  \"schnitte\": [");
                var schnittTeile = _scherenschnitte.Select(s =>
                    $"{{\"seite\":{s.Seite},\"fraktion\":{s.YFraction.ToString(System.Globalization.CultureInfo.InvariantCulture)}}}");
                sb.Append(string.Join(",", schnittTeile));
                sb.AppendLine("],");

                // Gelöschte Teile (Seite+TeilIdx)
                sb.Append("  \"geloeschteParts\": [");
                var partTeile = _gelöschteParts.Select(p => $"{{\"seite\":{p.Seite},\"teil\":{p.Teil}}}");
                sb.Append(string.Join(",", partTeile));
                sb.AppendLine("],");

                // Gelöschte Seiten
                sb.Append("  \"geloeschteSeiten\": [");
                sb.Append(string.Join(",", _gelöschteSeiten));
                sb.AppendLine("]");

                sb.AppendLine("}");

                IO.File.WriteAllText(jsonPfad, sb.ToString(), System.Text.Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine($"[SAVE] SchnittState → {jsonPfad}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SAVE] SchnittState-Fehler (nicht kritisch): {ex.Message}");
                // Nicht kritisch — kein MessageBox, Exception nach oben
                throw;
            }
        }

        private static string EscapeJsonString(string s)
            => s.Replace("\\", "\\\\").Replace("\"", "\\\"");

        /// <summary>
        /// Zentrale Speichermethode: schreibt _bearbeitet.pdf UND SchnittState-JSON.
        /// Setzt _hatUngespeicherteÄnderungen = false NUR bei vollständigem Erfolg beider Schritte.
        /// Gibt true zurück wenn alles geklappt hat.
        /// </summary>
        private bool SpeichereGesamtzustand()
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0) return false;

            // Schritt 1: PDF vollständig in Memory bauen
            byte[] neueBytes;
            try
            {
                using var ms = new IO.MemoryStream();
                SpeicherInStream(ms);
                neueBytes = ms.ToArray();
                if (neueBytes.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[GESAMTZUSTAND] FEHLER: SpeicherInStream lieferte 0 Bytes");
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogException(ex, "SpeichereGesamtzustand/Build");
                return false;
            }

            // Schritt 2: PDF auf Disk schreiben
            string bearbeitetPfad = BearbeitetPfadFür(_pdfPfad);
            try
            {
                IO.File.WriteAllBytes(bearbeitetPfad, neueBytes);
            }
            catch (Exception ex)
            {
                LogException(ex, "SpeichereGesamtzustand/PDF");
                return false;
            }
            _pdfBytes = neueBytes;

            // Schritt 3: SchnittState-JSON schreiben
            try
            {
                SpeichereSchnittState();
            }
            catch (Exception ex)
            {
                LogException(ex, "SpeichereGesamtzustand/JSON");
                return false;  // PDF wurde geschrieben, JSON nicht → kein Dirty-Reset
            }

            // Nur bei vollständigem Erfolg dirty-Flags zurücksetzen
            _hatUngespeicherteÄnderungen = false;
            _hatSitzungsÄnderungen = false;
            System.Diagnostics.Debug.WriteLine($"[GESAMTZUSTAND] OK → {bearbeitetPfad}");
            return true;
        }

        /// <summary>
        /// Lädt Schnittlinien und gelöschte Bereiche aus der JSON-Zustandsdatei wieder her.
        /// Muss NACH dem vollständigen Aufbau von _seitenBilder aufgerufen werden (in LadePdf/Dispatcher).
        /// Setzt _scherenschnitte, _gelöschteParts, _gelöschteSeiten — ohne ZeicheCanvas aufzurufen.
        /// _ausgewählteParts wird NICHT persistiert: temporärer UI-Zustand pro Sitzung.
        /// </summary>
        /// <param name="nurSchnittlinien">
        /// Wenn true: nur _scherenschnitte laden, _gelöschteParts und _gelöschteSeiten weglassen.
        /// Wird gesetzt wenn _bearbeitet.pdf geladen wurde — gelöschte Teile sind dort bereits
        /// physisch eingearbeitet, ein erneutes Setzen von _gelöschteParts würde Doppelt-Komposit erzeugen.
        /// </param>
        private void LadeSchnittState(bool nurSchnittlinien = false)
        {
            if (_pdfPfad == null) return;
            string jsonPfad = _pdfPfad + ".edit.json";
            if (!IO.File.Exists(jsonPfad)) return;

            try
            {
                string json = IO.File.ReadAllText(jsonPfad, System.Text.Encoding.UTF8);
                int n = _seitenBilder.Count;

                // Konsistenzprüfung: Meta abgleichen — bei Mismatch wird State nicht angewendet
                if (!PrüfeJsonMeta(json, _pdfPfad))
                {
                    System.Diagnostics.Debug.WriteLine("[LOAD-STATE] Konsistenz-Mismatch: JSON passt nicht zur aktuellen PDF – State wird ignoriert");
                    return;
                }

                // Atomar: erst in lokale Puffer parsen, dann in einem Zug übernehmen.
                // So entsteht kein inkonsistenter Teilzustand wenn ein Abschnitt defekt ist.
                var schnittePuffer = new List<(int Seite, double YFraction)>();
                var partsPuffer    = new HashSet<(int Seite, int Teil)>();
                var seitenPuffer   = new HashSet<int>();

                // Schnitte wiederherstellen
                string schnittAbschnitt = ExtrahiereJsonArray(json, "schnitte");
                if (schnittAbschnitt != null)
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(
                        schnittAbschnitt,
                        @"\{""seite""\s*:\s*(\d+)\s*,\s*""fraktion""\s*:\s*([0-9.eE+\-]+)\s*\}");
                    foreach (System.Text.RegularExpressions.Match m in matches)
                    {
                        int seite = int.Parse(m.Groups[1].Value);
                        double frak = double.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
                        if (seite >= 0 && seite < n && frak > 0.0 && frak < 1.0)
                            schnittePuffer.Add((seite, frak));
                    }
                }

                // Gelöschte Parts wiederherstellen
                string partsAbschnitt = ExtrahiereJsonArray(json, "geloeschteParts");
                if (partsAbschnitt != null)
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(
                        partsAbschnitt,
                        @"\{""seite""\s*:\s*(\d+)\s*,\s*""teil""\s*:\s*(\d+)\s*\}");
                    foreach (System.Text.RegularExpressions.Match m in matches)
                    {
                        int seite = int.Parse(m.Groups[1].Value);
                        int teil  = int.Parse(m.Groups[2].Value);
                        if (seite >= 0 && seite < n)
                            partsPuffer.Add((seite, teil));
                    }
                }

                // Gelöschte Seiten wiederherstellen
                string seitenAbschnitt = ExtrahiereJsonArray(json, "geloeschteSeiten");
                if (seitenAbschnitt != null)
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(seitenAbschnitt, @"\d+");
                    foreach (System.Text.RegularExpressions.Match m in matches)
                    {
                        int seite = int.Parse(m.Value);
                        if (seite >= 0 && seite < n)
                            seitenPuffer.Add(seite);
                    }
                }

                // Atomar übernehmen
                foreach (var s in schnittePuffer) _scherenschnitte.Add(s);
                if (!nurSchnittlinien)
                {
                    // Nur anwenden wenn Original-PDF geladen wurde.
                    // Bei _bearbeitet.pdf sind gelöschte Parts bereits physisch eingearbeitet —
                    // ein erneutes Setzen würde beim Rendering Doppelt-Komposit erzeugen.
                    foreach (var p in partsPuffer)  _gelöschteParts.Add(p);
                    foreach (var s in seitenPuffer) _gelöschteSeiten.Add(s);
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[LOAD-STATE] OK (nurSchnittlinien={nurSchnittlinien}): {_scherenschnitte.Count} Schnitte, " +
                    $"{_gelöschteParts.Count} gelöschte Parts, " +
                    $"{_gelöschteSeiten.Count} gelöschte Seiten aus: {IO.Path.GetFileName(jsonPfad)}");
            }
            catch (Exception ex)
            {
                // Sauber zurückfallen: teilweise befüllten State verwerfen
                _scherenschnitte.Clear();
                _gelöschteParts.Clear();
                _gelöschteSeiten.Clear();
                System.Diagnostics.Debug.WriteLine($"[LOAD-STATE] Fehler (nicht kritisch): {ex.Message}");
                // Editor zeigt unbearbeitete Original-PDF ohne State
            }
        }

        /// <summary>
        /// Prüft ob das Meta-Objekt im JSON zur aktuell geladenen PDF passt.
        /// Fehlendes Meta (altes JSON-Format) gilt als kompatibel.
        /// Gibt false zurück wenn Dateiname oder Größe nicht übereinstimmen.
        /// </summary>
        private static bool PrüfeJsonMeta(string json, string pdfPfad)
        {
            int metaIdx = json.IndexOf("\"meta\"", StringComparison.OrdinalIgnoreCase);
            if (metaIdx < 0) return true; // kein Meta → altes Format, kompatibel laden

            int objStart = json.IndexOf('{', metaIdx + 6);
            if (objStart < 0) return true;
            int objEnd = json.IndexOf('}', objStart + 1);
            if (objEnd   < 0) return true;
            string metaObj = json.Substring(objStart + 1, objEnd - objStart - 1);

            // Dateiname prüfen
            var dateiMatch = System.Text.RegularExpressions.Regex.Match(
                metaObj, @"""datei""\s*:\s*""([^""]+)""");
            if (dateiMatch.Success)
            {
                string jsonDatei = dateiMatch.Groups[1].Value;
                string aktDatei  = IO.Path.GetFileName(pdfPfad);
                if (!string.Equals(jsonDatei, aktDatei, StringComparison.OrdinalIgnoreCase))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[LOAD-STATE] Meta-Mismatch Dateiname: JSON={jsonDatei} / aktuell={aktDatei}");
                    return false;
                }
            }

            // Dateigröße prüfen
            var groesseMatch = System.Text.RegularExpressions.Regex.Match(
                metaObj, @"""groesse""\s*:\s*(\d+)");
            if (groesseMatch.Success && long.TryParse(groesseMatch.Groups[1].Value, out long jsonGroesse))
            {
                var fi = new IO.FileInfo(pdfPfad);
                if (fi.Exists && fi.Length != jsonGroesse)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[LOAD-STATE] Meta-Mismatch Größe: JSON={jsonGroesse} Bytes / aktuell={fi.Length} Bytes");
                    return false;
                }
            }

            return true;
        }

        // ── Reflow-Debug-Renderer (Schritt 4): Probe-Renderpfad ──────────────────────

        /// <summary>
        /// true = Ctrl+Shift+R wurde gedrückt, Canvas zeigt ReflowResult statt altem Modell.
        /// false = normaler ZeicheCanvas()-Pfad (Default).
        /// </summary>
        private bool _reflowDebugModus = false;

        /// <summary>
        /// Bereinigt sämtlichen Maus-/Drag-Zustand beim Wechsel in/aus dem Reflow-Debug-Modus.
        /// Verhindert: hängende PdfCanvas-Captures, stale CropDrag-Handler, veraltete Drag-Flags.
        /// </summary>
        private void ResetMausZustandBeimModuswechsel()
        {
            // Crop-Drag: CaptureMouse freigeben + Handler entfernen
            if (PdfCanvas.IsMouseCaptured) PdfCanvas.ReleaseMouseCapture();
            PdfCanvas.MouseMove         -= CropCanvas_MouseMove;
            PdfCanvas.MouseLeftButtonUp -= CropCanvas_MouseUp;

            // Drag-Flags zurücksetzen (blatt-Drag, Schnitt-Drag)
            _dragAktiv           = false;
            _dragQuellIdx        = -1;
            _gezogeneCropSeite   = null;
            _schnittDragAktiv    = false;
            _gezogenesSchnittIdx = -1;
        }

        /// <summary>
        /// Rendert den Canvas auf Basis des aktuellen ReflowResult (Debug/Test-Pfad).
        /// Ersetzt ZeicheCanvas() NICHT — wird nur via _reflowDebugModus aktiviert.
        ///
        /// Visuelle Konventionen:
        ///   - Jede OutputPage = grauer Hintergrund mit roter Seitennummer
        ///   - Bitmap-Blöcke = CroppedBitmap aus _seitenBilder, mit BlkId-Label
        ///   - Leerzeilen = grünes Rechteck
        ///   - Rote Linie = Seitenumbruch-Grenze
        ///   - Blaue Linie = Block-Anfang (bei Splits sichtbar)
        /// </summary>
        private void ZeicheCanvasReflow()
        {
            try
            {
            ZeicheCanvasReflowIntern();
            }
            catch (Exception ex)
            {
                LogException(ex, "ZeicheCanvasReflow");
                TxtInfo.Text = $"[REFLOW-DEBUG] Fehler: {ex.Message}";
            }
        }

        private void ZeicheCanvasReflowIntern()
        {
            var result = DebugReflowAusContentBlocks();

            PdfCanvas.Children.Clear();
            if (result.Pages.Count == 0)
            {
                TxtInfo.Text = "[REFLOW-DEBUG] Keine Blöcke — Strg+Shift+R zum Beenden";
                return;
            }

            const double seitenAbstand = 20.0;
            const double offsetX       = 20.0;
            double currentY = seitenAbstand;

            for (int pi = 0; pi < result.Pages.Count; pi++)
            {
                var    page  = result.Pages[pi];
                double pageW = page.WidthPx;
                double pageH = page.MaxHeightPx;

                // Seiten-Hintergrund — IsHitTestVisible=false: rein visuell,
                // kein Mausrad-/Klick-Blocking (folgt Konvention aus ZeicheCanvas)
                var bg = new Rectangle
                {
                    Width            = pageW,
                    Height           = pageH,
                    Fill             = new SolidColorBrush(Color.FromRgb(242, 242, 242)),
                    Stroke           = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                    StrokeThickness  = 1,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(bg, offsetX);
                Canvas.SetTop(bg, currentY);
                PdfCanvas.Children.Add(bg);

                // Seiten-Nummer (oben links)
                var pageLabel = new TextBlock
                {
                    Text             = $"Seite {pi + 1}  [{page.Blocks.Count} Block(e), " +
                                       $"gefüllt {page.FilledHeightPx:F0}/{pageH:F0} px]",
                    FontSize         = 11,
                    FontWeight       = FontWeights.Bold,
                    Foreground       = new SolidColorBrush(Color.FromRgb(190, 0, 0)),
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(pageLabel, offsetX + 4);
                Canvas.SetTop(pageLabel, currentY + 2);
                PdfCanvas.Children.Add(pageLabel);

                // PlacedBlocks rendern
                foreach (var pb in page.Blocks)
                {
                    double blockY = currentY + pb.YOffset;

                    // Block-Anfangs-Linie (blau, dünn) — bei geteilten Blöcken sichtbar
                    var startLinie = new Line
                    {
                        X1               = offsetX,
                        Y1               = blockY,
                        X2               = offsetX + pageW,
                        Y2               = blockY,
                        Stroke           = new SolidColorBrush(Color.FromArgb(100, 0, 80, 200)),
                        StrokeThickness  = 1,
                        IsHitTestVisible = false
                    };
                    PdfCanvas.Children.Add(startLinie);

                    if (pb.Block.IsLeerzeile)
                    {
                        // Leerzeile = grünes Rechteck
                        var rect = new Rectangle
                        {
                            Width            = pageW - 4,
                            Height           = pb.HeightPx,
                            Fill             = new SolidColorBrush(Color.FromArgb(80, 0, 180, 0)),
                            Stroke           = new SolidColorBrush(Color.FromArgb(160, 0, 140, 0)),
                            StrokeThickness  = 1,
                            IsHitTestVisible = false
                        };
                        Canvas.SetLeft(rect, offsetX + 2);
                        Canvas.SetTop(rect, blockY);
                        PdfCanvas.Children.Add(rect);

                        var lbl = new TextBlock
                        {
                            Text             = $"Leerzeile {pb.HeightPx:F0}px  (B{pb.Block.BlockId})",
                            FontSize         = 9,
                            Foreground       = new SolidColorBrush(Color.FromRgb(0, 110, 0)),
                            IsHitTestVisible = false
                        };
                        Canvas.SetLeft(lbl, offsetX + 6);
                        Canvas.SetTop(lbl, blockY + 2);
                        PdfCanvas.Children.Add(lbl);
                    }
                    else if (pb.Block.SourcePageIdx >= 0 && pb.Block.SourcePageIdx < _seitenBilder.Count)
                    {
                        var srcBmp = _seitenBilder[pb.Block.SourcePageIdx];
                        int srcH   = srcBmp.PixelHeight;
                        int srcW   = srcBmp.PixelWidth;

                        // Fraktionen zuerst auf [0,1] klemmen
                        double fracO = Math.Max(0.0, Math.Min(1.0, pb.SrcFracOben));
                        double fracU = Math.Max(0.0, Math.Min(1.0, pb.SrcFracUnten));

                        if (fracU <= fracO)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[REFLOW-CROP] Block B{pb.Block.BlockId} ungültig: " +
                                $"SrcFrac {pb.SrcFracOben:F4}–{pb.SrcFracUnten:F4} nach Clamp {fracO:F4}–{fracU:F4} — übersprungen");
                            continue;
                        }

                        // cropY und cropEndY als Pixelpositionen — cropH als Differenz,
                        // damit kein unabhängiges Runden auf Oben+Höhe cropY+cropH > srcH erzeugt
                        int cropY    = (int)(fracO * srcH);
                        int cropEndY = (int)(fracU * srcH);
                        int cropH    = cropEndY - cropY;

                        // Finale Sicherheitsklammer
                        cropY = Math.Max(0, Math.Min(cropY, srcH - 1));
                        cropH = Math.Max(1, Math.Min(cropH, srcH - cropY));

                        if (cropH <= 0)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[REFLOW-CROP] Block B{pb.Block.BlockId} cropH={cropH} nach Clamp — übersprungen");
                            continue;
                        }

                        var cropped = new CroppedBitmap(srcBmp, new Int32Rect(0, cropY, srcW, cropH));
                        var img = new Image
                        {
                            Source           = cropped,
                            Width            = pageW,
                            Height           = pb.HeightPx,
                            Stretch          = Stretch.Fill,
                            IsHitTestVisible = false   // rein visuell — Mausrad bleibt beim ScrollViewer
                        };
                        Canvas.SetLeft(img, offsetX);
                        Canvas.SetTop(img, blockY);
                        PdfCanvas.Children.Add(img);

                        // Block-ID Label (rechts oben, halb-transparent)
                        var blockLbl = new TextBlock
                        {
                            Text             = $"B{pb.Block.BlockId}",
                            FontSize         = 9,
                            Foreground       = new SolidColorBrush(Color.FromArgb(220, 0, 60, 200)),
                            Background       = new SolidColorBrush(Color.FromArgb(140, 255, 255, 255)),
                            IsHitTestVisible = false
                        };
                        Canvas.SetLeft(blockLbl, offsetX + pageW - 30);
                        Canvas.SetTop(blockLbl, blockY + 2);
                        PdfCanvas.Children.Add(blockLbl);
                    }
                }

                // Rote Trennlinie am Seitenende
                var trenn = new Line
                {
                    X1               = offsetX - 6,
                    Y1               = currentY + pageH,
                    X2               = offsetX + pageW + 6,
                    Y2               = currentY + pageH,
                    Stroke           = new SolidColorBrush(Color.FromRgb(200, 0, 0)),
                    StrokeThickness  = 2,
                    IsHitTestVisible = false
                };
                PdfCanvas.Children.Add(trenn);

                currentY += pageH + seitenAbstand;
            }

            double totalW = result.Pages.Count > 0
                ? result.Pages.Max(p => p.WidthPx) + offsetX * 2
                : 400;
            PdfCanvas.Width  = Math.Max(totalW, 400);
            PdfCanvas.Height = Math.Max(currentY, 200);

            TxtInfo.Text = $"[REFLOW-DEBUG] {result.Pages.Count} Ausgabe-Seite(n) — Strg+Shift+R beendet Debug-Modus";
        }

        // ── Reflow-Brücke (Schritt 1–3): Konvertierung Altmodell → ContentBlocks ────

        /// <summary>
        /// Konvertiert das aktuelle fraktionsbasierte Modell (_scherenschnitte, _gelöschteParts)
        /// in eine flache ContentBlock-Liste.
        ///
        /// Reihenfolge: entspricht _seitenReihenfolge (falls aktiv), sonst aufsteigend.
        /// Gelöschte Seiten (_gelöschteSeiten) werden übersprungen.
        /// Gelöschte Teile (_gelöschteParts) werden als IsDeleted=true markiert — nicht entfernt.
        ///
        /// Die Methode hat keine Seiteneffekte auf das Altmodell.
        /// </summary>
        private List<ContentBlock> KonvertiereAltesModellZuBlöcken()
        {
            var blöcke     = new List<ContentBlock>();
            int nextId     = 0;
            var reihenfolge = _seitenReihenfolge
                              ?? Enumerable.Range(0, _seitenBilder.Count).ToList();

            foreach (int si in reihenfolge)
            {
                if (_gelöschteSeiten.Contains(si)) continue;
                if (si < 0 || si >= _seitenBilder.Count) continue;

                var teilGrenzen = GetTeilGrenzen(si);

                for (int t = 0; t < teilGrenzen.Count; t++)
                {
                    var (fracOben, fracUnten) = teilGrenzen[t];
                    bool gelöscht = _gelöschteParts.Contains((si, t));

                    blöcke.Add(new ContentBlock
                    {
                        BlockId       = nextId++,
                        SourcePageIdx = si,
                        FracOben      = fracOben,
                        FracUnten     = fracUnten,
                        IsDeleted     = gelöscht,
                        ExtraHeightPx = 0.0
                    });
                }
            }

            return blöcke;
        }

        /// <summary>
        /// Bestimmt den Einfügeindex in _contentBlocks für eine neue Leerzeile,
        /// ober- oder unterhalb des t-ten Bitmap-Blocks von Seite si.
        ///
        /// Strategie: Bitmap-Blöcke von si werden der Reihe nach gezählt
        /// (Leerzeilen-Blöcke mit SourcePageIdx=-1 werden dabei übersprungen,
        /// bleiben aber an ihrer Position). Der t-te Bitmap-Block ist der Anker:
        ///   oberhalb = true  → direkt vor ihm einfügen
        ///   oberhalb = false → direkt nach ihm einfügen
        ///
        /// Fallback: ans Ende der si-Section (oder Listenende).
        /// </summary>
        private int FindReflowEinfügeIndex(int si, int t, bool oberhalb)
        {
            if (_contentBlocks == null) return 0;

            int bitmapCount = 0;
            for (int i = 0; i < _contentBlocks.Count; i++)
            {
                var b = _contentBlocks[i];
                // Nur echte Bitmap-Blöcke von Seite si als Anker verwenden.
                // Leerzeilen-Blöcke (SourcePageIdx == -1) und Blöcke anderer Seiten überspringen.
                if (b.SourcePageIdx != si || b.IsLeerzeile)
                    continue;

                if (bitmapCount == t)
                    return oberhalb ? i : i + 1;

                bitmapCount++;
            }

            // Fallback: nach dem letzten Block von si einfügen
            for (int i = _contentBlocks.Count - 1; i >= 0; i--)
            {
                if (_contentBlocks[i].SourcePageIdx == si)
                    return i + 1;
            }
            return _contentBlocks.Count;
        }

        /// <summary>
        /// Sucht in _contentBlocks den ersten echten Bitmap-Block von Seite <paramref name="si"/>,
        /// dessen Fraktionsbereich den Schnitt bei <paramref name="yFrac"/> enthält.
        ///
        /// Bedingungen (alle müssen gelten):
        ///   - SourcePageIdx == si  und  SourcePageIdx >= 0
        ///   - !IsLeerzeile
        ///   - !IsDeleted
        ///   - FracOben &lt; yFrac &lt; FracUnten  (echtes Inneres, nicht exakt auf Grenze)
        ///
        /// Gibt den Index in _contentBlocks zurück, oder -1 wenn kein passender Block gefunden.
        /// </summary>
        private int FindBlockFürSchnitt(int si, double yFrac)
        {
            if (_contentBlocks == null) return -1;
            for (int i = 0; i < _contentBlocks.Count; i++)
            {
                var b = _contentBlocks[i];
                if (b.SourcePageIdx != si || b.SourcePageIdx < 0) continue;
                if (b.IsLeerzeile)  continue;
                if (b.IsDeleted)    continue;
                if (yFrac > b.FracOben && yFrac < b.FracUnten)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Setzt IsDeleted am t-ten Bitmap-ContentBlock von Seite si.
        /// Leerzeilen-Blöcke (SourcePageIdx == -1) werden beim Zählen übersprungen.
        /// Keine Wirkung wenn _contentBlocks null ist oder kein passender Block gefunden wird.
        /// </summary>
        private void SetzeContentBlockGelöscht(int si, int t, bool gelöscht)
        {
            if (_contentBlocks == null) return;
            int bitmapCount = 0;
            foreach (var b in _contentBlocks)
            {
                if (b.SourcePageIdx != si || b.IsLeerzeile) continue;
                if (bitmapCount == t)
                {
                    b.IsDeleted = gelöscht;
                    System.Diagnostics.Debug.WriteLine(
                        $"[REFLOW-SYNC] B{b.BlockId} Seite={si} Teil={t} → IsDeleted={gelöscht}");
                    return;
                }
                bitmapCount++;
            }
            System.Diagnostics.Debug.WriteLine(
                $"[REFLOW-SYNC] WARNUNG: Kein ContentBlock für Seite={si} Teil={t} gefunden");
        }

        /// <summary>
        /// Setzt GapArt und GapMm auf dem ContentBlock für (si, t).
        /// Parallele Struktur zu SetzeContentBlockGelöscht.
        /// </summary>
        private void SetzeContentBlockGapInfo(int si, int t, GapModus modus, double mm)
        {
            if (_contentBlocks == null) return;
            int bitmapCount = 0;
            foreach (var b in _contentBlocks)
            {
                if (b.SourcePageIdx != si || b.IsLeerzeile) continue;
                if (bitmapCount == t)
                {
                    b.GapArt = modus;
                    b.GapMm  = mm;
                    System.Diagnostics.Debug.WriteLine(
                        $"[GAP-SYNC] B{b.BlockId} Seite={si} Teil={t} → GapArt={modus} GapMm={mm:F1}");
                    return;
                }
                bitmapCount++;
            }
            System.Diagnostics.Debug.WriteLine(
                $"[GAP-SYNC] WARNUNG: Kein ContentBlock für Seite={si} Teil={t} gefunden");
        }

        /// <summary>
        /// Teilt den ContentBlock, der den Schnitt bei <paramref name="yFrac"/> auf Seite <paramref name="si"/>
        /// enthält, in zwei eigenständige Blöcke auf.
        ///
        /// Block A: FracOben..yFrac   (oben)
        /// Block B: yFrac..FracUnten  (unten, erbt IsDeleted vom Original)
        ///
        /// Leerzeilen-Blöcke werden niemals gesplittet.
        /// Gibt true zurück wenn ein Split stattgefunden hat.
        /// </summary>
        private bool SplitContentBlockBeiSchnitt(int si, double yFrac)
        {
            int idx = FindBlockFürSchnitt(si, yFrac);
            if (idx < 0) return false;

            var orig = _contentBlocks[idx];

            var blockA = new ContentBlock
            {
                BlockId       = _nextBlockId++,
                SourcePageIdx = si,
                FracOben      = orig.FracOben,
                FracUnten     = yFrac,
                IsDeleted     = false,
                ExtraHeightPx = 0.0
            };
            var blockB = new ContentBlock
            {
                BlockId       = _nextBlockId++,
                SourcePageIdx = si,
                FracOben      = yFrac,
                FracUnten     = orig.FracUnten,
                IsDeleted     = orig.IsDeleted,
                ExtraHeightPx = 0.0
            };

            _contentBlocks.RemoveAt(idx);
            _contentBlocks.Insert(idx, blockB);
            _contentBlocks.Insert(idx, blockA);

            System.Diagnostics.Debug.WriteLine(
                $"[SPLIT] Seite={si} yFrac={yFrac:F3} → B{blockA.BlockId}({blockA.FracOben:F3}–{blockA.FracUnten:F3}) + B{blockB.BlockId}({blockB.FracOben:F3}–{blockB.FracUnten:F3})");

            return true;
        }

        /// <summary>
        /// Führt einen Reflow-Lauf auf Basis von _contentBlocks aus.
        /// Gibt das ReflowResult zurück. Ändert NICHTS am Editor-Zustand.
        /// </summary>
        private ReflowResult DebugReflowAusContentBlocks()
        {
            if (_seitenBilder.Count == 0 || _pdfPfad == null || _contentBlocks == null)
                return new ReflowResult();

            var heights = _seitenBilder.Select(b => (double)b.PixelHeight).ToArray();
            double pageH = _seitenBilder.Max(b => (double)b.PixelHeight);
            double pageW = _seitenBilder.Max(b => (double)b.PixelWidth);

            var result = ReflowEngine.RunReflow(_contentBlocks, heights, pageH, pageW);

            System.Diagnostics.Debug.WriteLine(
                ReflowEngine.DebugBeschreibung(result, _contentBlocks));

            return result;
        }

        /// <summary>
        /// Kompatibilitäts-Wrapper: konvertiert das Altmodell frisch in Blöcke und führt Reflow aus.
        /// Nur noch für Vergleichs-/Diagnosezwecke verwenden. Intern bevorzugt: DebugReflowAusContentBlocks().
        /// </summary>
        private ReflowResult DebugReflowAusAltmodell()
        {
            if (_seitenBilder.Count == 0 || _pdfPfad == null) return new ReflowResult();

            var blöcke  = KonvertiereAltesModellZuBlöcken();
            var heights = _seitenBilder.Select(b => (double)b.PixelHeight).ToArray();
            double pageH = _seitenBilder.Max(b => (double)b.PixelHeight);
            double pageW = _seitenBilder.Max(b => (double)b.PixelWidth);

            var result = ReflowEngine.RunReflow(blöcke, heights, pageH, pageW);

            System.Diagnostics.Debug.WriteLine(
                ReflowEngine.DebugBeschreibung(result, blöcke));

            return result;
        }

        /// <summary>
        /// Extrahiert den Inhalt zwischen [ und ] für einen JSON-Schlüssel.
        /// Beispiel: {"schnitte":[...]} → gibt den Inhalt zwischen den eckigen Klammern zurück.
        /// </summary>
        private static string ExtrahiereJsonArray(string json, string schlüssel)
        {
            string suche = $"\"{schlüssel}\"";
            int keyIdx = json.IndexOf(suche, StringComparison.OrdinalIgnoreCase);
            if (keyIdx < 0) return null;
            int start = json.IndexOf('[', keyIdx + suche.Length);
            if (start < 0) return null;
            int end   = json.IndexOf(']', start + 1);
            if (end   < 0) return null;
            return json.Substring(start + 1, end - start - 1);
        }

        /// <summary>
        /// Baut die bearbeitete PDF komplett in den übergebenen MemoryStream —
        /// KEIN Dateizugriff auf _pdfPfad. Identische Logik wie SpeicherNachPfad.
        /// </summary>
        private void SpeicherInStream(IO.MemoryStream zielStream)
        {
            System.Diagnostics.Debug.WriteLine($"[SPEICHER-STREAM] Start, _pdfBytes={(_pdfBytes != null ? _pdfBytes.Length + " Bytes" : "null")}");

            var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
            var sichtbar    = reihenfolge.Where(i => !_gelöschteSeiten.Contains(i) && i < _seitenBilder.Count).ToList();

            var pdfOut = new PdfSharp.Pdf.PdfDocument();

            // Quell-PDF öffnen — IMMER via _pdfBytes, nie _pdfPfad
            PdfSharp.Pdf.PdfDocument pdfIn = null;
            if (_pdfBytes != null)
            {
                try { pdfIn = PdfReader.Open(new IO.MemoryStream(_pdfBytes), PdfDocumentOpenMode.Import); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[SPEICHER-STREAM] PdfReader.Open(bytes) FEHLER: {ex.Message}"); }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[SPEICHER-STREAM] WARNUNG: _pdfBytes ist null, kein pdfIn verfügbar");
            }

            // Seitengröße — via _pdfBytes oder A4-Fallback
            (double pageWPts, double pageHPts) = _pdfBytes != null
                ? HolePdfSeitenGrösse(_pdfBytes)
                : (595.28, 841.89); // A4 Fallback wenn _pdfBytes nicht geladen

            foreach (int si in sichtbar)
            {
                bool hatKomposit        = _kompositBilder.ContainsKey(si);
                bool hatSchnitte        = _scherenschnitte.Any(s => s.Seite == si);
                bool hatGelöschteTeile  = _gelöschteParts.Any(p => p.Seite == si);
                bool istExternEingefügt = _eingefügteSeitenInfo.ContainsKey(si);
                bool hatCrop            = si < _cropLinks.Length
                    && (_cropLinks[si] > 0 || _cropRechts[si] > 0 || _cropOben[si] > 0 || _cropUnten[si] > 0);

                if (hatKomposit || (hatSchnitte && hatGelöschteTeile))
                {
                    BitmapSource exportBmp;
                    if (hatKomposit)
                    {
                        exportBmp = _kompositBilder[si];
                    }
                    else
                    {
                        var grenzen = GetTeilGrenzen(si);
                        var sichtbareTeilIndizes = Enumerable.Range(0, grenzen.Count)
                            .Where(t2 => !_gelöschteParts.Contains((si, t2)))
                            .ToList();
                        exportBmp = sichtbareTeilIndizes.Count > 0
                            ? ErzeugeKompositBild(si, sichtbareTeilIndizes)
                            : null;
                    }

                    if (exportBmp != null)
                    {
                        int origPixH = si < _seitenBilder.Count ? _seitenBilder[si].PixelHeight : exportBmp.PixelHeight;
                        BitmapSource finalExportBmp = exportBmp;
                        if (exportBmp.PixelHeight > origPixH)
                        {
                            var crop = new CroppedBitmap(exportBmp, new Int32Rect(0, 0, exportBmp.PixelWidth, origPixH));
                            crop.Freeze();
                            finalExportBmp = crop;
                        }

                        using var ms = new IO.MemoryStream();
                        var enc = new PngBitmapEncoder();
                        enc.Frames.Add(BitmapFrame.Create(finalExportBmp));
                        enc.Save(ms);
                        ms.Position = 0;
                        using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                        var seite = pdfOut.AddPage();
                        seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(pageWPts);
                        seite.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHPts);
                        using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                        gfx.DrawImage(xImg, 0, 0, pageWPts, pageHPts);
                    }
                    else if (pdfIn != null && si < pdfIn.PageCount)
                    {
                        pdfOut.AddPage(pdfIn.Pages[si]);
                    }
                }
                else if (istExternEingefügt)
                {
                    var bmp = _seitenBilder[si];
                    using var ms = new IO.MemoryStream();
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(bmp));
                    enc.Save(ms);
                    ms.Position = 0;
                    using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                    var (extW, extH) = HolePdfSeitenGrösse(_eingefügteSeitenInfo[si].Pfad);
                    var seite = pdfOut.AddPage();
                    seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(extW);
                    seite.Height = PdfSharp.Drawing.XUnit.FromPoint(extH);
                    using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                    gfx.DrawImage(xImg, 0, 0, extW, extH);
                }
                else if (pdfIn != null && si < pdfIn.PageCount)
                {
                    var seite = pdfOut.AddPage(pdfIn.Pages[si]);
                    if (hatCrop)
                    {
                        double cL = _cropLinks[si], cR = _cropRechts[si];
                        double cO = _cropOben[si],  cU = _cropUnten[si];
                        double x1 = cL * pageWPts;
                        double x2 = (1.0 - cR) * pageWPts;
                        double y1 = cU * pageHPts;
                        double y2 = (1.0 - cO) * pageHPts;
                        seite.CropBox = new PdfSharp.Pdf.PdfRectangle(
                            new PdfSharp.Drawing.XPoint(x1, y1),
                            new PdfSharp.Drawing.XPoint(x2, y2));
                    }
                }
                else
                {
                    // Eingefügte leere Seite: weißes Bitmap
                    var bmp = _seitenBilder[si];
                    using var ms = new IO.MemoryStream();
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(bmp));
                    enc.Save(ms);
                    ms.Position = 0;
                    using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                    var seite = pdfOut.AddPage();
                    seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(pageWPts);
                    seite.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHPts);
                    using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                    gfx.DrawImage(xImg, 0, 0, pageWPts, pageHPts);
                }
            }

            pdfIn?.Dispose();
            pdfOut.Save(zielStream, false);
            System.Diagnostics.Debug.WriteLine($"[SPEICHER-STREAM] Fertig: {zielStream.Length} Bytes");
        }

        private void SpeicherePdfMitÄnderungen()
        {
            if (_pdfPfad == null || _seitenBilder.Count == 0) return;

            // Prüfen ob überhaupt Änderungen vorliegen
            bool hatÄnderungen = _gelöschteSeiten.Count > 0
                || _seitenReihenfolge != null
                || _scherenschnitte.Count > 0
                || _gelöschteParts.Count > 0
                || _kompositBilder.Count > 0
                || _cropLinks.Any(v => v > 0) || _cropRechts.Any(v => v > 0)
                || _cropOben.Any(v => v > 0)  || _cropUnten.Any(v => v > 0)
                || _eingefügteSeitenInfo.Count > 0;

            if (!hatÄnderungen)
            {
                TxtInfo.Text = "Keine Änderungen zum Speichern";
                return;
            }

            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Title            = "Bearbeitete PDF speichern",
                Filter           = "PDF-Dateien|*.pdf",
                FileName         = IO.Path.GetFileNameWithoutExtension(_pdfPfad) + "_bearbeitet.pdf",
                InitialDirectory = IO.Path.GetDirectoryName(_pdfPfad) ?? ""
            };
            if (saveDlg.ShowDialog() != true) return;
            string zielPfad = saveDlg.FileName;
            SpeicherNachPfad(zielPfad, autoSave: false);
        }

        private void SpeicherNachPfad(string zielPfad, bool autoSave)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] SpeicherNachPfad Start: zielPfad={zielPfad}, autoSave={autoSave}");
                if (!autoSave) TxtInfo.Text = "Speichere PDF …";

                var reihenfolge = _seitenReihenfolge ?? Enumerable.Range(0, _seitenBilder.Count).ToList();
                var sichtbar    = reihenfolge.Where(i => !_gelöschteSeiten.Contains(i) && i < _seitenBilder.Count).ToList();

                var pdfOut = new PdfSharp.Pdf.PdfDocument();

                // Quell-PDF öffnen (für Original-Seiten) — immer via _pdfBytes (kein Datei-Handle auf _pdfPfad!)
                PdfSharp.Pdf.PdfDocument pdfIn = null;
                if (_pdfBytes != null)
                {
                    try { pdfIn = PdfReader.Open(new IO.MemoryStream(_pdfBytes, writable: false), PdfDocumentOpenMode.Import); }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] PdfReader.Open(bytes) FEHLER: {ex.Message}"); }
                }
                else if (IO.File.Exists(_pdfPfad))
                {
                    try { pdfIn = PdfReader.Open(_pdfPfad, PdfDocumentOpenMode.Import); }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] PdfReader.Open(pfad) FEHLER: {ex.Message}"); }
                }

                System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] pdfIn={pdfIn != null}, sichtbare Seiten={sichtbar.Count}");
                // Seitengröße via _pdfBytes (kein Datei-Handle)
                (double pageWPts, double pageHPts) = _pdfBytes != null
                    ? HolePdfSeitenGrösse(_pdfBytes)
                    : HolePdfSeitenGrösse(_pdfPfad);

                foreach (int si in sichtbar)
                {
                    bool hatKomposit        = _kompositBilder.ContainsKey(si);
                    bool hatSchnitte        = _scherenschnitte.Any(s => s.Seite == si);
                    bool hatGelöschteTeile  = _gelöschteParts.Any(p => p.Seite == si);
                    bool istExternEingefügt = _eingefügteSeitenInfo.ContainsKey(si);
                    bool hatCrop            = si < _cropLinks.Length
                        && (_cropLinks[si] > 0 || _cropRechts[si] > 0 || _cropOben[si] > 0 || _cropUnten[si] > 0);

                    if (hatKomposit || (hatSchnitte && hatGelöschteTeile))
                    {
                        // Seite mit gelöschten/zusammengeschobenen Teilen:
                        // → IMMER als EINE Seite in ORIGINALFORMAT exportieren.
                        // Gelöschte Bereiche erscheinen als weißer Leerraum.
                        BitmapSource exportBmp;
                        if (hatKomposit)
                        {
                            exportBmp = _kompositBilder[si];
                        }
                        else
                        {
                            // Komposit on-the-fly: sichtbare Teile + weiße Fläche für gelöschte
                            var grenzen = GetTeilGrenzen(si);
                            var sichtbareTeilIndizes = Enumerable.Range(0, grenzen.Count)
                                .Where(t2 => !_gelöschteParts.Contains((si, t2)))
                                .ToList();
                            exportBmp = sichtbareTeilIndizes.Count > 0
                                ? ErzeugeKompositBild(si, sichtbareTeilIndizes)
                                : null;
                        }

                        if (exportBmp != null)
                        {
                            // SEITENFORMAT-INVARIANTE: Ausgabe-Seite hat EXAKT pageWPts × pageHPts.
                            // Das Bitmap wird auf die volle Seitenfläche gezeichnet.
                            // Wenn Komposit kleiner als Original → weißer Leerraum unten (bereits in ErzeugeKompositBild gepaddet).
                            // Wenn Komposit größer als Original → darf nicht vorkommen (FügeLeerzeileEin übernimmt Überlauf auf neue Seite).
                            // Zur Sicherheit: Bitmap auf origH begrenzen bevor Export.
                            int origPixH = si < _seitenBilder.Count ? _seitenBilder[si].PixelHeight : exportBmp.PixelHeight;
                            BitmapSource finalExportBmp = exportBmp;
                            if (exportBmp.PixelHeight > origPixH)
                            {
                                // Bitmap auf origH abschneiden (Invariante erzwingen)
                                var crop = new CroppedBitmap(exportBmp, new Int32Rect(0, 0, exportBmp.PixelWidth, origPixH));
                                crop.Freeze();
                                finalExportBmp = crop;
                                System.Diagnostics.Debug.WriteLine($"[SPEICHERN] Seite {si}: Bitmap auf origH={origPixH} begrenzt (war {exportBmp.PixelHeight})");
                            }

                            using var ms = new IO.MemoryStream();
                            var enc = new PngBitmapEncoder();
                            enc.Frames.Add(BitmapFrame.Create(finalExportBmp));
                            enc.Save(ms);
                            ms.Position = 0;
                            using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                            var seite = pdfOut.AddPage();
                            seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(pageWPts);
                            seite.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHPts);
                            using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                            gfx.DrawImage(xImg, 0, 0, pageWPts, pageHPts);
                        }
                        else if (pdfIn != null && si < pdfIn.PageCount)
                        {
                            pdfOut.AddPage(pdfIn.Pages[si]);
                        }
                    }
                    else if (istExternEingefügt)
                    {
                        // Extern eingefügte Seite als Bitmap
                        var bmp = _seitenBilder[si];
                        using var ms = new IO.MemoryStream();
                        var enc = new PngBitmapEncoder();
                        enc.Frames.Add(BitmapFrame.Create(bmp));
                        enc.Save(ms);
                        ms.Position = 0;
                        using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                        var (extW, extH) = HolePdfSeitenGrösse(_eingefügteSeitenInfo[si].Pfad);
                        var seite = pdfOut.AddPage();
                        seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(extW);
                        seite.Height = PdfSharp.Drawing.XUnit.FromPoint(extH);
                        using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                        gfx.DrawImage(xImg, 0, 0, extW, extH);
                    }
                    else if (pdfIn != null && si < pdfIn.PageCount)
                    {
                        // Original-Seite, ggf. mit CropBox
                        var seite = pdfOut.AddPage(pdfIn.Pages[si]);
                        if (hatCrop)
                        {
                            double cL = _cropLinks[si], cR = _cropRechts[si];
                            double cO = _cropOben[si],  cU = _cropUnten[si];
                            double x1 = cL * pageWPts;
                            double x2 = (1.0 - cR) * pageWPts;
                            double y1 = cU * pageHPts;
                            double y2 = (1.0 - cO) * pageHPts;
                            seite.CropBox = new PdfSharp.Pdf.PdfRectangle(
                                new PdfSharp.Drawing.XPoint(x1, y1),
                                new PdfSharp.Drawing.XPoint(x2, y2));
                        }
                    }
                    else
                    {
                        // Eingefügte leere Seite: weißes Bitmap
                        var bmp = _seitenBilder[si];
                        using var ms = new IO.MemoryStream();
                        var enc = new PngBitmapEncoder();
                        enc.Frames.Add(BitmapFrame.Create(bmp));
                        enc.Save(ms);
                        ms.Position = 0;
                        using var xImg = PdfSharp.Drawing.XImage.FromStream(ms);
                        var seite = pdfOut.AddPage();
                        seite.Width  = PdfSharp.Drawing.XUnit.FromPoint(pageWPts);
                        seite.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHPts);
                        using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(seite);
                        gfx.DrawImage(xImg, 0, 0, pageWPts, pageHPts);
                    }
                }

                pdfIn?.Dispose();
                System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] pdfOut.Save({zielPfad}) startet");
                pdfOut.Save(zielPfad);
                _pdfBytes = IO.File.ReadAllBytes(zielPfad);
                _hatUngespeicherteÄnderungen = false;
                SpeichereSchnittState();
                string msg = autoSave
                    ? $"\u2714 Auto-gespeichert: {IO.Path.GetFileName(zielPfad)}"
                    : $"\u2714 PDF gespeichert: {IO.Path.GetFileName(zielPfad)}";
                TxtInfo.Text = msg;
                AppZustand.Instanz.SetzeStatus(msg);
                System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] SpeicherNachPfad erfolgreich: {zielPfad}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BUG3-SPEICHERN] FEHLER: {ex.Message}");
                LogException(ex, "SpeicherNachPfad");
                if (!autoSave)
                    MessageBox.Show($"Fehler beim Speichern:\n{App.GetExceptionKette(ex)}",
                        "Speicher-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtInfo.Text = "Fehler beim Speichern";
            }
        }
    }
}
