using Docnet.Core;
using Docnet.Core.Models;
using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Modules.Dokumente
{
    public partial class DokumentePanel : UserControl
    {
        // ── Felder ────────────────────────────────────────────────────────────

        private string? _aktiverDateipfad;
        private string? _projektPfad;
        private string  _cacheDir = "";
        private bool    _dokumentGeladen;
        private bool    _panelBereit;
        private string  _filterTyp = "Alle";
        private int     _baumTiefe = 0;   // 0 = alle Ebenen

        // Word-Vorschau
        private List<BitmapSource> _wordSeitenBilder = new();
        private double             _wordZoomFaktor   = 1.0;
        private const int          WordRenderBreite  = 800;

        private string?          _wordBasisPdf;
        private int              _wordRenderBreite = WordRenderBreite;
        private DispatcherTimer? _wordZoomTimer;
        private DispatcherTimer? _selektionDebounce;
        private string?          _selektionPfadPending;
        private CancellationTokenSource? _wordVorschauCts;
        private CancellationTokenSource? _wordZoomCts;

        // CacheBasis und CacheVersion → PdfCache

        private Thread?                  _vorThread;
        private CancellationTokenSource? _vorCts;
        private DispatcherTimer?         _abdeckungTimer;
        private readonly FileWatcherService  _fileWatcher;
        private readonly OrdnerWatcherService _ordnerWatcher;

        // Alle pdfium-Zugriffe laufen über AppZustand.RenderSem,
        // das auch PdfSchnittEditor.LadePdf nutzt → keine parallelen nativen Zugriffe.
        // Word-COM läuft ohne Semaphore (eigene wordApp-Instanz, pdfium-unabhängig).

        // Monoton steigender Zähler: jeder LadeVorschau-Aufruf erhöht ihn.
        // Hintergrund-Threads prüfen vor jedem UI-Update, ob sie noch aktuell sind.
        private volatile int _ladeGeneration;

        // UI-Sperre während Ladevorgang
        private bool             _uiGesperrt;
        private DispatcherTimer? _ladeFallbackTimer;

        // Wird gesetzt wenn nach about:blank automatisch neu geladen werden soll
        private string? _autoRefreshPfad;

        // Mehrfachauswahl im Baum (Ctrl+Klick / Shift+Klick)
        private readonly HashSet<string> _baumMehrfachAuswahl = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string? _baumAuswahlAnker;
        private static readonly Brush _mehrfachHintergrund = new SolidColorBrush(Color.FromArgb(80, 7, 99, 191));

        // Drag & Drop im Baum
        private Point _dragStartPunkt;
        private bool  _dragAktiv;

        public event Action<string?>? DateiAusgewählt;

        // ── Initialisierung ───────────────────────────────────────────────────

        public DokumentePanel()
        {
            InitializeComponent();

            if (Einstellungen.Instanz.DokumentAnsicht == AnsichtModus.Liste)
            {
                RbListe.IsChecked         = true;
                DokumentenBaum.Visibility = Visibility.Collapsed;
                DateiListe.Visibility     = Visibility.Visible;
            }

            foreach (var f in new[] {
                "Alle Dateien",
                "Word (.doc, .docx)",
                "Excel (.xls, .xlsx)",
                "PDF (.pdf)",
                "Bilder (.jpg, .png …)" })
                CbFilter.Items.Add(f);
            CbFilter.SelectedIndex = 0;

            foreach (var t in new[] { "1 Ebene", "2 Ebenen", "3 Ebenen", "Alle" })
                CbBaumTiefe.Items.Add(t);
            CbBaumTiefe.SelectedIndex = 3; // "Alle"

            _panelBereit = true;

            AktualisiereProjectComboBox();

            _fileWatcher = new FileWatcherService(Dispatcher);
            _fileWatcher.DateiGeändert += OnDateiGeändert;

            _ordnerWatcher = new OrdnerWatcherService(Dispatcher);
            _ordnerWatcher.OrdnerGeändert += AktualisiereNurStruktur;

            // UI-Freigabe wenn PdfSchnittEditor SetzeLaden(false) signalisiert
            AppZustand.Instanz.LadeZustandGeändert += aktiv => { if (!aktiv) GibUI(); };
        }

        // ── Öffentliche Aktionen ──────────────────────────────────────────────

        /// <summary>
        /// Speichert den aktuellen Sitzungszustand (Projekt, Datei, Editor-Zustand).
        /// </summary>
        public Core.SitzungsZustand SitzungSpeichern()
        {
            // Editor-Zustand wenn PDF-Vorschau aktiv
            var s = PdfEditor.Visibility == System.Windows.Visibility.Visible
                ? PdfEditor.SitzungSpeichern()
                : new Core.SitzungsZustand();
            s.ProjektPfad = _projektPfad;
            s.AktiveDatei = _aktiverDateipfad;
            return s;
        }

        /// <summary>
        /// Stellt einen gespeicherten Sitzungszustand wieder her (ohne Dialog).
        /// Projekt wird direkt geladen, zuletzt aktive Datei wird neu geöffnet.
        /// </summary>
        public void SitzungWiederherstellen(Core.SitzungsZustand sitzung)
        {
            if (sitzung == null) return;

            // Projekt laden (ohne Dialog, Existenz prüfen)
            if (!string.IsNullOrEmpty(sitzung.ProjektPfad) &&
                System.IO.Directory.Exists(sitzung.ProjektPfad))
            {
                _projektPfad = sitzung.ProjektPfad;
                var hash  = ((uint)sitzung.ProjektPfad!.ToLowerInvariant().GetHashCode()).ToString("X8");
                _cacheDir = System.IO.Path.Combine(Infrastructure.PdfCache.CacheBasis, hash);
                System.IO.Directory.CreateDirectory(_cacheDir);
                AppZustand.Instanz.SetzeProjekt(_projektPfad);
                _ordnerWatcher.Starte(_projektPfad);
                ProjektZurListeHinzufügen(_projektPfad);
                AktualisiereDokumentListe();
                // Keine Vorkonvertierung beim Session-Restore (würde Startzeit verlängern)
            }

            // Editor-Zustand vorbereiten BEVOR die Datei geladen wird
            PdfEditor.SitzungVorbereiten(sitzung);

            // Aktive Datei laden (Existenz prüfen)
            if (!string.IsNullOrEmpty(sitzung.AktiveDatei) &&
                System.IO.File.Exists(sitzung.AktiveDatei))
            {
                LadeVorschau(sitzung.AktiveDatei!);
            }
        }

        public void ProjektLaden()
        {
            var pfad = OrdnerDialog.Zeigen(
                startPfad: Einstellungen.Instanz.StandardPfad ?? "",
                titel:     "Projektordner wählen",
                besitzer:  Window.GetWindow(this));

            if (string.IsNullOrWhiteSpace(pfad)) return;

            _projektPfad = pfad;

            var hash = ((uint)pfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            _cacheDir = Path.Combine(PdfCache.CacheBasis, hash);
            Directory.CreateDirectory(_cacheDir);

            AppZustand.Instanz.SetzeProjekt(_projektPfad);
            _ordnerWatcher.Starte(_projektPfad);
            ProjektZurListeHinzufügen(_projektPfad);
            AktualisiereDokumentListe();
            StartVorkonvertierung(new DirectoryInfo(_projektPfad), _cacheDir);
        }

        // ── Projektverwaltung (P4) ────────────────────────────────────────────

        private bool _cbProjekteAktualisierung; // verhindert Rekursion bei SelectionChanged

        private void ProjektZurListeHinzufügen(string pfad)
        {
            var liste = Einstellungen.Instanz.ProjektPfade;
            // Pfad normalisiert (lowercase) auf Duplikate prüfen
            if (!liste.Any(p => string.Equals(p, pfad, StringComparison.OrdinalIgnoreCase)))
                liste.Add(pfad);
            Einstellungen.Instanz.Speichern();
            AktualisiereProjectComboBox();
        }

        private void AktualisiereProjectComboBox()
        {
            _cbProjekteAktualisierung = true;
            CbProjekte.Items.Clear();

            // Nur existierende Pfade anzeigen
            foreach (var pfad in Einstellungen.Instanz.ProjektPfade.Where(Directory.Exists))
            {
                var item = new ComboBoxItem
                {
                    Content = Path.GetFileName(pfad.TrimEnd(Path.DirectorySeparatorChar)),
                    Tag     = pfad,
                    ToolTip = pfad
                };
                CbProjekte.Items.Add(item);
            }

            // Aktuelles Projekt vorauswählen
            if (_projektPfad != null)
            {
                foreach (ComboBoxItem ci in CbProjekte.Items)
                {
                    if (string.Equals(ci.Tag as string, _projektPfad, StringComparison.OrdinalIgnoreCase))
                    {
                        CbProjekte.SelectedItem = ci;
                        break;
                    }
                }
            }
            _cbProjekteAktualisierung = false;
        }

        private void CbProjekte_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_cbProjekteAktualisierung) return;
            if (CbProjekte.SelectedItem is not ComboBoxItem ci) return;
            if (ci.Tag is not string pfad) return;
            if (string.Equals(pfad, _projektPfad, StringComparison.OrdinalIgnoreCase)) return;

            _projektPfad = pfad;
            var hash = ((uint)pfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            _cacheDir = Path.Combine(PdfCache.CacheBasis, hash);
            Directory.CreateDirectory(_cacheDir);

            AppZustand.Instanz.SetzeProjekt(_projektPfad);
            _ordnerWatcher.Starte(_projektPfad);
            AktualisiereDokumentListe();
        }

        private void BtnProjektHinzufügen_Click(object sender, RoutedEventArgs e)
        {
            var pfad = OrdnerDialog.Zeigen(
                startPfad: Einstellungen.Instanz.StandardPfad ?? "",
                titel:     "Projektordner hinzufügen",
                besitzer:  Window.GetWindow(this));

            if (string.IsNullOrWhiteSpace(pfad)) return;

            ProjektZurListeHinzufügen(pfad);

            // Automatisch auf das neue Projekt wechseln
            _projektPfad = pfad;
            var hash = ((uint)pfad.ToLowerInvariant().GetHashCode()).ToString("X8");
            _cacheDir = Path.Combine(PdfCache.CacheBasis, hash);
            Directory.CreateDirectory(_cacheDir);

            AppZustand.Instanz.SetzeProjekt(_projektPfad);
            _ordnerWatcher.Starte(_projektPfad);
            AktualisiereDokumentListe();
            StartVorkonvertierung(new DirectoryInfo(_projektPfad), _cacheDir);
        }

        private void BtnProjektEntfernen_Click(object sender, RoutedEventArgs e)
        {
            if (CbProjekte.SelectedItem is not ComboBoxItem ci) return;
            if (ci.Tag is not string pfad) return;

            var name = Path.GetFileName(pfad.TrimEnd(Path.DirectorySeparatorChar));
            if (MessageBox.Show(
                    $"Projekt \"{name}\" aus der Liste entfernen?\n\n(Dateien werden nicht gelöscht.)",
                    "Projekt entfernen",
                    MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
                != MessageBoxResult.Yes) return;

            Einstellungen.Instanz.ProjektPfade.RemoveAll(
                p => string.Equals(p, pfad, StringComparison.OrdinalIgnoreCase));
            Einstellungen.Instanz.Speichern();

            // Wenn aktuelles Projekt entfernt: erstes verbleibendes laden oder leeren
            if (string.Equals(pfad, _projektPfad, StringComparison.OrdinalIgnoreCase))
            {
                var erster = Einstellungen.Instanz.ProjektPfade.FirstOrDefault(Directory.Exists);
                if (erster != null)
                {
                    _projektPfad = erster;
                    var hash = ((uint)erster.ToLowerInvariant().GetHashCode()).ToString("X8");
                    _cacheDir = Path.Combine(PdfCache.CacheBasis, hash);
                    Directory.CreateDirectory(_cacheDir);
                    AppZustand.Instanz.SetzeProjekt(_projektPfad);
                    _ordnerWatcher.Starte(_projektPfad);
                    AktualisiereDokumentListe();
                }
                else
                {
                    _projektPfad = null;
                    _ordnerWatcher.Stoppe();
                    DokumentenBaum.Items.Clear();
                    DateiListe.Items.Clear();
                    AppZustand.Instanz.SetzeStatus("Kein Projekt geladen.");
                }
            }

            AktualisiereProjectComboBox();
        }

        public void InWordÖffnen()
        {
            if (_aktiverDateipfad is null) return;

            var ext = Path.GetExtension(_aktiverDateipfad);

            // AxisVM-Projektdateien (.axs) nicht öffnen – würde AxisVM starten
            if (DateiTypen.IstGesperrteExtension(ext))
            {
                AppZustand.Instanz.SetzeStatus(
                    $"{ext}-Dateien werden im Statik-Manager nicht geöffnet.",
                    StatusLevel.Warn);
                return;
            }

            if (DateiTypen.IstPdfDatei(ext))
            {
                // PDF → blockweise als Bilder in Word exportieren
                PdfEditor.ExportierenNachWord();
            }
            else
            {
                // Word-/andere Dateien direkt mit zugewiesener Anwendung öffnen
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName        = _aktiverDateipfad,
                        UseShellExecute = true
                    });
                    AppZustand.Instanz.SetzeStatus("Geöffnet: " + Path.GetFileName(_aktiverDateipfad));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Öffnen:\n" + ex.Message, "Fehler",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Konvertiert eine PDF-Datei in ein bearbeitbares Word-Dokument (Klon).
        /// Word 2013+ analysiert den Inhalt:
        ///   – Text bleibt als echter, editierbarer Text erhalten.
        ///   – Grafiken und nicht konvertierbare Elemente werden als Bilder eingebettet.
        /// Das Ergebnis wird gecacht; bei unveränderter PDF wird der vorhandene Klon direkt geöffnet.
        /// </summary>
        private void ErstelleUndÖffneWordKlon(string pdfPfad)
        {
            var zielPfad = PdfCache.GetWordKlonPfad(pdfPfad, _cacheDir);

            // Bereits vorhandenen Klon direkt öffnen (sofern aktuell)
            if (PdfCache.CacheGültig(zielPfad, pdfPfad))
            {
                WordPdfService.ÖffneInWord(zielPfad);
                return;
            }

            AppZustand.Instanz.SetzeStatus("Starte Word für PDF-Konvertierung …");

            var thread = new Thread(() =>
            {
                // PDF in einen temp-Pfad ohne Sonderzeichen kopieren,
                // damit COM-Interop nicht an Umlauten/Leerzeichen scheitert.
                string tempPdf = Path.Combine(
                    Path.GetTempPath(),
                    "sm_konvert_" + Guid.NewGuid().ToString("N") + ".pdf");

                Word.Application? wordApp = null;
                Word.Document?    wordDoc = null;
                try
                {
                    File.Copy(pdfPfad, tempPdf, overwrite: true);

                    Dispatcher.BeginInvoke(new Action(() =>
                        AppZustand.Instanz.SetzeStatus("Word analysiert PDF …")));

                    // Word sichtbar öffnen – so sieht der Nutzer den Konvertierungsfortschritt
                    // und kann mit eventuellen Word-Dialogen interagieren.
                    wordApp = new Word.Application { Visible = true };
                    wordApp.DisplayAlerts             = Word.WdAlertLevel.wdAlertsAll;
                    wordApp.Options.ConfirmConversions = false;

                    Dispatcher.BeginInvoke(new Action(() =>
                        AppZustand.Instanz.SetzeStatus("Word lädt und analysiert PDF …")));

                    // Word 2013+ wandelt das PDF via „PDF Reflow" um:
                    // Text → echte Absätze, Grafiken → eingebettete Bilder.
                    wordDoc = wordApp.Documents.Open(
                        FileName:           tempPdf,
                        ReadOnly:           false,
                        AddToRecentFiles:   false,
                        ConfirmConversions: false,
                        Format:             Word.WdOpenFormat.wdOpenFormatAuto);

                    Dispatcher.BeginInvoke(new Action(() =>
                        AppZustand.Instanz.SetzeStatus("Speichere Word-Klon …")));

                    // Zielverzeichnis anlegen
                    var dir = Path.GetDirectoryName(zielPfad);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                    wordDoc.SaveAs2(zielPfad, Word.WdSaveFormat.wdFormatXMLDocument);

                    if (!File.Exists(zielPfad))
                        throw new Exception("SaveAs2 hat keine Datei erzeugt. Ziel: " + zielPfad);

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        AppZustand.Instanz.SetzeStatus(
                            "Word-Klon erstellt: " + Path.GetFileName(zielPfad));
                        WordPdfService.ÖffneInWord(zielPfad);
                    }));
                }
                catch (Exception ex)
                {
                    // Vollständige Fehlermeldung inkl. HResult für Diagnose
                    var msg = ex.Message;
                    if (ex is System.Runtime.InteropServices.COMException com)
                        msg += $"\n\nCOM HResult: 0x{com.HResult:X8}";

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        AppZustand.Instanz.SetzeStatus("Fehler bei Konvertierung.", StatusLevel.Error);
                        MessageBox.Show(
                            "PDF → Word Konvertierung fehlgeschlagen:\n\n" + msg
                            + "\n\nVoraussetzung: Microsoft Word 2013 oder neuer.",
                            "Konvertierungsfehler", MessageBoxButton.OK, MessageBoxImage.Error);
                    }));
                }
                finally
                {
                    try { wordDoc?.Close(SaveChanges: false); } catch { }
                    try { wordApp?.Quit(SaveChanges: false); } catch { }
                    if (wordDoc != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc); } catch { }
                    if (wordApp != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    try { File.Delete(tempPdf); } catch { }
                }
            })
            {
                IsBackground = true,
                Name         = "PdfZuWordKlon"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        // ÖffneInWord → WordPdfService.ÖffneInWord

        public void Bereinigen()
        {
            _vorCts?.Cancel();
            _fileWatcher.Dispose();
            _ordnerWatcher.Dispose();
        }

        // ── Hintergrundkonvertierung (Word → Basis-PDF) ───────────────────────

        private void StartVorkonvertierung(DirectoryInfo root, string cacheDir)
        {
            _vorCts?.Cancel();
            var cts = new CancellationTokenSource();
            _vorCts = cts;

            _vorThread = new Thread(() => WordInteropService.WortDateienBatchZuPdf(root, cacheDir, cts.Token))
            {
                IsBackground = true,
                Name         = "StatikVorkonvertierung"
            };
            _vorThread.SetApartmentState(ApartmentState.STA);
            _vorThread.Start();
        }

        // VorkonvertierungTask → WordInteropService.WortDateienBatchZuPdf

        // ── Ansicht-Umschalter ────────────────────────────────────────────────

        private void Ansicht_Geändert(object sender, RoutedEventArgs e)
        {
            if (!_panelBereit) return;

            var modus = RbBaum.IsChecked == true ? AnsichtModus.Baum : AnsichtModus.Liste;
            Einstellungen.Instanz.DokumentAnsicht = modus;
            Einstellungen.Instanz.Speichern();

            DokumentenBaum.Visibility = modus == AnsichtModus.Baum ? Visibility.Visible : Visibility.Collapsed;
            DateiListe.Visibility     = modus == AnsichtModus.Liste ? Visibility.Visible : Visibility.Collapsed;

            _aktiverDateipfad = null;
            _dokumentGeladen  = false;
            DateiAusgewählt?.Invoke(null);
            ZeigeBrowser(mitAbdeckung: false);
            WordVorschau.Navigate("about:blank");

            if (_projektPfad != null) AktualisiereDokumentListe();
        }

        // ── Datei-Typ-Filter ──────────────────────────────────────────────────

        private void CbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_panelBereit) return;

            switch (CbFilter.SelectedIndex)
            {
                case 1: _filterTyp = "Word";  break;
                case 2: _filterTyp = "Excel"; break;
                case 3: _filterTyp = "PDF";   break;
                case 4: _filterTyp = "Bild";  break;
                default: _filterTyp = "Alle"; break;
            }

            if (_projektPfad != null) AktualisiereDokumentListe();
        }

        private void CbBaumTiefe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_panelBereit) return;
            _baumTiefe = CbBaumTiefe.SelectedIndex == 3 ? 0 : CbBaumTiefe.SelectedIndex + 1;
            if (_projektPfad != null) AktualisiereDokumentListe();
        }

        // ── Dokumentenliste ───────────────────────────────────────────────────

        private void AktualisiereDokumentListe()
        {
            if (_projektPfad is null) return;

            // Ausstehende Selektion und Datei-Watcher-Refresh abbrechen,
            // damit kein Dokument das nicht zum neuen Filter passt nachgeladen wird.
            _selektionDebounce?.Stop();
            _selektionPfadPending = null;
            _autoRefreshPfad      = null;
            _baumMehrfachAuswahl.Clear();

            _aktiverDateipfad = null;
            _dokumentGeladen  = false;
            DateiAusgewählt?.Invoke(null);
            ZeigeBrowser(mitAbdeckung: false);
            WordVorschau.Navigate("about:blank");

            var root = new DirectoryInfo(_projektPfad);

            if (RbBaum.IsChecked == true) ZeigeBaum(root);
            else                          ZeigeListe(root);

            AppZustand.Instanz.SetzeStatus($"{ZähleDateien(root)} Datei(en) gefunden.");
        }

        /// <summary>
        /// Baut nur den Baum/die Liste neu – ohne aktive Vorschau oder Auswahl zu löschen.
        /// Wird vom OrdnerWatcher aufgerufen, damit externe Dateiänderungen die Vorschau nicht unterbrechen.
        /// </summary>
        private void AktualisiereNurStruktur()
        {
            if (_projektPfad is null) return;
            _baumMehrfachAuswahl.Clear();
            var root = new DirectoryInfo(_projektPfad);
            if (RbBaum.IsChecked == true) ZeigeBaum(root);
            else                          ZeigeListe(root);
            AppZustand.Instanz.SetzeStatus($"{ZähleDateien(root)} Datei(en) gefunden.");
        }

        private void ZeigeBaum(DirectoryInfo root)
        {
            DokumentenBaum.Items.Clear();
            var rootItem = new TreeViewItem { Header = "📁 " + root.Name, IsExpanded = true };
            FülleBaumItem(rootItem, root);
            DokumentenBaum.Items.Add(rootItem);
        }

        private void ZeigeListe(DirectoryInfo root)
        {
            DateiListe.Items.Clear();

            foreach (var file in root.EnumerateFiles("*.*", SearchOption.AllDirectories)
                                     .Where(f => PasstZuFilter(f.Extension))
                                     .OrderBy(f => f.Name))
            {
                var ordnerVoll = file.DirectoryName ?? "";
                var relOrdner  = ordnerVoll.StartsWith(root.FullName, StringComparison.OrdinalIgnoreCase)
                    ? ordnerVoll.Substring(root.FullName.Length).TrimStart(Path.DirectorySeparatorChar)
                    : ordnerVoll;
                if (string.IsNullOrEmpty(relOrdner)) relOrdner = "\\";

                DateiListe.Items.Add(new DateiEintrag
                {
                    Dateiname     = file.Name,
                    OrdnerRelativ = relOrdner,
                    VollerPfad    = file.FullName
                });
            }
        }

        private void FülleBaumItem(TreeViewItem parent, DirectoryInfo dir, int tiefe = 0)
        {
            IEnumerable<DirectoryInfo> unterordner;
            try   { unterordner = dir.GetDirectories().OrderBy(d => d.Name); }
            catch (Exception ex) when (ex is DirectoryNotFoundException || ex is UnauthorizedAccessException || ex is PathTooLongException)
            { unterordner = Enumerable.Empty<DirectoryInfo>(); }

            foreach (var sub in unterordner)
            {
                if (!EnthältDateien(sub)) continue;
                var item = new TreeViewItem
                {
                    Header     = "📁 " + sub.Name,
                    IsExpanded = _baumTiefe == 0 || tiefe < _baumTiefe - 1,
                    Tag        = sub.FullName
                };
                FülleBaumItem(item, sub, tiefe + 1);
                parent.Items.Add(item);
            }

            IEnumerable<FileInfo> dateien;
            try   { dateien = dir.GetFiles().Where(f => PasstZuFilter(f.Extension)).OrderBy(f => f.Name); }
            catch (Exception ex) when (ex is DirectoryNotFoundException || ex is UnauthorizedAccessException || ex is PathTooLongException)
            { dateien = Enumerable.Empty<FileInfo>(); }

            foreach (var file in dateien)
            {
                parent.Items.Add(new TreeViewItem
                {
                    Header = DateiTypen.DateiIcon(file.Extension) + " " + file.Name,
                    Tag    = file.FullName
                });
            }
        }

        private bool EnthältDateien(DirectoryInfo dir)
        {
            try
            {
                return dir.EnumerateFiles("*.*", SearchOption.AllDirectories).Any(f => PasstZuFilter(f.Extension));
            }
            catch (Exception ex) when (ex is DirectoryNotFoundException || ex is UnauthorizedAccessException || ex is PathTooLongException)
            {
                return false;
            }
        }

        private int ZähleDateien(DirectoryInfo dir)
        {
            try
            {
                return dir.EnumerateFiles("*.*", SearchOption.AllDirectories).Count(f => PasstZuFilter(f.Extension));
            }
            catch (Exception ex) when (ex is DirectoryNotFoundException || ex is UnauthorizedAccessException || ex is PathTooLongException)
            {
                return 0;
            }
        }

        // ── Filter- und Typ-Hilfsmethoden ─────────────────────────────────────

        private bool PasstZuFilter(string ext)
        {
            var e = ext.ToLowerInvariant();
            switch (_filterTyp)
            {
                case "Word":  return e == ".doc" || e == ".docx";
                case "Excel": return e == ".xls" || e == ".xlsx" || e == ".xlsm";
                case "PDF":   return e == ".pdf";
                case "Bild":  return e == ".jpg" || e == ".jpeg" || e == ".png"
                                  || e == ".gif" || e == ".bmp"  || e == ".tif" || e == ".tiff";
                default:      return true;
            }
        }

        // DateiIcon, IstWordDatei, IstPdfDatei, IstBildDatei → DateiTypen

        // ── Auswahl ───────────────────────────────────────────────────────────

        private void StarteSelektionDebounce(string pfad)
        {
            if (_uiGesperrt) return;
            _selektionPfadPending = pfad;
            if (_selektionDebounce == null)
            {
                _selektionDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
                _selektionDebounce.Tick += (_, _) =>
                {
                    _selektionDebounce.Stop();
                    var p = _selektionPfadPending;
                    if (p == null) return;
                    try { LadeVorschau(p); }
                    catch (Exception ex)
                    {
                        Logger.Fehler("SelektionDebounce", ex.Message);
                        AppZustand.Instanz.SetzeStatus("Fehler: " + ex.Message, StatusLevel.Error);
                    }
                };
            }
            _selektionDebounce.Stop();
            _selektionDebounce.Start();
        }

        private void DokumentenBaum_SelectedItemChanged(object sender,
            RoutedPropertyChangedEventArgs<object> e)
        {
            if (DokumentenBaum.SelectedItem is not TreeViewItem item) return;
            if (item.Tag is not string pfad) return;
            StarteSelektionDebounce(pfad);
        }

        private void DateiListe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DateiListe.SelectedItem is not DateiEintrag eintrag) return;
            StarteSelektionDebounce(eintrag.VollerPfad);
        }

        private void LadeVorschau(string pfad)
        {
            _ladeGeneration++;
            _aktiverDateipfad = pfad;
            _dokumentGeladen  = false;
            DateiAusgewählt?.Invoke(pfad);
            _fileWatcher.Starte(pfad);
            SperreUI();

            try
            {
                switch (DocumentRoutingService.ErmittleVorschauTyp(pfad))
                {
                    case VorschauTyp.SchnittEditor:
                        // PDF → Schnitt-Editor: laufende Word-Threads abbrechen,
                        // damit kein paralleler pdfium-Zugriff entsteht.
                        _wordZoomCts?.Cancel();
                        _wordVorschauCts?.Cancel();
                        ZeigeSchnittEditor();
                        AppZustand.Instanz.SetzeStatus("Lade PDF: " + Path.GetFileName(pfad) + " …");
                        PdfEditor.LadePdf(pfad);
                        // GibUI() kommt via AppZustand.LadeZustandGeändert aus PdfSchnittEditor
                        break;
                    case VorschauTyp.WordVorschau:
                        ZeigeWordInfo(pfad);
                        AppZustand.Instanz.SetzeStatus("Word: " + Path.GetFileName(pfad));
                        // GibUI() kommt aus den Dispatcher-Callbacks in StarteWordVorschauRendern
                        break;
                    case VorschauTyp.Browser:
                        ZeigeBrowser(mitAbdeckung: false);
                        AppZustand.Instanz.SetzeStatus("Lade: " + Path.GetFileName(pfad) + " …");
                        WordVorschau.Navigate(new Uri(pfad));
                        GibUI(); // synchron: Navigation ist fire-and-forget
                        break;
                    case VorschauTyp.JsonVorschau:
                        ZeigeBrowser(mitAbdeckung: false);
                        ZeigeJsonVorschau(pfad);
                        GibUI(); // synchron: kein Hintergrundladevorgang
                        break;
                    case VorschauTyp.KeinVorschau:
                        ZeigeBrowser(mitAbdeckung: false);
                        ZeigeKeinVorschauHinweis(pfad);
                        GibUI(); // synchron: kein Hintergrundladevorgang
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Fehler("LadeVorschau", App.GetExceptionKette(ex));
                AppZustand.Instanz.SetzeStatus("Fehler beim Laden: " + ex.Message, StatusLevel.Error);
                try { ZeigeBrowser(mitAbdeckung: false); } catch { }
                try
                {
                    WordVorschau.NavigateToString(HtmlSeite(
                        "<b>Fehler beim Laden der Vorschau</b><br><br>"
                        + HtmlEncode(ex.Message),
                        bodyStyle: "font-family:Segoe UI;padding:30px;color:#c00"));
                }
                catch { }
                GibUI(); // Sicherheitsnetz: Sperre auch bei unerwarteten Ausnahmen aufheben
            }
        }

        // ── UI-Sperre ─────────────────────────────────────────────────────────

        private void SperreUI()
        {
            if (_uiGesperrt) return;
            _uiGesperrt = true;
            LinkeSeite.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;

            // Fallback: Sperre nach 30 s automatisch aufheben (Schutz vor Hängern)
            _ladeFallbackTimer?.Stop();
            _ladeFallbackTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
            _ladeFallbackTimer.Tick += (_, _) => { _ladeFallbackTimer!.Stop(); GibUI(); };
            _ladeFallbackTimer.Start();
        }

        private void GibUI()
        {
            _ladeFallbackTimer?.Stop();
            if (!_uiGesperrt) return;
            _uiGesperrt = false;
            LinkeSeite.IsEnabled = true;
            Mouse.OverrideCursor = null;
        }

        private void ZeigeSchnittEditor()
        {
            WordVorschau.Navigate("about:blank");
            WordVorschau.Visibility    = Visibility.Collapsed;
            AbdeckungsPanel.Visibility = Visibility.Collapsed;
            WordInfoPanel.Visibility   = Visibility.Collapsed;
            PdfEditor.Visibility       = Visibility.Visible;
        }

        private void ZeigeBrowser(bool mitAbdeckung)
        {
            PdfEditor.Visibility       = Visibility.Collapsed;
            WordInfoPanel.Visibility   = Visibility.Collapsed;
            WordVorschau.Visibility    = Visibility.Visible;
            AbdeckungsPanel.Visibility = mitAbdeckung ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ZeigeWordInfo(string pfad)
        {
            WordVorschau.Navigate("about:blank");
            WordVorschau.Visibility    = Visibility.Collapsed;
            AbdeckungsPanel.Visibility = Visibility.Collapsed;
            PdfEditor.Visibility       = Visibility.Collapsed;
            TxtWordDateiname.Text      = Path.GetFileName(pfad);
            TxtWordLadeStatus.Text     = "Vorschau wird geladen …";
            _wordZoomTimer?.Stop();
            _wordZoomCts?.Cancel();
            _wordVorschauCts?.Cancel();
            _wordBasisPdf     = null;
            _wordRenderBreite = WordRenderBreite;
            WordSeitenPanel.Children.Clear();
            _wordSeitenBilder.Clear();
            WordInfoPanel.Visibility   = Visibility.Visible;
            StarteWordVorschauRendern(pfad);
        }

        private void StarteWordVorschauRendern(string pfad)
        {
            _wordVorschauCts?.Cancel();
            var cts = new CancellationTokenSource();
            _wordVorschauCts = cts;
            var token = cts.Token;
            int myGen = _ladeGeneration;   // Generation bei Aufruf festhalten

            var cacheDir = string.IsNullOrEmpty(_cacheDir)
                ? Path.Combine(Path.GetTempPath(), "StatikManager", "preview")
                : _cacheDir;

            App.LogFehler("WordVorschau", $"[DEBUG] Starte: {Path.GetFileName(pfad)}");

            var t = new Thread(() =>
            {
                // ── Phase 1: Word COM — kein Semaphore ───────────────────────────
                // WortDateiZuPdf ist nicht abbrechbar, blockiert aber kein pdfium.
                // Jede Instanz nutzt eine eigene Word.Application → parallele Aufrufe ok.
                try
                {
                    if (token.IsCancellationRequested) return;

                    var basisPdf = PdfCache.GetBasisPdfPfad(pfad, cacheDir);

                    if (!PdfCache.CacheGültig(basisPdf, pfad))
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                            TxtWordLadeStatus.Text = "Erzeuge Vorschau (Word wird gestartet) …";
                        }));
                        Directory.CreateDirectory(Path.GetDirectoryName(basisPdf)!);
                        if (token.IsCancellationRequested) { App.LogFehler("WordVorschau", $"[DEBUG] Abgebrochen vor Word: {Path.GetFileName(pfad)}"); return; }
                        WordInteropService.WortDateiZuPdf(pfad, basisPdf);
                    }

                    // Frühzeitiger Ausstieg direkt nach Word COM: Semaphore wird nie erworben
                    if (token.IsCancellationRequested) { App.LogFehler("WordVorschau", $"[DEBUG] Abgebrochen nach Word: {Path.GetFileName(pfad)}"); return; }
                    if (!File.Exists(basisPdf)) return;

                    // ── Phase 2: pdfium Rendering — mit Semaphore ─────────────────
                    try { AppZustand.RenderSem.Wait(token); }
                    catch (OperationCanceledException)
                    {
                        Logger.Info("WordVorschau", $"[Sem] Abgebrochen beim Warten: {Path.GetFileName(pfad)}");
                        return;
                    }
                    try
                    {
                    try
                    {
                        if (token.IsCancellationRequested) return;

                        var bilder = new List<BitmapSource>();
                        var lib             = DocLib.Instance;
                        using var docReader = lib.GetDocReader(
                            basisPdf, new PageDimensions(WordRenderBreite, WordRenderBreite * 2));
                        int n = docReader.GetPageCount();

                        for (int i = 0; i < n; i++)
                        {
                            if (token.IsCancellationRequested) return;
                            try
                            {
                                using var pageReader = docReader.GetPageReader(i);
                                var raw = pageReader.GetImage();
                                int w = pageReader.GetPageWidth(), h = pageReader.GetPageHeight();
                                if (raw == null || w <= 0 || h <= 0 || raw.Length < w * h * 4) continue;

                                PdfRenderer.KompositioniereGegenWeiss(raw, w, h);
                                var bmp = BitmapSource.Create(w, h, 96, 96,
                                    PixelFormats.Bgra32, null, raw, w * 4);
                                bmp.Freeze();
                                bilder.Add(bmp);

                                if (i == 0)
                                {
                                    var ersteSeite = bmp;
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                                        _wordSeitenBilder = new List<BitmapSource> { ersteSeite };
                                        BaueWordSeitenPanel();
                                        TxtWordLadeStatus.Text = $"Lade … (1/{n})";
                                        Dispatcher.BeginInvoke(
                                            new Action(() => AktualisiereWordZoom(WordAnsicht.Seitenbreite)),
                                            System.Windows.Threading.DispatcherPriority.Loaded);
                                    }));
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Warn("WordVorschau", $"Seite {i + 1} übersprungen: {ex.Message}");
                            }
                        }

                        if (token.IsCancellationRequested) return;

                        var basisPdfFinal = basisPdf;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                            try
                            {
                                _wordBasisPdf     = basisPdfFinal;
                                _wordSeitenBilder = bilder;
                                BaueWordSeitenPanel();
                                TxtWordLadeStatus.Text = bilder.Count > 0
                                    ? $"{bilder.Count} Seite(n)"
                                    : "Vorschau nicht verfügbar";
                                App.LogFehler("WordVorschau", $"[DEBUG] Fertig: {Path.GetFileName(pfad)}, {bilder.Count} Seiten");
                                Dispatcher.BeginInvoke(
                                    new Action(() => AktualisiereWordZoom(WordAnsicht.Seitenbreite)),
                                    System.Windows.Threading.DispatcherPriority.Loaded);
                            }
                            finally { GibUI(); }
                        }));
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex)
                    {
                        App.LogFehler("StarteWordVorschauRendern", App.GetExceptionKette(ex));
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                            TxtWordLadeStatus.Text = "Vorschau nicht verfügbar";
                            GibUI();
                        }));
                    }
                    }
                    finally
                    {
                        AppZustand.RenderSem.Release();
                        Logger.Info("WordVorschau", $"[Sem] Freigegeben: {Path.GetFileName(pfad)}");
                    }
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    // Fehler in Phase 1 (Word COM) — kein Semaphore beteiligt
                    App.LogFehler("StarteWordVorschauRendern/WordCOM", App.GetExceptionKette(ex));
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (token.IsCancellationRequested || myGen != _ladeGeneration) return;
                        TxtWordLadeStatus.Text = "Vorschau nicht verfügbar";
                        GibUI();
                    }));
                }
            })
            { IsBackground = true, Name = "WordVorschauRendern" };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        private void BaueWordSeitenPanel()
        {
            WordSeitenPanel.Children.Clear();
            foreach (var bmp in _wordSeitenBilder)
            {
                var img = new Image { Source = bmp, Stretch = Stretch.None };
                RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
                var shadow = new System.Windows.Media.Effects.DropShadowEffect
                {
                    ShadowDepth = 3, BlurRadius = 8, Opacity = 0.35,
                    Color = System.Windows.Media.Colors.Black
                };
                var border = new Border
                {
                    Background    = Brushes.White,
                    Margin        = new Thickness(0, 0, 0, 16),
                    Effect        = shadow
                };
                border.Child = img;
                WordSeitenPanel.Children.Add(border);
            }
        }

        private void AktualisiereWordZoom(WordAnsicht ansicht)
        {
            if (_wordSeitenBilder.Count == 0) return;

            double verfügbareBreite = Math.Max(1, WordScrollViewer.ActualWidth  - 48);
            double verfügbareHöhe   = Math.Max(1, WordScrollViewer.ActualHeight - 48);
            double seitenBreite     = _wordSeitenBilder[0].PixelWidth;
            double seitenHöhe       = _wordSeitenBilder[0].PixelHeight;

            double zoom;
            switch (ansicht)
            {
                case WordAnsicht.Prozent100:
                    zoom = 1.0;
                    break;
                case WordAnsicht.Seitenbreite:
                    zoom = seitenBreite > 0 ? verfügbareBreite / seitenBreite : 1.0;
                    break;
                case WordAnsicht.EineSeite:
                    double zW = seitenBreite > 0 ? verfügbareBreite / seitenBreite : 1.0;
                    double zH = seitenHöhe   > 0 ? verfügbareHöhe   / seitenHöhe  : 1.0;
                    zoom = Math.Min(zW, zH);
                    break;
                case WordAnsicht.MehrereSeiten:
                    // ~2,5 Seiten pro Zeile sichtbar
                    zoom = seitenBreite > 0 ? verfügbareBreite / (seitenBreite * 2.5) : 0.4;
                    break;
                default:
                    zoom = 1.0;
                    break;
            }

            SetzeWordZoom(Math.Max(0.05, Math.Min(4.0, zoom)));
        }

        private void SetzeWordZoom(double zoom)
        {
            _wordZoomFaktor = zoom;
            // ScaleTransform relativ zur aktuellen Renderbreite (1.0 wenn neu gerendert)
            double scale = zoom * WordRenderBreite / _wordRenderBreite;
            WordZoomTransform.ScaleX = scale;
            WordZoomTransform.ScaleY = scale;
            TxtWordZoom.Text = $"{zoom * 100:0}%";

            // Debounce: nach kurzer Pause mit höherer Auflösung neu rendern
            if (_wordZoomTimer == null)
            {
                _wordZoomTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
                _wordZoomTimer.Tick += WordZoomTimer_Tick;
            }
            _wordZoomTimer.Stop();
            _wordZoomTimer.Start();
        }

        private void WordZoomTimer_Tick(object? sender, EventArgs e)
        {
            _wordZoomTimer!.Stop();
            if (_wordBasisPdf == null || !File.Exists(_wordBasisPdf) || _wordSeitenBilder.Count == 0) return;

            // Ziel-Renderbreite: mindestens Basis, maximal 3200 px
            int ziel = Math.Min(3200, Math.Max(WordRenderBreite,
                (int)Math.Round(WordRenderBreite * _wordZoomFaktor)));

            // Kein Neu-Render wenn Änderung < 15 %
            if (Math.Abs(ziel - _wordRenderBreite) < (int)(0.15 * _wordRenderBreite)) return;

            _wordZoomCts?.Cancel();
            var cts = new CancellationTokenSource();
            _wordZoomCts = cts;

            var basisPdf  = _wordBasisPdf;
            var pfad      = _aktiverDateipfad;
            var zielLocal = ziel;
            var token     = cts.Token;

            int myGen = _ladeGeneration;
            var t = new Thread(() => RendereWordSeiten(basisPdf, zielLocal, pfad!, token, myGen))
            { IsBackground = true, Name = "WordZoomRerender" };
            t.Start();
        }

        private void RendereWordSeiten(string basisPdf, int renderBreite, string pfad,
                                       CancellationToken token = default, int gen = 0)
        {
            try { AppZustand.RenderSem.Wait(token); }
            catch (OperationCanceledException) { return; }
            try
            {
            try
            {
                if (token.IsCancellationRequested) return;
                var bilder = new List<BitmapSource>();
                var lib = DocLib.Instance;
                using var docReader = lib.GetDocReader(
                    basisPdf, new PageDimensions(renderBreite, renderBreite * 2));
                int n = docReader.GetPageCount();

                for (int i = 0; i < n; i++)
                {
                    if (token.IsCancellationRequested) return;
                    try
                    {
                        using var pageReader = docReader.GetPageReader(i);
                        var raw = pageReader.GetImage();
                        int w = pageReader.GetPageWidth(), h = pageReader.GetPageHeight();
                        if (raw == null || w <= 0 || h <= 0 || raw.Length < w * h * 4) continue;

                        PdfRenderer.KompositioniereGegenWeiss(raw, w, h);
                        var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, raw, w * 4);
                        bmp.Freeze();
                        bilder.Add(bmp);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn("RendereWordSeiten", $"Seite {i + 1} übersprungen: {ex.Message}");
                    }
                }

                if (bilder.Count == 0 || token.IsCancellationRequested) return;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (token.IsCancellationRequested || _aktiverDateipfad != pfad || gen != _ladeGeneration) return;
                    _wordSeitenBilder  = bilder;
                    _wordRenderBreite  = renderBreite;
                    BaueWordSeitenPanel();
                    // ScaleTransform neu berechnen (sollte ~1.0 sein)
                    double scale = _wordZoomFaktor * WordRenderBreite / _wordRenderBreite;
                    WordZoomTransform.ScaleX = scale;
                    WordZoomTransform.ScaleY = scale;
                }));
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                App.LogFehler("RendereWordSeiten", App.GetExceptionKette(ex));
                Dispatcher.BeginInvoke(new Action(() =>
                    AppZustand.Instanz.SetzeStatus("Fehler: Zoom-Vorschau konnte nicht aktualisiert werden.", StatusLevel.Error)));
            }
            }
            finally { AppZustand.RenderSem.Release(); }
        }

        private void WordScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == 0) return;
            e.Handled = true;
            double neuerZoom = e.Delta > 0
                ? Math.Min(4.0, _wordZoomFaktor * 1.12)
                : Math.Max(0.05, _wordZoomFaktor / 1.12);
            SetzeWordZoom(neuerZoom);
        }

        private void BtnWordZoomMinus_Click(object sender, RoutedEventArgs e)
            => SetzeWordZoom(Math.Max(0.05, _wordZoomFaktor / 1.25));

        private void BtnWordZoomPlus_Click(object sender, RoutedEventArgs e)
            => SetzeWordZoom(Math.Min(4.0, _wordZoomFaktor * 1.25));

        private void BtnWord100_Click(object sender, RoutedEventArgs e)
            => SetzeWordZoom(1.0);

        private void BtnWordEineSeite_Click(object sender, RoutedEventArgs e)
            => AktualisiereWordZoom(WordAnsicht.EineSeite);

        private void BtnWordBreite_Click(object sender, RoutedEventArgs e)
            => AktualisiereWordZoom(WordAnsicht.Seitenbreite);

        private void BtnWordMehrere_Click(object sender, RoutedEventArgs e)
            => AktualisiereWordZoom(WordAnsicht.MehrereSeiten);

        /// <summary>
        /// Gibt den Pfad zum anzuzeigenden PDF zurück.
        ///
        /// Schritt 1: Basis-PDF ermitteln.
        ///   – Word-Datei  → Word exportiert unsichtbar als PDF (gecacht).
        ///   – PDF-Datei   → direkt verwenden.
        ///
        /// Schritt 2: Abdeckung einbrennen (falls aktiv).
        ///   PdfSharp zeichnet weiße Bänder an Kopf und/oder Fußbereich
        ///   JEDER Seite in den Seiteninhalt ein.
        ///   Die Bänder sind Teil des PDFs – nicht im Fenster verankert.
        /// </summary>
        private string BestimmePdfPfad(string quellPfad)
        {
            bool   kopf   = ChkKopf.IsChecked == true;
            bool   fuss   = ChkFuss.IsChecked == true;
            double kopfMm = kopf ? ParseMm(TxtKopfMm.Text, 20) : 0;
            double fussMm = fuss ? ParseMm(TxtFussMm.Text, 20) : 0;

            // ── Schritt 1: Basis-PDF ──────────────────────────────────────────
            string basePdf;
            if (DateiTypen.IstWordDatei(Path.GetExtension(quellPfad)))
            {
                basePdf = PdfCache.GetBasisPdfPfad(quellPfad, _cacheDir);
                if (!PdfCache.CacheGültig(basePdf, quellPfad))
                    WordInteropService.WortDateiZuPdf(quellPfad, basePdf);
            }
            else
            {
                basePdf = quellPfad;   // PDF direkt
            }

            // ── Schritt 2: Abdeckung einbrennen ──────────────────────────────
            if (kopf || fuss)
            {
                var covPfad = PdfCache.GetCoveredPdfPfad(quellPfad, _cacheDir, kopfMm, fussMm);
                if (!PdfCache.CacheGültig(covPfad, quellPfad))
                    PdfCoverService.AbdeckePdfSeiten(basePdf, covPfad, kopfMm, fussMm);
                return covPfad;
            }

            return basePdf;
        }

        private void ZeigeKeinVorschauHinweis(string pfad)
        {
            var name = HtmlEncode(Path.GetFileName(pfad));
            var ext  = Path.GetExtension(pfad).ToLowerInvariant();

            string hinweis = DateiTypen.IstGesperrteExtension(ext)
                ? "<p style='color:#c00;font-size:12px'>&#9888; AxisVM-Modelldateien werden nicht ge&ouml;ffnet, "
                  + "um einen unbeabsichtigten Start von AxisVM zu verhindern.</p>"
                : "";

            WordVorschau.NavigateToString(HtmlSeite(
                "<h3 style='color:#333'>Keine Vorschau verf&uuml;gbar</h3>"
                + "<p>Die Datei <b>" + name + "</b> kann nicht angezeigt werden.</p>"
                + hinweis
                + "<p style='color:#aaa;font-size:12px'>Dateityp: " + ext + "</p>",
                bodyStyle: "font-family:Segoe UI,sans-serif;padding:40px;color:#555"));
            _dokumentGeladen = true;
        }

        private void ZeigeJsonVorschau(string pfad)
        {
            string formatiert;
            try
            {
                var raw = File.ReadAllText(pfad, System.Text.Encoding.UTF8);
                formatiert = FormatierJson(raw);
            }
            catch (Exception ex)
            {
                formatiert = "(Datei konnte nicht gelesen werden: " + ex.Message + ")";
            }

            WordVorschau.NavigateToString(
                "<!DOCTYPE html><html><head>"
                + "<meta http-equiv='Content-Type' content='text/html; charset=utf-16'><style>"
                + "html,body{margin:0;padding:0;height:100%;background:#1e1e1e;}"
                + "pre{font-family:Consolas,'Courier New',monospace;font-size:12px;"
                + "color:#d4d4d4;margin:0;padding:16px;white-space:pre-wrap;"
                + "word-break:break-all;line-height:1.5;}"
                + "</style></head><body><pre>" + HtmlEncode(formatiert) + "</pre></body></html>");
            _dokumentGeladen = true;
        }

        /// <summary>
        /// Einfacher JSON-Formatter ohne externe Bibliothek.
        /// Ignoriert Whitespace außerhalb von Strings und baut Einrückung nach.
        /// </summary>
        private static string FormatierJson(string raw)
        {
            var sb    = new System.Text.StringBuilder(raw.Length * 2);
            int level = 0;
            bool inStr = false;

            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];

                // String-Zustand nachführen (Escape-Sequenzen korrekt überspringen)
                if (c == '\\' && inStr) { sb.Append(c); if (i + 1 < raw.Length) sb.Append(raw[++i]); continue; }
                if (c == '"') { inStr = !inStr; sb.Append(c); continue; }

                if (inStr) { sb.Append(c); continue; }

                switch (c)
                {
                    case '{': case '[':
                        sb.Append(c); sb.AppendLine();
                        level++;
                        sb.Append(' ', level * 2);
                        break;
                    case '}': case ']':
                        sb.AppendLine();
                        level = Math.Max(0, level - 1);
                        sb.Append(' ', level * 2);
                        sb.Append(c);
                        break;
                    case ',':
                        sb.Append(c); sb.AppendLine();
                        sb.Append(' ', level * 2);
                        break;
                    case ':':
                        sb.Append(": ");
                        break;
                    case ' ': case '\t': case '\r': case '\n':
                        break; // Whitespace außerhalb Strings wird neu erzeugt
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Kodiert alle Zeichen, die in HTML-Kontext unsicher sind (inkl. Non-ASCII),
        /// als numerische HTML-Entities. Verhindert Encoding-Fehler bei NavigateToString.
        /// </summary>
        private static string HtmlEncode(string s)
        {
            var sb = new System.Text.StringBuilder(s.Length + 32);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '&':  sb.Append("&amp;");  break;
                    case '<':  sb.Append("&lt;");   break;
                    case '>':  sb.Append("&gt;");   break;
                    case '"':  sb.Append("&quot;"); break;
                    default:
                        if (c > 127) sb.Append("&#").Append((int)c).Append(';');
                        else         sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Erzeugt eine vollständige HTML-Seite mit korrekter charset-Deklaration
        /// für NavigateToString (utf-16, da C#-Strings intern UTF-16 sind).
        /// </summary>
        private static string HtmlSeite(string inhalt, string bodyStyle = "")
        {
            var style = string.IsNullOrEmpty(bodyStyle) ? "" : " style='" + bodyStyle + "'";
            return "<!DOCTYPE html><html><head>"
                + "<meta http-equiv='Content-Type' content='text/html; charset=utf-16'>"
                + "</head><body" + style + ">" + inhalt + "</body></html>";
        }

        // ── Navigations-Kontrolle ─────────────────────────────────────────────

        private void WordVorschau_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
        {
            if (e.Uri == null) return;
            if (e.Uri.ToString() == "about:blank") return;

            // Nur bei HTML- und Bild-Dateien relevant: In-Page-Klicks (Links, eingebettete Bilder)
            // würden den Browser intern auf eine andere URL navigieren, was die Anzeige "einklemmt"
            // (der Tree zeigt noch die Original-Datei, der Browser zeigt aber die verlinkte Ressource).
            // Lösung: Jede Navigation, die nicht zu _aktiverDateipfad passt, wird unterbunden.
            if (_aktiverDateipfad == null) return;
            var ext = Path.GetExtension(_aktiverDateipfad).ToLowerInvariant();
            if (!DateiTypen.IstHtmlDatei(ext) && !DateiTypen.IstBildDatei(ext)) return;

            var erwartet = new Uri(_aktiverDateipfad);
            if (!string.Equals(e.Uri.LocalPath, erwartet.LocalPath, StringComparison.OrdinalIgnoreCase))
                e.Cancel = true;
        }

        // ── Vorschau geladen ──────────────────────────────────────────────────

        private void WordVorschau_LoadCompleted(object sender, NavigationEventArgs e)
        {
            if (e.Uri == null || e.Uri.ToString() == "about:blank")
            {
                // Wenn nach dem about:blank ein Auto-Refresh ausstehend ist:
                // Jetzt erst das neue PDF laden – Browser hat alten Cache komplett verworfen.
                if (_autoRefreshPfad != null)
                {
                    var pfad = _autoRefreshPfad;
                    _autoRefreshPfad = null;
                    StarteNeuladungImHintergrund(pfad);
                }
                return;
            }

            _dokumentGeladen = true;
            AppZustand.Instanz.SetzeStatus("Vorschau: " + Path.GetFileName(_aktiverDateipfad ?? ""));
        }

        /// <summary>
        /// Erzeugt das PDF auf einem STA-Hintergrundthread (kein UI-Blockieren)
        /// und navigiert danach auf dem UI-Thread zur neuen Datei.
        /// </summary>
        private void StarteNeuladungImHintergrund(string pfad)
        {
            AppZustand.Instanz.SetzeStatus("PDF wird neu erzeugt …");

            var cacheDir = _cacheDir;
            bool   kopf   = ChkKopf.IsChecked == true;
            bool   fuss   = ChkFuss.IsChecked == true;
            double kopfMm = kopf ? ParseMm(TxtKopfMm.Text, 20) : 0;
            double fussMm = fuss ? ParseMm(TxtFussMm.Text, 20) : 0;

            var thread = new Thread(() =>
            {
                var zielPdf = WordPdfService.BerechneNeuenZielPdf(pfad, cacheDir, kopf, fuss, kopfMm, fussMm);

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_aktiverDateipfad != pfad) return; // Nutzer hat inzwischen andere Datei gewählt
                    if (zielPdf == null)
                    {
                        AppZustand.Instanz.SetzeStatus("Fehler beim Erzeugen der Vorschau.", StatusLevel.Error);
                        return;
                    }
                    WordVorschau.Navigate(new Uri(zielPdf));
                }));
            })
            {
                IsBackground = true,
                Name         = "StatikAutoRefresh"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        // ── Abdeckungs-Steuerung ──────────────────────────────────────────────

        private void Abdeckung_Changed(object sender, RoutedEventArgs e)
        {
            // Debounce 400 ms – verhindert Neugenerierung bei jeder Zeicheneingabe
            _abdeckungTimer?.Stop();
            _abdeckungTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _abdeckungTimer.Tick += (s2, e2) =>
            {
                _abdeckungTimer?.Stop();
                AktualisiereAbdeckung();
            };
            _abdeckungTimer.Start();
        }

        private void AktualisiereAbdeckung()
        {
            if (!_dokumentGeladen || _aktiverDateipfad == null) return;

            var ext = Path.GetExtension(_aktiverDateipfad).ToLowerInvariant();
            if (!DateiTypen.IstWordDatei(ext) && !DateiTypen.IstPdfDatei(ext)) return;

            try
            {
                var zielPdf = BestimmePdfPfad(_aktiverDateipfad);
                var aktUri  = WordVorschau.Source;
                if (aktUri == null ||
                    !aktUri.LocalPath.Equals(zielPdf, StringComparison.OrdinalIgnoreCase))
                {
                    _dokumentGeladen = false;
                    WordVorschau.Navigate(new Uri(zielPdf));
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("AktualisiereAbdeckung", ex.Message);
                AppZustand.Instanz.SetzeStatus("Fehler beim Aktualisieren der Abdeckbänder.", StatusLevel.Error);
            }
        }

        // ── PDF-Erzeugung ─────────────────────────────────────────────────────

        // ErstelleBasisPdf  → WordInteropService.WortDateiZuPdf
        // AbdeckePdfSeiten  → PdfCoverService.AbdeckePdfSeiten

        // ── Dateiüberwachung (Auto-Refresh bei Speichern) ─────────────────────
        // FileSystemWatcher + Debounce → FileWatcherService

        private void OnDateiGeändert()
        {
            // Wird vom FileWatcherService nach 2-s-Debounce auf dem UI-Thread aufgerufen.
            if (_aktiverDateipfad == null) return;

            PdfCache.LöscheCacheFürDatei(_aktiverDateipfad, _cacheDir);
            AppZustand.Instanz.SetzeStatus("Datei geändert – Vorschau wird aktualisiert …");

            // Erst zu about:blank navigieren – das verwirft den IE-Cache
            // für die bisherige PDF-URL vollständig.
            // LoadCompleted erkennt das Flag und startet die Neuladung.
            _autoRefreshPfad = _aktiverDateipfad;
            _dokumentGeladen = false;
            WordVorschau.Navigate("about:blank");
        }

        // LöscheCacheFürDatei → PdfCache.LöscheCacheFürDatei

        // GetBasisPdfPfad, GetCoveredPdfPfad, CacheGültig → PdfCache

        // ── Hilfsmethoden ────────────────────────────────────────────────────

        private static double ParseMm(string text, double fallback)
            => double.TryParse(text.Replace(",", "."),
                   System.Globalization.NumberStyles.Any,
                   System.Globalization.CultureInfo.InvariantCulture,
                   out double val) ? val : fallback;

        // ── TreeView: horizontales Auto-Scrollen verhindern ───────────────────

        /// <summary>
        /// Verhindert, dass WPF beim Selektieren eines TreeViewItems automatisch
        /// horizontal scrollt. Vertikales Scrollen bleibt erhalten.
        /// </summary>
        private void DokumentenBaum_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
        {
            // Aktuelle horizontale Scrollposition merken,
            // dann nach dem WPF-internen Scroll wieder herstellen.
            var sv = FindScrollViewer(DokumentenBaum);
            if (sv == null) return;
            double gespeicherterHOffset = sv.HorizontalOffset;

            Dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Loaded,
                new Action(() => sv.ScrollToHorizontalOffset(gespeicherterHOffset)));
        }

        /// <summary>
        /// Durchsucht den Visual Tree rekursiv nach dem ersten ScrollViewer.
        /// </summary>
        private static ScrollViewer? FindScrollViewer(DependencyObject parent)
        {
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is ScrollViewer sv) return sv;
                var found = FindScrollViewer(child);
                if (found != null) return found;
            }
            return null;
        }

        // ── Kontextmenü ───────────────────────────────────────────────────────

        private void DokumentenBaum_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            // Aus dem geklickten Element das nächstliegende TreeViewItem heraussuchen
            var element = e.OriginalSource as System.Windows.DependencyObject;
            while (element != null && !(element is TreeViewItem))
                element = System.Windows.Media.VisualTreeHelper.GetParent(element);

            var treeItem = element as TreeViewItem;
            string? pfad = treeItem?.Tag as string;

            var menu = new ContextMenu();

            // Mehrfachauswahl hat Vorrang
            if (_baumMehrfachAuswahl.Count > 1)
            {
                var dateien = _baumMehrfachAuswahl.Where(p => File.Exists(p)).ToList();
                var ordner  = _baumMehrfachAuswahl.Where(p => Directory.Exists(p)).ToList();
                int gesamt  = dateien.Count + ordner.Count;
                var itemLöschen = new MenuItem { Header = $"🗑 {gesamt} Elemente löschen …" };
                itemLöschen.Click += (s, ev) => LöscheMehrfachAuswahl(dateien, ordner);
                menu.Items.Add(itemLöschen);
            }
            else
            {
                bool istOrdner = pfad != null && Directory.Exists(pfad);
                bool istDatei  = pfad != null && File.Exists(pfad);

                if (istOrdner)
                {
                    var itemLöschen = new MenuItem { Header = "🗑 Ordner löschen …" };
                    itemLöschen.Click += (s, ev) => OrdnerLöschen(pfad!);
                    menu.Items.Add(itemLöschen);
                }
                else
                {
                    var itemLöschen = new MenuItem
                    {
                        Header    = "🗑 Datei löschen",
                        IsEnabled = istDatei
                    };
                    itemLöschen.Click += (s, ev) => DateiLöschen(pfad);
                    menu.Items.Add(itemLöschen);
                }
            }

            if (_projektPfad != null)
            {
                menu.Items.Add(new Separator());
                var itemNeu = new MenuItem { Header = "📄 Neues Word-Dokument …" };
                itemNeu.Click += (s, ev) => NeuesWordDokument();
                menu.Items.Add(itemNeu);
            }

            DokumentenBaum.ContextMenu = menu;
        }

        private void DateiListe_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            var pfade = DateiListe.SelectedItems
                .OfType<DateiEintrag>()
                .Select(d => d.VollerPfad)
                .ToList();

            var menu = new ContextMenu();

            string löschLabel = pfade.Count == 1 ? "🗑 Datei löschen"
                              : pfade.Count  > 1 ? $"🗑 {pfade.Count} Dateien löschen"
                              :                    "🗑 Datei löschen";

            var itemLöschen = new MenuItem { Header = löschLabel, IsEnabled = pfade.Count > 0 };
            itemLöschen.Click += (s, ev) => DateienLöschen(pfade);
            menu.Items.Add(itemLöschen);

            if (_projektPfad != null)
            {
                menu.Items.Add(new Separator());
                var itemNeu = new MenuItem { Header = "📄 Neues Word-Dokument …" };
                itemNeu.Click += (s, ev) => NeuesWordDokument();
                menu.Items.Add(itemNeu);
            }

            DateiListe.ContextMenu = menu;
        }

        // ── Word-Info-Panel ───────────────────────────────────────────────────

        private void BtnWordÖffnen_Click(object sender, RoutedEventArgs e)
        {
            if (_aktiverDateipfad == null) return;

            // Gesperrte Dateitypen (z.B. .axs) nie mit Shell öffnen
            var ext = Path.GetExtension(_aktiverDateipfad);
            if (DateiTypen.IstGesperrteExtension(ext))
            {
                AppZustand.Instanz.SetzeStatus(
                    $"{ext}-Dateien werden im Statik-Manager nicht geöffnet.",
                    StatusLevel.Warn);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = _aktiverDateipfad,
                    UseShellExecute = true
                });
                AppZustand.Instanz.SetzeStatus("Geöffnet: " + Path.GetFileName(_aktiverDateipfad));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Öffnen:\n" + ex.Message, "Fehler",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Baum-Mehrfachauswahl (Ctrl+Klick / Shift+Klick / Delete) ─────────

        private void DokumentenBaum_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;

            _dragStartPunkt = e.GetPosition(null);

            var element = e.OriginalSource as DependencyObject;
            while (element != null && !(element is TreeViewItem))
                element = VisualTreeHelper.GetParent(element);

            var item = element as TreeViewItem;
            if (item?.Tag is not string pfad) return;

            bool ctrl  = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
            bool shift = (Keyboard.Modifiers & ModifierKeys.Shift)   != 0;

            if (ctrl)
            {
                if (_baumMehrfachAuswahl.Contains(pfad))
                    _baumMehrfachAuswahl.Remove(pfad);
                else
                {
                    _baumMehrfachAuswahl.Add(pfad);
                    _baumAuswahlAnker = pfad;
                }
                AktualisiereTreeViewHervorhebung();
                e.Handled = true;
            }
            else if (shift && _baumAuswahlAnker != null)
            {
                var alleItems = OrdneTreeItemsFlach(DokumentenBaum).ToList();
                int ankerIdx = alleItems.FindIndex(i =>
                    string.Equals(i.Tag as string, _baumAuswahlAnker, StringComparison.OrdinalIgnoreCase));
                int zielIdx  = alleItems.FindIndex(i =>
                    string.Equals(i.Tag as string, pfad, StringComparison.OrdinalIgnoreCase));

                if (ankerIdx >= 0 && zielIdx >= 0)
                {
                    int start = Math.Min(ankerIdx, zielIdx);
                    int ende  = Math.Max(ankerIdx, zielIdx);
                    _baumMehrfachAuswahl.Clear();
                    for (int i = start; i <= ende; i++)
                        if (alleItems[i].Tag is string p)
                            _baumMehrfachAuswahl.Add(p);
                }
                AktualisiereTreeViewHervorhebung();
                e.Handled = true;
            }
            else
            {
                // Normaler Klick: Mehrfachauswahl aufheben
                _baumMehrfachAuswahl.Clear();
                _baumAuswahlAnker = pfad;
                AktualisiereTreeViewHervorhebung();
            }
        }

        private void DokumentenBaum_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Delete) return;
            e.Handled = true;

            if (_baumMehrfachAuswahl.Count > 0)
            {
                var dateien = _baumMehrfachAuswahl.Where(p => File.Exists(p)).ToList();
                var ordner  = _baumMehrfachAuswahl.Where(p => Directory.Exists(p)).ToList();
                LöscheMehrfachAuswahl(dateien, ordner);
            }
            else if (DokumentenBaum.SelectedItem is TreeViewItem sel && sel.Tag is string pfad)
            {
                if      (File.Exists(pfad))      DateiLöschen(pfad);
                else if (Directory.Exists(pfad)) OrdnerLöschen(pfad);
            }
        }

        private void DateiListe_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Delete) return;
            e.Handled = true;
            var pfade = DateiListe.SelectedItems.OfType<DateiEintrag>().Select(d => d.VollerPfad).ToList();
            if (pfade.Count > 0) DateienLöschen(pfade);
        }

        private void LöscheMehrfachAuswahl(IList<string> dateien, IList<string> ordner)
        {
            int gesamt = dateien.Count + ordner.Count;
            if (gesamt == 0) return;

            string frage = gesamt == 1
                ? $"Element wirklich löschen?\n\n{Path.GetFileName(dateien.Concat(ordner).First())}"
                : $"{gesamt} Elemente (Dateien und/oder Ordner) wirklich löschen?";

            if (MessageBox.Show(frage, "Löschen bestätigen",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No)
                != MessageBoxResult.Yes) return;

            var fehler = new List<string>();
            bool vorschauBetroffen = false;

            foreach (var pfad in dateien)
            {
                try
                {
                    File.Delete(pfad);
                    if (string.Equals(_aktiverDateipfad, pfad, StringComparison.OrdinalIgnoreCase))
                        vorschauBetroffen = true;
                }
                catch (Exception ex) { fehler.Add($"{Path.GetFileName(pfad)}: {ex.Message}"); }
            }

            foreach (var pfad in ordner)
            {
                try
                {
                    if (_aktiverDateipfad != null &&
                        _aktiverDateipfad.StartsWith(pfad + Path.DirectorySeparatorChar,
                            StringComparison.OrdinalIgnoreCase))
                        vorschauBetroffen = true;
                    Directory.Delete(pfad, recursive: true);
                }
                catch (Exception ex) { fehler.Add($"{Path.GetFileName(pfad)}: {ex.Message}"); }
            }

            _baumMehrfachAuswahl.Clear();

            if (vorschauBetroffen)
            {
                _aktiverDateipfad = null;
                _dokumentGeladen  = false;
                DateiAusgewählt?.Invoke(null);
                ZeigeBrowser(mitAbdeckung: false);
                WordVorschau.Navigate("about:blank");
            }

            if (_projektPfad != null) AktualisiereDokumentListe();

            int gelöscht = gesamt - fehler.Count;
            AppZustand.Instanz.SetzeStatus(gelöscht > 0 ? $"{gelöscht} Element(e) gelöscht." : "Nichts gelöscht.");

            if (fehler.Count > 0)
                MessageBox.Show(
                    "Folgende Elemente konnten nicht gelöscht werden:\n\n" + string.Join("\n", fehler),
                    "Fehler beim Löschen", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void AktualisiereTreeViewHervorhebung(ItemsControl? parent = null)
        {
            parent ??= DokumentenBaum;
            foreach (var obj in parent.Items)
            {
                if (parent.ItemContainerGenerator.ContainerFromItem(obj) is not TreeViewItem item) continue;
                if (item.Tag is string pfad)
                    item.Background = _baumMehrfachAuswahl.Contains(pfad) ? _mehrfachHintergrund : null;
                AktualisiereTreeViewHervorhebung(item);
            }
        }

        private IEnumerable<TreeViewItem> OrdneTreeItemsFlach(ItemsControl parent)
        {
            foreach (var obj in parent.Items)
            {
                if (parent.ItemContainerGenerator.ContainerFromItem(obj) is not TreeViewItem item) continue;
                if (item.Tag is string) yield return item;
                if (item.IsExpanded)
                    foreach (var child in OrdneTreeItemsFlach(item))
                        yield return child;
            }
        }

        // ── Drag & Drop im Baum ───────────────────────────────────────────────

        private void DokumentenBaum_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _dragAktiv) return;

            var delta = e.GetPosition(null) - _dragStartPunkt;
            if (Math.Abs(delta.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(delta.Y) < SystemParameters.MinimumVerticalDragDistance) return;

            var element = e.OriginalSource as DependencyObject;
            while (element != null && !(element is TreeViewItem))
                element = VisualTreeHelper.GetParent(element);

            var item = element as TreeViewItem;
            if (item?.Tag is not string pfad) return;

            // Nicht den Stammordner verschieben
            if (string.Equals(pfad, _projektPfad, StringComparison.OrdinalIgnoreCase)) return;

            _dragAktiv = true;
            DragDrop.DoDragDrop(item, pfad, DragDropEffects.Move);
            _dragAktiv = false;
        }

        private void DokumentenBaum_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.StringFormat))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var element = e.OriginalSource as DependencyObject;
            while (element != null && !(element is TreeViewItem))
                element = VisualTreeHelper.GetParent(element);

            var targetItem = element as TreeViewItem;
            string? zielPfad = targetItem?.Tag as string;

            if (zielPfad == null || !Directory.Exists(zielPfad))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            string sourcePfad    = (string)e.Data.GetData(DataFormats.StringFormat);
            string? sourceEltern = Path.GetDirectoryName(sourcePfad);

            // Kein Drop auf sich selbst oder den eigenen Elternordner
            bool gleichePfad   = string.Equals(zielPfad, sourcePfad,   StringComparison.OrdinalIgnoreCase);
            bool gleicheEltern = string.Equals(zielPfad, sourceEltern, StringComparison.OrdinalIgnoreCase);
            // Kein Drop in eigenen Unterordner (Ordner in sich selbst verschieben)
            bool zielIstUnter  = zielPfad.StartsWith(sourcePfad + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

            e.Effects = (gleichePfad || gleicheEltern || zielIstUnter)
                ? DragDropEffects.None
                : DragDropEffects.Move;
            e.Handled = true;
        }

        private void DokumentenBaum_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.StringFormat)) return;

            string sourcePfad = (string)e.Data.GetData(DataFormats.StringFormat);
            if (string.IsNullOrEmpty(sourcePfad)) return;

            var element = e.OriginalSource as DependencyObject;
            while (element != null && !(element is TreeViewItem))
                element = VisualTreeHelper.GetParent(element);

            var targetItem = element as TreeViewItem;
            string? zielOrdner = targetItem?.Tag as string;

            if (zielOrdner == null || !Directory.Exists(zielOrdner)) return;

            string name     = Path.GetFileName(sourcePfad);
            string zielPfad = Path.Combine(zielOrdner, name);

            if (string.Equals(sourcePfad, zielPfad, StringComparison.OrdinalIgnoreCase)) return;

            if (File.Exists(zielPfad) || Directory.Exists(zielPfad))
            {
                MessageBox.Show($"Ziel existiert bereits:\n{zielPfad}",
                    "Verschieben nicht möglich", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                bool istDatei  = File.Exists(sourcePfad);
                bool istOrdner = Directory.Exists(sourcePfad);

                if      (istDatei)  File.Move(sourcePfad, zielPfad);
                else if (istOrdner) Directory.Move(sourcePfad, zielPfad);

                // Aktive Vorschau-Pfad anpassen wenn betroffen
                if (_aktiverDateipfad != null)
                {
                    if (istDatei && string.Equals(_aktiverDateipfad, sourcePfad, StringComparison.OrdinalIgnoreCase))
                        _aktiverDateipfad = zielPfad;
                    else if (istOrdner && _aktiverDateipfad.StartsWith(
                                sourcePfad + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                        _aktiverDateipfad = zielPfad + _aktiverDateipfad.Substring(sourcePfad.Length);
                }

                AktualisiereDokumentListe();
                AppZustand.Instanz.SetzeStatus($"Verschoben: {name} → {Path.GetFileName(zielOrdner)}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Verschieben:\n" + ex.Message,
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Datei löschen ─────────────────────────────────────────────────────

        // Einzelauftrag (Rechtsklick im Baum) → leitet an Batch-Methode weiter
        private void DateiLöschen(string? pfad)
        {
            if (pfad == null) return;
            DateienLöschen(new List<string> { pfad });
        }

        // Mülleimer-Button in jeder selektierten Zeile → löscht immer die gesamte Auswahl
        private void BtnZeileLöschen_Click(object sender, RoutedEventArgs e)
        {
            var pfade = DateiListe.SelectedItems
                .OfType<DateiEintrag>()
                .Select(d => d.VollerPfad)
                .ToList();

            // Fallback: DataContext der geklickten Zeile, falls Selektion leer
            if (pfade.Count == 0 && sender is Button btn && btn.DataContext is DateiEintrag eintrag)
                pfade.Add(eintrag.VollerPfad);

            DateienLöschen(pfade);
        }

        // Kernmethode: Batch-Löschung mit Bestätigungsdialog und Fehlersammlung
        private void DateienLöschen(IList<string> pfade)
        {
            if (pfade == null || pfade.Count == 0) return;

            // Bestätigungsdialog — Default-Fokus: Nein
            string frage = pfade.Count == 1
                ? $"Datei wirklich löschen?\n\n{Path.GetFileName(pfade[0])}"
                : $"{pfade.Count} Dateien wirklich löschen?";

            var antwort = MessageBox.Show(
                frage, "Löschen bestätigen",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (antwort != MessageBoxResult.Yes) return;

            var fehler = new List<string>();
            bool vorschauBetroffen = false;

            foreach (var pfad in pfade)
            {
                try
                {
                    if (File.Exists(pfad))
                        File.Delete(pfad);

                    if (string.Equals(_aktiverDateipfad, pfad, StringComparison.OrdinalIgnoreCase))
                        vorschauBetroffen = true;
                }
                catch (Exception ex)
                {
                    fehler.Add($"{Path.GetFileName(pfad)}: {ex.Message}");
                }
            }

            // Vorschau leeren, wenn aktive Datei betroffen
            if (vorschauBetroffen)
            {
                _aktiverDateipfad = null;
                _dokumentGeladen  = false;
                DateiAusgewählt?.Invoke(null);
                ZeigeBrowser(mitAbdeckung: false);
                WordVorschau.Navigate("about:blank");
            }

            if (_projektPfad != null) AktualisiereDokumentListe();

            int gelöscht = pfade.Count - fehler.Count;
            AppZustand.Instanz.SetzeStatus(
                gelöscht > 0 ? $"{gelöscht} Datei(en) gelöscht." : "Keine Datei gelöscht.");

            if (fehler.Count > 0)
                MessageBox.Show(
                    "Folgende Dateien konnten nicht gelöscht werden:\n\n" + string.Join("\n", fehler),
                    "Fehler beim Löschen", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void OrdnerLöschen(string ordnerPfad)
        {
            var name = Path.GetFileName(ordnerPfad);
            var antwort = MessageBox.Show(
                $"Position \"{name}\" und alle Inhalte unwiderruflich löschen?",
                "Ordner löschen bestätigen",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (antwort != MessageBoxResult.Yes) return;

            // Vorschau leeren, wenn aktive Datei im gelöschten Ordner liegt
            if (_aktiverDateipfad != null &&
                _aktiverDateipfad.StartsWith(ordnerPfad + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                _aktiverDateipfad = null;
                _dokumentGeladen  = false;
                DateiAusgewählt?.Invoke(null);
                ZeigeBrowser(mitAbdeckung: false);
                WordVorschau.Navigate("about:blank");
            }

            try
            {
                Directory.Delete(ordnerPfad, recursive: true);
                if (_projektPfad != null) AktualisiereDokumentListe();
                AppZustand.Instanz.SetzeStatus($"Ordner gelöscht: {name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Löschen:\n" + ex.Message, "Fehler",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Neues Word-Dokument ───────────────────────────────────────────────

        private void NeuesWordDokument()
        {
            if (_projektPfad == null) return;

            var dlg = new SaveFileDialog
            {
                Title            = "Neues Word-Dokument erstellen",
                Filter           = "Word-Dokument (*.docx)|*.docx",
                DefaultExt       = "docx",
                InitialDirectory = _projektPfad,
                FileName         = "Neues Dokument"
            };
            if (dlg.ShowDialog(Window.GetWindow(this)) != true) return;

            var zielPfad = dlg.FileName;
            AppZustand.Instanz.SetzeStatus("Erstelle Word-Dokument …");

            var thread = new Thread(() =>
            {
                Word.Application? wordApp = null;
                Word.Document?    wordDoc = null;
                try
                {
                    wordApp = new Word.Application { Visible = false };
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wordDoc = wordApp.Documents.Add();
                    wordDoc.SaveAs2(zielPfad, Word.WdSaveFormat.wdFormatXMLDocument);

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (_projektPfad != null) AktualisiereDokumentListe();
                        AppZustand.Instanz.SetzeStatus(
                            $"Erstellt: {Path.GetFileName(zielPfad)}");
                    }));
                }
                catch (Exception ex)
                {
                    var msg = ex.Message;
                    Dispatcher.BeginInvoke(new Action(() =>
                        MessageBox.Show("Fehler beim Erstellen:\n" + msg,
                            "Fehler", MessageBoxButton.OK, MessageBoxImage.Error)));
                }
                finally
                {
                    try { wordDoc?.Close(SaveChanges: false); } catch { }
                    try { wordApp?.Quit(); } catch { }
                    if (wordDoc != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc); } catch { }
                    if (wordApp != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                }
            })
            { IsBackground = true, Name = "NeuesWordDoc" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        // ── Datenklasse ───────────────────────────────────────────────────────

        private enum WordAnsicht { Prozent100, EineSeite, Seitenbreite, MehrereSeiten }

        private class DateiEintrag
        {
            public string Dateiname     { get; set; } = "";
            public string OrdnerRelativ { get; set; } = "";
            public string VollerPfad    { get; set; } = "";
        }
    }
}
