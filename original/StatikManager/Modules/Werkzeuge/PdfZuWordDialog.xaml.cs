using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace StatikManager.Modules.Werkzeuge
{
    /// <summary>
    /// Konvertiert PDF-Dateien in bearbeitbare Word-Dokumente.
    /// Word 2013+ öffnet PDFs nativ: Text bleibt Text, Grafiken werden als Bilder eingebettet.
    /// </summary>
    public partial class PdfZuWordDialog : Window
    {
        private CancellationTokenSource? _cts;
        private bool _läuft;

        public PdfZuWordDialog()
        {
            InitializeComponent();
        }

        // ── Dateiauswahl ──────────────────────────────────────────────────────

        private void BtnPdfSuchen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title           = "PDF-Datei auswählen",
                Filter          = "PDF-Dateien (*.pdf)|*.pdf|Alle Dateien (*.*)|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog(this) != true) return;

            TxtPdfPfad.Text = dlg.FileName;

            // Zieldatei vorbelegen wenn noch leer
            if (string.IsNullOrEmpty(TxtWordPfad.Text))
                TxtWordPfad.Text = Path.Combine(
                    Path.GetDirectoryName(dlg.FileName) ?? "",
                    Path.GetFileNameWithoutExtension(dlg.FileName) + ".docx");
        }

        private void BtnWordSuchen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title           = "Word-Zieldatei festlegen",
                Filter          = "Word-Dokument (*.docx)|*.docx",
                DefaultExt      = "docx",
                OverwritePrompt = true
            };
            if (!string.IsNullOrEmpty(TxtPdfPfad.Text))
            {
                dlg.InitialDirectory = Path.GetDirectoryName(TxtPdfPfad.Text);
                dlg.FileName         = Path.GetFileNameWithoutExtension(TxtPdfPfad.Text);
            }
            if (dlg.ShowDialog(this) == true)
                TxtWordPfad.Text = dlg.FileName;
        }

        // ── Konvertieren / Abbrechen ──────────────────────────────────────────

        private async void BtnKonvertieren_Click(object sender, RoutedEventArgs e)
        {
            if (_läuft) { _cts?.Cancel(); return; }
            if (!EingabenPrüfen()) return;

            string pdfPfad  = TxtPdfPfad.Text;
            string wordPfad = TxtWordPfad.Text;

            UiAktivieren(false);
            _cts = new CancellationTokenSource();

            var fortschritt = new Progress<(int wert, string text)>(p =>
            {
                Fortschritt.Value = p.wert;
                TxtStatus.Text    = p.text;
            });

            try
            {
                await KonvertierenAsync(pdfPfad, wordPfad, fortschritt, _cts.Token);

                if (!_cts.IsCancellationRequested)
                {
                    Fortschritt.Value = 100;
                    TxtStatus.Text    = "Fertig.";

                    var antwort = MessageBox.Show(
                        $"Erfolgreich konvertiert:\n{wordPfad}\n\nDatei jetzt öffnen?",
                        "Fertig", MessageBoxButton.YesNo, MessageBoxImage.Information);

                    if (antwort == MessageBoxResult.Yes)
                        try { Process.Start(new ProcessStartInfo(wordPfad) { UseShellExecute = true }); }
                        catch { /* ignorieren */ }
                }
                else
                {
                    TxtStatus.Text    = "Abgebrochen.";
                    Fortschritt.Value = 0;
                }
            }
            catch (OperationCanceledException)
            {
                TxtStatus.Text    = "Abgebrochen.";
                Fortschritt.Value = 0;
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Fehler: " + ex.Message;
                MessageBox.Show(
                    "Fehler bei der Konvertierung:\n\n" + ex.Message
                    + "\n\nHinweis: Word 2013 oder neuer wird benötigt.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _cts?.Dispose();
                _cts = null;
                UiAktivieren(true);
            }
        }

        private void BtnSchliessen_Click(object sender, RoutedEventArgs e)
        {
            if (_läuft) { _cts?.Cancel(); return; }
            Close();
        }

        // ── Konvertierungslogik ───────────────────────────────────────────────

        /// <summary>
        /// Öffnet das PDF in Word (Word 2013+ unterstützt PDF als Eingabeformat),
        /// analysiert Inhalt automatisch und speichert als bearbeitbares DOCX.
        /// Läuft auf einem STA-Hintergrundthread damit der UI-Thread frei bleibt.
        /// </summary>
        private static Task KonvertierenAsync(
            string pdfPfad, string wordPfad,
            IProgress<(int, string)> fortschritt,
            CancellationToken ct)
        {
            var tcs = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);

            var thread = new Thread(() =>
            {
                Word.Application? wordApp = null;
                Word.Document?    wordDoc = null;

                try
                {
                    ct.ThrowIfCancellationRequested();
                    fortschritt.Report((20, "Starte Word …"));

                    wordApp = new Word.Application { Visible = false };
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                    ct.ThrowIfCancellationRequested();
                    fortschritt.Report((40, "Öffne PDF und analysiere Inhalt …"));

                    // Word 2013+ öffnet PDFs direkt via "PDF Reflow":
                    // Text wird erkannt, Bilder als InlineShapes eingebettet.
                    wordDoc = wordApp.Documents.Open(
                        FileName:          pdfPfad,
                        ReadOnly:          false,
                        AddToRecentFiles:  false,
                        Visible:           false,
                        ConfirmConversions: false);

                    ct.ThrowIfCancellationRequested();
                    fortschritt.Report((80, "Speichere als Word-Dokument …"));

                    // Zieldatei-Verzeichnis sicherstellen
                    var zielDir = Path.GetDirectoryName(wordPfad);
                    if (!string.IsNullOrEmpty(zielDir))
                        Directory.CreateDirectory(zielDir);

                    wordDoc.SaveAs2(wordPfad, Word.WdSaveFormat.wdFormatXMLDocument);

                    fortschritt.Report((100, "Fertig."));
                    tcs.TrySetResult(true);
                }
                catch (OperationCanceledException) { tcs.TrySetCanceled(); }
                catch (Exception ex)               { tcs.TrySetException(ex); }
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
            {
                IsBackground = true,
                Name         = "PdfZuWord_STA"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        // ── Hilfsmethoden ────────────────────────────────────────────────────

        private bool EingabenPrüfen()
        {
            if (string.IsNullOrWhiteSpace(TxtPdfPfad.Text))
            {
                MessageBox.Show("Bitte wählen Sie eine PDF-Quelldatei aus.",
                    "Eingabe fehlt", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            if (!File.Exists(TxtPdfPfad.Text))
            {
                MessageBox.Show("Die angegebene PDF-Datei wurde nicht gefunden.",
                    "Datei fehlt", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(TxtWordPfad.Text))
            {
                MessageBox.Show("Bitte legen Sie eine Word-Zieldatei fest.",
                    "Eingabe fehlt", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        private void UiAktivieren(bool aktiv)
        {
            _läuft                  = !aktiv;
            TxtPdfPfad.IsEnabled    = aktiv;
            TxtWordPfad.IsEnabled   = aktiv;
            BtnPdfSuchen.IsEnabled  = aktiv;
            BtnWordSuchen.IsEnabled = aktiv;
            BtnKonvertieren.Content = aktiv ? "Konvertieren" : "Abbrechen";
            BtnSchliessen.Content   = aktiv ? "Schließen"    : "Abbrechen";
            BtnSchliessen.IsCancel  = aktiv;
            if (aktiv) Fortschritt.Value = 0;
        }
    }
}
