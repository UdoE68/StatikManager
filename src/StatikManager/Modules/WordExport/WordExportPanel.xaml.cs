// src/StatikManager/Modules/WordExport/WordExportPanel.xaml.cs
using StatikManager.Core;
using StatikManager.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace StatikManager.Modules.WordExport
{
    public partial class WordExportPanel : System.Windows.Controls.UserControl
    {
        private readonly DispatcherTimer _statusTimer;
        private string? _projektPfad;
        private readonly Action<string?> _projektGeaendertHandler;

        public WordExportPanel()
        {
            InitializeComponent();

            // Status-Poll-Timer: alle 3 Sekunden Word-Status prüfen
            _statusTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            _statusTimer.Tick += (_, _) => AktualisiereWordStatus();
            _statusTimer.Start();

            // Auf Projektwechsel reagieren
            _projektGeaendertHandler = pfad =>
            {
                _projektPfad = pfad;
                LadePositionen();
            };
            AppZustand.Instanz.ProjektGeändert += _projektGeaendertHandler;

            // Einstellungen laden
            var ein = Einstellungen.Instanz;
            ChkMitUeberschrift.IsChecked = ein.WordExportMitUeberschrift;
            ChkMitMassstab.IsChecked     = ein.WordExportMitMassstab;
            WaehleBildbreiteComboBox(ein.WordExportBildbreite);

            AktualisiereWordStatus();
        }

        // ── Word-Status ───────────────────────────────────────────────────

        private void AktualisiereWordStatus()
        {
            bool bereit = WordEinfuegenService.IstWordBereit();
            StatusKreis.Fill   = bereit ? Brushes.Green : Brushes.Red;
            TxtWordStatus.Text = bereit ? "Word verbunden" : "Word nicht verbunden";

            var pfad = WordEinfuegenService.GetAktiveDokumentPfad();
            TxtDokumentPfad.Text = pfad != null
                ? System.IO.Path.GetFileName(pfad)
                : "";
        }

        // ── Dokument-Aktionen ─────────────────────────────────────────────

        internal void BtnNeuErstellen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var vorlage = Einstellungen.Instanz.WordVorlagen
                    .Find(v => v.Standard && v.PfadGültig)?.Pfad;
                WordEinfuegenService.ErstelleDokument(vorlage);
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Erstellen: " + ex.Message,
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        internal void BtnOeffnen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Word-Dokument öffnen",
                Filter = "Word-Dokumente|*.docx;*.doc|Alle Dateien|*.*"
            };

            var sitzung = SitzungsZustand.Laden();
            if (!string.IsNullOrEmpty(sitzung.WordExportLetztesDokument))
                dlg.InitialDirectory = Path.GetDirectoryName(sitzung.WordExportLetztesDokument);

            if (dlg.ShowDialog() != true) return;

            try
            {
                WordEinfuegenService.OeffneDokument(dlg.FileName);

                sitzung.WordExportLetztesDokument = dlg.FileName;
                SitzungsZustand.Speichern(sitzung);
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Öffnen: " + ex.Message,
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Positionen laden ──────────────────────────────────────────────

        internal void BtnAktualisieren_Click(object sender, RoutedEventArgs e)
            => LadePositionen();

        private void LadePositionen()
        {
            var positionen = new List<PositionViewModel>();

            if (!string.IsNullOrEmpty(_projektPfad))
            {
                string statistikOrdner = Path.Combine(_projektPfad, "Statik");
                if (Directory.Exists(statistikOrdner))
                {
                    foreach (var posOrdner in Directory.GetDirectories(statistikOrdner, "Pos_*"))
                    {
                        var pos = LesePosition(posOrdner);
                        if (pos != null) positionen.Add(pos);
                    }
                }
            }

            if (positionen.Count == 0)
            {
                TxtKeinePositionen.Visibility = Visibility.Visible;
                PositionenTree.Visibility     = Visibility.Collapsed;
                TxtKeinePositionen.Text       = string.IsNullOrEmpty(_projektPfad)
                    ? "Kein Projekt geladen."
                    : "Keine Positionen gefunden.";
            }
            else
            {
                TxtKeinePositionen.Visibility = Visibility.Collapsed;
                PositionenTree.Visibility     = Visibility.Visible;
                PositionenTree.ItemsSource    = positionen;
            }
        }

        private PositionViewModel? LesePosition(string ordnerPfad)
        {
            try
            {
                string jsonPfad = Path.Combine(ordnerPfad, "position.json");
                if (!File.Exists(jsonPfad)) return null;

                string jsonText = File.ReadAllText(jsonPfad, System.Text.Encoding.UTF8);

                var pos = new PositionViewModel
                {
                    Id         = LeseJsonFeld(jsonText, "id",   ""),
                    Name       = LeseJsonFeld(jsonText, "name", ""),
                    OrdnerPfad = ordnerPfad
                };

                // Ausschnitte aus dem JSON-Array parsen (einfache Regex-freie Extraktion)
                int ausschnitteStart = jsonText.IndexOf("\"ausschnitte\"", StringComparison.Ordinal);
                if (ausschnitteStart >= 0)
                {
                    int arrayStart = jsonText.IndexOf('[', ausschnitteStart);
                    int arrayEnd   = jsonText.LastIndexOf(']');
                    if (arrayStart >= 0 && arrayEnd > arrayStart)
                    {
                        string arrayText = jsonText.Substring(arrayStart + 1, arrayEnd - arrayStart - 1);
                        // Einzelne Objekte aufteilen – suche { ... } Blöcke
                        int pos2 = 0;
                        while (pos2 < arrayText.Length)
                        {
                            int blockStart = arrayText.IndexOf('{', pos2);
                            if (blockStart < 0) break;

                            int depth = 0;
                            int blockEnd = blockStart;
                            for (int i = blockStart; i < arrayText.Length; i++)
                            {
                                if (arrayText[i] == '{') depth++;
                                else if (arrayText[i] == '}') { depth--; if (depth == 0) { blockEnd = i; break; } }
                            }

                            string block = arrayText.Substring(blockStart, blockEnd - blockStart + 1);

                            string dateiname = LeseJsonFeld(block, "dateiname", "");
                            string pngPfad   = Path.Combine(ordnerPfad, "daten", dateiname);

                            // Maßstab aus Ausschnitt-JSON lesen
                            string massstab  = "";
                            string jsonDatei = Path.ChangeExtension(pngPfad, ".json");
                            if (File.Exists(jsonDatei))
                            {
                                try
                                {
                                    string aText = File.ReadAllText(jsonDatei, System.Text.Encoding.UTF8);
                                    massstab = LeseJsonFeld(aText, "massstab", "");
                                }
                                catch { }
                            }

                            string nrStr = LeseJsonFeld(block, "nr", "0");
                            int nr = 0;
                            int.TryParse(nrStr, out nr);

                            pos.Ausschnitte.Add(new AusschnittViewModel
                            {
                                Nr           = nr,
                                Ueberschrift = LeseJsonFeld(block, "ueberschrift", ""),
                                PngPfad      = pngPfad,
                                Massstab     = massstab
                            });

                            pos2 = blockEnd + 1;
                        }
                    }
                }

                return pos;
            }
            catch (Exception ex)
            {
                Logger.Warn("WordExport", "Position lesen fehlgeschlagen: " + ex.Message);
                return null;
            }
        }

        // ── Einfügen ──────────────────────────────────────────────────────

        internal void BtnEinfuegen_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Button btn) return;
            if (btn.Tag is not AusschnittViewModel ausschnitt) return;

            if (!WordEinfuegenService.IstWordBereit())
            {
                MessageBox.Show("Bitte zuerst ein Word-Dokument öffnen.",
                    "Word", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!File.Exists(ausschnitt.PngPfad))
            {
                MessageBox.Show("PNG-Datei nicht gefunden:\n" + ausschnitt.PngPfad,
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                WordEinfuegenService.EinfuegenAnCursor(
                    pngPfad:          ausschnitt.PngPfad,
                    ueberschrift:     ausschnitt.Ueberschrift,
                    massstab:         ausschnitt.Massstab,
                    bildbreiteOption: Einstellungen.Instanz.WordExportBildbreite,
                    mitUeberschrift:  ChkMitUeberschrift.IsChecked == true,
                    mitMassstab:      ChkMitMassstab.IsChecked == true);

                AppZustand.Instanz.SetzeStatus(
                    "Eingefügt: " + ausschnitt.Ueberschrift);
                AktualisiereWordStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Einfügen:\n" + ex.Message,
                    "Word", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Optionen ──────────────────────────────────────────────────────

        internal void CbBildbreite_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CbBildbreite.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            if (!Enum.TryParse<BildbreiteOption>(item.Tag?.ToString(), out var option)) return;

            Einstellungen.Instanz.WordExportBildbreite = option;
            Einstellungen.Instanz.Speichern();
        }

        internal void Optionen_Geaendert(object sender, RoutedEventArgs e)
        {
            var ein = Einstellungen.Instanz;
            ein.WordExportMitUeberschrift = ChkMitUeberschrift.IsChecked == true;
            ein.WordExportMitMassstab     = ChkMitMassstab.IsChecked == true;
            ein.Speichern();
        }

        private void WaehleBildbreiteComboBox(BildbreiteOption option)
        {
            foreach (System.Windows.Controls.ComboBoxItem item in CbBildbreite.Items)
            {
                if (item.Tag?.ToString() == option.ToString())
                {
                    CbBildbreite.SelectedItem = item;
                    return;
                }
            }
            CbBildbreite.SelectedIndex = 0; // Fallback: Seitenbreite
        }

        // ── JSON-Hilfsparser (ohne externe Bibliothek) ────────────────────

        /// <summary>Liest einen einfachen String- oder Zahlwert aus einem JSON-Fragment.</summary>
        private static string LeseJsonFeld(string json, string key, string fallback)
        {
            if (string.IsNullOrEmpty(json)) return fallback;

            string suche = "\"" + key + "\"";
            int idx = json.IndexOf(suche, StringComparison.Ordinal);
            if (idx < 0) return fallback;

            int doppelpunkt = json.IndexOf(':', idx + suche.Length);
            if (doppelpunkt < 0) return fallback;

            int pos = doppelpunkt + 1;
            while (pos < json.Length &&
                   (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n'))
                pos++;

            if (pos >= json.Length) return fallback;

            if (json[pos] == '"')
            {
                // String-Wert
                int start = pos + 1;
                int end   = start;
                while (end < json.Length)
                {
                    if (json[end] == '"' && (end == 0 || json[end - 1] != '\\')) break;
                    end++;
                }
                if (end >= json.Length) return fallback;
                return json.Substring(start, end - start);
            }
            else
            {
                // Zahl oder bool – bis Komma, }, ] oder Whitespace
                int start = pos;
                int end   = pos;
                while (end < json.Length &&
                       json[end] != ',' && json[end] != '}' && json[end] != ']' &&
                       json[end] != ' ' && json[end] != '\r' && json[end] != '\n')
                    end++;
                return json.Substring(start, end - start).Trim();
            }
        }

        // ── Cleanup ───────────────────────────────────────────────────────

        public void Bereinigen()
        {
            AppZustand.Instanz.ProjektGeändert -= _projektGeaendertHandler;
            _statusTimer.Stop();
        }
    }
}
