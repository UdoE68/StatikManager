using StatikManager.Core;
using StatikManager.Modules.Bildschnitt;
using StatikManager.Modules.Dokumente;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager
{
    public partial class MainWindow : Window
    {
        private readonly ModulManager _modulManager = new();
        private readonly System.Collections.Generic.Dictionary<string, UIElement> _panels
            = new System.Collections.Generic.Dictionary<string, UIElement>();

        private readonly System.Windows.Threading.DispatcherTimer _resetTimer = new()
        {
            Interval = System.TimeSpan.FromSeconds(4)
        };

        public MainWindow()
        {
            InitializeComponent();

            // ── Versionsanzeige (Titel + Statusleiste) ───────────────────────
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();

                // Versionsnummer aus <Version> im csproj → z. B. "1.0.0"
                var version = asm.GetName().Version?.ToString(3) ?? "–";

                // InformationalVersion = "dd.MM.yyyy HH:mm"  oder  "dd.MM.yyyy HH:mm+abc1234"
                var infoVer = (asm.GetCustomAttributes(
                        typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
                    as System.Reflection.AssemblyInformationalVersionAttribute[])
                    ?[0].InformationalVersion ?? "";

                // Format: "1.0.0|dd.MM.yyyy HH:mm"  oder  "1.0.0|dd.MM.yyyy HH:mm|githash"
                var parts     = infoVer.Split('|');
                var buildZeit = parts.Length > 1 ? parts[1] : "–";
                var gitHash   = parts.Length > 2 ? parts[2] : "";

#if DEBUG
                const string Konfig = "DEBUG";
#else
                const string Konfig = "RELEASE";
#endif
                var gitAnzeige = gitHash.Length > 0 ? $"  •  {gitHash}" : "";
                Title           = $"Statik-Manager v{version} [{Konfig}] – Build {buildZeit}{gitAnzeige}";
                TxtVersion.Text = $"v{version}  •  {Konfig}  •  Build {buildZeit}{gitAnzeige}";
            }

            // Shell mit AppZustand verbinden
            _resetTimer.Tick += (_, _) =>
            {
                _resetTimer.Stop();
                AppZustand.Instanz.SetzeStatus("Bereit");
                AppZustand.Instanz.ResetProgress();
            };

            AppZustand.Instanz.StatusGeändert     += AktualisiereStatusAnzeige;
            AppZustand.Instanz.FortschrittGeändert += AktualisiereProgress;
            AppZustand.Instanz.ProjektGeändert     += pfad => TxtProjektInfo.Text =
                pfad != null ? "Projekt: " + System.IO.Path.GetFileName(pfad) : "";

            // ── Module registrieren ──────────────────────────────────────────
            // Neues Modul hinzufügen: einfach eine weitere Zeile hier.
            _modulManager.Registrieren(new DokumenteModul());
            _modulManager.Registrieren(new BildschnittModul());

            // ── Module in die Shell integrieren ─────────────────────────────
            IntegriereModule();

            // ── Sitzung nach vollständiger UI-Initialisierung laden ──────────
            Loaded += (_, _) =>
            {
                var sitzung = Core.SitzungsZustand.Laden();
                if (HauptInhalt.Content is Modules.Dokumente.DokumentePanel panel)
                    panel.SitzungWiederherstellen(sitzung);
            };
        }

        private void IntegriereModule()
        {
            int menüIndex = HauptMenü.Items.IndexOf(MenüHilfe);
            bool ersterPanel = true;

            foreach (var modul in _modulManager.Module)
            {
                // Alle Panels erstellen (nicht nur erstes), damit Modul-Wechsel möglich ist
                var panel = modul.ErstellePanel();
                _panels[modul.Id] = panel;
                if (ersterPanel) { HauptInhalt.Content = panel; ersterPanel = false; }

                // Menüeintrag vor "Hilfe" einfügen
                var menüEintrag = modul.ErzeugeMenüEintrag();
                if (menüEintrag != null)
                {
                    HauptMenü.Items.Insert(menüIndex, menüEintrag);
                    menüIndex++;
                }

                // Werkzeugleiste befüllen
                foreach (var element in modul.ErzeugeWerkzeugleistenEinträge())
                    WerkzeugLeiste.Children.Add(element);
            }

            // Modul-Wechsel-Event (z.B. Doppelklick im DokumentePanel → BildschnittPanel)
            AppZustand.Instanz.ModulWechselAngefordert += (modulId, pfad) =>
            {
                if (!_panels.TryGetValue(modulId, out var panel)) return;
                HauptInhalt.Content = panel;
                if (panel is Modules.Bildschnitt.BildschnittPanel bp)
                    bp.LadeDatei(pfad);
            };
        }

        // ── Status-Anzeige ────────────────────────────────────────────────────

        private void AktualisiereStatusAnzeige(string text, StatusLevel level)
        {
            TxtStatus.Text = text;
            switch (level)
            {
                case StatusLevel.Error:
                    TxtStatus.Foreground       = (System.Windows.Media.Brush)FindResource("Farbe.StatusFehler");
                    StatusIndikator.Background = (System.Windows.Media.Brush)FindResource("Farbe.StatusFehler");
                    StatusIndikator.Visibility = Visibility.Visible;
                    break;
                case StatusLevel.Warn:
                    TxtStatus.Foreground       = (System.Windows.Media.Brush)FindResource("Farbe.StatusWarn");
                    StatusIndikator.Background = (System.Windows.Media.Brush)FindResource("Farbe.StatusWarn");
                    StatusIndikator.Visibility = Visibility.Visible;
                    break;
                default:
                    TxtStatus.ClearValue(System.Windows.Controls.TextBlock.ForegroundProperty);
                    StatusIndikator.Visibility = Visibility.Collapsed;
                    break;
            }

            // Auto-Reset: nur Info-Meldungen, nicht "Bereit" selbst (verhindert Schleife)
            _resetTimer.Stop();
            if (level == StatusLevel.Info && text != "Bereit")
                _resetTimer.Start();
        }

        private void AktualisiereProgress(int aktuell, int gesamt)
        {
            if (gesamt <= 0)
            {
                StatusProgress.Visibility = Visibility.Collapsed;
                return;
            }
            StatusProgress.Maximum    = gesamt;
            StatusProgress.Value      = aktuell;
            StatusProgress.Visibility = Visibility.Visible;
        }

        // ── Menü: Datei ──────────────────────────────────────────────────────

        private void MenüProjektLaden_Click(object sender, RoutedEventArgs e)
        {
            // Weiterleitung an das Dokumente-Modul (generisch über Interface möglich,
            // hier direkt da "Projekt laden" im Datei-Menü Konvention ist)
            var modul = _modulManager.FindeModul("dokumente") as DokumenteModul;
            // Alternativ: IModul um eine "AktionAusführen(string)" Methode erweitern
            // Für jetzt: der Werkzeugleisten-Button des Moduls übernimmt diese Aufgabe.
            // Dieser Klick delegiert an das Panel des Moduls über den gleichen Weg.
            if (HauptInhalt.Content is Modules.Dokumente.DokumentePanel panel)
                panel.ProjektLaden();
        }

        private void MenüEinstellungen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Modules.EinstellungsDialog.EinstellungenFenster { Owner = this };
            dlg.ShowDialog();
        }

        private void MenüStandardpfad_Click(object sender, RoutedEventArgs e)
        {
            var pfad = OrdnerDialog.Zeigen(
                startPfad: Einstellungen.Instanz.StandardPfad ?? "",
                titel:     "Standardpfad festlegen – Startordner beim Öffnen von Projekten",
                besitzer:  this);

            if (string.IsNullOrWhiteSpace(pfad)) return;

            Einstellungen.Instanz.StandardPfad = pfad;
            Einstellungen.Instanz.Speichern();
            AppZustand.Instanz.SetzeStatus("Standardpfad: " + pfad);
        }

        private void MenüBeenden_Click(object sender, RoutedEventArgs e) => Close();

        // ── Menü: Hilfe ───────────────────────────────────────────────────────

        private void MenüÜber_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Statik-Manager\nVersion 1.0\n\nDokumentenverwaltung für Statik-Projekte.",
                "Über Statik-Manager",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        // ── Aufräumen ─────────────────────────────────────────────────────────

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (HauptInhalt.Content is Modules.Dokumente.DokumentePanel panel)
            {
                if (!panel.PdfEditor.FrageObSpeichern())
                {
                    e.Cancel = true;
                    return;
                }
            }
            base.OnClosing(e);
        }

        protected override void OnClosed(System.EventArgs e)
        {
            // Sitzung speichern bevor Ressourcen freigegeben werden
            if (HauptInhalt.Content is Modules.Dokumente.DokumentePanel panel)
                Core.SitzungsZustand.Speichern(panel.SitzungSpeichern());

            base.OnClosed(e);
            _modulManager.AllesBereinigen();
        }
    }
}
