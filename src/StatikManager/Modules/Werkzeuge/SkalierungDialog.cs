using System;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Modules.Werkzeuge
{
    internal enum SkalierungWahl { Verkleinern, Originalgröße, Abbrechen }

    /// <summary>
    /// Einfacher Code-only-Dialog für die Skalierungsentscheidung beim Word-Export.
    /// Wird angezeigt wenn das Bild größer als der Zielbereich der Vorlage ist.
    /// </summary>
    internal class SkalierungDialog : Window
    {
        public SkalierungWahl Wahl { get; private set; } = SkalierungWahl.Abbrechen;

        public SkalierungDialog(int prozent, double bildB_pt, double bildH_pt,
                                double zielB_pt, double zielH_pt)
        {
            Title  = "Bild zu groß für Vorlage";
            Width  = 470;
            SizeToContent = SizeToContent.Height;
            ResizeMode    = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;

            var outer = new DockPanel { Margin = new Thickness(16) };

            var txt = new TextBlock
            {
                Text = $"Das Bild passt nicht vollständig in die Vorlage.\n" +
                       $"Erforderliche Skalierung: {prozent} %\n\n" +
                       $"Bild:     {bildB_pt:F0} × {bildH_pt:F0} pt\n" +
                       $"Zielbereich: {zielB_pt:F0} × {zielH_pt:F0} pt\n\n" +
                       $"Wie soll eingefügt werden?",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14)
            };
            DockPanel.SetDock(txt, Dock.Top);
            outer.Children.Add(txt);

            var buttons = new StackPanel
            {
                Orientation         = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            buttons.Children.Add(Btn("Verkleinern (proportional)",
                () => { Wahl = SkalierungWahl.Verkleinern;   DialogResult = true; }));
            buttons.Children.Add(Btn("Originalgröße (rechts abschneiden)",
                () => { Wahl = SkalierungWahl.Originalgröße; DialogResult = true; }));
            buttons.Children.Add(Btn("Abbrechen",
                () => { Wahl = SkalierungWahl.Abbrechen;     DialogResult = false; }));

            DockPanel.SetDock(buttons, Dock.Bottom);
            outer.Children.Add(buttons);

            Content = outer;
        }

        private static Button Btn(string text, Action onClick)
        {
            var btn = new Button
            {
                Content  = text,
                Padding  = new Thickness(10, 5, 10, 5),
                Margin   = new Thickness(6, 0, 0, 0),
                MinWidth = 80
            };
            btn.Click += (_, _) => onClick();
            return btn;
        }
    }
}
