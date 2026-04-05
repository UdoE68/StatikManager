using StatikManager.Core;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace StatikManager.Modules.Bildschnitt
{
    public class BildschnittModul : IModul
    {
        private BildschnittPanel? _panel;

        public string Id   => "bildschnitt";
        public string Name => "Bildschnitt";

        public UIElement ErstellePanel()
        {
            _panel = new BildschnittPanel();
            return _panel;
        }

        public MenuItem? ErzeugeMenüEintrag()
        {
            var menü = new MenuItem { Header = "_Bildschnitt" };
            var itemBildLaden = new MenuItem { Header = "Bild laden …" };
            itemBildLaden.Click += (_, _) => _panel?.BtnBildLaden_Click(itemBildLaden, new RoutedEventArgs());
            menü.Items.Add(itemBildLaden);
            return menü;
        }

        public IEnumerable<FrameworkElement> ErzeugeWerkzeugleistenEinträge()
        {
            yield break;
        }

        public void Bereinigen() => _panel?.Bereinigen();
    }
}
