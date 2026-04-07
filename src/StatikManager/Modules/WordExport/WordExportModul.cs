// src/StatikManager/Modules/WordExport/WordExportModul.cs
using StatikManager.Core;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace StatikManager.Modules.WordExport
{
    public class WordExportModul : IModul
    {
        private WordExportPanel? _panel;

        public string Id   => "wordexport";
        public string Name => "Word-Export";

        public UIElement ErstellePanel()
        {
            _panel = new WordExportPanel();
            return _panel;
        }

        public MenuItem? ErzeugeMenüEintrag()
        {
            var menü = new MenuItem { Header = "_Word-Export" };

            var itemNeu = new MenuItem { Header = "Neues Word-Dokument …" };
            itemNeu.Click += (_, _) => _panel?.BtnNeuErstellen_Click(itemNeu, new RoutedEventArgs());

            var itemOeffnen = new MenuItem { Header = "Word-Dokument öffnen …" };
            itemOeffnen.Click += (_, _) => _panel?.BtnOeffnen_Click(itemOeffnen, new RoutedEventArgs());

            menü.Items.Add(itemNeu);
            menü.Items.Add(itemOeffnen);
            return menü;
        }

        public IEnumerable<FrameworkElement> ErzeugeWerkzeugleistenEinträge()
        {
            yield break;
        }

        public void Bereinigen() => _panel?.Bereinigen();
    }
}
