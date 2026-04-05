# StatikManager – Code-Patterns

Bewiesene Patterns die im Projekt verwendet werden.

---

## FileSystemWatcher mit Debounce (UI-Thread)

```csharp
// OrdnerWatcherService.cs
_watcher = new FileSystemWatcher(ordnerPfad)
{
    NotifyFilter          = NotifyFilters.FileName | NotifyFilters.DirectoryName,
    IncludeSubdirectories = true,
    EnableRaisingEvents   = true,
    InternalBufferSize    = 65536
};
_watcher.Created += handler;
_watcher.Deleted += handler;
_watcher.Renamed += renamed;
// Changed NICHT abonnieren!

private void OnFsEvent()
{
    _dispatcher.BeginInvoke(new Action(() =>
    {
        _debounceTimer?.Stop();
        _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _debounceTimer.Tick += (s, ev) => { _debounceTimer.Stop(); OrdnerGeaendert?.Invoke(); };
        _debounceTimer.Start();
    }));
}
```

---

## WebBrowser NavigateToString mit deutschen Umlauten

```csharp
private static string HtmlEncode(string s)
{
    var sb = new StringBuilder(s.Length + 32);
    foreach (char c in s)
    {
        switch (c)
        {
            case '&': sb.Append("&amp;");  break;
            case '<': sb.Append("&lt;");   break;
            case '>': sb.Append("&gt;");   break;
            case '"': sb.Append("&quot;"); break;
            default:
                if (c > 127) sb.Append("&#").Append((int)c).Append(';');
                else         sb.Append(c);
                break;
        }
    }
    return sb.ToString();
}

private static string HtmlSeite(string inhalt, string bodyStyle = "")
{
    var style = string.IsNullOrEmpty(bodyStyle) ? "" : " style='" + bodyStyle + "'";
    return "<!DOCTYPE html><html><head>"
        + "<meta http-equiv='Content-Type' content='text/html; charset=utf-16'>"
        + "</head><body" + style + ">" + inhalt + "</body></html>";
}
```

---

## COM-Zugriff auf Word (STA-Thread)

```csharp
var thread = new Thread(() =>
{
    Word.Application? wordApp = null;
    try
    {
        wordApp = new Word.Application { Visible = false };
        wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        // ... COM-Arbeit ...
    }
    finally
    {
        try { wordDoc?.Close(SaveChanges: false); } catch { }
        try { wordApp?.Quit(); } catch { }
        if (wordApp != null)
            Marshal.ReleaseComObject(wordApp);
    }
}) { IsBackground = true, Name = "WordThread" };
thread.SetApartmentState(ApartmentState.STA);
thread.Start();
```

---

## pdfium Semaphore

```csharp
// Zugriff immer ueber AppZustand.RenderSem:
await AppZustand.Instanz.RenderSem.WaitAsync(ct);
try
{
    // pdfium-Arbeit (Docnet.Core)
}
finally
{
    AppZustand.Instanz.RenderSem.Release();
}
```

---

## INotifyPropertyChanged fuer DataGrid-Binding

```csharp
private sealed class MeinItem : INotifyPropertyChanged
{
    private bool _sichtbar = true;

    public bool Sichtbar
    {
        get => _sichtbar;
        set { _sichtbar = value; Notify(nameof(Sichtbar)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Notify(string name) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
```

---

## DataGrid mit Single-Click Checkbox

```xml
<DataGridTemplateColumn Header="✓" Width="36" CanUserResize="False">
    <DataGridTemplateColumn.CellTemplate>
        <DataTemplate>
            <CheckBox IsChecked="{Binding Sichtbar, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </DataTemplate>
    </DataGridTemplateColumn.CellTemplate>
</DataGridTemplateColumn>
```

---

## HTML zu PDF via Edge Headless

```csharp
private static string? SucheEdgePfad()
{
    var kandidaten = new[]
    {
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    };
    return kandidaten.FirstOrDefault(File.Exists);
}

// Aufruf im Hintergrund-Task:
var htmlUri = new Uri(htmlPfad).AbsoluteUri;
var args = $"--headless --no-sandbox --print-to-pdf=\"{pdfPfad}\" \"{htmlUri}\"";
var psi = new ProcessStartInfo(edgePfad, args) { CreateNoWindow = true, UseShellExecute = false };
using var proc = Process.Start(psi);
bool erfolg = (proc?.WaitForExit(30_000) ?? false) && File.Exists(pdfPfad);
```

---

## TreeView Multi-Select (Ctrl+Klick / Shift+Klick)

```csharp
// Felder:
private readonly HashSet<string> _baumMehrfachAuswahl = new(StringComparer.OrdinalIgnoreCase);
private string? _baumAuswahlAnker;
private static readonly Brush _mehrfachHintergrund = new SolidColorBrush(Color.FromArgb(80, 7, 99, 191));

// PreviewMouseLeftButtonDown:
if (ctrl)
{
    if (_baumMehrfachAuswahl.Contains(pfad)) _baumMehrfachAuswahl.Remove(pfad);
    else { _baumMehrfachAuswahl.Add(pfad); _baumAuswahlAnker = pfad; }
    AktualisiereTreeViewHervorhebung();
    e.Handled = true;
}

// Hervorhebung:
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
```

---

## WPF ToolBarTray (mehrzeilig, kein Gripper)

```xml
<ToolBarTray Background="{DynamicResource Farbe.Werkzeugleiste}" IsLocked="True">
  <ToolBar Band="0" BandIndex="0" Background="{DynamicResource Farbe.Werkzeugleiste}">
    <Button Content="Aktion" ToolTip="Beschreibung" Height="26" Padding="8,0" Click="..."/>
    <Separator/>
    <ToggleButton x:Name="BtnModus" Height="26" Padding="8,0"
                  Checked="Modus_Checked" Unchecked="Modus_Unchecked">
        <ToggleButton.Style>
            <Style TargetType="ToggleButton">
                <Setter Property="Content" Value="&#x2702; Normal"/>
                <Style.Triggers>
                    <Trigger Property="IsChecked" Value="True">
                        <Setter Property="Background" Value="#CC3333"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="Content" Value="&#x2702; Aktiv – Esc beenden"/>
                    </Trigger>
                </Style.Triggers>
            </Style>
        </ToggleButton.Style>
    </ToggleButton>
  </ToolBar>
  <ToolBar Band="1" BandIndex="0"><!-- zweite Zeile --></ToolBar>
</ToolBarTray>
```
- `Band` = Zeilennummer (0-basiert), `BandIndex` = Position in der Zeile
- `IsLocked="True"` blendet Gripper-Handles aus
- Standard-Button: `Height="26" Padding="8,0"` | Icon-Button: `Width="26" Height="26"`
- `TxtInfo`-Statustext am Ende: `<TextBlock x:Name="TxtInfo" VerticalAlignment="Center" FontSize="11" FontStyle="Italic" Foreground="#555"/>`

### Unicode-Icons (XML-Entity)
`&#x2702;` Schere | `&#x270F;` Stift | `&#x1F5D1;` Papierkorb | `&#x2714;` Haken
`&#x2716;` X | `&#x2795;` Plus | `&#x21A9;` Zurück | `&#x1F4E4;` Export

---

## Einstellungen XML-Serialisierung

```csharp
[XmlRoot("Einstellungen")]
public sealed class Einstellungen
{
    private static Einstellungen? _instanz;
    public static Einstellungen Instanz => _instanz ??= Laden();

    [XmlArray("ProjektEintraege")]
    [XmlArrayItem("Projekt")]
    public List<ProjektEintrag> ProjektEintraege { get; set; } = new();

    public void Speichern()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DateiPfad)!);
        using var fs = File.Create(DateiPfad);
        new XmlSerializer(typeof(Einstellungen)).Serialize(fs, this);
    }
}
```
