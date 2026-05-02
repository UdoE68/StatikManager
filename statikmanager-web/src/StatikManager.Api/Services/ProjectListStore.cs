using System.Text.Json;
using StatikManager.Api.Contracts.Projects;

namespace StatikManager.Api.Services;

/// <summary>Liest und schreibt %APPDATA%\StatikManagerWeb\projekte.json.</summary>
public sealed class ProjectListStore
{
    private readonly object _lock = new();
    private readonly string _filePath;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    public ProjectListStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var dir = Path.Combine(appData, "StatikManagerWeb");
        try
        {
            Directory.CreateDirectory(dir);
        }
        catch (Exception)
        {
            /* TryWriteAll erstellt bei Bedarf erneut */
        }

        _filePath = Path.Combine(dir, "projekte.json");
    }

    public IReadOnlyList<SavedProjectDto> GetAll()
    {
        lock (_lock)
        {
            return ReadAll().ToList();
        }
    }

    /// <summary>Ordner muss existieren. Keine Duplikate (Pfad vergleich case-insensitive).</summary>
    public string? TryAdd(string? pathRaw, string? name)
    {
        if (string.IsNullOrWhiteSpace(pathRaw))
            return "Pfad darf nicht leer sein.";

        string full;
        try
        {
            full = Path.GetFullPath(pathRaw.Trim());
        }
        catch (Exception)
        {
            return "Der Pfad ist ungültig.";
        }

        if (!Directory.Exists(full))
            return "Der Pfad existiert nicht oder ist kein Ordner.";

        lock (_lock)
        {
            var list = ReadAll();
            foreach (var item in list)
            {
                if (PathsEqual(item.FullPath, full))
                    return "Dieses Projekt ist bereits gespeichert.";
            }

            var shortName = string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            list.Add(new SavedProjectDto(full, shortName));
            var writeErr = TryWriteAll(list);
            if (writeErr is not null)
                return writeErr;
        }

        return null;
    }

    /// <summary>Entfernt anhand des Pfads (wie in der Liste gespeichert).</summary>
    public string? TryRemove(string? pathRaw)
    {
        if (string.IsNullOrWhiteSpace(pathRaw))
            return "Pfad fehlt.";

        string full;
        try
        {
            full = Path.GetFullPath(pathRaw.Trim());
        }
        catch (Exception)
        {
            return "Der Pfad ist ungültig.";
        }

        lock (_lock)
        {
            var list = ReadAll();
            var idx = list.FindIndex(p => PathsEqual(p.FullPath, full));
            if (idx < 0)
                return "Projekt ist nicht in der Liste.";

            list.RemoveAt(idx);
            var writeErr = TryWriteAll(list);
            if (writeErr is not null)
                return writeErr;
        }

        return null;
    }

    private static bool PathsEqual(string a, string b)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(a),
                Path.GetFullPath(b),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception)
        {
            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }

    private List<SavedProjectDto> ReadAll()
    {
        if (!File.Exists(_filePath))
            return [];

        try
        {
            var json = File.ReadAllText(_filePath);
            if (string.IsNullOrWhiteSpace(json))
                return [];

            var parsed = JsonSerializer.Deserialize<List<SavedProjectDto>>(json, JsonOptions);
            return parsed ?? [];
        }
        catch (Exception)
        {
            return [];
        }
    }

    private string? TryWriteAll(List<SavedProjectDto> list)
    {
        try
        {
            var dir = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(list, JsonOptions);
            File.WriteAllText(_filePath, json);
            return null;
        }
        catch (Exception ex)
        {
            return $"Projektliste konnte nicht gespeichert werden: {ex.Message}";
        }
    }
}
