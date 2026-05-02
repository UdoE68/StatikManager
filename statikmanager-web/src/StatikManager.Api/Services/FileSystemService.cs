using System.IO;
using StatikManager.Api.Contracts.Browse;
using StatikManager.Api.Contracts.Session;

namespace StatikManager.Api.Services;

public sealed class FileSystemService : IFileSystemService
{
    private readonly object _lock = new();
    private string? _rootPath;

    public SessionResponse GetSession()
    {
        lock (_lock)
        {
            return new SessionResponse(_rootPath);
        }
    }

    public string? TrySetRoot(string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath))
            return "Pfad darf nicht leer sein.";

        var trimmed = rootPath.Trim();
        string full;
        try
        {
            full = Path.GetFullPath(trimmed);
        }
        catch (Exception)
        {
            return "Der Pfad ist ungültig.";
        }

        if (Directory.Exists(full))
        {
            lock (_lock)
            {
                _rootPath = full;
            }

            return null;
        }

        if (File.Exists(full))
            return "Pfad muss ein Verzeichnis sein.";

        return "Der Pfad existiert nicht.";
    }

    public string? TryBrowse(string? relativePath, out BrowseResponse? response)
    {
        response = null;

        string? root;
        lock (_lock)
        {
            root = _rootPath;
        }

        if (string.IsNullOrEmpty(root))
            return "Kein Projekt gewählt.";

        string rootFull;
        try
        {
            rootFull = Path.GetFullPath(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        }
        catch (Exception)
        {
            return "Interner Fehler: Projekt-Root ist ungültig.";
        }

        string targetFull;
        try
        {
            var relNormalized = NormalizeRelativeInput(relativePath);
            targetFull = string.IsNullOrEmpty(relNormalized)
                ? rootFull
                : Path.GetFullPath(Path.Combine(rootFull, relNormalized));
        }
        catch (Exception)
        {
            return "Der Pfad ist ungültig.";
        }

        if (!LiegtUnterhalbOderIstRoot(rootFull, targetFull))
            return "Zugriff außerhalb des Projektordners ist nicht erlaubt.";

        if (File.Exists(targetFull))
            return "Pfad ist eine Datei, kein Ordner.";

        if (!Directory.Exists(targetFull))
            return "Der Ordner existiert nicht.";

        var entries = new List<BrowseEntry>();

        foreach (var dir in Directory.GetDirectories(targetFull))
        {
            var name = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (string.IsNullOrEmpty(name))
                continue;

            var relOut = ToApiRelativePath(rootFull, dir);
            entries.Add(new BrowseEntry(
                name,
                relOut,
                IsDirectory: true,
                SizeBytes: null,
                ModifiedUtc: Directory.GetLastWriteTimeUtc(dir)));
        }

        foreach (var file in Directory.GetFiles(targetFull))
        {
            var name = Path.GetFileName(file);
            if (string.IsNullOrEmpty(name))
                continue;

            var fi = new FileInfo(file);
            var relOut = ToApiRelativePath(rootFull, file);
            entries.Add(new BrowseEntry(
                name,
                relOut,
                IsDirectory: false,
                SizeBytes: fi.Length,
                ModifiedUtc: fi.LastWriteTimeUtc));
        }

        entries.Sort(static (a, b) =>
        {
            if (a.IsDirectory != b.IsDirectory)
                return a.IsDirectory ? -1 : 1;
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        });

        response = new BrowseResponse(entries);
        return null;
    }

    /// <summary>
    /// Nur Verzeichnis-separator-normalisierte relative Segmente, ohne führenden Slash.
    /// </summary>
    private static string NormalizeRelativeInput(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
            return "";

        var s = relativePath.Trim().Replace('/', Path.DirectorySeparatorChar)
            .TrimStart(Path.DirectorySeparatorChar);

        var parts = s.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? "" : Path.Combine(parts);
    }

    private static bool LiegtUnterhalbOderIstRoot(string rootFull, string candidateFull)
    {
        try
        {
            rootFull = Path.GetFullPath(rootFull);
            candidateFull = Path.GetFullPath(candidateFull);
        }
        catch
        {
            return false;
        }

        if (string.Equals(rootFull, candidateFull, StringComparison.OrdinalIgnoreCase))
            return true;

        var prefix = rootFull.EndsWith(Path.DirectorySeparatorChar)
            ? rootFull
            : rootFull + Path.DirectorySeparatorChar;

        return candidateFull.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
               && candidateFull.Length > prefix.Length;
    }

    private static string ToApiRelativePath(string rootFull, string absolutePath)
    {
        var rel = Path.GetRelativePath(rootFull, absolutePath);
        return rel.Replace(Path.DirectorySeparatorChar, '/')
            .Replace(Path.AltDirectorySeparatorChar, '/');
    }
}
