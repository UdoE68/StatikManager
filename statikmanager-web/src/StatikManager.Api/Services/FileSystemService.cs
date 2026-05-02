using System.IO;
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
}
