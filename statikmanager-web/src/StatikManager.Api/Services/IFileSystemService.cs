using StatikManager.Api.Contracts.Browse;
using StatikManager.Api.Contracts.Session;

namespace StatikManager.Api.Services;

public interface IFileSystemService
{
    SessionResponse GetSession();

    /// <summary>
    /// Setzt das Root nach Validierung. Gibt null zurück bei Erfolg, sonst eine deutschsprachige Fehlermeldung.
    /// </summary>
    string? TrySetRoot(string rootPath);

    /// <summary>
    /// Listet den angegebenen Unterordner relativ zum Root. Leer/null = Root.
    /// </summary>
    string? TryBrowse(string? relativePath, out BrowseResponse? response);
}
