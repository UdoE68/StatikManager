using StatikManager.Api.Contracts.Session;

namespace StatikManager.Api.Services;

public interface IFileSystemService
{
    SessionResponse GetSession();

    /// <summary>
    /// Setzt das Root nach Validierung. Gibt null zurück bei Erfolg, sonst eine deutschsprachige Fehlermeldung.
    /// </summary>
    string? TrySetRoot(string rootPath);
}
