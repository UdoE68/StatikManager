namespace StatikManager.Api.Contracts.Browse;

/// <summary>Ein Eintrag im aktuellen Verzeichnis (relativ zum Projekt-Root).</summary>
public sealed record BrowseEntry(
    string Name,
    string RelativePath,
    bool IsDirectory,
    long? SizeBytes,
    DateTime ModifiedUtc);
