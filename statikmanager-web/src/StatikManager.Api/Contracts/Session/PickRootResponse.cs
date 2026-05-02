namespace StatikManager.Api.Contracts.Session;

/// <summary>
/// Ergebnis des Ordnerdialogs. <see cref="RootPath"/> ist null bei Abbruch.
/// </summary>
public sealed record PickRootResponse(string? RootPath);
