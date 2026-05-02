namespace StatikManager.Api.Contracts.Session;

/// <summary>Setzt das Projekt-Root (nur im Speicher).</summary>
public sealed record SetRootRequest(string? RootPath);
