namespace StatikManager.Api.Contracts.Session;

/// <summary>Aktuelle Session: gesetztes Projekt-Root oder null.</summary>
public sealed record SessionResponse(string? RootPath);
