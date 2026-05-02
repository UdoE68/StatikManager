namespace StatikManager.Api.Contracts.Projects;

/// <summary>Gespeichertes Projekt (persistiert in projekte.json).</summary>
public sealed record SavedProjectDto(string Path, string? Name);
