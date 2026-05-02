using System.Text.Json.Serialization;

namespace StatikManager.Api.Contracts.Projects;

/// <summary>Gespeichertes Projekt (persistiert in projekte.json). JSON-Feld weiterhin „path“.</summary>
public sealed record SavedProjectDto(
    [property: JsonPropertyName("path")] string FullPath,
    string? Name);
