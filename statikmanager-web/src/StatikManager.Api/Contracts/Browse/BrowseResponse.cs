namespace StatikManager.Api.Contracts.Browse;

public sealed record BrowseResponse(IReadOnlyList<BrowseEntry> Entries);
