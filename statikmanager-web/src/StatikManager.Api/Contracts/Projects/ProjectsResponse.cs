namespace StatikManager.Api.Contracts.Projects;

public sealed record ProjectsResponse(IReadOnlyList<SavedProjectDto> Projects);
