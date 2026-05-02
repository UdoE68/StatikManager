namespace StatikManager.Api.Contracts.File;

public sealed record FileMetaResponse(
    string RelativePath,
    string Name,
    FileKind Kind,
    long SizeBytes,
    DateTime ModifiedUtc,
    string MimeType);
