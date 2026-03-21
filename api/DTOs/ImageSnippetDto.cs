namespace SnippetsApi.DTOs;

/// <summary>
/// Summary returned by GET /api/image-snippets (no base64 payload).
/// Fetch full image via GET /api/image-snippets/{id}/content.
/// </summary>
public record ImageSnippetDto(
    int Id,
    string Title,
    string? Tags,
    string MimeType,
    DateTime UpdatedUtc);

/// <summary>
/// Full image content returned by GET /api/image-snippets/{id}/content.
/// </summary>
public record ImageSnippetContentDto(
    int Id,
    string ImageBase64,
    string MimeType);
