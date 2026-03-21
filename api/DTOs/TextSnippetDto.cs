namespace SnippetsApi.DTOs;

/// <param name="Id">Database identifier.</param>
/// <param name="Title">Display title.</param>
/// <param name="BodyText">Full text content.</param>
/// <param name="Tags">Comma-separated tags, may be null.</param>
/// <param name="UpdatedUtc">Last-modified timestamp (UTC).</param>
public record TextSnippetDto(
    int Id,
    string Title,
    string BodyText,
    string? Tags,
    DateTime UpdatedUtc);
