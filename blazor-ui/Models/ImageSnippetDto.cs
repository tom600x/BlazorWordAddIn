namespace BlazorUI.Models;

/// <summary>
/// Summary row for an image snippet (from GET /api/image-snippets).
/// The full Base64 payload is fetched on-demand via /api/image-snippets/{id}/content.
/// </summary>
public class ImageSnippetDto
{
    public int Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public string? Tags { get; set; }
    public string MimeType { get; set; } = "image/png";
    public DateTime UpdatedUtc { get; set; }
}

/// <summary>Full image content returned by GET /api/image-snippets/{id}/content.</summary>
public class ImageSnippetContentDto
{
    public int Id { get; set; }
    public string ImageBase64 { get; set; } = string.Empty;
    public string MimeType { get; set; } = "image/png";
}
