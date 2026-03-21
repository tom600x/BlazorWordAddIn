namespace SnippetsApi.Models;

public class ImageSnippet
{
    public int Id { get; set; }

    /// <summary>UPN of the user who owns this snippet.</summary>
    public string OwnerUpn { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    /// <summary>Base64-encoded image content (no data: URI prefix).</summary>
    public string ImageBase64 { get; set; } = string.Empty;

    /// <summary>MIME type, e.g. image/png or image/jpeg.</summary>
    public string MimeType { get; set; } = "image/png";

    /// <summary>Comma-separated tags for filtering.</summary>
    public string? Tags { get; set; }

    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
}
