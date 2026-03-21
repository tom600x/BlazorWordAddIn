namespace SnippetsApi.Models;

public class TextSnippet
{
    public int Id { get; set; }

    /// <summary>UPN of the user who owns this snippet.</summary>
    public string OwnerUpn { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public string BodyText { get; set; } = string.Empty;

    /// <summary>Comma-separated tags for filtering.</summary>
    public string? Tags { get; set; }

    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
}
