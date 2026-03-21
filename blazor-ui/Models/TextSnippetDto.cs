namespace BlazorUI.Models;

/// <summary>A text snippet returned by GET /api/text-snippets.</summary>
public class TextSnippetDto
{
    public int Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public string BodyText { get; set; } = string.Empty;
    public string? Tags { get; set; }
    public DateTime UpdatedUtc { get; set; }
}
