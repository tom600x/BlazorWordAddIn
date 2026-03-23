using Microsoft.EntityFrameworkCore;
using SnippetsApi.Data;
using SnippetsApi.DTOs;

namespace SnippetsApi.Services;

/// <summary>
/// Data access layer for text and image snippets.
/// All queries scope results to the requesting user's UPN.
/// </summary>
public class SnippetRepository
{
    private readonly AppDbContext _db;

    public SnippetRepository(AppDbContext db)
    {
        _db = db;
    }

    /// <summary>
    /// Returns all TextSnippets owned by the authenticated user.
    /// </summary>
    public async Task<List<TextSnippetDto>> GetTextSnippetsForUserAsync(string upn)
    {
        return await _db.TextSnippets
            .Where(s => s.OwnerUpn == upn)
            .OrderBy(s => s.Title)
            .Select(s => new TextSnippetDto(s.Id, s.Title, s.BodyText, s.Tags, s.UpdatedUtc))
            .ToListAsync();
    }

    /// <summary>
    /// Returns ImageSnippet summaries (no base64 payload) owned by the authenticated user.
    /// </summary>
    public async Task<List<ImageSnippetDto>> GetImageSnippetsForUserAsync(string upn)
    {
        return await _db.ImageSnippets
            .Where(s => s.OwnerUpn == upn)
            .OrderBy(s => s.Title)
            .Select(s => new ImageSnippetDto(s.Id, s.Title, s.Tags, s.MimeType, s.UpdatedUtc))
            .ToListAsync();
    }

    /// <summary>
    /// Returns the full base64 image content for a specific snippet,
    /// ensuring the requesting user owns it.
    /// </summary>
    public async Task<ImageSnippetContentDto?> GetImageContentAsync(int id, string upn)
    {
        var snippet = await _db.ImageSnippets
            .Where(s => s.Id == id && s.OwnerUpn == upn)
            .FirstOrDefaultAsync();

        if (snippet is null) return null;

        return new ImageSnippetContentDto(snippet.Id, snippet.ImageBase64, snippet.MimeType);
    }
}
