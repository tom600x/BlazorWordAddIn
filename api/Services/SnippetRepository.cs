using Microsoft.EntityFrameworkCore;
using SnippetsApi.Data;
using SnippetsApi.DTOs;

namespace SnippetsApi.Services;

/// <summary>
/// Data access layer for text and image snippets.
/// All queries scope results to the requesting user's UPN, including snippets
/// shared with them via the AllowedUsers table.
/// </summary>
public class SnippetRepository
{
    private readonly AppDbContext _db;

    public SnippetRepository(AppDbContext db)
    {
        _db = db;
    }

    /// <summary>
    /// Returns all TextSnippets the user owns OR has been granted access to.
    /// </summary>
    public async Task<List<TextSnippetDto>> GetTextSnippetsForUserAsync(string upn)
    {
        // Collect owner UPNs that have shared with this user
        var allowedOwners = await _db.AllowedUsers
            .Where(a => a.AllowedUpn == upn)
            .Select(a => a.OwnerUpn)
            .ToListAsync();

        return await _db.TextSnippets
            .Where(s => s.OwnerUpn == upn || allowedOwners.Contains(s.OwnerUpn))
            .OrderBy(s => s.Title)
            .Select(s => new TextSnippetDto(s.Id, s.Title, s.BodyText, s.Tags, s.UpdatedUtc))
            .ToListAsync();
    }

    /// <summary>
    /// Returns ImageSnippet summaries (no base64 payload) visible to the user.
    /// </summary>
    public async Task<List<ImageSnippetDto>> GetImageSnippetsForUserAsync(string upn)
    {
        var allowedOwners = await _db.AllowedUsers
            .Where(a => a.AllowedUpn == upn)
            .Select(a => a.OwnerUpn)
            .ToListAsync();

        return await _db.ImageSnippets
            .Where(s => s.OwnerUpn == upn || allowedOwners.Contains(s.OwnerUpn))
            .OrderBy(s => s.Title)
            .Select(s => new ImageSnippetDto(s.Id, s.Title, s.Tags, s.MimeType, s.UpdatedUtc))
            .ToListAsync();
    }

    /// <summary>
    /// Returns the full base64 image content for a specific snippet,
    /// ensuring the requesting user is authorized to access it.
    /// </summary>
    public async Task<ImageSnippetContentDto?> GetImageContentAsync(int id, string upn)
    {
        var allowedOwners = await _db.AllowedUsers
            .Where(a => a.AllowedUpn == upn)
            .Select(a => a.OwnerUpn)
            .ToListAsync();

        var snippet = await _db.ImageSnippets
            .Where(s => s.Id == id && (s.OwnerUpn == upn || allowedOwners.Contains(s.OwnerUpn)))
            .FirstOrDefaultAsync();

        if (snippet is null) return null;

        return new ImageSnippetContentDto(snippet.Id, snippet.ImageBase64, snippet.MimeType);
    }
}
