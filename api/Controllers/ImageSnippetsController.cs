using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SnippetsApi.DTOs;
using SnippetsApi.Extensions;
using SnippetsApi.Services;

namespace SnippetsApi.Controllers;

[ApiController]
[Route("api/image-snippets")]
[Authorize]
public class ImageSnippetsController : ControllerBase
{
    private readonly SnippetRepository _repo;

    public ImageSnippetsController(SnippetRepository repo)
    {
        _repo = repo;
    }

    /// <summary>
    /// Returns image snippet summaries (no base64 payload) visible to the user.
    /// The full image is fetched lazily via GET /{id}/content.
    /// </summary>
    [HttpGet]
    public async Task<ActionResult<IEnumerable<ImageSnippetDto>>> GetAsync()
    {
        var upn = User.GetUpn();
        if (string.IsNullOrEmpty(upn))
            return Unauthorized("Unable to determine user identity from token claims.");

        var snippets = await _repo.GetImageSnippetsForUserAsync(upn);
        return Ok(snippets);
    }

    /// <summary>
    /// Returns the full base64 image content for a single snippet.
    /// Called just before inserting the image into Word to avoid loading all images upfront.
    /// </summary>
    [HttpGet("{id:int}/content")]
    public async Task<ActionResult<ImageSnippetContentDto>> GetContentAsync(int id)
    {
        var upn = User.GetUpn();
        if (string.IsNullOrEmpty(upn))
            return Unauthorized("Unable to determine user identity from token claims.");

        var content = await _repo.GetImageContentAsync(id, upn);
        if (content is null)
            return NotFound();

        return Ok(content);
    }
}
