using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SnippetsApi.DTOs;
using SnippetsApi.Extensions;
using SnippetsApi.Services;

namespace SnippetsApi.Controllers;

[ApiController]
[Route("api/text-snippets")]
[Authorize]
public class TextSnippetsController : ControllerBase
{
    private readonly SnippetRepository _repo;

    public TextSnippetsController(SnippetRepository repo)
    {
        _repo = repo;
    }

    /// <summary>
    /// Returns all text snippets owned by the authenticated user.
    /// </summary>
    [HttpGet]
    public async Task<ActionResult<IEnumerable<TextSnippetDto>>> GetAsync()
    {
        var upn = User.GetUpn();
        if (string.IsNullOrEmpty(upn))
            return Unauthorized("Unable to determine user identity from token claims.");

        var snippets = await _repo.GetTextSnippetsForUserAsync(upn);
        return Ok(snippets);
    }
}
