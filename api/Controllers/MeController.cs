using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SnippetsApi.DTOs;
using SnippetsApi.Extensions;

namespace SnippetsApi.Controllers;

/// <summary>
/// Returns identity information extracted from the validated Office SSO token.
/// The Blazor UI uses this to display the signed-in user's name / UPN.
/// </summary>
[ApiController]
[Route("api/me")]
[Authorize]
public class MeController : ControllerBase
{
    [HttpGet]
    public ActionResult<UserInfoDto> Get()
    {
        var upn = User.GetUpn();
        if (string.IsNullOrEmpty(upn))
            return Unauthorized("Unable to determine user identity from token claims.");

        return Ok(new UserInfoDto(
            Upn: upn,
            DisplayName: User.GetDisplayName(),
            Oid: User.GetOid(),
            Tid: User.GetTid()));
    }
}
