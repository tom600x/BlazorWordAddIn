using System.Security.Claims;

namespace SnippetsApi.Extensions;

/// <summary>
/// Extension methods that extract standard claims from an Office SSO JWT.
/// Office SSO tokens (v2.0) use claim names that differ from the legacy WS-Federation names.
/// </summary>
public static class ClaimsPrincipalExtensions
{
    /// <summary>
    /// Returns the user's UPN / email, checking multiple claim name variants.
    /// v2.0 tokens use "preferred_username"; legacy tokens may use "upn".
    /// </summary>
    public static string? GetUpn(this ClaimsPrincipal principal)
        => principal.FindFirstValue("preferred_username")
        ?? principal.FindFirstValue("upn")
        ?? principal.FindFirstValue(ClaimTypes.Upn)
        ?? principal.FindFirstValue(ClaimTypes.Email)
        ?? principal.FindFirstValue("email")
        // Last-resort: use the object ID so JWT-valid tokens at least get through.
        // Snippets are keyed by UPN so this user will see an empty list rather than an error.
        ?? principal.FindFirstValue("oid")
        ?? principal.FindFirstValue("http://schemas.microsoft.com/identity/claims/objectidentifier");

    /// <summary>
    /// Returns the Entra object ID ("oid" claim).
    /// </summary>
    public static string? GetOid(this ClaimsPrincipal principal)
        => principal.FindFirstValue("oid")
        ?? principal.FindFirstValue("http://schemas.microsoft.com/identity/claims/objectidentifier");

    /// <summary>
    /// Returns the tenant ID ("tid" claim).
    /// </summary>
    public static string? GetTid(this ClaimsPrincipal principal)
        => principal.FindFirstValue("tid")
        ?? principal.FindFirstValue("http://schemas.microsoft.com/identity/claims/tenantid");

    /// <summary>
    /// Returns the display name from the "name" claim.
    /// </summary>
    public static string? GetDisplayName(this ClaimsPrincipal principal)
        => principal.FindFirstValue("name")
        ?? principal.FindFirstValue(ClaimTypes.GivenName)
        ?? principal.FindFirstValue(ClaimTypes.Name);
}
