namespace BlazorUI.Models;

/// <summary>Identifies the signed-in user as returned by GET /api/me.</summary>
public class UserInfoDto
{
    public string Upn { get; set; } = string.Empty;
    public string? DisplayName { get; set; }
    public string? Oid { get; set; }
    public string? Tid { get; set; }
}
