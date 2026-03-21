namespace SnippetsApi.Models;

/// <summary>
/// Grants a second user read access to all of an owner's snippets.
/// Composite primary key: (OwnerUpn, AllowedUpn).
/// </summary>
public class AllowedUser
{
    public string OwnerUpn { get; set; } = string.Empty;

    public string AllowedUpn { get; set; } = string.Empty;

    public DateTime GrantedUtc { get; set; } = DateTime.UtcNow;
}
