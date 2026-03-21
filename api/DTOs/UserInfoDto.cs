namespace SnippetsApi.DTOs;

/// <param name="Upn">User principal name (email).</param>
/// <param name="DisplayName">Display name from the 'name' claim.</param>
/// <param name="Oid">Entra object ID.</param>
/// <param name="Tid">Tenant ID.</param>
public record UserInfoDto(
    string Upn,
    string? DisplayName,
    string? Oid,
    string? Tid);
