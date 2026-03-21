using System.Text.Json;
using BlazorUI.Models;

namespace BlazorUI.Services;

/// <summary>
/// Calls the backend API using the browser's native fetch() (via JS interop) to avoid
/// Blazor WASM BrowserHttpHandler issues inside Office task pane iframes.
/// </summary>
public class ApiService
{
    private readonly OfficeInteropService _officeInterop;

    private static readonly JsonSerializerOptions JsonOptions =
        new() { PropertyNameCaseInsensitive = true };

    // Cache the token within the Blazor session.
    // Token lifetime is typically 1 hour.
    private string? _cachedToken;

    public ApiService(OfficeInteropService officeInterop)
    {
        _officeInterop = officeInterop;
    }

    /// <summary>
    /// Forces re-acquisition of the Office SSO token on the next API call.
    /// Call this when a previous request failed with 401.
    /// </summary>
    public void InvalidateToken() => _cachedToken = null;

    /// <summary>
    /// Injects a token obtained externally (e.g. via MSAL fallback dialog).
    /// </summary>
    public void SetCachedToken(string token) => _cachedToken = token;

    // -------------------------------------------------------------------------
    // API methods
    // -------------------------------------------------------------------------

    /// <summary>Gets the signed-in user's identity from the backend (GET /api/me).</summary>
    public async Task<UserInfoDto?> GetMeAsync()
    {
        await EnsureTokenAsync();
        var json = await _officeInterop.FetchJsonAsync("api/me", _cachedToken);
        return JsonSerializer.Deserialize<UserInfoDto>(json, JsonOptions);
    }

    /// <summary>Gets the text snippets accessible to the signed-in user.</summary>
    public async Task<List<TextSnippetDto>> GetTextSnippetsAsync()
    {
        await EnsureTokenAsync();
        var json = await _officeInterop.FetchJsonAsync("api/text-snippets", _cachedToken);
        return JsonSerializer.Deserialize<List<TextSnippetDto>>(json, JsonOptions) ?? new();
    }

    /// <summary>Gets image snippet summaries (without the full Base64 payload).</summary>
    public async Task<List<ImageSnippetDto>> GetImageSnippetsAsync()
    {
        await EnsureTokenAsync();
        var json = await _officeInterop.FetchJsonAsync("api/image-snippets", _cachedToken);
        return JsonSerializer.Deserialize<List<ImageSnippetDto>>(json, JsonOptions) ?? new();
    }

    /// <summary>
    /// Fetches the full Base64 image content for a single image snippet.
    /// Lazy-loaded when the user clicks "Insert" to avoid large initial payloads.
    /// </summary>
    public async Task<ImageSnippetContentDto?> GetImageContentAsync(int id)
    {
        await EnsureTokenAsync();
        var json = await _officeInterop.FetchJsonAsync($"api/image-snippets/{id}/content", _cachedToken);
        return JsonSerializer.Deserialize<ImageSnippetContentDto>(json, JsonOptions);
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Ensures _cachedToken is populated. If not already cached, attempts Office SSO.
    /// Throws <see cref="OfficeAuthException"/> if no token can be obtained.
    /// </summary>
    private async Task EnsureTokenAsync()
    {
        if (!string.IsNullOrEmpty(_cachedToken))
            return;

        // GetOfficeTokenAsync throws OfficeAuthException when SSO is unavailable
        // (e.g. Word Online, error 13006). The caller shows the Sign In button.
        _cachedToken = await _officeInterop.GetOfficeTokenAsync();
    }
}
