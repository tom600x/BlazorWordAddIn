using System.Net.Http.Headers;
using System.Net.Http.Json;
using BlazorUI.Models;

namespace BlazorUI.Services;

/// <summary>
/// Calls the backend API, automatically acquiring and attaching the Office SSO token
/// as a Bearer header before each request.
/// </summary>
public class ApiService
{
    private readonly HttpClient _http;
    private readonly OfficeInteropService _officeInterop;

    // Cache the token within the Blazor session to avoid repeated getAccessToken calls.
    // Token lifetime is typically 1 hour; a real app would check expiry (the JWT "exp" claim).
    private string? _cachedToken;

    // Serialize token acquisition – Office.auth.getAccessToken() only allows one call at a time.
    private readonly SemaphoreSlim _tokenLock = new(1, 1);

    public ApiService(HttpClient http, OfficeInteropService officeInterop)
    {
        _http = http;
        _officeInterop = officeInterop;
    }

    /// <summary>
    /// Forces re-acquisition of the Office SSO token on the next API call.
    /// Call this when a previous request failed with 401.
    /// </summary>
    public void InvalidateToken() => _cachedToken = null;

    // -------------------------------------------------------------------------
    // API methods
    // -------------------------------------------------------------------------

    /// <summary>Gets the signed-in user's identity from the backend (GET /api/me).</summary>
    public async Task<UserInfoDto?> GetMeAsync()
    {
        await SetAuthHeaderAsync();
        return await GetAsync<UserInfoDto>("api/me");
    }

    /// <summary>Gets the text snippets accessible to the signed-in user.</summary>
    public async Task<List<TextSnippetDto>> GetTextSnippetsAsync()
    {
        await SetAuthHeaderAsync();
        return await GetAsync<List<TextSnippetDto>>("api/text-snippets")
               ?? new List<TextSnippetDto>();
    }

    /// <summary>Gets image snippet summaries (without the full Base64 payload).</summary>
    public async Task<List<ImageSnippetDto>> GetImageSnippetsAsync()
    {
        await SetAuthHeaderAsync();
        return await GetAsync<List<ImageSnippetDto>>("api/image-snippets")
               ?? new List<ImageSnippetDto>();
    }

    /// <summary>
    /// Fetches the full Base64 image content for a single image snippet.
    /// Lazy-loaded when the user clicks "Insert" to avoid large initial payloads.
    /// </summary>
    public async Task<ImageSnippetContentDto?> GetImageContentAsync(int id)
    {
        await SetAuthHeaderAsync();
        return await GetAsync<ImageSnippetContentDto>($"api/image-snippets/{id}/content");
    }

    /// <summary>
    /// GET helper that reads the response body on error so the actual server
    /// error message is surfaced instead of a generic "500 Internal Server Error".
    /// </summary>
    private async Task<T?> GetAsync<T>(string url)
    {
        var response = await _http.GetAsync(url);
        if (response.IsSuccessStatusCode)
            return await response.Content.ReadFromJsonAsync<T>();

        var body = await response.Content.ReadAsStringAsync();
        // Truncate long error bodies for display
        if (body.Length > 500) body = body[..500];
        throw new HttpRequestException(
            $"{(int)response.StatusCode} {response.ReasonPhrase}: {body}",
            null,
            response.StatusCode);
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Acquires the token once and caches it. Safe to call from parallel tasks.
    /// </summary>
    public async Task EnsureTokenAsync()
    {
        if (!string.IsNullOrEmpty(_cachedToken)) return;

        await _tokenLock.WaitAsync();
        try
        {
            if (string.IsNullOrEmpty(_cachedToken))
                _cachedToken = await _officeInterop.GetOfficeTokenAsync();
        }
        finally { _tokenLock.Release(); }
    }

    private async Task SetAuthHeaderAsync()
    {
        await EnsureTokenAsync();

        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", _cachedToken);
    }
}
