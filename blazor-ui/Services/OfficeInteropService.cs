using System.Text.Json;
using Microsoft.JSInterop;

namespace BlazorUI.Services;

/// <summary>
/// Provides strongly-typed wrappers around the office interop JavaScript functions
/// defined in wwwroot/js/officeInterop.js.
/// </summary>
public class OfficeInteropService
{
    private readonly IJSRuntime _js;

    public OfficeInteropService(IJSRuntime js) => _js = js;

    /// <summary>
    /// Acquires an Office SSO token for the add-in's backend API.
    /// The token's audience is the Application ID URI (api://HOST/CLIENT_ID).
    /// </summary>
    /// <returns>Raw JWT string on success.</returns>
    /// <exception cref="OfficeAuthException">When Office cannot issue the token.</exception>
    public async Task<string> GetOfficeTokenAsync()
    {
        // If we're not hosted inside an Office client, return a clear error instead
        bool inOffice = await IsInOfficeAsync();
        if (!inOffice)
        {
            throw new OfficeAuthException(0, "Not running inside an Office host. Office SSO is unavailable.");
        }

        var result = await _js.InvokeAsync<JsonElement>("officeInterop.getOfficeToken");

        if (result.GetProperty("success").GetBoolean())
        {
            return result.GetProperty("token").GetString()
                ?? throw new OfficeAuthException(0, "Token was null despite success flag.");
        }

        int code = result.TryGetProperty("errorCode", out var codeEl) ? codeEl.GetInt32() : 0;
        string msg = result.TryGetProperty("errorMessage", out var msgEl)
            ? msgEl.GetString() ?? "Unknown SSO error"
            : "Unknown SSO error";

        throw new OfficeAuthException(code, msg);
    }

    /// <summary>
    /// Acquires an access token via MSAL interactive login using the Office Dialog API.
    /// Called as a fallback when <see cref="GetOfficeTokenAsync"/> fails (e.g. in Word Online).
    /// </summary>
    /// <returns>Raw JWT string on success.</returns>
    /// <exception cref="OfficeAuthException">When the MSAL dialog flow fails or the user cancels.</exception>
    public async Task<string> GetMsalTokenAsync()
    {
        var result = await _js.InvokeAsync<JsonElement>("officeInterop.getMsalTokenViaDialog");

        if (result.GetProperty("success").GetBoolean())
        {
            return result.GetProperty("token").GetString()
                ?? throw new OfficeAuthException(0, "MSAL token was null despite success flag.");
        }

        int code = result.TryGetProperty("errorCode", out var codeEl) ? codeEl.GetInt32() : 0;
        string msg = result.TryGetProperty("errorMessage", out var msgEl)
            ? msgEl.GetString() ?? "MSAL sign-in failed."
            : "MSAL sign-in failed.";

        throw new OfficeAuthException(code, msg);
    }

    /// <summary>
    /// Makes a GET request via the browser's native fetch() API, bypassing Blazor's BrowserHttpHandler.
    /// Returns the response body as a JSON string.
    /// </summary>
    /// <exception cref="HttpRequestException">On network error or non-2xx HTTP status.</exception>
    public async Task<string> FetchJsonAsync(string url, string? token)
    {
        var result = await _js.InvokeAsync<JsonElement>("officeInterop.fetchJson", url, (object?)token);

        if (result.GetProperty("ok").GetBoolean())
            return result.GetProperty("body").GetString() ?? "null";

        // fetch() threw (network-level failure: CORS, offline, etc.)
        if (result.TryGetProperty("error", out var errEl) && errEl.ValueKind != JsonValueKind.Null)
        {
            var msg = errEl.GetString() ?? "Unknown fetch error";
            throw new HttpRequestException(msg);
        }

        // HTTP error response — include the body and X-Auth-Error header for diagnostics.
        var status = result.GetProperty("status").GetInt32();
        var body = result.TryGetProperty("body", out var bodyEl) && bodyEl.ValueKind != JsonValueKind.Null
            ? bodyEl.GetString() ?? ""
            : "";
        // fetchJson also captures X-Auth-Error as a console warning; body should contain it too.
        var detail = string.IsNullOrWhiteSpace(body) ? $"HTTP {status} (no body)" : $"HTTP {status}: {body}";
        throw new HttpRequestException(detail, null, (System.Net.HttpStatusCode)status);
    }

    /// <summary>Inserts <paramref name="text"/> at the current Word selection.</summary>
    public async Task<bool> InsertTextAsync(string text)
    {
        var result = await _js.InvokeAsync<JsonElement>("officeInterop.insertText", text);
        return result.GetProperty("success").GetBoolean();
    }

    /// <summary>
    /// Inserts a Base64-encoded image at the current Word selection as an inline picture.
    /// </summary>
    /// <param name="base64">Raw Base64 (no data URI prefix).</param>
    public async Task<bool> InsertImageAsync(string base64)
    {
        var result = await _js.InvokeAsync<JsonElement>("officeInterop.insertImage", base64);
        return result.GetProperty("success").GetBoolean();
    }

    /// <summary>Returns true when the page is loaded inside an Office host.</summary>
    public async Task<bool> IsInOfficeAsync()
    {
        try
        {
            return await _js.InvokeAsync<bool>("officeInterop.isInOffice");
        }
        catch (JSException)
        {
            // JS interop failed (script not loaded or function missing) — treat as not in Office.
            return false;
        }
        catch (Exception)
        {
            // Be conservative: on any unexpected error, report not-in-office so UI can proceed.
            return false;
        }
    }
}

/// <summary>Wraps an Office SSO error code and message.</summary>
public class OfficeAuthException : Exception
{
    public int ErrorCode { get; }

    public OfficeAuthException(int code, string message) : base(message)
        => ErrorCode = code;

    /// <summary>
    /// User-facing guidance based on common Office SSO error codes.
    /// </summary>
    public string UserMessage => ErrorCode switch
    {
        13001 => "You are not signed in to Office. Please sign in and try again.",
        13002 => "You need to grant permission to this add-in. Please try again and accept the consent prompt.",
        13003 => "Microsoft personal accounts are not supported. Please sign in with your work or school account.",
        13005 => "This add-in is not configured for single sign-on. Contact your administrator.",
        13006 => "There was an Office client error. Try signing out of Office and signing back in.",
        13007 => "Another sign-in operation is in progress. Please wait and try again.",
        13008 => "The sign-in request timed out. Please try again.",
        13012 => "The add-in is not properly configured for SSO. Contact your administrator.",
        _ => $"Sign-in failed (code {ErrorCode}): {Message}"
    };
}
