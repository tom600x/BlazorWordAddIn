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
        => await _js.InvokeAsync<bool>("officeInterop.isInOffice");
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
