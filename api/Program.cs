using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Identity.Web;
using Microsoft.EntityFrameworkCore;
using SnippetsApi.Data;
using SnippetsApi.Services;

var builder = WebApplication.CreateBuilder(args);

// ---------------------------------------------------------------------------
// Authentication — Microsoft.Identity.Web validates the Office SSO JWT.
//
// The token issued by Office.auth.getAccessToken has:
//   aud  = the Application ID URI  (api://HOSTNAME/CLIENT_ID)
//   iss  = https://login.microsoftonline.com/{tid}/v2.0
//   upn / preferred_username = the signed-in user's UPN
//
// For GCC tenants the issuer uses the COMMERCIAL Azure AD endpoint
// (login.microsoftonline.com), not login.microsoftonline.us.
// For GCC High / DoD, update AzureAd:Instance accordingly.
// ---------------------------------------------------------------------------
builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApi(
        jwtOptions =>
        {
            builder.Configuration.GetSection("AzureAd").Bind(jwtOptions);
            jwtOptions.Events = new JwtBearerEvents
            {
                OnAuthenticationFailed = ctx =>
                {
                    // Surface the real validation failure reason in a response header
                    // so front-end diagnostics can show it.
                    var msg = ctx.Exception?.Message ?? "unknown";
                    if (msg.Length > 500) msg = msg[..500];
                    ctx.Response.Headers.Append("X-Auth-Error", msg);
                    return Task.CompletedTask;
                }
            };
        },
        msOptions =>
        {
            builder.Configuration.GetSection("AzureAd").Bind(msOptions);
        });

builder.Services.AddAuthorization();

// ---------------------------------------------------------------------------
// CORS — allow the Blazor task pane to call this API.
// In production both are served from the same App Service (same origin),
// so CORS is not strictly needed. It IS needed for local dev when the Blazor
// dev-server runs on a different port than the API.
// ---------------------------------------------------------------------------
builder.Services.AddCors(options =>
{
    // Word Online loads the task pane through Microsoft's gateway infrastructure,
    // so the browser's request origin may not match tomwordaddin.azurewebsites.net.
    // Bearer token auth provides the security; we allow any origin here.
    options.AddPolicy("AddInPolicy", policy =>
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod()
              .WithExposedHeaders("X-Auth-Error"));
});

// ---------------------------------------------------------------------------
// Data access — Azure SQL with Managed Identity.
//
// Connection string uses "Authentication=Active Directory Default" which
// instructs Microsoft.Data.SqlClient to use DefaultAzureCredential.
// This works for:
//   • System-assigned Managed Identity on App Service (production)
//   • Azure CLI / developer credentials on dev machine (local dev with Azure SQL)
//   • Standard SQL auth connection string (local dev with SQL Server / LocalDB)
//
// For local dev with LocalDB, set "Trusted_Connection=True" and omit
// "Authentication=Active Directory Default".
// ---------------------------------------------------------------------------
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlServer(
        builder.Configuration.GetConnectionString("DefaultConnection"),
        sqlOptions => sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 3,
            maxRetryDelay: TimeSpan.FromSeconds(5),
            errorNumbersToAdd: null)));

builder.Services.AddScoped<SnippetRepository>();

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

var app = builder.Build();

// ---------------------------------------------------------------------------
// Middleware pipeline
// ---------------------------------------------------------------------------

// Return exception details in the response body (helps remote diagnostics).
app.UseExceptionHandler(errorApp =>
{
    errorApp.Run(async context =>
    {
        context.Response.StatusCode = 500;
        context.Response.ContentType = "text/plain";
        var feature = context.Features.Get<Microsoft.AspNetCore.Diagnostics.IExceptionHandlerFeature>();
        var msg = feature?.Error?.ToString() ?? "Unknown server error";
        if (msg.Length > 2000) msg = msg[..2000];
        await context.Response.WriteAsync(msg);
    });
});

// Serve Blazor WASM static files (output from BlazorUI project reference).
// The _framework/ directory and index.html come from here.
app.UseBlazorFrameworkFiles();

// Serve /js/*.js with no-cache so the browser always revalidates.
// This prevents stale officeInterop.js from being served after deploys.
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        var path = ctx.Context.Request.Path.Value ?? "";
        // Don't interfere with _framework files (they have content-hash filenames already)
        if (!path.StartsWith("/_framework/"))
        {
            ctx.Context.Response.Headers["Cache-Control"] = "no-cache, must-revalidate";
        }
    }
});

app.UseHttpsRedirection();
app.UseCors("AddInPolicy");

app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

// Fallback to index.html for Blazor client-side routing
// (all non-/api/* requests serve the SPA shell)
app.MapFallbackToFile("index.html");

app.Run();
