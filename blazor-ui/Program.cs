using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using BlazorUI;
using BlazorUI.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

// Root components
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

// When hosted by the API project, the base address IS the API base address.
// All /api/* calls are relative, so this works for same-origin deployment.
// For local dev (standalone mode), ApiBaseUrl is overridden by appsettings.Development.json.
var apiBaseUrl = builder.Configuration["ApiBaseUrl"]
    ?? builder.HostEnvironment.BaseAddress;

builder.Services.AddScoped(_ => new HttpClient
{
    BaseAddress = new Uri(apiBaseUrl)
});

builder.Services.AddScoped<OfficeInteropService>();
builder.Services.AddScoped<ApiService>();

await builder.Build().RunAsync();
