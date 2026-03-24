<#
.SYNOPSIS
    Starts the API/Blazor app and sideloads the Word add-in for local development.

.DESCRIPTION
    1. Registers the localhost manifest in the Windows registry so Word can load it.
    2. Builds and launches the API project (which also hosts the Blazor WASM frontend)
       on https://localhost:7001.
    3. Opens Word if requested.

    When you close the script (Ctrl+C), the API is stopped and the manifest is
    unregistered from the registry.

.PREREQUISITES
    - .NET 8 SDK
    - Node.js / npm (for npx)
    - Office desktop (Word) installed
    - Run once first:  npx office-addin-dev-certs install
    - Run once first:  dotnet dev-certs https --trust
#>

param(
    [string]$Manifest = "add-in\manifest.localhost.xml",
    [int]$TimeoutSeconds = 60
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$manifestPath = Join-Path $root $Manifest

# ── 0. Fix WebView2 data-directory permission error ─────────────────
# Office's OneAuth component creates a WebView2 cache next to WINWORD.EXE in
# Program Files.  If the directory doesn't exist the user gets:
#   "Microsoft Edge can't read and write to its data directory"
# Fix: create the directory once (requires elevation).
$wv2Dir = "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE.OneAuth.WebView2\EBWebView"
if (-not (Test-Path $wv2Dir)) {
    Write-Host "Creating OneAuth WebView2 data directory (requires elevation) ..." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -Wait -ArgumentList @(
        "-NoProfile", "-Command",
        "New-Item -Path '$wv2Dir' -ItemType Directory -Force | Out-Null;" +
        "`$acl = Get-Acl (Split-Path '$wv2Dir');" +
        "`$rule = New-Object System.Security.AccessControl.FileSystemAccessRule('BUILTIN\Users','FullControl','ContainerInherit,ObjectInherit','None','Allow');" +
        "`$acl.AddAccessRule(`$rule);" +
        "Set-Acl (Split-Path '$wv2Dir') `$acl"
    )
    if (Test-Path $wv2Dir) {
        Write-Host "WebView2 data directory created." -ForegroundColor Green
    } else {
        Write-Host "WARNING: Could not create WebView2 directory. You may see a data-directory error in Word." -ForegroundColor Red
    }
}

# ── 1. Register the manifest via the Windows registry ───────────────
Write-Host "Registering add-in manifest in Windows registry ..." -ForegroundColor Cyan
npx -y office-addin-dev-settings sideload $manifestPath
Write-Host "Manifest registered." -ForegroundColor Green

# ── 2. Start the API ────────────────────────────────────────────────
Write-Host "Starting API on https://localhost:7001 ..." -ForegroundColor Cyan
$apiJob = Start-Job -ScriptBlock {
    param($dir)
    Set-Location $dir
    dotnet run --project api\Api.csproj --launch-profile Api
} -ArgumentList $root

# ── 3. Wait for the server to be ready ──────────────────────────────
Write-Host "Waiting for API to respond ..." -ForegroundColor Cyan
$ready = $false
$elapsed = 0
while (-not $ready -and $elapsed -lt $TimeoutSeconds) {
    Start-Sleep -Seconds 2
    $elapsed += 2
    try {
        $response = Invoke-WebRequest -Uri "https://localhost:7001" -UseBasicParsing -TimeoutSec 5 -ErrorAction SilentlyContinue
        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500) {
            $ready = $true
        }
    } catch {
        # Server not up yet — keep waiting
    }
    Write-Host "  ... $elapsed s" -ForegroundColor DarkGray
}

if (-not $ready) {
    Write-Host "API did not start within $TimeoutSeconds seconds." -ForegroundColor Red
    Write-Host "Check for build errors:" -ForegroundColor Red
    Receive-Job -Job $apiJob
    Stop-Job -Job $apiJob
    Remove-Job -Job $apiJob
    npx -y office-addin-dev-settings unregister $manifestPath 2>$null
    exit 1
}

Write-Host "API is running on https://localhost:7001" -ForegroundColor Green

Write-Host ""
Write-Host "=== Add-in is sideloaded. Word should already be open with the add-in loaded. ===" -ForegroundColor Yellow
Write-Host "=== Look for 'Word Snippets (Dev)' in the Home ribbon. ===" -ForegroundColor Yellow
Write-Host "=== Press Ctrl+C to stop the API and unregister the add-in. ===" -ForegroundColor Yellow
Write-Host ""

# ── 5. Keep running until Ctrl+C ────────────────────────────────────
try {
    while ($true) {
        Start-Sleep -Seconds 5
        # Surface any API errors
        if ($apiJob.State -eq 'Failed') {
            Write-Host "API process failed:" -ForegroundColor Red
            Receive-Job -Job $apiJob
            break
        }
    }
} finally {
    # ── Cleanup ─────────────────────────────────────────────────────
    Write-Host "`nStopping API ..." -ForegroundColor Cyan
    Stop-Job -Job $apiJob -ErrorAction SilentlyContinue
    Remove-Job -Job $apiJob -ErrorAction SilentlyContinue

    Write-Host "Unregistering add-in manifest ..." -ForegroundColor Cyan
    npx -y office-addin-dev-settings unregister $manifestPath 2>$null

    Write-Host "Done." -ForegroundColor Green
}
