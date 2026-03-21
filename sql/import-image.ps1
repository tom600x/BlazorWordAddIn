<#
.SYNOPSIS
    Inserts an image file from disk into dbo.ImageSnippets as a base64-encoded record.

.DESCRIPTION
    Reads the image file, base64-encodes it, and inserts a row into dbo.ImageSnippets
    using a parameterized query.  Works with both LocalDB (dev) and Azure SQL (prod).

.PARAMETER ImagePath
    Full path to the image file on disk (e.g. C:\images\logo.png).

.PARAMETER Title
    Display name for the snippet (shown in the task pane).

.PARAMETER OwnerUpn
    UPN of the snippet owner (e.g. alice@contoso.com).

.PARAMETER Tags
    Optional comma-separated tags (e.g. "logo,header,branding").

.PARAMETER ConnectionString
    SQL connection string.
    LocalDB default:  "Server=(localdb)\mssqllocaldb;Database=BlazorWordAddIn;Trusted_Connection=True;"
    Azure SQL example: "Server=myserver.database.windows.net;Database=BlazorWordAddIn;Authentication=Active Directory Default;"

.EXAMPLE
    # LocalDB (dev)
    .\import-image.ps1 `
        -ImagePath  "C:\images\agency-logo.png" `
        -Title      "Agency Logo" `
        -OwnerUpn   "alice@contoso.com" `
        -Tags       "logo,header"

.EXAMPLE
    # Azure SQL
    .\import-image.ps1 `
        -ImagePath        "C:\images\agency-logo.png" `
        -Title            "Agency Logo" `
        -OwnerUpn         "alice@contoso.com" `
        -Tags             "logo,header" `
        -ConnectionString "Server=myserver.database.windows.net;Database=BlazorWordAddIn;Authentication=Active Directory Default;"
#>
param(
    [Parameter(Mandatory)][string]$ImagePath,
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][string]$OwnerUpn,
    [string]$Tags = $null,
    [string]$ConnectionString = "Server=(localdb)\mssqllocaldb;Database=BlazorWordAddIn;Trusted_Connection=True;"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Validate the file ──────────────────────────────────────────────────────────
if (-not (Test-Path $ImagePath)) {
    Write-Error "File not found: $ImagePath"
    exit 1
}

$extension = [System.IO.Path]::GetExtension($ImagePath).ToLower()
$mimeMap = @{
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.jpeg' = 'image/jpeg'
    '.gif'  = 'image/gif'
    '.bmp'  = 'image/bmp'
    '.webp' = 'image/webp'
    '.svg'  = 'image/svg+xml'
}
$mimeType = $mimeMap[$extension]
if (-not $mimeType) {
    Write-Error "Unsupported file extension '$extension'. Supported: $($mimeMap.Keys -join ', ')"
    exit 1
}

# ── Base64-encode the file ─────────────────────────────────────────────────────
$bytes    = [System.IO.File]::ReadAllBytes($ImagePath)
$base64   = [System.Convert]::ToBase64String($bytes)
$fileSize = [math]::Round($bytes.Length / 1KB, 1)
Write-Host "File: $ImagePath ($fileSize KB, $mimeType)"

# ── Insert into dbo.ImageSnippets ─────────────────────────────────────────────
Add-Type -AssemblyName System.Data

$conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
try {
    $conn.Open()

    $sql = @"
INSERT INTO dbo.ImageSnippets (OwnerUpn, Title, ImageBase64, MimeType, Tags)
VALUES (@OwnerUpn, @Title, @ImageBase64, @MimeType, @Tags);
SELECT SCOPE_IDENTITY() AS NewId;
"@

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $cmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OwnerUpn",   $OwnerUpn)))  | Out-Null
    $cmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Title",      $Title)))     | Out-Null
    $cmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ImageBase64",$base64)))    | Out-Null
    $cmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@MimeType",   $mimeType)))  | Out-Null

    $tagsParam = New-Object System.Data.SqlClient.SqlParameter("@Tags", [System.Data.SqlDbType]::NVarChar, 500)
    $tagsParam.Value = if ($Tags) { $Tags } else { [System.DBNull]::Value }
    $cmd.Parameters.Add($tagsParam) | Out-Null

    $newId = $cmd.ExecuteScalar()
    Write-Host "Inserted: Id=$newId  Title='$Title'  Owner=$OwnerUpn"
}
finally {
    $conn.Close()
}
