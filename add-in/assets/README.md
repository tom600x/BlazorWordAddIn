# Add-in Icon Assets

Place the following icon files in this directory before deploying:

| File          | Size    | Notes                                |
|---------------|---------|--------------------------------------|
| `icon-16.png` | 16×16   | Small toolbar icon                   |
| `icon-32.png` | 32×32   | Medium toolbar icon                  |
| `icon-80.png` | 80×80   | Large icon (get-started card, store) |

## Creating placeholder icons (PowerShell)

If you need quick placeholder icons for testing, you can use the .NET System.Drawing API or
download the Office Add-in Samples icons from:

  https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Assets

## Requirements

- Format: PNG (transparent background preferred)
- Must be served over HTTPS from `https://REPLACE_APP_HOSTNAME/assets/`
- The manifest references these via `https://REPLACE_APP_HOSTNAME/assets/icon-*.png`
- Place these files in `api/wwwroot/assets/` for the hosted deployment

## GCC note

Ensure the icon files are served from a domain that is accessible within the GCC network
boundary. When hosted on Azure App Service (commercial), verify network egress rules permit
the M365 Office clients to download these assets.
