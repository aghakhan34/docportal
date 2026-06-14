<#
=====================================================================
  Setup-Columns.ps1
  Creates the 4 columns the Document Portal needs on the "Documents"
  library of  https://intlexp.sharepoint.com/sites/Corporate

  Columns created (idempotent — safe to re-run, skips any that exist):
    • Division    Choice   (6 division labels)
    • Department  Text
    • Status      Choice   (Published / Draft / Review)
    • Starred     Yes/No

  ---------------------------------------------------------------------
  PREREQUISITES (one time)
    1. Install the module (run once, any PowerShell 7+ window):
         Install-Module PnP.PowerShell -Scope CurrentUser

    2. PnP needs an app (client) ID to sign in interactively. You can
       reuse the app registration you already made for the portal:
         clientId = 911c89fb-4db5-4523-9d4a-2f32e5e21d31
       On that app registration, in Entra → Authentication:
         • Add a platform → "Mobile and desktop applications"
           with redirect URI:  http://localhost
         • Set "Allow public client flows" = Yes
       (This is separate from the SPA redirect the web app uses — an
        app can have both. If you'd rather not touch it, register a
        dedicated PnP app or use the Graph SDK alternative in README.)

  RUN IT
       pwsh ./Setup-Columns.ps1
    You'll get a browser sign-in. Use an account with permission to
    edit the Corporate site's library.
  ---------------------------------------------------------------------
#>

# ---- settings ---------------------------------------------------------
$SiteUrl  = 'https://intlexp.sharepoint.com/sites/Corporate'
$ClientId = '911c89fb-4db5-4523-9d4a-2f32e5e21d31'
$Library  = 'Documents'

$DivisionChoices = @(
    'Corporate','Operations','Finance',
    'Human Resources','IT & Technology','Legal & Compliance'
)
$StatusChoices = @('Published','Draft','Review')
# ----------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

function Test-Column {
    param([string]$Name)
    try   { Get-PnPField -List $Library -Identity $Name -ErrorAction Stop | Out-Null; return $true }
    catch { return $false }
}

function Add-ChoiceColumn {
    param([string]$Name, [string[]]$Choices)
    if (Test-Column $Name) { Write-Host "  • $Name already exists — skipping" -ForegroundColor Yellow; return }
    Add-PnPField -List $Library -DisplayName $Name -InternalName $Name `
        -Type Choice -Choices $Choices -AddToDefaultView | Out-Null
    Write-Host "  + $Name (Choice) created" -ForegroundColor Green
}

function Add-TextColumn {
    param([string]$Name)
    if (Test-Column $Name) { Write-Host "  • $Name already exists — skipping" -ForegroundColor Yellow; return }
    Add-PnPField -List $Library -DisplayName $Name -InternalName $Name `
        -Type Text -AddToDefaultView | Out-Null
    Write-Host "  + $Name (Text) created" -ForegroundColor Green
}

function Add-BooleanColumn {
    param([string]$Name)
    if (Test-Column $Name) { Write-Host "  • $Name already exists — skipping" -ForegroundColor Yellow; return }
    Add-PnPField -List $Library -DisplayName $Name -InternalName $Name `
        -Type Boolean -AddToDefaultView | Out-Null
    Write-Host "  + $Name (Yes/No) created" -ForegroundColor Green
}

Write-Host "Creating columns on '$Library' ..." -ForegroundColor Cyan
Add-ChoiceColumn  -Name 'Division'   -Choices $DivisionChoices
Add-TextColumn    -Name 'Department'
Add-ChoiceColumn  -Name 'Status'     -Choices $StatusChoices
Add-BooleanColumn -Name 'Starred'

Write-Host "`nDone. Verifying:" -ForegroundColor Cyan
Get-PnPField -List $Library |
    Where-Object { $_.InternalName -in 'Division','Department','Status','Starred' } |
    Select-Object Title, InternalName, TypeAsString |
    Format-Table -AutoSize

Disconnect-PnPOnline
Write-Host "`nAll set. Now flip mode to 'sharepoint' in config.js and reload the app." -ForegroundColor Green
