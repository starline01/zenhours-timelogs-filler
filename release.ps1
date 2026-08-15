<#
    Publishes a change to every PC.

    Tampermonkey only pulls an update when @version goes UP, so forgetting to
    bump it is the one mistake that silently ships nothing. This bumps it,
    commits, and pushes in one step.

        .\release.ps1                          # 1.2.1 -> 1.2.2, default message
        .\release.ps1 -Message "fix lunch col" # same, with your commit message
        .\release.ps1 -Version 1.3.0           # set an exact version
#>
param(
    [string]$Version,
    [string]$Message = "Update Zenhours DTR Filler"
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$file = Join-Path $PSScriptRoot 'zenhours-dtr-filler.user.js'
if (-not (Test-Path $file)) { throw "Cannot find zenhours-dtr-filler.user.js next to this script." }

$content = Get-Content $file -Raw
if ($content -notmatch '(?m)^//\s*@version\s+(\S+)\s*$') { throw "No @version line found in the userscript." }
$current = $Matches[1]

if (-not $Version) {
    $parts = $current.Split('.')
    if ($parts.Count -lt 3) { throw "Version '$current' is not major.minor.patch - pass -Version explicitly." }
    $parts[2] = [string]([int]$parts[2] + 1)
    $Version = $parts -join '.'
}

if ([string]::IsNullOrWhiteSpace($Version)) { throw "Resolved an empty version." }
Write-Host "Version $current -> $Version" -ForegroundColor Cyan

# Write without a BOM: a byte-order mark ahead of the // ==UserScript== header
# can stop Tampermonkey recognising the metadata block.
$updated = $content -replace '(?m)^(//\s*@version\s+)\S+\s*$', ('${1}' + $Version)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($file, $updated, $utf8NoBom)

# Sanity check: syntax must be valid before it reaches anyone's browser.
$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
    node --check $file
    if ($LASTEXITCODE -ne 0) {
        [System.IO.File]::WriteAllText($file, $content, $utf8NoBom)   # roll the bump back
        throw "Syntax check failed - version reverted, nothing pushed."
    }
    Write-Host "Syntax OK" -ForegroundColor Green
}

git add -A
if ($?) {
    git commit -m "$Message (v$Version)"
    if ($?) {
        git push
        if ($?) {
            Write-Host ""
            Write-Host "Pushed v$Version." -ForegroundColor Green
            Write-Host "Each PC picks it up on its next update check (Tampermonkey default: ~daily)."
            Write-Host "To pull it now on a PC: Tampermonkey dashboard -> Check for userscript updates."
            Write-Host "raw.githubusercontent.com caches ~5 min, so it is not instant."
        }
    }
}
