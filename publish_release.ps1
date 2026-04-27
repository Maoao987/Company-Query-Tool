param(
    [ValidateSet("fix", "release")]
    [string]$Mode = "fix",
    [string]$Version = "",
    [string]$Repo = "Maoao987/Company-Query-Tool",
    [switch]$SkipBuild,
    [switch]$SkipPush,
    [switch]$SkipRelease
)

$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

function Get-CurrentVersion {
    return (Get-Content -Path (Join-Path $PSScriptRoot "version.txt") -Raw -Encoding UTF8).Trim()
}

function Set-CurrentVersion([string]$NewVersion) {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText((Join-Path $PSScriptRoot "version.txt"), "$NewVersion`n", $utf8NoBom)
}

function Get-BumpedPatchVersion([string]$CurrentVersion) {
    $parts = $CurrentVersion.Split(".")
    if ($parts.Count -lt 3) {
        throw "Invalid version format: $CurrentVersion"
    }
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $patch = [int]$parts[2] + 1
    return "$major.$minor.$patch"
}

function Get-ChangedFiles {
    $lines = git status --porcelain
    $files = @()
    foreach ($line in $lines) {
        if (-not $line) { continue }
        $path = $line.Substring(3).Trim()
        if ($path) { $files += $path }
    }
    return $files | Sort-Object -Unique
}

function Get-AreaText([string[]]$Files) {
    $areas = New-Object System.Collections.Generic.List[string]
    $joined = ($Files -join "`n")
    if ($joined -match "(^|`n)app\.py($|`n)") { $areas.Add("Update Streamlit UI and input validation") }
    if ($joined -match "(^|`n)company_query\.py($|`n)") { $areas.Add("Update company, stock, ETF, and fund query logic") }
    if ($joined -match "(^|`n)pdf_report\.py($|`n)") { $areas.Add("Update PDF report content") }
    if ($joined -match "(^|`n)findbiz_scraper\.py($|`n)") { $areas.Add("Update findbiz scraping flow") }
    if ($joined -match "(^|`n)web_snapshot\.py($|`n)") { $areas.Add("Update source snapshot PDF flow") }
    if ($joined -match "(^|`n)(install\.ps1|CompanyQueryToolSetup\.iss|start\.bat|start_hidden\.vbs)($|`n)") { $areas.Add("Update Windows installer and startup flow") }
    if ($joined -match "(^|`n)(update_manager\.py|update_config\.json)($|`n)") { $areas.Add("Update app update flow") }
    if ($joined -match "(^|`n)(version\.txt|version\.iss\.inc|sync_version\.ps1)($|`n)") { $areas.Add("Update version metadata") }
    if ($joined -match "(^|`n)README\.md($|`n)") { $areas.Add("Update README documentation") }
    if ($joined -match "(^|`n)tests/") { $areas.Add("Add or update tests") }
    if ($areas.Count -eq 0) { $areas.Add("Routine maintenance and fixes") }
    return $areas
}

function New-ReleaseNotes([string]$TargetVersion, [string]$ReleaseMode, [string[]]$Files) {
    $areas = Get-AreaText $Files
    $headline = if ($ReleaseMode -eq "fix") {
        "Same-version fix with refreshed installer and zip assets."
    } else {
        "New version release with feature updates, maintenance fixes, and fresh installer assets."
    }

    $bulletLines = $areas | ForEach-Object { "- $_" }
    $fileLines = $Files | ForEach-Object { "- ``$_``" }
    if (-not $fileLines) { $fileLines = @("- ``No source changes detected``") }

    return @"
# v$TargetVersion

$headline

## Highlights
$($bulletLines -join "`r`n")

## Changed Files
$($fileLines -join "`r`n")

## Install
- Download and run ``CompanyQueryToolSetup.exe``.
- Existing installations can be upgraded in place.
"@
}

function Invoke-BuildArtifacts {
    powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "sync_version.ps1")
    & "C:\Users\aschy\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (Join-Path $PSScriptRoot "CompanyQueryToolSetup.iss")
    cmd /c "cd /d $PSScriptRoot && if not exist dist mkdir dist && if exist $PSScriptRoot\dist\Company_Query_Tool_Setup.zip del /f /q $PSScriptRoot\dist\Company_Query_Tool_Setup.zip && tar.exe -a -c -f $PSScriptRoot\dist\Company_Query_Tool_Setup.zip app.py company_query.py findbiz_scraper.py Install.bat install.ps1 pdf_report.py requirements.txt start.bat start_hidden.vbs update_manager.py update_config.json version.txt web_snapshot.py CompanyQueryToolSetup.iss version.iss.inc sync_version.ps1 publish_release.ps1 README.md tests vendor wheelhouse"
}

$currentVersion = Get-CurrentVersion
$targetVersion = if ($Version) { $Version.Trim() } elseif ($Mode -eq "release") { Get-BumpedPatchVersion $currentVersion } else { $currentVersion }

if (-not $targetVersion) {
    throw "Unable to determine target version."
}

if ($targetVersion -ne $currentVersion) {
    Set-CurrentVersion $targetVersion
}

$changedFiles = Get-ChangedFiles

if (-not $SkipBuild) {
    Invoke-BuildArtifacts
}

$notesDir = Join-Path $PSScriptRoot "dist"
New-Item -ItemType Directory -Force -Path $notesDir | Out-Null
$notesPath = Join-Path $notesDir "release_notes_v$targetVersion.md"
$notesContent = New-ReleaseNotes -TargetVersion $targetVersion -ReleaseMode $Mode -Files $changedFiles
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($notesPath, $notesContent, $utf8NoBom)

git add .
$hasStagedChanges = $true
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    $hasStagedChanges = $false
}

if ($hasStagedChanges) {
    $commitMessage = if ($Mode -eq "release") { "release: v$targetVersion" } else { "fix: refresh v$targetVersion package and notes" }
    git commit -m $commitMessage
}

if (-not $SkipPush) {
    git push origin main
}

if (-not $SkipRelease) {
    $tag = "v$targetVersion"
    $exePath = Join-Path $PSScriptRoot "dist\CompanyQueryToolSetup.exe"
    $zipPath = Join-Path $PSScriptRoot "dist\Company_Query_Tool_Setup.zip"

    gh release view $tag --repo $Repo | Out-Null 2>$null
    if ($LASTEXITCODE -eq 0) {
        gh release upload $tag $exePath $zipPath --clobber --repo $Repo
        gh release edit $tag --title $tag --notes-file $notesPath --repo $Repo
    } else {
        gh release create $tag $exePath $zipPath --title $tag --notes-file $notesPath --repo $Repo --latest
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Mode: $Mode"
Write-Host "Version: $targetVersion"
Write-Host "Notes: $notesPath"
