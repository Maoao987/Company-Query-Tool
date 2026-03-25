$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$versionFile = Join-Path $root "version.txt"
$includeFile = Join-Path $root "version.iss.inc"

if (-not (Test-Path $versionFile)) {
    throw "version.txt not found: $versionFile"
}

$version = (Get-Content -Path $versionFile -Raw -Encoding UTF8).Trim()
if (-not $version) {
    throw "version.txt is empty"
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText(
    $includeFile,
    "#define MyAppVersion `"$version`"`n",
    $utf8NoBom
)

Write-Host "Synced installer version to $version" -ForegroundColor Green
