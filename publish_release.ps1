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
    if ($joined -match "(^|`n)app\.py($|`n)") { $areas.Add("介面排版、按鈕和操作流程又被我順手梳了一遍") }
    if ($joined -match "(^|`n)company_query\.py($|`n)") { $areas.Add("股票代號查詢現在支援特別股代號格式（如 00981A），純數字代號查詢不受影響") }
    if ($joined -match "(^|`n)pdf_report\.py($|`n)") { $areas.Add("PDF 報告又更像正式文件，不像半夜趕出來的草稿") }
    if ($joined -match "(^|`n)findbiz_scraper\.py($|`n)") { $areas.Add("公司登記資料抓取穩定度再往上拉") }
    if ($joined -match "(^|`n)web_snapshot\.py($|`n)") { $areas.Add("來源快照流程再補強，列印留存更安心") }
    if ($joined -match "(^|`n)(install\.ps1|CompanyQueryToolSetup\.iss|start\.bat|start_hidden\.vbs)($|`n)") { $areas.Add("安裝和啟動流程多修幾刀，讓它更像成品而不是挑機器運氣") }
    if ($joined -match "(^|`n)(update_manager\.py|update_config\.json)($|`n)") { $areas.Add("更新器變聰明了，該跳提示的時候不再裝神秘") }
    if ($joined -match "(^|`n)(version\.txt|version\.iss\.inc|sync_version\.ps1)($|`n)") { $areas.Add("版本號終於站好隊，程式內外不再各講各話") }
    if ($areas.Count -eq 0) { $areas.Add("這次主要是例行維護，把工具再打磨得順一點") }
    return $areas
}

function New-ReleaseNotes([string]$TargetVersion, [string]$ReleaseMode, [string[]]$Files) {
    $areas = Get-AreaText $Files
    $headline = if ($ReleaseMode -eq "fix") {
        "這次不是大改版，是把該補的洞補好，順便把小妖精趕出去。"
    } else {
        "新版本報到，功能有進步，Bug 有減肥，工具心情也比較穩定。"
    }

    $bulletLines = $areas | ForEach-Object { "- $_" }
    $fileLines = $Files | ForEach-Object { "- ``$_``" }
    if (-not $fileLines) { $fileLines = @("- ``（本次沒有偵測到未提交檔案）``") }

    return @"
# v$TargetVersion

$headline

## 這次更新了什麼
$($bulletLines -join "`r`n")

## 這次動到的檔案
$($fileLines -join "`r`n")

## 溫馨提醒
- 如果你看到的是同版號修復版，代表這次是原地補強，不是另外開新版本。
- 如果你看到的是新版本號，那就是正式升級，可以放心更新。
- 如果更新後還能挑出問題，歡迎繼續丟給我，表示這工具還有進化空間。
"@
}

function Invoke-BuildArtifacts {
    powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "sync_version.ps1")
    & "C:\Users\aschy\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (Join-Path $PSScriptRoot "CompanyQueryToolSetup.iss")
    cmd /c "cd /d $PSScriptRoot && if not exist dist mkdir dist && if exist $PSScriptRoot\dist\Company_Query_Tool_Setup.zip del /f /q $PSScriptRoot\dist\Company_Query_Tool_Setup.zip && tar.exe -a -c -f $PSScriptRoot\dist\Company_Query_Tool_Setup.zip app.py company_query.py findbiz_scraper.py Install.bat install.ps1 pdf_report.py requirements.txt start.bat start_hidden.vbs update_manager.py update_config.json version.txt web_snapshot.py CompanyQueryToolSetup.iss version.iss.inc sync_version.ps1 vendor wheelhouse"
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
