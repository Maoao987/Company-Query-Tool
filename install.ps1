#Requires -Version 5.0
# Company Query Tool - Windows Installer
param(
    [string]$AppDir = $PSScriptRoot,  # Install.bat passes -AppDir "%~dp0"
    [switch]$NoDesktopShortcut,
    [switch]$NoPause
)

trap {
    Write-Host ""
    Write-Host "  [ERR] $_" -ForegroundColor Red
    Write-Host "  Please screenshot this window and contact support." -ForegroundColor Red
    Wait-ForEnter "`nPress Enter to exit"
    exit 1
}

$ErrorActionPreference = "Stop"
$Host.UI.RawUI.WindowTitle = "Company Query Tool - Installer"

# Fixed install target - always pure ASCII, always writable
$InstallDir = Join-Path $env:LOCALAPPDATA "CompanyQueryTool"
$VenvDir    = Join-Path $InstallDir ".venv"
$PyVersion  = "3.12.7"
$PyMinMinor = 10
$PyMaxMinor = 13
$ManagedPyMajor = 3
$ManagedPyMinor = 12

function Step($n,$t,$msg){ Write-Host "`n[$n/$t] $msg" -ForegroundColor Cyan }
function OK($msg)  { Write-Host "  [OK]  $msg" -ForegroundColor Green  }
function Warn($msg){ Write-Host "  [!!]  $msg" -ForegroundColor Yellow }
function Wait-ForEnter($msg) {
    if ($NoPause) { return }
    try { Read-Host $msg | Out-Null } catch {}
}
function Get-NativeArchitecture {
    if ($env:PROCESSOR_ARCHITEW6432) {
        return $env:PROCESSOR_ARCHITEW6432.ToUpperInvariant()
    }
    if ($env:PROCESSOR_ARCHITECTURE) {
        return $env:PROCESSOR_ARCHITECTURE.ToUpperInvariant()
    }
    if ([Environment]::Is64BitOperatingSystem) {
        return "AMD64"
    }
    return "X86"
}
function Get-PythonDownloadInfo {
    $arch = Get-NativeArchitecture
    if ($arch -eq "ARM64") {
        return @{
            Arch           = "ARM64"
            VendorSubdir   = "arm64"
            FileName       = "python-$PyVersion-arm64.exe"
            WheelSubdir    = ""
            Url            = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-arm64.exe"
        }
    }
    if ($arch -eq "X86") {
        return @{
            Arch           = "x86"
            VendorSubdir   = "x86"
            FileName       = "python-$PyVersion.exe"
            WheelSubdir    = ""
            Url            = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion.exe"
        }
    }
    return @{
        Arch           = "x64"
        VendorSubdir   = "x64"
        FileName       = "python-$PyVersion-amd64.exe"
        WheelSubdir    = "py312-win_amd64"
        Url            = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-amd64.exe"
    }
}
function Resolve-RealPythonExe([string]$candidate) {
    try {
        $realExe = (& $candidate -c "import os, sys; print(os.path.realpath(sys.executable))" 2>$null | Select-Object -First 1).Trim()
        if ($realExe -and (Test-Path $realExe)) { return $realExe }
    } catch {}
    return $candidate
}
function Get-PythonVersionInfo([string]$exe) {
    try {
        $ver = & $exe --version 2>&1
        if ($ver -match "Python (\d+)\.(\d+)\.(\d+)") {
            return @{
                Major = [int]$Matches[1]
                Minor = [int]$Matches[2]
                Patch = [int]$Matches[3]
                Text  = $Matches[0]
            }
        }
        if ($ver -match "Python (\d+)\.(\d+)") {
            return @{
                Major = [int]$Matches[1]
                Minor = [int]$Matches[2]
                Patch = 0
                Text  = $Matches[0]
            }
        }
    } catch {}
    return $null
}
function Test-SupportedPython([hashtable]$info) {
    return $info -and $info.Major -eq 3 -and $info.Minor -ge $PyMinMinor -and $info.Minor -le $PyMaxMinor
}
function Get-PythonBitness([string]$exe) {
    try {
        $bits = (& $exe -c "import struct; print(struct.calcsize('P') * 8)" 2>$null | Select-Object -First 1).Trim()
        if ($bits -match "^\d+$") { return [int]$bits }
    } catch {}
    return 0
}
function Get-ManagedPythonPath([hashtable]$downloadInfo) {
    if ($downloadInfo.Arch -eq "ARM64") {
        return Join-Path $env:LOCALAPPDATA "Programs\Python\Python312-arm64\python.exe"
    }
    if ($downloadInfo.Arch -eq "x64") {
        return Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"
    }
    return ""
}
function Test-ManagedPython([string]$exePath, [hashtable]$downloadInfo) {
    if (-not $exePath -or -not (Test-Path $exePath)) {
        return $false
    }
    $info = Get-PythonVersionInfo $exePath
    if (-not $info) {
        return $false
    }
    if ($info.Major -ne $ManagedPyMajor -or $info.Minor -ne $ManagedPyMinor) {
        return $false
    }
    if ($downloadInfo.Arch -eq "x64") {
        return (Get-PythonBitness $exePath) -eq 64
    }
    return $true
}
function Test-PreferredOfflinePython([hashtable]$pythonInfo, [string]$exePath, [hashtable]$downloadInfo, [string]$wheelDir) {
    if (-not $wheelDir -or -not (Test-Path $wheelDir)) {
        return $true
    }
    if ($downloadInfo.Arch -ne "x64") {
        return $true
    }
    if (-not $pythonInfo -or $pythonInfo.Major -ne 3 -or $pythonInfo.Minor -ne 12) {
        return $false
    }
    return (Get-PythonBitness $exePath) -eq 64
}
function Stop-AppProcesses {
    Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Path -and $_.Path.StartsWith($InstallDir, [System.StringComparison]::OrdinalIgnoreCase)
    } | Stop-Process -Force -ErrorAction SilentlyContinue
}
function Ensure-BundledManagedPython([hashtable]$downloadInfo, [string]$bundledInstaller) {
    $managedExe = Get-ManagedPythonPath $downloadInfo
    if (Test-ManagedPython $managedExe $downloadInfo) {
        $info = Get-PythonVersionInfo $managedExe
        OK "Managed runtime ready: $($info.Text) : $managedExe"
        return @{
            Exe  = $managedExe
            Info = $info
        }
    }
    if (-not (Test-Path $bundledInstaller)) {
        Write-Host "  [ERR] Missing bundled Python installer: $bundledInstaller" -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }

    Warn "Tool runtime is missing or incompatible. Reinstalling bundled Python $PyVersion for $($downloadInfo.Arch)..."
    $installer = Join-Path $env:TEMP ("companyquery_python_" + $downloadInfo.FileName)
    Copy-Item $bundledInstaller $installer -Force
    $proc = Start-Process $installer `
        -ArgumentList "/quiet InstallAllUsers=0 PrependPath=0 Include_pip=1 Include_tcltk=0 Include_test=0" `
        -Wait -PassThru
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) {
        Write-Host "  [ERR] Bundled Python install failed (code $($proc.ExitCode))" -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }

    if (-not (Test-ManagedPython $managedExe $downloadInfo)) {
        Write-Host "  [ERR] Bundled Python was not installed correctly: $managedExe" -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }

    $info = Get-PythonVersionInfo $managedExe
    OK "Managed runtime installed: $($info.Text) : $managedExe"
    return @{
        Exe  = $managedExe
        Info = $info
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Company Query Tool - Windows Installer"  -ForegroundColor Cyan
Write-Host "  Install path: $InstallDir"               -ForegroundColor DarkCyan
Write-Host "==========================================" -ForegroundColor Cyan

# ── STEP 1: Find Python 3.10+ (skip WindowsApps stub) ──────────────────
Step 1 5 "Checking Python installation..."

function Find-Python {
    $candidates = @()
    foreach ($cmd in @("py","python","python3")) {
        $c = Get-Command $cmd -ErrorAction SilentlyContinue
        if ($c) { $candidates += $c.Source }
    }
    foreach ($base in @("$env:LOCALAPPDATA\Programs\Python","C:\Python","$env:ProgramFiles\Python")) {
        if (Test-Path $base) {
            Get-ChildItem $base -Filter "python.exe" -Recurse -Depth 2 -ErrorAction SilentlyContinue |
                ForEach-Object { $candidates += $_.FullName }
        }
    }
    foreach ($exe in ($candidates | Select-Object -Unique)) {
        if (-not $exe -or -not (Test-Path $exe)) { continue }
        # Skip WindowsApps shim if it can't actually run
        if ($exe -like "*WindowsApps*") {
            $null = & $exe -c "import sys; sys.exit(0)" 2>&1
            if ($LASTEXITCODE -ne 0) { continue }
        }
        $realExe = Resolve-RealPythonExe $exe
        $info = Get-PythonVersionInfo $realExe
        if (Test-SupportedPython $info) {
            return @{
                Exe  = $realExe
                Info = $info
            }
        }
    }
    return $null
}

$pyDownload = Get-PythonDownloadInfo
Write-Host "  Detected native architecture: $($pyDownload.Arch)" -ForegroundColor Gray
$BundledPythonInstaller = Join-Path $AppDir ("vendor\python\" + $pyDownload.VendorSubdir + "\" + $pyDownload.FileName)
$BundledWheelDir = if ($pyDownload.WheelSubdir) { Join-Path $AppDir ("wheelhouse\" + $pyDownload.WheelSubdir) } else { "" }
$UseManagedBundledRuntime = $pyDownload.Arch -eq "x64" -and (Test-Path $BundledPythonInstaller) -and $BundledWheelDir -and (Test-Path $BundledWheelDir)

if ($UseManagedBundledRuntime) {
    Warn "Using bundled managed runtime for maximum compatibility on x64 Windows."
    $Python = Ensure-BundledManagedPython $pyDownload $BundledPythonInstaller
} else {
    $Python = Find-Python
}

if (-not $UseManagedBundledRuntime -and $Python -and -not (Test-PreferredOfflinePython $Python.Info $Python.Exe $pyDownload $BundledWheelDir)) {
    Warn "Found $($Python.Info.Text), but bundled offline packages require Python 3.12 x64. Installing bundled Python instead..."
    $Python = $null
}

if (-not $UseManagedBundledRuntime -and -not $Python) {
    Warn "Compatible Python not found. Installing Python $PyVersion for $($pyDownload.Arch)..."
    $installer = Join-Path $env:TEMP "python_installer.exe"
    if (Test-Path $BundledPythonInstaller) {
        Write-Host "  Using bundled Python installer: $BundledPythonInstaller" -ForegroundColor Gray
        Copy-Item $BundledPythonInstaller $installer -Force
    } else {
        try {
            Write-Host "  Downloading (~25 MB), please wait..." -ForegroundColor Yellow
            $ProgressPreference = "SilentlyContinue"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $pyDownload.Url -OutFile $installer -UseBasicParsing
            $ProgressPreference = "Continue"
            OK "Download complete"
        } catch {
            Write-Host "  [ERR] Download failed: $_" -ForegroundColor Red
            Write-Host "  Visit https://www.python.org/downloads/ and install Python manually." -ForegroundColor Yellow
            Start-Process "https://www.python.org/downloads/"
            Wait-ForEnter "`nPress Enter to exit"; exit 1
        }
    }
    Write-Host "  Installing Python (~1 min)..." -ForegroundColor Yellow
    $proc = Start-Process $installer `
        -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=0 Include_test=0" `
        -Wait -PassThru
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) {
        Write-Host "  [ERR] Python install failed (code $($proc.ExitCode))" -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }
    $Python = Find-Python
    if (-not $Python) {
        foreach ($guess in @(
            "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python312-arm64\python.exe"
        )) {
            if (Test-Path $guess) {
                $info = Get-PythonVersionInfo $guess
                if (Test-SupportedPython $info) {
                    $Python = @{ Exe = $guess; Info = $info }
                    break
                }
            }
        }
    }
    if (-not $Python) {
        Write-Host "  [ERR] Cannot find Python after install. Close and re-run Install.bat." -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }
    OK "Python installed: $($Python.Info.Text) : $($Python.Exe)"
} elseif (-not $UseManagedBundledRuntime) {
    OK "Found $($Python.Info.Text) : $($Python.Exe)"
}

$PythonExe = $Python.Exe

# ── STEP 2: Copy app files to InstallDir ───────────────────────────────
Step 2 5 "Copying app files to install directory..."

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

$appFiles = @("app.py","company_query.py","findbiz_scraper.py","pdf_report.py","requirements.txt","web_snapshot.py","update_manager.py","update_config.json","version.txt")
$missing  = 0

# Determine where the source files are
# $AppDir is passed by Install.bat as the folder where Install.bat lives (the extracted ZIP folder)
# When running directly, $AppDir = $PSScriptRoot (the folder this ps1 lives in)
$srcCandidates = @($AppDir)
# Also try the location of the original (non-temp) script as fallback
$srcCandidates += $PSScriptRoot

foreach ($f in $appFiles) {
    $copied = $false
    foreach ($src in ($srcCandidates | Select-Object -Unique)) {
        $srcFile = Join-Path $src $f
        if (Test-Path $srcFile) {
            Copy-Item $srcFile $InstallDir -Force
            Write-Host "  copied: $f" -ForegroundColor Gray
            $copied = $true; break
        }
    }
    if (-not $copied) {
        Write-Host "  [ERR] Cannot find: $f" -ForegroundColor Red
        $missing++
    }
}

if ($missing -gt 0) {
    Write-Host ""
    Write-Host "  [ERR] $missing file(s) not found. Make sure all files are extracted from the ZIP." -ForegroundColor Red
    Wait-ForEnter "Press Enter to exit"; exit 1
}
OK "All app files copied to: $InstallDir"

# ── STEP 3: Create virtual environment ─────────────────────────────────
Step 3 5 "Creating virtual environment (.venv)..."

# Rebuild the tool's private environment on every install so mismatched
# dependencies never leak across upgrades or machine changes.
Stop-AppProcesses
if (Test-Path $VenvDir) {
    Write-Host "  Removing old venv..." -ForegroundColor Gray
    Remove-Item $VenvDir -Recurse -Force
}
& $PythonExe -m venv $VenvDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [ERR] venv creation failed" -ForegroundColor Red
    Wait-ForEnter "Press Enter to exit"; exit 1
}
OK "Virtual environment created: $VenvDir"

$VenvPip = Join-Path $VenvDir "Scripts\pip.exe"
$VenvPy  = Join-Path $VenvDir "Scripts\python.exe"

# ── STEP 4: Install packages ────────────────────────────────────────────
Step 4 5 "Installing packages (~2-5 min first time, please wait)..."

$env:PIP_DISABLE_PIP_VERSION_CHECK = "1"

$reqFile = Join-Path $InstallDir "requirements.txt"
Write-Host "  Installing from: $reqFile" -ForegroundColor Gray

$installArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--retries", "3")
$usedBundledWheelhouse = $false

if ($BundledWheelDir -and (Test-Path $BundledWheelDir)) {
    $wheelCount = (Get-ChildItem $BundledWheelDir -File -ErrorAction SilentlyContinue | Measure-Object).Count
    if ($wheelCount -gt 0) {
        Write-Host "  Using bundled wheelhouse: $BundledWheelDir" -ForegroundColor Gray
        $offlineArgs = $installArgs + @("--no-index", "--find-links", $BundledWheelDir, "-r", $reqFile)
        $installProc = Start-Process $VenvPy `
            -ArgumentList $offlineArgs `
            -Wait -PassThru -NoNewWindow
        if ($installProc.ExitCode -eq 0) {
            $usedBundledWheelhouse = $true
        } else {
            if ($UseManagedBundledRuntime) {
                Write-Host "  [ERR] Bundled wheelhouse install failed in managed runtime mode." -ForegroundColor Red
                Write-Host "  This setup package may be incomplete or corrupted. Please re-download the installer package." -ForegroundColor Yellow
                Wait-ForEnter "Press Enter to exit"; exit 1
            }
            Warn "Bundled wheelhouse install failed; retrying online."
        }
    }
}

if (-not $usedBundledWheelhouse) {
    $onlineArgs = $installArgs + @("-r", $reqFile)
    $installProc = Start-Process $VenvPy `
        -ArgumentList $onlineArgs `
        -Wait -PassThru -NoNewWindow
    if ($installProc.ExitCode -ne 0) {
        Write-Host "  [ERR] Package install failed" -ForegroundColor Red
        Wait-ForEnter "Press Enter to exit"; exit 1
    }
}
OK "All packages installed"

# ── STEP 5: Streamlit config + shortcuts ────────────────────────────────
Step 5 5 "Configuring Streamlit and creating shortcuts..."

$stDir = Join-Path $env:USERPROFILE ".streamlit"
New-Item -ItemType Directory -Force -Path $stDir | Out-Null
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText(
    (Join-Path $stDir "credentials.toml"),
    "[general]`nemail = `"`"`n",
    $utf8NoBom
)

$localSt = Join-Path $InstallDir ".streamlit"
New-Item -ItemType Directory -Force -Path $localSt | Out-Null
[System.IO.File]::WriteAllText(
    (Join-Path $localSt "config.toml"),
    "[browser]`ngatherUsageStats = false`n`n[server]`nheadless = true`n`n[client]`ntoolbarMode = `"minimal`"`n`n[theme]`nbase = `"light`"`nprimaryColor = `"#1f4b84`"`nbackgroundColor = `"#f4f8fc`"`nsecondaryBackgroundColor = `"#ffffff`"`ntextColor = `"#172033`"`n",
    $utf8NoBom
)

$stExe   = Join-Path $VenvDir    "Scripts\streamlit.exe"
$appPy   = Join-Path $InstallDir "app.py"
$startBat = @"
@echo off
setlocal
cd /d "$InstallDir"

if /I not "%~1"=="--foreground" (
    if exist "%~dp0start_hidden.vbs" (
        wscript.exe "%~dp0start_hidden.vbs"
        exit /b
    )
)

if exist ".venv\Scripts\streamlit.exe" (
    start "" http://localhost:8501
    "$stExe" run "$appPy" --server.headless true --browser.gatherUsageStats false
    if /I "%~1"=="--foreground" exit /b %ERRORLEVEL%
    pause
) else (
    echo.
    echo  [!] Please run "Install.bat" first!
    echo      Please double-click "Install.bat" in this folder.
    echo.
    if /I "%~1"=="--foreground" exit /b 1
    pause
)
"@
$startVbs = @"
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = appDir & "\start.bat"
shell.Run "cmd.exe /c """ & batPath & """ --foreground", 0, False
"@

[System.IO.File]::WriteAllText((Join-Path $InstallDir "start.bat"), $startBat, [System.Text.Encoding]::ASCII)
[System.IO.File]::WriteAllText((Join-Path $InstallDir "start_hidden.vbs"), $startVbs, [System.Text.Encoding]::ASCII)

$desktop = [Environment]::GetFolderPath("Desktop")
if (-not $NoDesktopShortcut -and $desktop -and (Test-Path $desktop)) {
    $legacyBat = Join-Path $desktop "CompanyQuery.bat"
    if (Test-Path $legacyBat) {
        Remove-Item $legacyBat -Force -ErrorAction SilentlyContinue
    }

    $shortcutPath = Join-Path $desktop "CompanyQuery.lnk"
    $wsh = New-Object -ComObject WScript.Shell
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = Join-Path $env:WINDIR "System32\wscript.exe"
    $shortcut.Arguments = "`"$(Join-Path $InstallDir "start_hidden.vbs")`""
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.IconLocation = (Join-Path $env:SystemRoot "System32\SHELL32.dll") + ",220"
    $shortcut.Save()
    OK "Desktop shortcut: CompanyQuery.lnk"
} elseif ($NoDesktopShortcut) {
    OK "Desktop shortcut skipped (installer mode)"
} else {
    Warn "Desktop folder not found; start.bat was still created in $InstallDir"
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  Double-click 'CompanyQuery' on Desktop to start." -ForegroundColor White
Write-Host "  The app now runs in the background without a terminal window." -ForegroundColor White
Write-Host "  Browser opens automatically: http://localhost:8501" -ForegroundColor White
Write-Host ""
Wait-ForEnter "Press Enter to close"
