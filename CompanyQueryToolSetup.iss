#define MyAppName "Company Query Tool"
#include "version.iss.inc"
#define MyAppPublisher "aschy"
#define MyAppExeName "start.bat"
#define MyBundleDir "{localappdata}\CompanyQueryToolBundle"
#define MyRuntimeDir "{localappdata}\CompanyQueryTool"

[Setup]
AppId={{3F6171D7-249A-48D9-AE85-EFEE8DF3E006}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={#MyBundleDir}
DefaultGroupName={#MyAppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x86compatible
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
OutputDir=dist
OutputBaseFilename=CompanyQueryToolSetup
UninstallDisplayIcon={#MyRuntimeDir}\{#MyAppExeName}
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: checkedonce

[Files]
Source: "app.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "company_query.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "findbiz_scraper.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "Install.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "install.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "pdf_report.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "start.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "start_hidden.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "update_manager.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "update_config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "web_snapshot.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "vendor\*"; DestDir: "{app}\vendor"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
Source: "wheelhouse\*"; DestDir: "{app}\wheelhouse"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{#MyRuntimeDir}\start_hidden.vbs"""; WorkingDir: "{#MyRuntimeDir}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{#MyRuntimeDir}\start_hidden.vbs"""; WorkingDir: "{#MyRuntimeDir}"; Tasks: desktopicon

[Run]
Filename: "{sys}\wscript.exe"; Parameters: """{#MyRuntimeDir}\start_hidden.vbs"""; Description: "Launch {#MyAppName}"; WorkingDir: "{#MyRuntimeDir}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{#MyRuntimeDir}"
Type: filesandordirs; Name: "{app}"

[Code]
var
  RuntimeInstallSucceeded: Boolean;

function RuntimeInstalled: Boolean;
begin
  Result := RuntimeInstallSucceeded and FileExists(ExpandConstant('{#MyRuntimeDir}\{#MyAppExeName}'));
end;

function RunRuntimeInstaller: Boolean;
var
  ResultCode: Integer;
  PowerShellPath, Params: String;
begin
  if IsWin64 then
    PowerShellPath := ExpandConstant('{sysnative}\WindowsPowerShell\v1.0\powershell.exe')
  else
    PowerShellPath := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
  Params :=
    '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{app}\install.ps1') + '"' +
    ' -AppDir "' + ExpandConstant('{app}') + '"' +
    ' -NoDesktopShortcut -NoPause';

  Log('Running runtime installer: ' + PowerShellPath + ' ' + Params);
  if not Exec(PowerShellPath, Params, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then begin
    Log('Failed to start runtime installer. Win32 error: ' + SysErrorMessage(ResultCode));
    Result := False;
    exit;
  end;

  Log('Runtime installer exit code: ' + IntToStr(ResultCode));
  Result := ResultCode = 0;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    WizardForm.StatusLabel.Caption :=
      'Preparing the Company Query Tool runtime. This can take a few minutes...';
    WizardForm.Update;
    RuntimeInstallSucceeded := RunRuntimeInstaller;
    if not RuntimeInstallSucceeded then
      RaiseException('Company Query Tool runtime setup failed. Please review the setup log for details.');
  end;
end;
