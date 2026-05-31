#define MyAppName "OutlookAI"
#define MyAppPublisher "SixFive7"
#define VstoRuntimeUrl "https://aka.ms/VSTORuntimeDownload"
#define NetFramework48Url "https://go.microsoft.com/fwlink/?LinkId=2085155"

[Setup]
AppId={{78AF2871-0CEB-4451-B80D-455552E37C91}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}\Setup
UninstallDisplayName={#MyAppName}
DisableWelcomePage=yes
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
DisableFinishedPage=yes
OutputDir=.
OutputBaseFilename=OutlookAI-v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
SetupMutex=OutlookAISetup
CloseApplications=yes
RestartApplications=no
CreateAppDir=yes

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Register add-in with Outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\OutlookAI.vsto|vstolocal"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookAI"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Description"; ValueData: "OutlookAI"; Flags: uninsdeletekey

; Prevent Outlook from disabling the add-in (boot + crash + demand = 0xB)
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "11"; Flags: uninsdeletevalue

[UninstallRun]
Filename: "certutil"; Parameters: "-user -delstore TrustedPublisher ""{app}\OutlookAI.cer"""; Flags: runhidden

[Code]
function IsVstoInstalled: Boolean;
var
  version: string;
begin
  Result :=
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Version', version) or
    RegQueryStringValue(HKLM, 'SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R', 'Version', version) or
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4', 'Version', version);
end;

function IsNetFramework48Installed: Boolean;
var
  release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', release)
    and (release >= 528040);
end;

function DownloadFile(const Url, DestPath: string; var ErrorDetail: string): Boolean;
var
  ResultCode: Integer;
  CmdLine: string;
  AnsiContent: AnsiString;
begin
  Result := False;
  ErrorDetail := '';

  CmdLine := 'try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; '
    + 'Invoke-WebRequest -Uri "' + Url + '" -OutFile "' + DestPath + '" -UseBasicParsing; '
    + 'exit 0 } catch { $_.Exception.Message | Out-File "' + DestPath + '.err" -Encoding utf8; exit 1 }';

  if not Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -Command "' + CmdLine + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    ErrorDetail := 'Could not start PowerShell';
    exit;
  end;

  if ResultCode <> 0 then
  begin
    if FileExists(DestPath + '.err') then
      if LoadStringFromFile(DestPath + '.err', AnsiContent) then
        ErrorDetail := Trim(String(AnsiContent));
    if ErrorDetail = '' then
      ErrorDetail := 'Download failed (exit code ' + IntToStr(ResultCode) + ')';
    DeleteFile(DestPath + '.err');
    exit;
  end;

  DeleteFile(DestPath + '.err');

  if not FileExists(DestPath) then
  begin
    ErrorDetail := 'Download completed but file was not created';
    exit;
  end;

  Result := True;
end;

procedure DownloadAndInstallVstoRuntime;
var
  TempPath: string;
  ErrorCode: Integer;
  ErrorDetail: string;
begin
  TempPath := ExpandConstant('{tmp}\vstor_redist.exe');

  WizardForm.StatusLabel.Caption := 'Downloading VSTO Runtime...';
  WizardForm.ProgressGauge.Style := npbstMarquee;

  if not DownloadFile('{#VstoRuntimeUrl}', TempPath, ErrorDetail) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('Could not download the VSTO Runtime:' + #13#10 +
      ErrorDetail + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.StatusLabel.Caption := 'Installing VSTO Runtime...';

  if not ShellExec('runas', TempPath, '/q /norestart', '', SW_HIDE, ewWaitUntilTerminated, ErrorCode) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('The VSTO Runtime is required but could not be installed (error ' + IntToStr(ErrorCode) + ').' + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.ProgressGauge.Style := npbstNormal;

  if not IsVstoInstalled then
  begin
    MsgBox('The VSTO Runtime installation did not complete successfully.' + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
  end;
end;

procedure InstallCertificate;
var
  ResultCode: Integer;
begin
  if not Exec('certutil', ExpandConstant('-f -user -addstore TrustedPublisher "{app}\OutlookAI.cer"'), '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    if not WizardSilent then
      MsgBox('Could not run certutil to install the signing certificate.' + #13#10 +
        'The add-in may not load in Outlook due to a trust error.' + #13#10 + #13#10 +
        'You can fix this manually by running:' + #13#10 +
        'certutil -f -user -addstore TrustedPublisher "' + ExpandConstant('{app}') + '\OutlookAI.cer"', mbError, MB_OK);
    exit;
  end;

  if (ResultCode <> 0) and (not WizardSilent) then
  begin
    MsgBox('Failed to install the signing certificate (certutil exit code ' + IntToStr(ResultCode) + ').' + #13#10 +
      'The add-in may not load in Outlook due to a trust error.' + #13#10 + #13#10 +
      'You can fix this manually by running:' + #13#10 +
      'certutil -f -user -addstore TrustedPublisher "' + ExpandConstant('{app}') + '\OutlookAI.cer"', mbError, MB_OK);
  end;
end;

procedure DownloadAndInstallNetFramework;
var
  TempPath: string;
  ErrorCode: Integer;
  ErrorDetail: string;
begin
  TempPath := ExpandConstant('{tmp}\ndp48-web.exe');

  WizardForm.StatusLabel.Caption := 'Downloading .NET Framework 4.8...';
  WizardForm.ProgressGauge.Style := npbstMarquee;

  if not DownloadFile('{#NetFramework48Url}', TempPath, ErrorDetail) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('Could not download .NET Framework 4.8:' + #13#10 +
      ErrorDetail + #13#10 + #13#10 +
      'Please install it from Windows Update or from:' + #13#10 +
      'https://dotnet.microsoft.com/download/dotnet-framework/net48' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.StatusLabel.Caption := 'Installing .NET Framework 4.8...';

  if not ShellExec('runas', TempPath, '/q /norestart', '', SW_HIDE, ewWaitUntilTerminated, ErrorCode) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('.NET Framework 4.8 is required but could not be installed (error ' + IntToStr(ErrorCode) + ').' + #13#10 + #13#10 +
      'Please install it from Windows Update or from:' + #13#10 +
      'https://dotnet.microsoft.com/download/dotnet-framework/net48' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.ProgressGauge.Style := npbstNormal;

  if not IsNetFramework48Installed then
  begin
    MsgBox('.NET Framework 4.8 installation did not complete successfully.' + #13#10 + #13#10 +
      'A restart may be required. Please restart your computer and then' + #13#10 +
      'run the OutlookAI installer again.', mbError, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not WizardSilent then
    begin
      if not IsNetFramework48Installed then
        DownloadAndInstallNetFramework;

      // VSTO Runtime requires .NET 4.8; skip if it is still
      // missing (e.g. a reboot is needed to finish .NET setup).
      if IsNetFramework48Installed and (not IsVstoInstalled) then
        DownloadAndInstallVstoRuntime;
    end;

    InstallCertificate;
  end;
end;
