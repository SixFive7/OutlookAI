#define MyAppName "OutlookAI"
#define MyAppPublisher "SixFive7"
#define VstoRuntimeUrl "https://aka.ms/VSTORuntimeDownload"
#define NetFramework48Url "https://go.microsoft.com/fwlink/?LinkId=2085155"
; The MCP server is a framework-dependent net10.0-windows console app. Its runtimeconfig
; asks for Microsoft.NETCore.App 10.0.0 only - it references no WinForms/WPF, so the BASE
; .NET runtime is enough and the (much larger) Desktop runtime is NOT required.
; Default roll-forward is Minor, i.e. any 10.x satisfies it but 11.x would not.
#define NetRuntime10Url "https://aka.ms/dotnet/10.0/dotnet-runtime-win-x64.exe"
#define NetRuntime10ManualUrl "https://dotnet.microsoft.com/download/dotnet/10.0"

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
; One recursive rule covers the whole payload: the VSTO publish output at the top level
; AND the MCP server, which the release workflow publishes into publish\McpServer\ so it
; lands at {app}\McpServer\. Deliberately NOT a second explicit entry - an entry naming
; publish\McpServer\* would make ISCC fail in the compile-only installer-validation gate
; (build.yml), which only creates a placeholder file in publish\.
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Where the app is installed, so the add-in can find {app}\McpServer\OutlookAI.McpServer.exe
; without guessing from its own assembly location (which sits under Application Files\<ver>\
; when installed, but directly in bin\Release for a developer build).
Root: HKCU; Subkey: "Software\OutlookAI"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"; Flags: uninsdeletevalue

; Register add-in with Outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\OutlookAI.vsto|vstolocal"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookAI"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Description"; ValueData: "OutlookAI"; Flags: uninsdeletekey

; Keep Outlook from auto-disabling the add-in after a slow start or a crash. What matters
; for the exemption is that the value NAME (the add-in id) is present in this list; the DWORD
; data is only a flag. Cover Outlook 2013 / 2016+ / future to match versions checked elsewhere.
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\17.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue

[UninstallRun]
Filename: "certutil"; Parameters: "-user -delstore TrustedPublisher OutlookAI"; Flags: runhidden
; Belt and braces: also remove the superseded certificate explicitly by thumbprint, in
; case the subject match above does not catch it.
Filename: "certutil"; Parameters: "-user -delstore TrustedPublisher E205060633DD7062D4F90033130542948A69D068"; Flags: runhidden

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

// Where the shared frameworks live. The x64 sharedhost key carries the dotnet root on
// machines that have any .NET 5+ installed; fall back to the default location. Note the
// sibling 'sharedfx' key is NOT reliable - it is absent on machines that do have the
// runtime (verified on the development machine, which has 10.0.10 installed), which is
// why detection probes the filesystem rather than the registry.
// HKLM64 / {commonpf64} deliberately: this setup is a 32-bit process, so a plain HKLM
// read would be redirected into Wow6432Node and a plain {commonpf} would point at the
// 32-bit Program Files - neither of which is where the x64 runtime lives. Both constants
// exist only on 64-bit Windows, hence the IsWin64 guard (a 32-bit Windows cannot run this
// x64 server at all, so reporting "not installed" there is correct).
function DotnetSharedFrameworkRoot: string;
var
  base: string;
begin
  if not IsWin64 then
  begin
    Result := '';
    exit;
  end;

  if RegQueryStringValue(HKLM64, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost', 'Path', base) and (base <> '') then
    Result := AddBackslash(base) + 'shared\Microsoft.NETCore.App'
  else
    Result := ExpandConstant('{commonpf64}\dotnet\shared\Microsoft.NETCore.App');
end;

// True when some Microsoft.NETCore.App 10.x is present. Exactly 10.x on purpose: the
// server's default roll-forward policy (Minor) will use any newer 10.x but will NOT
// accept a future 11.x, so accepting ">= 10" here would report a satisfied prerequisite
// on a machine where the server cannot start.
function IsNetRuntime10Installed: Boolean;
var
  FindRec: TFindRec;
  root: string;
begin
  Result := False;
  root := DotnetSharedFrameworkRoot;
  if (root = '') or (not DirExists(root)) then
    exit;

  if FindFirst(AddBackslash(root) + '*', FindRec) then
  begin
    try
      repeat
        if ((FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0)
          and (Copy(FindRec.Name, 1, 3) = '10.') then
        begin
          Result := True;
          exit;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

// The MCP server is spawned once per agent session, so several copies of
// {app}\McpServer\OutlookAI.McpServer.exe are typically running and holding their own
// image file open when an update lands. CloseApplications=yes alone does not reliably
// deal with them (Restart Manager has no window to ask), so stop them explicitly, BEFORE
// any file is replaced. Matching is by executable path under {app} - never by image name
// alone, so a developer build running from a source tree is left alone. Sessions whose
// server is stopped this way simply spawn a fresh one on their next mail call; nothing is
// persisted in the server process, so no state is lost. Best effort throughout: a machine
// where this fails still gets the normal in-use handling from CloseApplications.
procedure StopRunningMcpServers;
var
  ResultCode: Integer;
  CmdLine: string;
  AppDir: string;
begin
  AppDir := ExpandConstant('{app}');
  if AppDir = '' then
    exit;

  CmdLine := 'try { Get-CimInstance Win32_Process -Filter ''Name=''''OutlookAI.McpServer.exe'''''' '
    + '| Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith(''' + AppDir + ''', '
    + '[StringComparison]::OrdinalIgnoreCase) } '
    + '| ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } } catch { }';

  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -Command "' + CmdLine + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
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

// An earlier OutlookAI signing certificate had its private key exposed publicly, so it
// must no longer be trusted. Delete it by THUMBPRINT, never by subject: it shares
// CN=OutlookAI with the current certificate, so a subject match would also remove the
// one just installed above. Best effort - on most machines it was never present, and a
// "not found" result is the normal case, so all failures are ignored silently.
procedure RemoveCompromisedCertificate;
var
  ResultCode: Integer;
  Executed: Boolean;
begin
  Executed := Exec('certutil', '-user -delstore TrustedPublisher E205060633DD7062D4F90033130542948A69D068', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
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

procedure DownloadAndInstallNetRuntime10;
var
  TempPath: string;
  ErrorCode: Integer;
  ErrorDetail: string;
begin
  TempPath := ExpandConstant('{tmp}\dotnet-runtime-10-win-x64.exe');

  WizardForm.StatusLabel.Caption := 'Downloading .NET 10 Runtime...';
  WizardForm.ProgressGauge.Style := npbstMarquee;

  if not DownloadFile('{#NetRuntime10Url}', TempPath, ErrorDetail) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('Could not download the .NET 10 Runtime:' + #13#10 +
      ErrorDetail + #13#10 + #13#10 +
      'The Outlook add-in will still work. The mail server that lets AI agents' + #13#10 +
      'search and read your mail needs it, and will not start without it.' + #13#10 + #13#10 +
      'You can install it later from:' + #13#10 +
      '{#NetRuntime10ManualUrl}', mbError, MB_OK);
    exit;
  end;

  WizardForm.StatusLabel.Caption := 'Installing .NET 10 Runtime...';

  if not ShellExec('runas', TempPath, '/quiet /norestart', '', SW_HIDE, ewWaitUntilTerminated, ErrorCode) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('The .NET 10 Runtime could not be installed (error ' + IntToStr(ErrorCode) + ').' + #13#10 + #13#10 +
      'The Outlook add-in will still work. The mail server that lets AI agents' + #13#10 +
      'search and read your mail needs it, and will not start without it.' + #13#10 + #13#10 +
      'You can install it later from:' + #13#10 +
      '{#NetRuntime10ManualUrl}', mbError, MB_OK);
    exit;
  end;

  WizardForm.ProgressGauge.Style := npbstNormal;

  if not IsNetRuntime10Installed then
  begin
    MsgBox('The .NET 10 Runtime installation did not complete successfully.' + #13#10 + #13#10 +
      'The Outlook add-in will still work. The mail server that lets AI agents' + #13#10 +
      'search and read your mail needs it, and will not start without it.' + #13#10 + #13#10 +
      'You can install it later from:' + #13#10 +
      '{#NetRuntime10ManualUrl}', mbError, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  // Before any file is written, not after: by ssPostInstall the copy has already been
  // attempted and would have failed on the in-use server executable.
  if CurStep = ssInstall then
    StopRunningMcpServers;

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

      // Needed only by the MCP server, never by the add-in itself - so a failure here is
      // reported and then tolerated; setup still completes and Outlook still gets the
      // add-in. Interactive runs only, like the two prerequisites above: a silent
      // auto-update runs unattended after Outlook closes, and the elevation prompt this
      // needs would sit there unanswered, blocking the update. On that path the add-in
      // detects the missing runtime instead and says so in OutlookAI Settings.
      if not IsNetRuntime10Installed then
        DownloadAndInstallNetRuntime10;
    end;

    InstallCertificate;
    // Runs on every install and every silent auto-update, so machines that still trust
    // the exposed certificate are cleaned up as they update.
    RemoveCompromisedCertificate;
  end;
end;
