#define MyAppName "OutlookAI"
#define MyAppPublisher "SixFive7"
; The VSTO runtime is CARRIED INSIDE this installer rather than fetched at install time -
; see the [Files] entry and InstallVstoRuntime. v3.0.1 fetched it from
; https://aka.ms/VSTORuntimeDownload, and that alias was silently repointed at the Download
; Center *page* (Content-Type text/html): on a clean machine setup saved 127 KB of HTML as
; vstor_redist.exe, executed it, and finished with no VSTO runtime and nothing to say why.
; Pointing at a direct CDN path instead would fix that particular rot but not the class of
; it - any URL can move, and this is the one prerequisite without which the add-in cannot
; load at all. A payload compiled into setup cannot 404.
; Where the failure messages send the user - a page to read, not a 40 MB direct download.
#define VstoRuntimeManualUrl "https://www.microsoft.com/download/details.aspx?id=105890"
; The MCP server is a framework-dependent net10.0-windows console app. Its runtimeconfig
; asks for Microsoft.NETCore.App 10.0.0 only - it references no WinForms/WPF, so the BASE
; .NET runtime is enough and the (much larger) Desktop runtime is NOT required.
; Default roll-forward is Minor, i.e. any 10.x satisfies it but 11.x would not.
; Still an aka.ms alias, and so exposed in principle to the same rot that broke the VSTO one -
; but it is left as-is deliberately: it currently redirects straight to the versioned build
; (builds.dotnet.microsoft.com/.../dotnet-runtime-10.0.10-win-x64.exe, 200, octet-stream), and
; the whole point of this alias is that it tracks the newest 10.x patch without an edit here.
; The MZ check in DownloadFile is what makes that safe - if it ever starts serving a page,
; setup now says so instead of running it.
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

; The Visual Studio 2010 Tools for Office Runtime redistributable, vstor_redist.exe
; 10.0.60917.00 (41,828,424 bytes, SHA-256
; CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18) - the package that
; registers the SOFTWARE\Microsoft\VSTO Runtime Setup\v4R key IsVstoInstalled probes.
;
; dontcopy, not a DestDir entry: this is a prerequisite installer, not part of the app.
; It is compressed into setup, extracted to {tmp} by ExtractTemporaryFile only on the
; machines that actually need it, and never lands in {app} - so nothing has to clean it up
; afterwards and the installed footprint is unchanged.
;
; NOT in git - it is a 40 MB third-party binary in a public repo. CI fetches it before
; compiling; see .github/workflows/release.yml (real payload, hash-verified) and
; build.yml (placeholder, since that gate only checks that the script compiles).
; Deliberately no skipifsourcedoesntexist: if the fetch step is ever removed or fails, the
; compile must break loudly here rather than quietly produce an installer whose VSTO
; prerequisite is missing - that silent-failure mode is exactly what shipped in v3.0.1.
Source: "Redist\vstor_redist.exe"; Flags: dontcopy

[Registry]
; Where the app is installed, so the add-in can find {app}\McpServer\OutlookAI.McpServer.exe
; without guessing from its own assembly location (which sits under Application Files\<ver>\
; when installed, but directly in bin\Release for a developer build).
;
; uninsdeletekey, not uninsdeletevalue: uninstall must take the WHOLE Software\OutlookAI key,
; including the two subkeys the add-in creates at runtime and that setup itself never writes -
; Tuning (with its Desired\ and Applied\ children) and Mcp. Nothing but OutlookAI's own
; bookkeeping lives under this key, so removing it loses no user data.
; It deliberately does NOT reach the values the add-in writes into Outlook's own hives
; (Office\16.0\Outlook\{Search, Cached Mode, PST} and the Policies mirror of Cached Mode):
; that is the user's Outlook configuration, and uninstalling never reverts it - see
; Services\OutlookTuningService.cs and the README.
; Only the uninstaller acts on this. A /SILENT auto-update re-runs setup and never runs the
; uninstaller, so tuning and registration state survive updates untouched.
Root: HKCU; Subkey: "Software\OutlookAI"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"; Flags: uninsdeletekey

; Register add-in with Outlook.
; `|vstolocal` = load in place, no ClickOnce. The VSTO runtime then resolves the add-in
; assemblies from the folder holding the .vsto - it does NOT look inside the ClickOnce
; `Application Files\<version>\` folder, and it does not undo the `.deploy` rename. So
; {app} must contain OutlookAI.vsto, OutlookAI.dll.manifest AND the un-suffixed assemblies
; side by side; the release workflow's "Flatten VSTO payload" step is what puts them there.
; Without it the add-in fails to load with FileNotFoundException and Outlook sets
; LoadBehavior=2 (the defect shipped in v2.3.3.141 through v3.0.0.319).
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
// Setup is a 32-bit process, so these HKLM reads pass through the WOW64 redirector and
// actually land in SOFTWARE\Wow6432Node\... - which is exactly where the VSTO runtime
// registers on 64-bit Windows, and what Microsoft's own detection guidance says to check.
// Hence no explicit 'Wow6432Node\' probe (it would name the same key a second time; the
// redirector does not redirect a path that already says Wow6432Node) and no HKLM64 probe:
// verified on a machine running 64-bit Outlook with the runtime installed and working, the
// native 64-bit view has no 'VSTO Runtime Setup' key at all, so a 64-bit-view probe could
// never turn a real miss into a hit - it could only ever add a false positive.
function IsVstoInstalled: Boolean;
var
  version: string;
begin
  Result :=
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Version', version) or
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

// An HTTP error page is still a *successful* download: the server answers 200, the bytes land
// at the requested name, and nothing about the result says "this is not an installer". That is
// precisely how v3.0.1 shipped broken - the VSTO alias began redirecting to a Download Center
// landing page, setup wrote 127 KB of HTML to vstor_redist.exe and ran it, and the failure was
// invisible. So prove the bytes are a Windows executable before anything executes them.
// Two cheap tests, both from one open handle: every PE image begins with the 'MZ' signature
// (0x4D 0x5A), and no prerequisite fetched here is remotely small - the smaller of the two is
// the ~30 MB .NET runtime - so a 1 MB floor rejects error pages and truncated transfers while
// staying far below any real payload. Failures are swallowed into a False result rather than
// raised: a file that cannot even be opened is not one to execute, and an escaping exception
// would abort the whole install instead of falling through to the "could not download" message.
function IsWindowsExecutable(const Path: string): Boolean;
var
  Stream: TFileStream;
  Header: AnsiString;
begin
  Result := False;

  try
    Stream := TFileStream.Create(Path, fmOpenRead or fmShareDenyNone);
    try
      // Read fills Header itself and returns how many bytes it actually got, so the two
      // indexes below can never run past the end of a file too short to hold a signature.
      if Stream.Size >= 1000000 then
        if Stream.Read(Header, 2) = 2 then
          Result := (Ord(Header[1]) = $4D) and (Ord(Header[2]) = $5A);
    finally
      Stream.Free;
    end;
  except
    Result := False;
  end;
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

  // Guards every prerequisite, not just the one that was caught rotting: this is the single
  // point both DownloadAndInstall* procedures go through. Delete the file so no later step can
  // execute it, then fail exactly like any other failed download - each caller already tells
  // the user which prerequisite it was and where to install it by hand.
  if not IsWindowsExecutable(DestPath) then
  begin
    DeleteFile(DestPath);
    ErrorDetail := 'The download did not return an installer - the server sent a web page or '
      + 'an incomplete file. The download link has probably changed.';
    exit;
  end;

  Result := True;
end;

// No download, and so no DownloadFile / IsWindowsExecutable guard: the bytes are the ones
// compiled into this installer, verified by hash when CI fetched them, and covered by
// setup's own CRC check on extraction. There is no network step left here to rot or to
// hand back a web page.
procedure InstallVstoRuntime;
var
  TempPath: string;
  ErrorCode: Integer;
  ErrorDetail: string;
  Extracted: Boolean;
begin
  TempPath := ExpandConstant('{tmp}\vstor_redist.exe');

  WizardForm.StatusLabel.Caption := 'Preparing VSTO Runtime...';
  WizardForm.ProgressGauge.Style := npbstMarquee;

  // Unpacks the payload from setup into {tmp}; Inno removes {tmp} on exit, so the 40 MB
  // copy needs no cleanup of ours. It signals failure by raising rather than returning, and
  // an uncaught exception here would abort the entire install - so catch it and report it
  // the way every other prerequisite failure is reported. Only machines that actually lack
  // the runtime ever reach this, which is also why the [Files] entry is listed last: with
  // SolidCompression, extracting it means decompressing everything before it, and putting
  // it first would push that cost onto every install instead of just these.
  Extracted := False;
  try
    ExtractTemporaryFile('vstor_redist.exe');
    Extracted := True;
  except
    ErrorDetail := GetExceptionMessage;
  end;

  if not Extracted then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('Could not unpack the VSTO Runtime installer:' + #13#10 +
      ErrorDetail + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeManualUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.StatusLabel.Caption := 'Installing VSTO Runtime...';

  if not ShellExec('runas', TempPath, '/q /norestart', '', SW_HIDE, ewWaitUntilTerminated, ErrorCode) then
  begin
    WizardForm.ProgressGauge.Style := npbstNormal;
    MsgBox('The VSTO Runtime is required but could not be installed (error ' + IntToStr(ErrorCode) + ').' + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeManualUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
    exit;
  end;

  WizardForm.ProgressGauge.Style := npbstNormal;

  if not IsVstoInstalled then
  begin
    MsgBox('The VSTO Runtime installation did not complete successfully.' + #13#10 + #13#10 +
      'Please install it manually from:' + #13#10 +
      '{#VstoRuntimeManualUrl}' + #13#10 + #13#10 +
      'Then restart the OutlookAI installer.', mbError, MB_OK);
  end;
end;

// TrustedPublisher only - the store Office's Customization Installer consults to decide
// whether an add-in may install without asking.
//
// Deliberately NOT the user's Root store as well. Adding this self-signed certificate there
// would complete its chain and retire Office's "Publisher cannot be verified" prompt, but
// Root is a protected store: certutil hands the import to crypt32, which puts up its own
// modal "Security Warning ... installing a certificate with an unconfirmed thumbprint is a
// security risk" dialog and waits, and -f does not suppress it. Measured on a clean Windows
// 11 VM: still blocked after 30 seconds, certificate never added. So it only trades one
// prompt for a more alarming one - and this runs on silent auto-updates too, where nobody is
// there to answer it, which would strand updates exactly the way v3.0.1 existed to fix.
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
      // .NET Framework 4.8 ships in-box with Windows 10 1903 and with every Windows 11, so
      // on any machine this add-in supports it is already there and this check simply passes.
      // Checked anyway rather than assumed: the add-in targets net48 and cannot load without
      // it, so on a machine where it really is absent, saying so plainly beats installing an
      // add-in that would then silently never load. No download here - the branch that used
      // to fetch and install 4.8 could not have run on any supported OS, and servicing an
      // in-box Windows component is Windows Update's job, not setup's.
      if not IsNetFramework48Installed then
        MsgBox('OutlookAI needs .NET Framework 4.8, which is not installed on this computer.' + #13#10 + #13#10 +
          'It is part of Windows 10 version 1903 and later, and of every Windows 11, so this' + #13#10 +
          'normally means Windows is older than OutlookAI supports.' + #13#10 + #13#10 +
          'Install .NET Framework 4.8 from Windows Update or from' + #13#10 +
          'https://dotnet.microsoft.com/download/dotnet-framework/net48,' + #13#10 +
          'then run the OutlookAI installer again.', mbError, MB_OK);

      // VSTO Runtime requires .NET 4.8; skip if it is missing, since the runtime
      // would refuse to install anyway.
      if IsNetFramework48Installed and (not IsVstoInstalled) then
        InstallVstoRuntime;

      // Needed only by the MCP server, never by the add-in itself - so a failure here is
      // reported and then tolerated; setup still completes and Outlook still gets the
      // add-in. Interactive runs only, like the prerequisite checks above: a silent
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
