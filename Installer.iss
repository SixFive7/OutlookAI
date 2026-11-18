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
; The MCP server is a framework-dependent net11.0-windows console app, and it runs on a PRIVATE
; copy of Microsoft's .NET 11 runtime (Q172, decided by the maintainer 2026-11-18). Its
; runtimeconfig asks for Microsoft.NETCore.App 11.0.0 only - the BASE runtime, not the Desktop one
; - and its apphost looks first in ..\..\dotnet, which is {localappdata}\OutlookAI\dotnet, and then
; at a machine-wide .NET 11 (McpServer\Directory.Build.props). SetUpServerRuntime fills that
; folder on every install and every silent update, per user and with no administrator rights.
; The runtime is never inside this installer: the updater refuses a download over 50 MB, and
; Microsoft's package is nearly 40 MB on its own.
; Where a failure sends the user to install it by hand: the same page as
; McpRegistrationService.DotnetRuntimeDownloadUrl, one from setup and one from OutlookAI
; Settings. Compared by Tools/Checks/check-pinned-constants.ps1.
#define ServerRuntimeManualUrl "https://dotnet.microsoft.com/download/dotnet/11.0"
; THE COMPROMISED SIGNING CERTIFICATE (decision 193 (b)): the one that signed v1.0.0.1 to
; v2.3.1.111, whose private key became public. Named by its SHA-1 thumbprint and by nothing
; else, and removed by thumbprint from these trust stores - the user's always, the computer's
; when setup runs with administrator rights - by RemoveCompromisedCertificate, on every install,
; update and uninstall. T1 InstallerCompromisedCertificateTests pins all three.
#define CompromisedCertThumbprint "E205060633DD7062D4F90033130542948A69D068"
#define CompromisedCertUserStores "TrustedPublisher TrustedPeople CA Root"
#define CompromisedCertMachineStores "TrustedPublisher TrustedPeople CA AuthRoot Root"
; ITS PUBLIC KEY, for the VSTO trust entries that name the key and not the certificate: the RSA
; modulus, 1024 bits, as base64 in exactly the form .NET writes it in the <Modulus> of the
; <RSAKeyValue> the VSTO runtime stores. A public key names nobody. Read 2026-11-17 from the
; public half of the certificate above, which is not in this repository. Removed by
; RemoveCompromisedTrustEntries on every install, update and uninstall; T1
; InstallerCompromisedTrustEntryTests pins it by its SHA-256.
#define CompromisedKeyModulus "uG6emg3rYVFe+OpNdO5oBsRtIb5MLBiCOKamIZ0V3QntBawNEC7948U8KeGAeiuUZwZ4pqGCRh4vIZ/Yv0GIz8DBjvuleWcckI+MSL1LZheVyBcBPm7UOmfd/i2aYAKqxdvtQrODDb81000zefQg4owVdYaweuT/uOvgryhnXjk="

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
; THE SAME NAME IS IN ThisAddIn.cs AS InstallerMutexName. The add-in opens this mutex at
; startup and, if it exists, skips ALL startup work - which is what stops it re-triggering the
; updater and spinning up processes that this installer then tears down mid-flight to swap the
; add-in files. Rename one side and the guard evaporates with no error at all, so the two are
; compared mechanically by Tools/Checks/check-pinned-constants.ps1 rather than by memory.
SetupMutex=OutlookAISetup
CloseApplications=yes
RestartApplications=no
CreateAppDir=yes

[Files]
; One recursive rule covers the whole payload: the VSTO publish output at the top level
; AND the MCP server, which the release build (Tools/Publish-Release.ps1, through
; Testbed/host/Publish-AddInPayload.ps1) publishes into publish\McpServer\ so it lands at
; {app}\McpServer\, and the helper that sets up the server's runtime, in publish\RuntimeSetup\
; and so {app}\RuntimeSetup\. Deliberately NOT more explicit entries - an entry naming
; publish\McpServer\* would make ISCC fail for the testbed build of the same script, whose
; installer carries neither.
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
; NOT in git - it is a 40 MB third-party binary in a public repo. It is staged media
; (Testbed/MEDIA.md), and Testbed/host/Publish-AddInPayload.ps1 - which builds every release
; and every testbed payload - holds it to a pinned SHA-256 and length before compiling.
; Deliberately no skipifsourcedoesntexist: if the staged copy is ever missing, the
; compile must break loudly here rather than quietly produce an installer whose VSTO
; prerequisite is missing - that silent-failure mode is exactly what shipped in v3.0.1.
Source: "Redist\vstor_redist.exe"; Flags: dontcopy

[Registry]
; Where the app is installed, so the add-in can find {app}\McpServer\OutlookAI.McpServer.exe
; without guessing from its own assembly location (which sits under Application Files\<ver>\
; when installed, but directly in bin\Release for a developer build).
;
; uninsdeletevalue, not uninsdeletekey: uninstall removes ONLY setup's own InstallDir value and
; leaves the rest of Software\OutlookAI standing. This reverses what v3.1.0 did, deliberately.
; The key stopped being pure bookkeeping when Prompts arrived: it now holds user-authored
; content - edited prompt text and the custom quick-button set - next to the Tuning (with its
; Desired\ and Applied\ children) and Mcp subkeys the add-in writes at runtime. Deleting a
; user's own prompts on uninstall would be data loss with no export path, so all settings
; survive an uninstall and are still there after a reinstall.
; It deliberately does NOT reach the values the add-in writes into Outlook's own hives
; (Office\16.0\Outlook\{Search, Cached Mode, PST} and the Policies mirror of Cached Mode):
; that is the user's Outlook configuration, and uninstalling never reverts it - see
; Services\OutlookTuningService.cs and the README.
; Only the uninstaller acts on this. A /SILENT auto-update re-runs setup and never runs the
; uninstaller, so every setting survives updates untouched either way.
Root: HKCU; Subkey: "Software\OutlookAI"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"; Flags: uninsdeletevalue

; Register add-in with Outlook.
; `|vstolocal` = load in place, no ClickOnce. The VSTO runtime then resolves the add-in
; assemblies from the folder holding the .vsto - it does NOT look inside the ClickOnce
; `Application Files\<version>\` folder, and it does not undo the `.deploy` rename. So
; {app} must contain OutlookAI.vsto, OutlookAI.dll.manifest AND the un-suffixed assemblies
; side by side; the release build's "Flatten VSTO payload" step is what puts them there.
; Without it the add-in fails to load with FileNotFoundException and Outlook sets
; LoadBehavior=2 (the defect shipped in v2.3.3.141 through v3.0.0.319).
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\OutlookAI.vsto|vstolocal"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookAI"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Description"; ValueData: "OutlookAI"; Flags: uninsdeletekey

; Keep Outlook from auto-disabling the add-in after a slow start or a crash. What matters
; for the exemption is that the value NAME (the add-in id) is present in this list; the DWORD
; data is only a flag. Cover Outlook 2013 / 2016+ / future.
;
; THE SET OF VERSIONS HERE MUST MATCH McpServer\OutlookAI.Core\AddIn\OfficeVersions.cs (OfficeVersions.Supported),
; which is the add-in's one list of the Office majors it knows about - the theme probe and the
; Outlook tuning both work from it. The order does not matter here (all three are written); the
; SET does, and Tools/Checks/check-pinned-constants.ps1 compares the two rather than leaving
; it to the comment that used to claim, without checking, that this "matched versions checked
; elsewhere".
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\17.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"; ValueData: "1"; Flags: uninsdeletevalue

[UninstallDelete]
; The mail server's private .NET runtime (Q172), which SetUpServerRuntime's helper put beside
; {app} in OutlookAI's per-user folder: the runtime, any work folder a run left behind, and the
; helper's log. Nothing else there - the audit log, signature backups and saved attachments are
; the user's, and stay.
Type: filesandordirs; Name: "{localappdata}\OutlookAI\dotnet"
Type: filesandordirs; Name: "{localappdata}\OutlookAI\dotnet.setup-*"
Type: files; Name: "{localappdata}\OutlookAI\dotnet-setup.log"

[UninstallRun]
Filename: "certutil"; Parameters: "-user -delstore TrustedPublisher OutlookAI"; Flags: runhidden
; The compromised certificate is removed by thumbprint from every trust store on uninstall as
; on install, and the VSTO trust entries for its key with it: CurUninstallStepChanged runs
; RemoveCompromisedTrustEntries and RemoveCompromisedCertificate.

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

// The MCP server is spawned once per agent session, so several copies of
// {app}\McpServer\OutlookAI.McpServer.exe are typically running and holding their own
// image file open when an update lands. Each of those may in turn have spawned an
// OutlookAI.ComHost.exe child, which holds ITS image open the same way - so both names
// must be stopped, or an update replaces the parent and silently leaves a stale child
// running the previous build. CloseApplications=yes alone does not reliably
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

  // Both image names, still matched by path under {app} so a developer build running
  // from a source tree is left alone. The COM host is stopped FIRST: stopping the parent
  // first would orphan it briefly, and an orphan still holds the file we are about to
  // replace.
  //
  // -OperationTimeoutSec is not optional here. This Exec is synchronous
  // (ewWaitUntilTerminated), it runs at ssInstall before any file is written, and the
  // updater passes /NOCANCEL - so on a machine whose CIM/WMI service is wedged, an
  // unbounded query hangs setup with no progress bar moving and no way for the user to
  // stop it. Bounded, the worst case is that a stale server keeps running and the normal
  // in-use handling from CloseApplications takes over, which is the same best-effort
  // outcome every other failure here already has.
  CmdLine := 'try { Get-CimInstance Win32_Process -OperationTimeoutSec 30 '
    + '| Where-Object { ($_.Name -eq ''OutlookAI.ComHost.exe'' -or $_.Name -eq ''OutlookAI.McpServer.exe'') '
    + '-and $_.ExecutablePath -and $_.ExecutablePath.StartsWith(''' + AppDir + ''', '
    + '[StringComparison]::OrdinalIgnoreCase) } '
    + '| Sort-Object { if ($_.Name -eq ''OutlookAI.ComHost.exe'') { 0 } else { 1 } } '
    + '| ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } } catch { }';

  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -Command "' + CmdLine + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

// No download: the bytes are the ones compiled into this installer, verified by hash when the
// release was built, and covered by setup's own CRC check on extraction. There is no network
// step here to rot or to hand back a web page.
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

// THE COMPROMISED SIGNING CERTIFICATE - the one that signed v1.0.0.1 to v2.3.1.111, whose
// private key became public - must no longer be trusted anywhere, so it is removed from every
// certificate TRUST store setup can write (decision 193 (b), 2026-11-17). Before that, from the
// user's TrustedPublisher store only.
//
//  * By THUMBPRINT, its one identifier here, never by subject: its subject is not OutlookAI's but
//    the name of the computer and user it was made on, which no file here may carry. No copy of it
//    is shipped, and none may enter the repository.
//  * The user's stores always - this per-user installer can write them. The computer's stores
//    only when setup runs with administrator rights, the one case it can write them. Trust
//    stores only: TrustedPublisher, TrustedPeople, CA (intermediates, which chains are built
//    from), Root, and the computer's AuthRoot. Never added to Disallowed: that is for the
//    maintainer's own computer, where it is done by hand.
//  * The computer's Root and AuthRoot only when the certificate is there, and not waited for.
//  * THE USER'S ROOT STORE ASKS. Windows puts every removal from it to the user: a window titled
//    "Root Certificate Store" - "Do you want to DELETE the following certificate from the Root
//    Store?", the certificate's subject, issuer, validity, serial number and thumbprints, Yes and
//    No, No the default - shown even when certutil runs hidden, and certutil waits for the answer
//    (still waiting after two minutes). Yes removes it (certutil exits 0), No keeps it (0x800704C7,
//    ERROR_CANCELLED). Measured 2026-11-17, Docs/live-tier-on-the-vm.md section 4.8. So that store
//    has its own procedure, RemoveCompromisedCertificateFromUserRoot, below.
//  * Best effort throughout: on most machines it is in none of these stores, and "not found" is
//    the normal answer, so every result is ignored. Install and uninstall run the same removal.
const
  UserRootCertificateKey = 'Software\Microsoft\SystemCertificates\Root\Certificates\{#CompromisedCertThumbprint}';

// THE USER'S ROOT STORE: only when the certificate is in it, and only when somebody is there to
// answer Windows' question.
//
//  * In it means in the user's own Root store: the certificate's key under UserRootCertificateKey,
//    which is where a certificate the user added to Root is kept (measured). The user's view of Root
//    also lists the computer's roots, which this user cannot remove - so the question is never put
//    about one of those, and the normal machine, the certificate in no Root store, starts no
//    certutil here at all.
//  * An interactive install or uninstall says why in one sentence first, then lets Windows ask, and
//    waits for the answer: the user is at the screen, and the setup window waits behind the question.
//  * A silent run - the add-in's own updates are /SILENT - never asks and never waits: it leaves the
//    certificate for the next interactive install or uninstall, which asks again. Windows' question
//    unexplained, after an update nobody watched, would look like an attack, and it would come back
//    after every update while it stood unanswered.
procedure RemoveCompromisedCertificateFromUserRoot;
var
  ResultCode: Integer;
  Silent: Boolean;
begin
  if not RegKeyExists(HKCU, UserRootCertificateKey) then
    exit;
  if IsUninstaller then
    Silent := UninstallSilent
  else
    Silent := WizardSilent;
  if Silent then
  begin
    Log('The compromised certificate is in the user''s Root store. A silent run leaves it for the next interactive install or uninstall, which asks.');
    exit;
  end;
  MsgBox('Windows will now ask, in a window titled Root Certificate Store, whether to delete an old OutlookAI certificate (thumbprint ' + Copy('{#CompromisedCertThumbprint}', 1, 8) + '...) from your trusted root certificates: its private key has leaked, so please answer Yes.', mbInformation, MB_OK);
  if Exec('certutil', '-user -delstore Root {#CompromisedCertThumbprint}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    Log('certutil -user -delstore Root exited ' + IntToStr(ResultCode) + ': 0 removed, -2147023673 (ERROR_CANCELLED) kept.');
end;

procedure DeleteCompromisedCertificate(const Scope, Store: String);
var
  ResultCode: Integer;
  Wait: TExecWait;
begin
  if (Scope = '-user ') and (Store = 'Root') then
  begin
    RemoveCompromisedCertificateFromUserRoot;
    exit;
  end;
  Wait := ewWaitUntilTerminated;
  if (Store = 'Root') or (Store = 'AuthRoot') then
  begin
    // certutil -store exits 0 when it lists the certificate, and non-zero when it is not there.
    if not Exec('certutil', Scope + '-store ' + Store + ' {#CompromisedCertThumbprint}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      exit;
    if ResultCode <> 0 then
      exit;
    Wait := ewNoWait;
  end;
  if not Exec('certutil', Scope + '-delstore ' + Store + ' {#CompromisedCertThumbprint}', '', SW_HIDE, Wait, ResultCode) then
    exit;
end;

// Each store in a space-separated list, in order.
procedure DeleteCompromisedCertificateFromStores(const Scope, Stores: String);
var
  Rest: String;
  Store: String;
  P: Integer;
begin
  Rest := Stores;
  while Rest <> '' do
  begin
    P := Pos(' ', Rest);
    if P = 0 then
    begin
      Store := Rest;
      Rest := '';
    end
    else
    begin
      Store := Copy(Rest, 1, P - 1);
      Rest := Copy(Rest, P + 1, Length(Rest));
    end;
    if Store <> '' then
      DeleteCompromisedCertificate(Scope, Store);
  end;
end;

procedure RemoveCompromisedCertificate;
begin
  DeleteCompromisedCertificateFromStores('-user ', '{#CompromisedCertUserStores}');
  if IsAdmin then
    DeleteCompromisedCertificateFromStores('', '{#CompromisedCertMachineStores}');
end;

// VSTO'S TRUST ENTRIES FOR THE COMPROMISED KEY. Clicking Install on Office's "Publisher cannot be
// verified" prompt for v1.0.0.1 to v2.3.1.111 left an entry under VstoInclusionKey\<guid>: Url, the
// location of the .vsto, and PublicKey, the key that signed it, as <RSAKeyValue>. From then on the
// VSTO runtime loads ANY add-in at that location whose manifests are signed with that key, with no
// prompt - it compares the location and the key and never looks at the certificate - and that key's
// private half is public. So every entry holding the compromised key goes, wherever it points, on
// every install, update and uninstall. Docs/live-tier-on-the-vm.md section 4.8 has the evidence.
//
//  * Matched on the public key alone: its modulus, found in the value once the value's white space
//    is out. The runtime reads the value with .NET's FromXmlString and compares the two keys'
//    bytes, so a namespace, a prefix, another element order or a line break inside the base64 is
//    still this key to it - and to this test. Never on the Url, which names a user's folders.
//  * Only this key: the entry the current release's own prompt wrote for the same folder, and every
//    other add-in's, stay. Removing an entry costs at most one more prompt; keeping this one lets
//    anything signed with the leaked key load from that folder unasked.
//  * Best effort: no key, no entry or no PublicKey is the normal case and changes nothing.
const
  VstoInclusionKey = 'Software\Microsoft\VSTO\Security\Inclusion';

// S with every space, tab, carriage return and line feed taken out - the white space .NET's
// base64 decoder skips.
function WithoutWhiteSpace(const S: String): String;
var
  I: Integer;
begin
  Result := '';
  for I := 1 to Length(S) do
    if (S[I] <> ' ') and (S[I] <> #9) and (S[I] <> #10) and (S[I] <> #13) then
      Result := Result + S[I];
end;

procedure RemoveCompromisedTrustEntries;
var
  Names: TArrayOfString;
  PublicKey: String;
  I: Integer;
begin
  if not RegGetSubkeyNames(HKCU, VstoInclusionKey, Names) then
    exit;
  for I := 0 to GetArrayLength(Names) - 1 do
    if RegQueryStringValue(HKCU, VstoInclusionKey + '\' + Names[I], 'PublicKey', PublicKey) then
      if Pos('{#CompromisedKeyModulus}', WithoutWhiteSpace(PublicKey)) > 0 then
        if RegDeleteKeyIncludingSubkeys(HKCU, VstoInclusionKey + '\' + Names[I]) then
          Log('Removed the VSTO trust entry ' + Names[I] + ', which trusted the compromised key.');
end;

// THE MAIL SERVER'S PRIVATE .NET RUNTIME (Q172). Runs on every install AND every silent update:
// the helper copied with the server needs no administrator rights, so unlike the machine-wide
// .NET 10 runtime setup used to install, an update nobody watches gets it too. It does nothing
// when the runtime is already in place; otherwise it downloads Microsoft's runtime package - or
// takes it from beside this installer, when it is there under Microsoft's own name - holds it to
// the pinned length and SHA-512, and unpacks it into {localappdata}\OutlookAI\dotnet
// (McpServer\OutlookAI.RuntimeSetup\PrivateRuntime.cs). Its exit code says how it went; its log,
// {localappdata}\OutlookAI\dotnet-setup.log, says why. A failure is reported on an interactive
// install and logged on a silent one - the add-in works without the mail server, OutlookAI
// Settings says what is missing, and the next install or update tries again.
// The folder goes in as '<src>\.' so that no backslash stands before the closing quote, where the
// command line would read it as an escaped quote.
procedure SetUpServerRuntime;
var
  Helper: string;
  ResultCode: Integer;
  Reason: string;
begin
  Helper := ExpandConstant('{app}\RuntimeSetup\OutlookAI.RuntimeSetup.exe');
  if not FileExists(Helper) then
  begin
    Log('No ' + Helper + ': this installer carries no mail server, so it sets up no runtime for one.');
    exit;
  end;

  if not WizardSilent then
  begin
    WizardForm.StatusLabel.Caption := 'Setting up the mail server''s .NET runtime...';
    WizardForm.ProgressGauge.Style := npbstMarquee;
  end;

  if not Exec(Helper, '--from "' + AddBackslash(ExpandConstant('{src}')) + '."', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    ResultCode := -1;

  if not WizardSilent then
    WizardForm.ProgressGauge.Style := npbstNormal;

  if ResultCode = 0 then
  begin
    Log('The mail server''s .NET runtime is in place.');
    exit;
  end;

  case ResultCode of
    10: Reason := 'It could not be downloaded.';
    11: Reason := 'The download was not the file Microsoft publishes, so it was not used.';
    12: Reason := 'It could not be unpacked.';
    13: Reason := 'It could not be put in place.';
  else
    Reason := 'Setting it up failed (code ' + IntToStr(ResultCode) + ').';
  end;
  Log('The mail server''s .NET runtime is not in place: ' + Reason + ' The helper''s log: '
    + ExpandConstant('{localappdata}\OutlookAI\dotnet-setup.log'));

  if not WizardSilent then
    MsgBox('The .NET runtime the mail server needs could not be set up. ' + Reason + #13#10 + #13#10 +
      'The Outlook add-in will still work. The mail server that lets AI agents' + #13#10 +
      'search and read your mail needs it, and will not start without it.' + #13#10 + #13#10 +
      'Run the installer again to retry, or install the .NET 11 runtime from:' + #13#10 +
      '{#ServerRuntimeManualUrl}', mbError, MB_OK);
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
    end;

    InstallCertificate;
    // Run on every install and every silent auto-update, so machines that still trust the
    // exposed key are cleaned up as they update: its VSTO trust entries, then its certificate.
    RemoveCompromisedTrustEntries;
    RemoveCompromisedCertificate;

    // Needed only by the mail server, never by the add-in itself, so a failure is reported and
    // then tolerated: setup still completes and Outlook still gets the add-in. Last, because it
    // is the one step that may wait on the network.
    SetUpServerRuntime;
  end;
end;

// Uninstalling removes the compromised key's VSTO trust entries and its certificate too: an
// uninstall is no reason to go on trusting it.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveCompromisedTrustEntries;
    RemoveCompromisedCertificate;
  end;
end;
