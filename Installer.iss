#define MyAppName "OutlookAI"
#define MyAppPublisher "SixFive7"

[Setup]
AppId={{78AF2871-0CEB-4451-B80D-455552E37C91}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}\Setup
DisableWelcomePage=yes
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
DisableFinishedPage=yes
OutputDir=.
OutputBaseFilename=OutlookAI-v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
Uninstallable=no
PrivilegesRequired=lowest
SetupMutex=OutlookAISetup
CloseApplications=no
RestartApplications=no
CreateAppDir=yes

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "{commoncf}\microsoft shared\VSTO\10.0\VSTOInstaller.exe"; Parameters: "/s /u ""{app}\OutlookAI.vsto"""; Flags: waituntilterminated runhidden; Check: IsVSTOInstalled
Filename: "{commoncf}\microsoft shared\VSTO\10.0\VSTOInstaller.exe"; Parameters: "/s /i ""{app}\OutlookAI.vsto"""; Flags: waituntilterminated

[Code]
function IsVSTOInstalled: Boolean;
begin
  Result := RegKeyExists(HKEY_CURRENT_USER, 'Software\Microsoft\Office\Outlook\Addins\OutlookAI');
end;
