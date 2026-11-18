# TODO

- [ ] **Take the personal data out of the public history (Q183 to Q196): rewrite the history at a
  quiet point and force-push.**
  The audit of 2026-11-17 and the maintainer's answers are in QUESTIONS.md's decision log; the
  inventory with the values stays in `.work/history-audit-2026-11/` of the maintainer's checkout,
  and no value goes into a committed file, this one included (AGENTS.md, Personal data). Done: the
  purge of the 198 commits left over from the earlier rewrites was asked of GitHub Support (Q183);
  the old signing key is out of his store and on his untrusted list (Q184); Q190's safeguards are in,
  with the identifiers in the private list only (Q196); and HEAD needs no value any more (Q189) -
  the one entry left in `Tools/Checks/personal-data.json`'s baseline is a measured test address
  that came back one character short. Left, in order:
  1. The rewrite (Q185 (b), Q187 (a), Q188 (a), Q191 (a)), only at a quiet point: the in-flight
     branches landed and work frozen; `git filter-repo` on a fresh `--mirror` clone, with the rules in
     `.work/history-rewrite/` of the maintainer's checkout; the audit's scanner run on the result,
     finding nothing, and the commit count unchanged; then the force-push of master and the moved
     tags - as updates of the same tag names, never a delete and a re-create - on his word; then every
     clone and worktree made again. His name and e-mail stay as author and committer (Q186 (a)).
  2. After the force-push: `gh api repos/SixFive7/OutlookAI/releases/latest` still names the newest
     release; a second Support request names the commits the rewrite replaced; the 522 old Actions run
     records are deleted and the leftover "Dependabot Updates" workflow is disabled through the API;
     and one normal commit repoints the commit IDs cited in tracked files, from filter-repo's commit
     map. Locally, the unpushed 3ee7ca8 still sits in the reflog until
     `git reflog expire --expire-unreachable=now --all` and `git gc --prune=now`.

- [ ] **Investigate whether OutlookAI can do anything about the Outlook hang of 2026-11-15: its MCP
  server cold-started `OUTLOOK.EXE -Embedding` through COM, that Outlook deadlocked 5 s into its
  start, and the user reported that Outlook would not start.**
  **Read this first.** This item was written on 2026-11-16 by an agent OUTSIDE this project: the user's
  general Claude Code session that diagnosed and fixed the hang on 2026-11-15. The user asked for it
  in these words: *"The todo should be to investigate if OutlookAI can do something about this issue.
  As you do not have the context the agent in the OutlookAI has you cannot make the correct
  tradeoffs. So restrict yourself to dumping all the knowledge you do have into the todo item whilst
  clearly stating that these are only guesses with not enough context on how everything works."*
  Apart from the source lines quoted in 1.15, that agent read nothing of OutlookAI. The only
  exceptions are TODO.md, AGENTS.md and `git log`, to place this item. It did not read the design,
  the decision log (QUESTIONS.md), the runbook or the tests. So:
  - **PART 1** holds OBSERVED FACTS only, each with its source and time.
  - **PART 2** holds that agent's INFERENCES about the causal chain. They are guesses.
  - **PART 5** holds DIRECTIONS. They are guesses made without knowing how OutlookAI works or why it does
    what it does. None of them is a recommendation; the OutlookAI agent makes the trade-offs.
  **The task:** investigate whether OutlookAI can prevent this, or detect it and recover, or at least
  report it. Decide what, if anything, to change, and record the decision. Nothing was changed in
  OutlookAI by the session that wrote this.
  **Everything needed is in this item, except identifying values.** Times are local CET (UTC+1)
  unless a time ends in Z (UTC). This repository is public, so every identifying value is replaced by
  a placeholder from the legend below:
  - e-mail addresses;
  - the Windows user and computer names;
  - mailbox, store and display names;
  - OST/NST file names;
  - the name of the user's own add-in.
  Those substitutions, plus the removal of blank lines and trailing spaces, are the ONLY changes made
  to the verbatim excerpts. Every technical value (PIDs, TIDs, times, addresses in memory, offsets,
  versions, sizes, store types) is unchanged.
  **The real values** are in `.work/outlook-hang-2026-11-15-identifiers.md`, with other machine
  details that help reproduce this on the user's workstation. That file exists only in the local
  clone on that workstation and is deliberately NOT committed (`.work/` is git-ignored), because it
  holds personal data. If it is missing, ask the user.
  **Pointers.** The evidence folder `C:\Users\<user>\Downloads\tmp-outlook-diag\` and the Claude
  transcripts named in PART 6 are optional pointers only. Downloads is auto-deleted after 31 days, so
  the folder is likely gone after about 2026-12-16.
  **Left out entirely:** the mailbox content the fixing session read afterwards (calendar entries),
  and third parties' names.
  **Placeholders.** The same token means the same value everywhere in this item.
  - `<user>`: the Windows user name, in paths such as `C:\Users\<user>\...` and in Claude's project
    folder names `C--Users-<user>-...`.
  - `<computer>`: the computer name.
  - `<mailbox-A>`: the SMTP address, which is also the account and store display name, of an Exchange
    account added to the profile (exchangeStoreType 4, cached mode, OST 9,433 MB, one NST). The
    healthy instance showed its Inbox at 17:03.
  - `<mailbox-B>`: the same for the profile's primary Exchange mailbox (exchangeStoreType 0, cached
    mode, OST 24,273 MB, two NSTs). The hidden Explorer windows of both COM-started instances (76888
    and 60368) showed its Inbox.
  - `<mailbox-C>`: the same for another added Exchange account (exchangeStoreType 4, cached mode, OST
    16 MB, one NST).
  - `<mailbox-D>`: the SMTP address of a fourth account. A cached store carries the same display name
    (exchangeStoreType 4, onlineOnly true, OST 16 MB, one NST).
  - `<store-D-online>`: the display name of `<mailbox-D>`'s delivery store, which is online-only
    (exchangeStoreType 1, not cached).
  - `<delegate-mailbox-1>` and `<delegate-mailbox-2>`: the display names of two delegate mailboxes
    (isDelegate true, exchangeStoreType 1, not cached, in the local index).
  - `<custom-addin>`: the name of the user's own VSTO add-in, loaded from a Debug build in a local
    repository (LoadBehavior 3).
  - `<mailbox-X>.ost`, `<mailbox-X>.nst` and `<mailbox-B>(2).nst`: Outlook's data files for those
    mailboxes, which Outlook names after the account.
  **Short version (INFERENCE, see PART 2).**
  1. At 16:27:39 an OutlookAI MCP server (PID 68964, serving an unrelated Claude session) ran
     `list_accounts` while no Outlook was running.
  2. It cold-started Outlook through COM, and DcomLaunch started `OUTLOOK.EXE -Embedding` (PID 76888)
     with no window.
  3. About 5 s later Outlook's UI thread deadlocked on itself. An incoming late-bound object-model call
     (most likely the MCP's connect sequence) was opening the Exchange offline store and held an
     OLMAPI32 lock. Under it, Outlook made an outgoing COM call to the Network List Service
     (`netprofm!PublicNetworkListManager::GetCost`).
  4. While that call waited, COM let a second call into the same UI thread: the **Teams Meeting
     Add-in**'s telemetry setup, asking the object model for the current user's SMTP address from a CLR
     thread-pool thread. That call went into EMSMDB32/OLMAPI32 and blocked in a non-pumping wait.
  5. 41 other threads queued on critical sections (31 of them in `OLMAPI32!HrCreateOfflineObj+0x28b`).
     CPU stayed at 0. The instance never showed a window, and Outlook did not start for the user.
  6. The instance stayed window-less and deadlocked from 16:27:44 until the diagnosing agent killed it
     at 17:02:02. Within 67 ms the same MCP server, with no client call pending, cold-started a new
     `OUTLOOK.EXE -Embedding` (PID 60368). That one booted fine and was turned into a normal visible
     Outlook by a shell launch.
  **PART 1 - OBSERVED FACTS (each with its source).**
  **1.1 Machine and software.** Sources: registry, file versions, `Get-AppxPackage`, `.lnk` targets.
  The diagnosing agent read these on 2026-11-15 between 16:47 and 17:00, as user `<computer>\<user>` at
  Medium integrity (BUILTIN\Administrators deny-only, i.e. UAC-filtered, not elevated).
  - Windows 11 Pro 10.0.26300, build 26300, InstallDate 8-11-2024 23:02:25, LastBootUpTime
    13-11-2026 11:14:07 (`Win32_OperatingSystem`).
  - Classic Outlook from Office LTSC 2021 Volume Click-to-Run, x64, 16.0.14334.20974. Raw values
    (`HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`, `...\ClickToRun\Updates`,
    `...\ClickToRun`, the service and the file):
  ```text
  VersionToReport         : 16.0.14334.20974
  ClientVersionToReport   : 16.0.14334.20974
  Platform                : x64
  ProductReleaseIds       : ProPlusSPLA2021Volume,VisioPro2021Volume
  UpdateChannel           : http://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e
  CDNBaseUrl              : http://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e
  AudienceId              : 5030841d-c919-4594-8d2d-84ae4f96e58e
  UpdatesEnabled          : True
  UpdateUrl               :
  UpdateToVersion         :
  UpdateDeadline          :
  InstallationPath        : C:\Program Files\Microsoft Office
  SharedComputerLicensing : 0
  --- Updates key
  UpdatesSkuToSkuBlocked                  : 0
  QueueUpdateParameters                   : displaylevel=False flt.useofficehelperaddon=unknown flt.useoutlookshareaddon=unknown flt.UseTeamsOnInstallConsumer=unknown flt.UseTeamsOnUpdateConsumer=unknown triggersource=SCHEDULEDTASK
  UpdateDetectionLastRunTime              : 13439186984173
  UpdateDeadlineMet                       : 1
  UpdatesThrottleValue                    : 390
  UpdatesAppliedTime                      : 13437050855135
  UpdateDeliveryMechanism                 : 5030841d-c919-4594-8d2d-84ae4f96e58e
  ApplyInitiated                          : 1
  DownloadTime                            : 13437017537987
  UpdateDownloadFailedCorruptionCheck     : 0
  UpdateToVersion                         :
  UpdatesReadyToApply                     : 0
  UpdatesBlockedTime                      : 0
  UpdatePackageApplied                    : 0
  --- ClickToRun root
  LastScenario         : UPDATE
  LastScenarioResult   : Success
  --- service
  Name      : ClickToRunSvc
  Status    : Running
  StartType : Automatic
  --- OUTLOOK.EXE file
  FullName       : C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE
  Length         : 42159352
  LastWriteTime  : 21-10-2026 02:52:08
  FileVersion    : 16.0.14334.20974
  ProductVersion : 16.0.14334.20974
  ```
  - New Outlook (olk.exe) is not in use.
    - `Get-AppxPackage -Name '*Outlook*'` returned only `MicrosoftWindows.NarratorExtension-OutlookWebApp`
      and `MicrosoftWindows.NarratorScript.Outlook`.
    - A folder `C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2023.1101.300_x64__8wekyb3d8bbwe`
      dated 15-11-2023 exists, but the package is not registered for the user.
    - HKCU `Software\Microsoft\Office\16.0\Outlook\Options\General` has `HideNewOutlookToggle : 1`.
    - HKCU `Software\Policies\Microsoft\Office\16.0\Outlook\Preferences` has
      `NewOutlookMigrationUserSetting : 0`.
    - HKCU `...\16.0\Outlook\Preferences` has no `UseNewOutlook` value.
    - No olk.exe was running.
  - Every Outlook shortcut targets classic Outlook, with no arguments:
  ```text
  C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk
     -> C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
  C:\Users\<user>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Outlook.lnk
     -> C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
  C:\Users\<user>\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Outlook.lnk
     -> C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
  ```
  - COM add-ins (registry `Software\Microsoft\Office\Outlook\Addins`, HKCU and HKLM;
    `HKLM\SOFTWARE\WOW6432Node\...` is absent) and Outlook's resiliency keys
    (HKCU `Software\Microsoft\Office\16.0\Outlook\Resiliency`; there is no `DisabledItems` key, and
    `CrashingAddinList` has no values). Verbatim:
  ```text
  --- Resiliency subkeys/values:
  HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList
  HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList
     TeamsAddin.Connect
     OutlookAI
  HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\NotificationReminderAddinData
     OutlookAI
     OutlookAI\dtype
  --- Add-ins (LoadBehavior):
  HKCU: | AccessAddin.DC | LoadBehavior=2 | Microsoft Access Outlook Add-in for Data Collection and Publishing
  HKCU: | ColleagueImport.ColleagueImportAddin | LoadBehavior=3 | Microsoft SharePoint Server Colleague Import Add-in
  HKCU: | <custom-addin> | LoadBehavior=3 | <custom-addin>
  HKCU: | OutlookAI | LoadBehavior=3 | OutlookAI
  HKCU: | TeamsAddin.FastConnect | LoadBehavior=3 | Microsoft Teams Meeting Add-in for Microsoft Office
  HKCU: | TeamViewerMeetingAddIn.AddIn | LoadBehavior=2 |
  HKCU: | UCAddin.LyncAddin.1 | LoadBehavior=0 |
  HKLM: | TeamViewerMeetingAddIn.AddIn | LoadBehavior=3 | TeamViewer Meeting Add-In
  [OutlookAI]
  Description  : OutlookAI
  FriendlyName : OutlookAI
  LoadBehavior : 3
  Manifest     : file:///C:\Users\<user>\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal
  [<custom-addin>]
  Description  : <custom-addin>
  FriendlyName : <custom-addin>
  LoadBehavior : 3
  Manifest     : file:///C:/Users/<user>/source/repos/<custom-addin>/bin/Debug/<custom-addin>.vsto|vstolocal
  [TeamsAddin.FastConnect]
  FriendlyName : Microsoft Teams Meeting Add-in for Microsoft Office
  LoadBehavior : 3
  Description  : Microsoft Teams Meeting Add-in for Microsoft Office
  ```
  - Teams Meeting Add-in **1.26.27102**, at
    `C:\Users\<user>\AppData\Local\Microsoft\TeamsMeetingAdd-in\1.26.27102\x64\`. In that folder,
    `Microsoft.Teams.AddinLoader.dll` and `OneAuth.dll` both have file version 1.0.26271.2.
  - Modules loaded in the hung OUTLOOK.EXE 76888 (`Get-Process -Id 76888 -Module`, 255 modules in all;
    selected lines, verbatim):
  ```text
  ModuleName                                                      FileName                                                                                                                                                             Company               Ver
  ----------                                                      --------                                                                                                                                                             -------               ---
  ColleagueImport.dll                                             C:\Program Files\Microsoft Office\root\Office16\ADDINS\ColleagueImport.dll                                                                                           Microsoft Corporation 16.0.14334.20090
  windhawk.dll                                                    C:\Program Files\Windhawk\Engine\1.7.3_3\64\windhawk.dll                                                                                                             Ramen Software        1.7.3
  MpOav.dll                                                       C:\ProgramData\Microsoft\Windows Defender\Platform\4.18.26100.4-0\MpOav.dll                                                                                          Microsoft Corporation 4.18.26100.4 (79b015c6fabb20b1ccd067f376e490f0bac19809)
  Microsoft.Teams.AddinLoader.dll                                 C:\Users\<user>\AppData\Local\Microsoft\TeamsMeetingAdd-in\1.26.27102\x64\Microsoft.Teams.AddinLoader.dll                                                              Microsoft Corporation 1.0.26271.2
  OneAuth.dll                                                     C:\Users\<user>\AppData\Local\Microsoft\TeamsMeetingAdd-in\1.26.27102\x64\OneAuth.dll                                                                                  Microsoft Corporation 1.0.26271.2
  Microsoft.Office.Tools.v4.0.Framework.ni.dll                    C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O2eb0cc9a#\932f431099a884ac946dcf42ce415a0a\Microsoft.Office.Tools.v4.0.Framework.ni.dll                    Microsoft Corporation 10.0.60917.0
  Microsoft.Office.Tools.Common.Implementation.ni.dll             C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O4a946565#\9c01546322fedf1b769480cadf29f0e9\Microsoft.Office.Tools.Common.Implementation.ni.dll             Microsoft Corporation 10.0.60917.0
  Microsoft.Office.Tools.Outlook.ni.dll                           C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O4b20f4dd#\493e4f7c9563737dd2bc16380ff936bb\Microsoft.Office.Tools.Outlook.ni.dll                           Microsoft Corporation 10.0.60917.0
  Microsoft.Office.Tools.Outlook.Implementation.ni.dll            C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O5263ae01#\f27e75d5270bfbf4e0a86021fa4be25f\Microsoft.Office.Tools.Outlook.Implementation.ni.dll            Microsoft Corporation 10.0.60917.0
  Microsoft.Office.Tools.ni.dll                                   C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O5949707a#\a3d94f264fbc44d27de7769c0020f015\Microsoft.Office.Tools.ni.dll                                   Microsoft Corporation 10.0.60917.0
  Microsoft.Office.Tools.Common.ni.dll                            C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.O854200f9#\b5f2ad97e0c0138139c6f79c8e76f0a8\Microsoft.Office.Tools.Common.ni.dll                            Microsoft Corporation 10.0.60917.0
  Microsoft.VisualStudio.Tools.Applications.ServerDocument.ni.dll C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.V1955d7fd#\2501d77ee42f4bfe540589711efa9cbe\Microsoft.VisualStudio.Tools.Applications.ServerDocument.ni.dll Microsoft Corporation 10.0.60917.0
  Microsoft.VisualStudio.Tools.Office.Runtime.ni.dll              C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.V28a60cc2#\526b8a692e01589c96df11137f26abda\Microsoft.VisualStudio.Tools.Office.Runtime.ni.dll              Microsoft Corporation 10.0.60917.0
  Microsoft.VisualStudio.Tools.Applications.Runtime.ni.dll        C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.V883708cb#\a47ce4784a812cc6c1d466ab3f301142\Microsoft.VisualStudio.Tools.Applications.Runtime.ni.dll        Microsoft Corporation 10.0.60917.0
  Microsoft.VisualStudio.Tools.Applications.Hosting.ni.dll        C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\Microsoft.Vf9a08577#\4b8ca356acfb0e105b10e33e9c5ab3b8\Microsoft.VisualStudio.Tools.Applications.Hosting.ni.dll        Microsoft Corporation 10.0.60917.0
  clr.dll                                                         C:\Windows\Microsoft.NET\Framework64\v4.0.30319\clr.dll                                                                                                              Microsoft Corporation 4.8.10008.0 built by: d5c29ab0c000001
  mscoreei.dll                                                    C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscoreei.dll                                                                                                         Microsoft Corporation 4.8.10001.0 built by: bd501acdc000004
  mscoree.dll                                                     C:\WINDOWS\SYSTEM32\mscoree.dll                                                                                                                                      Microsoft Corporation 10.0.26100.3624 (WinBuild.160101.0800)
  netprofm.dll                                                    C:\WINDOWS\System32\netprofm.dll                                                                                                                                     Microsoft Corporation 10.0.26100.1 (WinBuild.160101.0800)
  npmproxy.dll                                                    C:\WINDOWS\System32\npmproxy.dll                                                                                                                                     Microsoft Corporation 10.0.26100.9711 (WinBuild.160101.0800)
  Windows.Networking.Connectivity.dll                             C:\Windows\System32\Windows.Networking.Connectivity.dll                                                                                                              Microsoft Corporation 10.0.26100.9818 (WinBuild.160101.0800)
  ```
    The list had no module whose name contains `OutlookAI` or `<custom-addin>`, although event 45 in 1.4
    says both of those VSTO add-ins loaded at 16:27:41. Why was not checked.
  - Other code in Outlook: Windhawk 1.7.3 injects `windhawk.dll` (see the list above). Which Windhawk
    mods are active was not checked.
  - The OutlookAI install that ran: MCP server
    `C:\Users\<user>\AppData\Local\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe`
    (`outlook_health.registration`, see 1.13), and the add-in manifest
    `file:///C:\Users\<user>\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal`. Which OutlookAI
    version or source revision that is was NOT checked.
  - Outlook data files at 17:08 (from 1.12): `<mailbox-B>.ost` 24,273 MB, `<mailbox-A>.ost`
    9,433 MB, `<mailbox-C>.ost` 16 MB, `<mailbox-D>.ost` 16 MB, and five `.nst`
    files of 16 MB each.
  **1.2 Timeline.** Each line names its source.
  - 2026-11-13 10:00:22: Outlook boot (event 45). [Application log]
  - 2026-11-13 10:38:01: WER `AppHangB1` and `Application Hang` 1002, "OUTLOOK.EXE version 16.0.14334.20974
    stopped interacting with Windows and was closed". Verbatim in 1.4. [Application log]
  - 2026-11-13 10:38:53 to 10:53:59: Outlook store reconciliations (id 30) and boots at 10:51:36 and
    10:53:53 (id 45). [Application log]
  - 2026-11-13 11:14:07: Windows boot. [Win32_OperatingSystem] At 11:14:21 svchost PID 3240 (netprofm)
    started, and at 11:14:22 OfficeClickToRun PID 5964. [Win32_Process]
  - 2026-11-13 11:32:36 and 11:41:45: Office alert "Outlook couldn't start last time. Safe mode could
    help ... Do you want to start in safe mode?" Verbatim in 1.4. [OAlerts log]
  - 2026-11-13 11:50:56 to 20:19:12: 10 more event-45 entries, 13 that day in total. Event 45 is logged
    when Outlook loads its add-ins, normally once per start. [Application log]
  - 2026-11-13 17:22:08, 17:22:19, 17:46:34 and 23:10:24, and 2026-11-14 18:51:20 and 18:56:58: six
    OutlookAI MCP servers started that were still alive on 2026-11-15 (see 1.3). [Win32_Process]
  - 2026-11-14: no `Outlook`-provider events at all. [Application log]
  - 2026-11-15 16:21:10: MCP server PID 68964 started, parent claude.exe 76924. [Win32_Process]
  - 15:22:35.568Z (16:22:35): Claude session 579e5f93 starts (an unrelated research task).
    [transcript, PART 6]
  - 15:27:21.454Z (16:27:21): its sub-agent a26ab951d6306c1a7 starts. [transcript]
  - **15:27:39.467Z (16:27:39.467): that sub-agent calls `mcp__outlookai__list_accounts` with input
    `{}`.** [transcript]
  - **16:27:39: OUTLOOK.EXE PID 76888 created, command line `"C:\Program Files\Microsoft
    Office\Root\Office16\OUTLOOK.EXE" -Embedding`, parent svchost.exe PID 2228 (DcomLaunch).**
    [Win32_Process] In the same second, MCP 68964's thread 23908 (0x5D64) and Outlook's UI thread
    73000 (0x11D28) were created. [thread StartTime]
  - 16:27:41: event 45, Outlook loaded seven add-ins, including OutlookAI (boot time 656 ms),
    <custom-addin> (469 ms) and the Teams Meeting Add-in (16 ms). Verbatim in 1.4.
    [Application log]
    Outlook thread 29760 (0x7440), the Teams add-in's caller in 1.7, was created at 16:27:41.
    [thread StartTime]
  - 16:27:43 to 16:27:46: event 32 catalog checkpoints for six stores. At 16:27:44, event 63
    "GetAppManifests succeeded". These are the last `Outlook`-provider events of that instance.
    [Application log]
  - About 16:27:44.7: the UI thread entered the wait it never left. This time is derived from the
    wait-chain tick in 1.6 and depends on reading that tick (an inference).
  - 15:27:44.832Z (16:27:44.832), 0.1 s later: a Claude Agent SDK session starts with cwd
    `C:\WINDOWS\system32` (entrypoint `sdk-cli`, version 2.1.325, permissionMode `bypassPermissions`,
    prompt `-`). It ends at 15:27:48.087Z. What started it is unknown. A process started by DcomLaunch,
    such as this OUTLOOK.EXE, would by default have that working directory; that is not verified.
    [transcript 6c9c2c07, PART 6]
  - 16:28:40, 16:28:56 and 16:36:54: further Outlook threads start and block on the UI thread's critical
    section. [wait chains, 1.6]
  - About 16:33:40: MCP thread 0x5D64's current wait began (derived from its wait-chain tick, an
    inference).
  - 15:44:15.221Z (16:44:15): the user asks a new Claude session (fd2ffacb) "why is my outlook not
    starting?". That session's MCP server 67204 had started at 16:44:09. [transcript, Win32_Process]
  - 16:47:12 to 16:58: read-only diagnosis by that session's sub-agent; its outputs are throughout PART 1.
  - 15:50:00.482Z (16:50:00): the user interrupts sub-agent a26ab951's pending tool call
    ("[Request interrupted by user for tool use]"). At 15:50:14.382Z session 579e5f93 starts a
    replacement sub-agent (a4de6896a59e9e3b0). At 15:50:45.355Z the user asks the diagnosing session to
    also fix Outlook. [transcripts]
  - 16:52:15: a fresh Network List Manager client gets answers from netprofm in 108 ms (1.9).
  - 16:58:37: `cdb -pv` non-invasive attach to PID 76888, then detach (`qd`). Stacks in 1.7.
  - **17:02:02.374: `Stop-Process -Id 76888 -Force`. PID 76888 was gone at 17:02:02.402.** [command output]
  - **17:02:02.469: MCP 68964 starts thread 75728. OUTLOOK.EXE PID 60368 is created in the same second,
    again `-Embedding`, parent svchost 2228.** [Get-Process threads, Win32_Process] No other MCP server
    started a thread between 17:01:55 and 17:02:10 (1.11).
  - 16:02:03.1967660Z and 16:02:03.2510478Z (17:02:03): OutlookAI's `tuning.lastReconcileUtc` and
    `registration.addInLastReconcileUtc`. [outlook_health at 17:04, 1.13]
  - 17:02:09.738 and 17:02:09.778: MCP 68964 starts threads 27600 and 73948.
  - About 17:02:15 (not timestamped): PID 60368 has 285 threads, no hung windows, no blocked wait
    chains, and hidden Explorer windows only (1.10).
  - **17:02:54.368: the diagnosing agent runs `explorer.exe "C:\Program Files\Microsoft
    Office\root\Office16\OUTLOOK.EXE"`. No new OUTLOOK.EXE remained; 60368 was still the only one.**
  - **17:03:01: PID 60368 shows a visible, not hung window "Inbox - <mailbox-A> - Outlook".**
  - 17:03:39 to 17:03:59: verification round 1 (1.12).
  - 16:04:27Z to 16:05:43Z (17:04:27 to 17:05:43): the diagnosing agent calls the outlookai tools
    `outlook_health`, `list_accounts` and `search` (7 times); all succeeded (1.13).
  - Between 17:05:43 and 17:08:19: read-only COM reads by the diagnosing agent (1.14).
  - 17:08:19 to 17:08:34: verification round 2 (1.12).
  - 17:11:24: the coordinating session confirms PID 60368 (`OUTLOOK.EXE -Embedding`, started 17:02:02)
    is the only Outlook process, with Responding=True and MainWindowTitle
    "Inbox - <mailbox-A> - Outlook" (1.14).
  **1.3 Processes, at about 16:48 on 2026-11-15.** Source: `Get-CimInstance Win32_Process`; the parent
  chain was built by the diagnosing agent's script. Verbatim:
  ```text
  PID     : 5964
  Name    : OfficeClickToRun.exe
  Session : 0
  Created : 13-11-2026 11:14:22
  PPID    : 1736
  Chain   : services.exe(1736) <- wininit.exe(2012) <- <gone:1768>
  Cmd     :
  PID     : 32724
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 13-11-2026 17:22:08
  PPID    : 5912
  Chain   : claude.exe(5912) <- bash.exe(38340) <- bash.exe(36084) <- bash.exe(14424) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 26812
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 13-11-2026 17:22:19
  PPID    : 35756
  Chain   : claude.exe(35756) <- bash.exe(44672) <- bash.exe(24356) <- bash.exe(37604) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 65532
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 13-11-2026 17:46:34
  PPID    : 64472
  Chain   : claude.exe(64472) <- Code.exe(6348) <- Code.exe(23604) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 31340
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 13-11-2026 23:10:24
  PPID    : 49344
  Chain   : claude.exe(49344) <- bash.exe(40380) <- bash.exe(66288) <- bash.exe(19192) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 37340
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 14-11-2026 18:51:20
  PPID    : 52100
  Chain   : claude.exe(52100) <- bash.exe(51296) <- bash.exe(58732) <- bash.exe(22076) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 73176
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 14-11-2026 18:56:58
  PPID    : 73536
  Chain   : claude.exe(73536) <- bash.exe(72920) <- bash.exe(72412) <- bash.exe(73360) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 68964
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 15-11-2026 16:21:10
  PPID    : 76924
  Chain   : claude.exe(76924) <- bash.exe(62420) <- bash.exe(77464) <- bash.exe(76328) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  PID     : 76888
  Name    : OUTLOOK.EXE
  Session : 1
  Created : 15-11-2026 16:27:39
  PPID    : 2228
  Chain   : svchost.exe(2228) <- services.exe(1736) <- wininit.exe(2012) <- <gone:1768>
  Cmd     : "C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE" -Embedding
  PID     : 67204
  Name    : OutlookAI.McpServer.exe
  Session : 1
  Created : 15-11-2026 16:44:09
  PPID    : 78308
  Chain   : claude.exe(78308) <- bash.exe(64640) <- bash.exe(77628) <- bash.exe(59280) <- WindowsTerminal.exe(28816) <- explorer.exe(12928) <- <gone:12876>
  Cmd     : C:\Users\<user>\AppData\Local/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe
  ```
  - **Eight OutlookAI MCP servers were alive**, one per claude.exe parent:
    - 32724 (claude 5912), 26812 (claude 35756), 65532 (claude 64472, under VS Code) and 31340
      (claude 49344), all from 13 Nov;
    - 37340 (claude 52100) and 73176 (claude 73536), from 14 Nov;
    - 68964 (claude 76924), the trigger, from 15 Nov 16:21;
    - 67204 (claude 78308), the diagnosing session's, from 15 Nov 16:44.
    The first six belong to Claude sessions from 13 and 14 Nov that were still running. Matching 76924
    to session 579e5f93 and 78308 to session fd2ffacb is by timing (an inference). The other claude.exe
    processes' session ids were not checked.
  - **The DcomLaunch host**, `tasklist /svc` (verbatim):
  ```text
  Image Name                     PID Services
  ========================= ======== ============================================
  svchost.exe                   2228 BrokerInfrastructure, DcomLaunch, PlugPlay,
                                     Power, SystemEventsBroker
  Image Name                     PID Services
  ========================= ======== ============================================
  OfficeClickToRun.exe          5964 ClickToRunSvc
  ```
  - **Integrity and elevation**, from `OpenProcessToken` + `GetTokenInformation` (TokenElevation,
    TokenElevationType, TokenIntegrityLevel) at about 16:48 and 17:02. Verbatim:
  ```text
  === PID 76888
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  === PID 68964
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 0
  === PID 67204
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 0
  === PID 32724
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 0
  === PID 65532
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 0
  ```
    PID 60368's token (Medium, not elevated) is in 1.10. The integrity of MCP servers 26812, 31340,
    37340 and 73176 was NOT checked.
  - **PID 76888 state at about 16:48.** `Get-Process` and a 6-second CPU and thread sample. Verbatim
    (`Responding : True` means nothing here, because the process had no main window):
  ```text
  Id                 : 76888
  ProcessName        : OUTLOOK
  StartTime          : 15-11-2026 16:27:39
  Responding         : True
  MainWindowHandle   : 0
  MainWindowTitle    :
  CPU                : 10,78125
  TotalProcessorTime : 00:00:10.7812500
  UserProcessorTime  : 00:00:06.0156250
  HandleCount        : 6729
  Threads            : 323
  WS_MB              : 351,3
  Private_MB         : 315,2
  Path               : C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE
  FileVersion        : 16.0.14334.20974
  CPU delta over 6s: 0,000s
  Main thread 73000 (0x11D28): state=Wait wait=UserRequest start=15-11-2026 16:27:39 cpu=00:00:01.8593750
  --- thread state summary
  Count Name
  ----- ----
    249 Wait/UserRequest
     59 Wait/Unknown
      9 Wait/EventPairLow
      1 Wait/ExecutionDelay
  --- threads that used CPU in the 6s window
  --- earliest-started threads
     Id Hex     StartTime           ThreadState   WaitReason TotalProcessorTime
     -- ---     ---------           -----------   ---------- ------------------
  73000 0x11D28 15-11-2026 16:27:39        Wait  UserRequest 00:00:01.8593750
  27900 0x6CFC  15-11-2026 16:27:39        Wait  UserRequest 00:00:00
  10380 0x288C  15-11-2026 16:27:39        Wait  UserRequest 00:00:00
  44028 0xABFC  15-11-2026 16:27:39        Wait  UserRequest 00:00:00
  74684 0x123BC 15-11-2026 16:27:39        Wait  UserRequest 00:00:00.0468750
  69468 0x10F5C 15-11-2026 16:27:39        Wait EventPairLow 00:00:00
  ```
  - **PID 3240**, which the UI thread's outer COM call pointed at (1.6). Verbatim, with the tick-count
    anchor used for the times derived in 1.6:
  ```text
  ProcessId       : 3240
  Name            : svchost.exe
  SessionId       : 0
  CreationDate    : 13-11-2026 11:14:21
  ParentProcessId : 1736
  ExecutablePath  :
  CommandLine     :
  Image Name                     PID Services
  ========================= ======== ============================================
  svchost.exe                   3240 netprofm
  Image Name:   svchost.exe
  PID:          3240
  Session Name: Services
  Session#:     0
  Mem Usage:    23.668 K
  Status:       Unknown
  User Name:    N/A
  CPU Time:     0:00:22
  Window Title: N/A
  TickCount64 now (ms): 192986017  -> ticks(15.625ms): 12351105  at 16:50:36.512
  ```
  **1.4 Event logs, verbatim.**
  The `Outlook` provider in the Application log, all of 2026-11-15. These were the last events from
  PID 76888, and none were logged after 16:27:46:
  ```text
  [2026-11-15 16:27:41] Outlook id=45 lvl=Information
  Outlook loaded the following add-in(s):
  Name: Microsoft Exchange Add-in
  Description: Exchange support for Unified Messaging, e-mail permission rules, and calendar availability.
  ProgID: UmOutlookAddin.FormRegionAddin
  GUID: {F959DBBB-3867-41F2-8E5F-3B8BEFAA81B3}
  Load Behavior: 3
  HKLM: 1
  Location: C:\Program Files\Microsoft Office\root\Office16\ADDINS\UmOutlookAddin.dll
  Boot Time (Milliseconds): 0
  Name: Microsoft Teams Meeting Add-in for Microsoft Office
  Description: Microsoft Teams Meeting Add-in for Microsoft Office
  ProgID: TeamsAddin.FastConnect
  GUID: {19A6E644-14E6-4A60-B8D7-DD20610A871D}
  Load Behavior: 3
  HKLM: 0
  Location: C:\Users\<user>\AppData\Local\Microsoft\TeamsMeetingAdd-in\1.26.27102\x64\Microsoft.Teams.AddinLoader.dll
  Boot Time (Milliseconds): 16
  Name: OutlookAI
  Description: OutlookAI
  ProgID: OutlookAI
  GUID: {00000000-0000-0000-0000-000000000000}
  Load Behavior: 3
  HKLM: 0
  Location: file:///C:\Users\<user>\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal
  Boot Time (Milliseconds): 656
  Name: Outlook Social Connector 2016
  Description: Connects to social networking sites and provides people, activity, and status information.
  ProgID: OscAddin.Connect
  GUID: {2163EB1F-3FD9-4212-A41F-81D1F933597F}
  Load Behavior: 3
  HKLM: 1
  Location: C:\Program Files\Microsoft Office\root\Office16\SOCIALCONNECTOR.DLL
  Boot Time (Milliseconds): 16
  Name: OneNote Notes about Outlook Items
  Description: Adds Send to OneNote and Notes about this Item buttons to the command bar
  ProgID: OneNote.OutlookAddin
  GUID: {93E5752E-B889-47C5-8545-654EE2533C64}
  Load Behavior: 3
  HKLM: 1
  Location: C:\Program Files\Microsoft Office\root\Office16\ONBttnOL.dll
  Boot Time (Milliseconds): 15
  Name: <custom-addin>
  Description: <custom-addin>
  ProgID: <custom-addin>
  GUID: {00000000-0000-0000-0000-000000000000}
  Load Behavior: 3
  HKLM: 0
  Location: file:///C:/Users/<user>/source/repos/<custom-addin>/bin/Debug/<custom-addin>.vsto|vstolocal
  Boot Time (Milliseconds): 469
  Name: Microsoft SharePoint Server Colleague Import Add-in
  Description: The Add-in allows Microsoft SharePoint Server to import colleague suggestions based on your Outlook content
  ProgID: ColleagueImport.ColleagueImportAddin
  GUID: {EFEF7FDB-0CED-4FB6-B3BB-3C50D39F4120}
  Load Behavior: 3
  HKLM: 0
  Location: C:\Program Files\Microsoft Office\root\Office16\ADDINS\ColleagueImport.dll
  Boot Time (Milliseconds): 16
  [2026-11-15 16:27:43] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-C>.ost has detected a catalog checkpoint.
  [2026-11-15 16:27:43] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-B>.ost has detected a catalog checkpoint.
  [2026-11-15 16:27:44] Outlook id=63 lvl=Information
  The Exchange web service request GetAppManifests succeeded.
  [2026-11-15 16:27:44] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-A>.ost has detected a catalog checkpoint.
  [2026-11-15 16:27:44] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-A>.nst has detected a catalog checkpoint.
  [2026-11-15 16:27:46] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-B>(2).nst has detected a catalog checkpoint.
  [2026-11-15 16:27:46] Outlook id=32 lvl=Information
  The store C:\Users\<user>\AppData\Local\Microsoft\Outlook\<mailbox-D>.ost has detected a catalog checkpoint.
  ```
  `Outlook`-provider history: event-45 (add-ins loaded, one per Outlook boot) times over the 7 days
  before 2026-11-15 17:00, plus per-day counts by event id:
  ```text
  event 45 (add-ins loaded) entries:
  [2026-11-09 14:47:47] Outlook id=45 lvl=Information
  [2026-11-09 18:42:01] Outlook id=45 lvl=Information
  [2026-11-11 01:57:33] Outlook id=45 lvl=Information
  [2026-11-12 04:14:50] Outlook id=45 lvl=Information
  [2026-11-12 06:23:05] Outlook id=45 lvl=Information
  [2026-11-12 23:12:05] Outlook id=45 lvl=Information
  [2026-11-13 10:00:22] Outlook id=45 lvl=Information
  [2026-11-13 10:51:36] Outlook id=45 lvl=Information
  [2026-11-13 10:53:53] Outlook id=45 lvl=Information
  [2026-11-13 11:50:56] Outlook id=45 lvl=Information
  [2026-11-13 11:52:25] Outlook id=45 lvl=Information
  [2026-11-13 11:53:14] Outlook id=45 lvl=Information
  [2026-11-13 11:54:05] Outlook id=45 lvl=Information
  [2026-11-13 11:54:17] Outlook id=45 lvl=Information
  [2026-11-13 11:55:26] Outlook id=45 lvl=Information
  [2026-11-13 13:05:07] Outlook id=45 lvl=Information
  [2026-11-13 19:41:27] Outlook id=45 lvl=Information
  [2026-11-13 20:01:59] Outlook id=45 lvl=Information
  [2026-11-13 20:19:12] Outlook id=45 lvl=Information
  [2026-11-15 16:27:41] Outlook id=45 lvl=Information
  per-day counts by event id (30 store reconciliation started, 32 catalog checkpoint, 36 search indexing error, 38 reconciliation completed, 45 add-ins loaded, 63 GetAppManifests succeeded):
  2026-11-09 id=30         x8
  2026-11-09 id=32         x16
  2026-11-09 id=38         x8
  2026-11-09 id=45         x2
  2026-11-09 id=63         x15
  2026-11-10 id=32         x8
  2026-11-10 id=63         x19
  2026-11-11 id=32         x5
  2026-11-11 id=45         x1
  2026-11-11 id=63         x9
  2026-11-12 id=30         x8
  2026-11-12 id=32         x11
  2026-11-12 id=36         x1
  2026-11-12 id=38         x8
  2026-11-12 id=45         x3
  2026-11-12 id=63         x22
  2026-11-13 id=30         x39
  2026-11-13 id=38         x17
  2026-11-13 id=45         x13
  2026-11-13 id=63         x21
  2026-11-15 id=32         x6
  2026-11-15 id=45         x1
  2026-11-15 id=63         x1
  ```
  Application Error, Application Hang and Windows Error Reporting events mentioning outlook or olk.exe
  in the same 7 days. These are the only two: there was no Application Error 1000 and no WER crash
  report.
  ```text
  [2026-11-13 10:38:01] Windows Error Reporting id=1001 lvl=Information
  Fault bucket , type 0
  Event Name: AppHangB1
  Response: Not available
  Cab Id: 0
  Problem signature:
  P1: OUTLOOK.EXE
  P2: 16.0.14334.20974
  P3: 6acc8a3f
  P4: 0f9a
  P5: 134217728
  P6:
  P7:
  P8:
  P9:
  P10:
  Attached files:
  \\?\C:\ProgramData\Microsoft\Windows\WER\Temp\WER.a156f898-38ff-4ed7-9958-13afe4f52a41.tmp.WERInternalMetadata.xml
  \\?\C:\ProgramData\Microsoft\Windows\WER\Temp\WER.0efdcb78-8fc3-48d6-a4dc-758d9887c0ad.tmp.csv
  \\?\C:\ProgramData\Microsoft\Windows\WER\Temp\WER.5e6803ac-57eb-40a8-a4a4-63b55e58aa5b.tmp.txt
  \\?\C:\Users\<user>\AppData\Local\Temp\{D54BF330-195B-427F-B8F1-69A9860C1E1A} - OProcSessId.dat
  \\?\C:\Users\<user>\AppData\Local\Temp\WER.09acb592-9f75-402f-a4d4-53e47ee690e8.tmp.appcompat.txt
  \\?\C:\ProgramData\Microsoft\Windows\WER\Temp\WER.8e8ec779-e6cc-4eff-a980-fcf3962531a2.tmp.xml
  These files may be available here:
  NULL
  Analysis symbol:
  Rechecking for solution: 0
  Report Id: 1692bbde-a3ee-43a8-b086-aeeb8ca713e5
  Report Status: 97
  Hashed bucket:
  Cab Guid: 0
  [2026-11-13 10:38:01] Application Hang id=1002 lvl=Error
  The program OUTLOOK.EXE version 16.0.14334.20974 stopped interacting with Windows and was closed. To see if more information about the problem is available, check the problem history in the Security and Maintenance control panel.
  ```
  Office alerts (log `OAlerts`), last 30 days: the Outlook start and error entries. The other eight
  entries were ordinary save/draft prompts in Outlook, Word and Excel and are omitted.
  ```text
  [2026-10-19 10:56:00] Microsoft Office 16 Alerts id=300 lvl=Information
  Microsoft Outlook
  The operation failed.
  P1: 300032
  P2: 16.0.14334.20918
  P3: 749mP4: 0x80040107
  [2026-11-04 15:32:24] Microsoft Office 16 Alerts id=300 lvl=Information
  Microsoft Outlook
  The operation failed.
  P1: 300032
  P2: 16.0.14334.20974
  P3: 749mP4: 0x80040107
  [2026-11-13 11:32:36] Microsoft Office 16 Alerts id=300 lvl=Information
  Microsoft Outlook
  Outlook couldn't start last time. Safe mode could help you troubleshoot the problem, but some features might not be available in this mode.
  Do you want to start in safe mode?
  P1: 700159
  P2: 16.0.14334.20974
  P3: xatgP4:
  [2026-11-13 11:41:45] Microsoft Office 16 Alerts id=300 lvl=Information
  Microsoft Outlook
  Outlook couldn't start last time. Safe mode could help you troubleshoot the problem, but some features might not be available in this mode.
  Do you want to start in safe mode?
  P1: 700159
  P2: 16.0.14334.20974
  P3: xatgP4:
  ```
  NCSI (`Microsoft-Windows-NCSI/Operational`), last 7 days, all entries:
  ```text
  [2026-11-12 07:33:17] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: None ChangeReason: SuspectArpProbeFailed)
  [2026-11-12 07:33:58] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Local ChangeReason: PassivePacketHops)
  [2026-11-12 07:33:59] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Internet ChangeReason: ActiveHttpProbeSucceeded)
  [2026-11-13 10:53:39] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Local ChangeReason: PassivePacketHops)
  [2026-11-13 10:53:39] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Internet ChangeReason: ActiveHttpProbeSucceeded)
  [2026-11-13 11:14:23] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Local ChangeReason: PassivePacketHops)
  [2026-11-13 11:14:23] Microsoft-Windows-NCSI id=4042 lvl=Information
  Capability change on {3d5aa5f8-d9f7-44f5-b73c-979c3802cf91} (0x600800B000000 Family: V4 Capability: Internet ChangeReason: ActiveHttpProbeSucceeded)
  ```
  **1.5 Windows of the hung PID 76888, about 16:48.** Source: `EnumWindows`, `IsWindowVisible`,
  `IsHungAppWindow`, `GetWindowRect`, `GetClassName` and `GetWindowText`. Every window is invisible.
  Those on the UI thread (0x12C58 = PID 76888, 0x11D28 = TID 73000, as in the `ViewDeferedNotifManager`
  title) report hung. The hidden Explorer `Inbox - <mailbox-B> - Outlook` existed. Verbatim:
  ```text
  === PID 76888
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 27
    hwnd=0xD01DE vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='HardwareMonitorWindowClass' title='HardwareMonitorWindow'
    hwnd=0x210F92 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='tooltips_class32' title=''
    hwnd=0x1B0DFA vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='ViewDeferedNotifManager' title='ViewDeferedNotifManager 0x12C58 0x11D28'
    hwnd=0x280D38 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='TNOTIFThisThreadSink' title='TNOTIFThisThreadSink 0x12C58 0x11D28'
    hwnd=0x2F0DFE vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='_WwO' title=''
    hwnd=0x3B0FB4 vis=False hung=True min=False owner=0x0 rect=(-8,-8,3848,2120) class='OpusApp' title='Word'
    hwnd=0x1B0ED6 vis=False hung=True min=False owner=0x480FEA rect=(0,0,29,16) class='WBPOPUP' title=''
    hwnd=0x2B218C vis=False hung=False min=False owner=0x0 rect=(0,0,136,39) class='Hidden Window of the IM Abstraction layer' title=''
    hwnd=0x430FBC vis=False hung=True min=False owner=0x0 rect=(0,0,136,39) class='rctrl_sbpwnd32' title=''
    hwnd=0x30203A vis=False hung=True min=False owner=0x0 rect=(0,0,16,16) class='rctrl_sbpwnd32' title=''
    hwnd=0x2A0466 vis=False hung=True min=False owner=0x0 rect=(0,0,16,16) class='rctrl_sbpwnd32' title=''
    hwnd=0x480FEA vis=False hung=True min=False owner=0x0 rect=(156,156,3036,1719) class='rctrl_renwnd32' title='Inbox - <mailbox-B> - Outlook'
    hwnd=0x4C20B2 vis=False hung=True min=False owner=0x0 rect=(0,0,16,16) class='MerengueNotifyWnd' title='Microsoft Outlook Social Connector'
    hwnd=0x2141FB6 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='WMS ST Notif Class' title='WMS ST Notif Window 00012C58 00011D28'
    hwnd=0x322024 vis=False hung=True min=False owner=0x0 rect=(0,0,16,16) class='OutlookAcctMgrNotificationWindow' title=''
    hwnd=0x4F0F7A vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='MsoStdCompMgr' title=''
    hwnd=0xD220F2 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='.NET-BroadcastEventWindow.4.0.0.0.2982bee.0' title='.NET-BroadcastEventWindow.4.0.0.0.2982bee.0'
    hwnd=0xDD1EB2 vis=False hung=True min=False owner=0x0 rect=(78,78,2958,1641) class='TeamsMeetingAddinWindow' title=''
    hwnd=0x800096 vis=False hung=False min=False owner=0x0 rect=(0,0,136,39) class='Outlook Notification Area Icon Window' title=''
    hwnd=0x16C09E2 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='OleDdeWndClass' title='DDE Server Window'
    hwnd=0x10F22E6 vis=False hung=True min=False owner=0x0 rect=(0,0,0,0) class='WMS Notif Engine:Dispatch Window Class' title='W'
    hwnd=0xE1071A vis=False hung=False min=False owner=0x0 rect=(0,0,136,39) class='OfficePowerManagerWindow' title='OfficePowerManagerWindow'
    hwnd=0x900AD8 vis=False hung=True min=False owner=0x0 rect=(32767,32767,98302,98302) class='ARC Event Window 00' title=''
    hwnd=0x392152 vis=False hung=True min=False owner=0x0 rect=(32767,32767,98302,98302) class='ARC Event Window 00' title=''
    hwnd=0xE05F4 vis=False hung=True min=False owner=0x0 rect=(3841,2161,3857,2177) class='mspim_wnd32' title='Microsoft Outlook'
    hwnd=0xB60112 vis=False hung=False min=False owner=0x800096 rect=(0,0,0,0) class='IME' title='Default IME'
    hwnd=0x3520EC vis=False hung=True min=False owner=0xE05F4 rect=(0,0,0,0) class='IME' title='Default IME'
  ```
  **1.6 Wait chains**, captured 16:50 to 16:51 with `GetThreadWaitChain` (flags
  `WCT_OUT_OF_PROC_FLAG | WCT_OUT_OF_PROC_CS_FLAG | WCT_OUT_OF_PROC_COM_FLAG`, COM callbacks
  registered via `RegisterWaitChainCOMCallback`), the API behind Resource Monitor's "Analyze wait
  chain".
  - `waitms` is printed raw from `ThreadObject.WaitTime`. The diagnosing agent's reading (an
    inference): it is the kernel tick count, in 15.625 ms ticks, at which the wait began, not a
    duration. A duration is impossible here, because 12263311 ms is longer than the process had
    existed.
  - With the anchor in 1.3 (tick 12351105 at 16:50:36.512), tick 12263311 gives about 16:27:44.7 for
    the UI thread, and tick 12286093 gives about 16:33:40 for MCP thread 0x5D64.
  - For PID 76888, all chains that contain more than the thread itself are below (317 threads were
    walked; the rest showed only their own node). `Com[Owned] -> Thread[pid=3240 tid=0 PidOnly]` means
    WCT saw an outgoing COM call owned by PID 3240 (netprofm) but could not see which thread there.
  ```text
  === PID 76888 OUTLOOK threads=317
    tid 73000 (0x11D28) start=16:27:39: Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 62952 (0xF5E8) start=16:27:39: Thread[pid=76888 tid=62952(0xF5E8) waitms=12263630 ctxsw=527 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 68656 (0x10C30) start=16:27:40: Thread[pid=76888 tid=68656(0x10C30) waitms=12263630 ctxsw=559 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 23592 (0x5C28) start=16:27:40: Thread[pid=76888 tid=23592(0x5C28) waitms=12263630 ctxsw=469 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 59712 (0xE940) start=16:27:41: Thread[pid=76888 tid=59712(0xE940) waitms=12263630 ctxsw=423 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 29760 (0x7440) start=16:27:41: Thread[pid=76888 tid=29760(0x7440) waitms=12263279 ctxsw=161 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 28728 (0x7038) start=16:27:41: Thread[pid=76888 tid=28728(0x7038) waitms=12263533 ctxsw=448 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=72648(0x11BC8) waitms=12263630 ctxsw=387 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 39728 (0x9B30) start=16:27:41: Thread[pid=76888 tid=39728(0x9B30) waitms=12263630 ctxsw=510 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 632 (0x278) start=16:27:42: Thread[pid=76888 tid=632(0x278) waitms=12263630 ctxsw=59 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 68784 (0x10CB0) start=16:27:42: Thread[pid=76888 tid=68784(0x10CB0) waitms=12263630 ctxsw=319 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 69776 (0x11090) start=16:27:42: Thread[pid=76888 tid=69776(0x11090) waitms=12263630 ctxsw=140 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 74836 (0x12454) start=16:27:42: Thread[pid=76888 tid=74836(0x12454) waitms=12263630 ctxsw=297 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 72648 (0x11BC8) start=16:27:42: Thread[pid=76888 tid=72648(0x11BC8) waitms=12263630 ctxsw=387 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 44020 (0xABF4) start=16:27:42: Thread[pid=76888 tid=44020(0xABF4) waitms=12263630 ctxsw=215 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 76896 (0x12C60) start=16:27:42: Thread[pid=76888 tid=76896(0x12C60) waitms=12263630 ctxsw=419 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 71888 (0x118D0) start=16:27:42: Thread[pid=76888 tid=71888(0x118D0) waitms=12263630 ctxsw=39 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 31044 (0x7944) start=16:27:42: Thread[pid=76888 tid=31044(0x7944) waitms=12263630 ctxsw=37 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 74816 (0x12440) start=16:27:43: Thread[pid=76888 tid=74816(0x12440) waitms=12263630 ctxsw=294 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 78436 (0x13264) start=16:27:43: Thread[pid=76888 tid=78436(0x13264) waitms=12263630 ctxsw=347 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 21932 (0x55AC) start=16:27:43: Thread[pid=76888 tid=21932(0x55AC) waitms=12263630 ctxsw=1138 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 57004 (0xDEAC) start=16:27:43: Thread[pid=76888 tid=57004(0xDEAC) waitms=12263630 ctxsw=315 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 75980 (0x128CC) start=16:27:43: Thread[pid=76888 tid=75980(0x128CC) waitms=12263630 ctxsw=75 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 64112 (0xFA70) start=16:27:43: Thread[pid=76888 tid=64112(0xFA70) waitms=12263630 ctxsw=77 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 62384 (0xF3B0) start=16:27:43: Thread[pid=76888 tid=62384(0xF3B0) waitms=12263630 ctxsw=200 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 74384 (0x12290) start=16:27:43: Thread[pid=76888 tid=74384(0x12290) waitms=12263284 ctxsw=148 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=72648(0x11BC8) waitms=12263630 ctxsw=387 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 2676 (0xA74) start=16:27:43: Thread[pid=76888 tid=2676(0xA74) waitms=12263630 ctxsw=74 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 12396 (0x306C) start=16:27:43: Thread[pid=76888 tid=12396(0x306C) waitms=12265080 ctxsw=151 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 69772 (0x1108C) start=16:27:43: Thread[pid=76888 tid=69772(0x1108C) waitms=12263630 ctxsw=70 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 78492 (0x1329C) start=16:27:44: Thread[pid=76888 tid=78492(0x1329C) waitms=12265080 ctxsw=109 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 30992 (0x7910) start=16:27:44: Thread[pid=76888 tid=30992(0x7910) waitms=12265080 ctxsw=62 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 76280 (0x129F8) start=16:27:45: Thread[pid=76888 tid=76280(0x129F8) waitms=12263630 ctxsw=77 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 61016 (0xEE58) start=16:27:52: Thread[pid=76888 tid=61016(0xEE58) waitms=12265080 ctxsw=38 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 77560 (0x12EF8) start=16:27:52: Thread[pid=76888 tid=77560(0x12EF8) waitms=12271101 ctxsw=45 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 56940 (0xDE6C) start=16:28:40: Thread[pid=76888 tid=56940(0xDE6C) waitms=12267892 ctxsw=11 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 77172 (0x12D74) start=16:28:40: Thread[pid=76888 tid=77172(0x12D74) waitms=12267892 ctxsw=2 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 77336 (0x12E18) start=16:28:56: Thread[pid=76888 tid=77336(0x12E18) waitms=12267892 ctxsw=1 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 70460 (0x1133C) start=16:28:56: Thread[pid=76888 tid=70460(0x1133C) waitms=12267892 ctxsw=2 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 76992 (0x12CC0) start=16:28:56: Thread[pid=76888 tid=76992(0x12CC0) waitms=12267892 ctxsw=1 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 67388 (0x1073C) start=16:28:56: Thread[pid=76888 tid=67388(0x1073C) waitms=12267892 ctxsw=1 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 75060 (0x12534) start=16:28:56: Thread[pid=76888 tid=75060(0x12534) waitms=12271101 ctxsw=15 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 77324 (0x12E0C) start=16:28:56: Thread[pid=76888 tid=77324(0x12E0C) waitms=12271101 ctxsw=7 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 73520 (0x11F30) start=16:28:56: Thread[pid=76888 tid=73520(0x11F30) waitms=12271101 ctxsw=20 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
    tid 58760 (0xE588) start=16:36:54: Thread[pid=76888 tid=58760(0xE588) waitms=12321521 ctxsw=70 Blocked] -> CriticalSection[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
  ```
  Every OutlookAI MCP server, Explorer and PID 3240, same API with only multi-node chains printed. Only
  68964 had one: thread 0x5D64, created 16:27:39, in a COM call to Outlook's UI thread. Verbatim:
  ```text
  === PID 26812 OutlookAI.McpServer threads=13
  === PID 31340 OutlookAI.McpServer threads=13
  === PID 32724 OutlookAI.McpServer threads=15
  === PID 37340 OutlookAI.McpServer threads=13
  === PID 65532 OutlookAI.McpServer threads=12
  === PID 67204 OutlookAI.McpServer threads=13
  === PID 68964 OutlookAI.McpServer threads=16
    tid 23908 (0x5D64) start=16:27:39: Thread[pid=68964 tid=23908(0x5D64) waitms=12286093 ctxsw=1633 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=76888 tid=73000(0x11D28) waitms=12263311 ctxsw=11257 Blocked] -> Com[Owned '' timeout=0] -> Thread[pid=3240 tid=0(0x0) waitms=0 ctxsw=0 PidOnly]
  === PID 73176 OutlookAI.McpServer threads=13
  === PID 12928 explorer threads=469
  === PID 3240 svchost threads=11
  ```
  **1.7 Stacks.** From `cdb.exe` 10.0.29680.1004 (WinDbg package 1.2611.4001.0), attached
  non-invasively (`-pv`) to PID 76888 at 16:58:37 with Microsoft public symbols, then detached
  (`qd`). Office's PDBs are public symbols only, so `OUTLOOK!`, `OLMAPI32!`, `EMSMDB32!` and `mso!` names
  are the nearest exports plus an offset, not real function names. The `Microsoft_Teams_MeetingAddin!`
  frames are managed frames the debugger resolved. Header lines, verbatim:
  ```text
  Microsoft (R) Windows Debugger Version 10.0.29680.1004 AMD64
  Copyright (c) Microsoft Corporation. All rights reserved.
  Symbol search path is: srv*C:\Users\<user>\Downloads\tmp-outlook-diag\sym*https://msdl.microsoft.com/download/symbols
  WARNING: Process 76888 is not attached as a debuggee
           The process can be examined but debug events will not be received
  (12c58.11d28): Wake debugger - code 80000007 (first chance)
  0:000> cdb: Reading initial command '.symopt+ 0x80000; ~~[0x11d28]s; .echo === MAIN THREAD 0x11D28 ===; kn 100; .echo === THREAD 0x7440 ===; ~~[0x7440]kn 60; .echo === ALL THREADS TOP FRAMES ===; ~*kn 8; qd'
  ```
  The UI thread 0x11D28, `kn 100`, verbatim. Read it bottom-up:
  - 0x74-0x6f: Outlook's message loop.
  - 0x6e-0x5c: an incoming COM call via `IDispatch::Invoke`, i.e. late-bound.
  - 0x5b-0x41: Outlook, OLMAPI32 and EMSMDB32 opening a store, down to `OLMAPI32!HrCreateOfflineObj`.
  - 0x40: `netprofm!PublicNetworkListManager::GetCost`.
  - 0x3f-0x2d: an outgoing COM call waiting in COM's modal loop (`CCliModalLoop`).
  - 0x2c-0x1e: a second incoming call, dispatched re-entrantly (`ReentrantSTAInvokeInApartment`) through
    the universal marshaler's stub (`CUnivStubWrapper`).
  - 0x1d-0x03: Outlook, OLMAPI32 and EMSMDB32 again, down to `OLMAPI32!StartSyncLogging` ->
    `GetMAPISessionID` -> `HrCreateSingletonQueueOlkAsyncTask` -> `ScInitMapiUtil`.
  - 0x02-0x00: `mso!Ordinal9147` -> `WaitForMultipleObjectsEx`.
  ```text
  === MAIN THREAD 0x11D28 ===
   # Child-SP          RetAddr               Call Site
  00 000000fa`a9af9458 00007ffd`db50ab53     ntdll!NtWaitForMultipleObjects+0x14
  01 000000fa`a9af9460 00007ffc`87f2d693     KERNELBASE!WaitForMultipleObjectsEx+0x123
  02 000000fa`a9af9750 00007ffc`b99f0534     mso!Ordinal9147+0x163
  03 000000fa`a9af97c0 00007ffc`b9a00195     OLMAPI32!ScInitMapiUtil+0x714
  04 000000fa`a9af9820 00007ffc`b99e78bd     OLMAPI32!HrCreateSingletonQueueOlkAsyncTask+0x615
  05 000000fa`a9af9930 00007ffc`b99e77cd     OLMAPI32!GetMAPISessionID+0xc9d
  06 000000fa`a9af9980 00007ffc`b9c8175f     OLMAPI32!GetMAPISessionID+0xbad
  07 000000fa`a9af99b0 00007ffc`8bb8af3f     OLMAPI32!StartSyncLogging+0x80cf
  08 000000fa`a9af99f0 00007ffc`8bb8d51d     EMSMDB32+0xaaf3f
  09 000000fa`a9af9b40 00007ffc`8bbb6cc2     EMSMDB32+0xad51d
  0a 000000fa`a9af9cf0 00007ffc`8bbb6b32     EMSMDB32!HrEmsuiInit+0x3c42
  0b 000000fa`a9af9d30 00007ffc`8bbb6c52     EMSMDB32!HrEmsuiInit+0x3ab2
  0c 000000fa`a9af9da0 00007ffc`8bbc9077     EMSMDB32!HrEmsuiInit+0x3bd2
  0d 000000fa`a9af9e10 00007ffc`8bc3569e     EMSMDB32!HrEmsuiInit+0x15ff7
  0e 000000fa`a9af9e70 00007ffc`8bc353ff     EMSMDB32!MSProviderInit+0x56f3e
  0f 000000fa`a9af9ed0 00007ffc`8bc352fa     EMSMDB32!MSProviderInit+0x56c9f
  10 000000fa`a9af9f40 00007ffc`b9a6c801     EMSMDB32!MSProviderInit+0x56b9a
  11 000000fa`a9af9fc0 00007ffc`b99eb4d9     OLMAPI32!HrConvertMAPIFormPropsToFDMProps+0x3061
  12 000000fa`a9afa0c0 00007ffc`b99f2e38     OLMAPI32!WrapStoreEntryID+0x1069
  13 000000fa`a9afa190 00007ff7`1a4d6d96     OLMAPI32!HrOpenABEntryWithExchangeContext+0xf8
  14 000000fa`a9afa210 00007ff7`1a4d6cc3     OUTLOOK!HrSetOutlookSpecialFolderEntryID+0x241056
  15 000000fa`a9afa2a0 00007ff7`1a4d6b81     OUTLOOK!HrSetOutlookSpecialFolderEntryID+0x240f83
  16 000000fa`a9afa2d0 00007ff7`1a4d69ef     OUTLOOK!HrSetOutlookSpecialFolderEntryID+0x240e41
  17 000000fa`a9afa3c0 00007ff7`19d6a518     OUTLOOK!HrSetOutlookSpecialFolderEntryID+0x240caf
  18 000000fa`a9afa3f0 00007ff7`19d6a050     OUTLOOK!FFolderSupportsUnicode+0x63b8
  19 000000fa`a9afa530 00007ff7`19cebb8c     OUTLOOK!FFolderSupportsUnicode+0x5ef0
  1a 000000fa`a9afa630 00007ff7`19e1f67a     OUTLOOK+0xdbb8c
  1b 000000fa`a9afa770 00007ff7`19e1f32d     OUTLOOK!GetOutlookSafeModeState+0x3da0a
  1c 000000fa`a9afa890 00007ff7`1a4d69c1     OUTLOOK!GetOutlookSafeModeState+0x3d6bd
  1d 000000fa`a9afa8f0 00007ffd`dd492d53     OUTLOOK!HrSetOutlookSpecialFolderEntryID+0x240c81
  1e 000000fa`a9afa930 00007ffd`dd40c60d     RPCRT4!Invoke+0x73
  1f 000000fa`a9afa980 00007ffd`dc8d450d     RPCRT4!NdrStubCall2+0x30d
  20 000000fa`a9afac30 00007ffd`dc73c4a0     combase!CStdStubBuffer_Invoke+0x14d [onecore\com\combase\ndr\ndrole\stub.cxx @ 1400]
  21 000000fa`a9afac70 00007ffd`dc8d3eee     OLEAUT32!CUnivStubWrapper::Invoke+0x50
  22 (Inline Function) --------`--------     combase!InvokeStubWithExceptionPolicyAndTracing::__l6::<lambda_c9f3956a20c9da92a64affc24fdd69ec>::operator()+0x2d [onecore\com\combase\dcomrem\channelb.cxx @ 1152]
  23 000000fa`a9afacd0 00007ffd`dc8d36b8     combase!ObjectMethodExceptionHandlingAction<<lambda_c9f3956a20c9da92a64affc24fdd69ec> >+0x4e [onecore\com\combase\dcomrem\excepn.hxx @ 94]
  24 (Inline Function) --------`--------     combase!InvokeStubWithExceptionPolicyAndTracing+0x199 [onecore\com\combase\dcomrem\channelb.cxx @ 1150]
  25 000000fa`a9afad30 00007ffd`dc95ad00     combase!DefaultStubInvoke+0x3c8 [onecore\com\combase\dcomrem\channelb.cxx @ 1219]
  26 000000fa`a9afaf00 00007ffd`dc8d5d42     combase!SyncServerCall::StubInvoke+0x40 [onecore\com\combase\dcomrem\ServerCall.hpp @ 791]
  27 (Inline Function) --------`--------     combase!StubInvoke+0x201 [onecore\com\combase\dcomrem\channelb.cxx @ 1484]
  28 000000fa`a9afaf40 00007ffd`dc8d4951     combase!ServerCall::ContextInvoke+0x3c2 [onecore\com\combase\dcomrem\ctxchnl.cxx @ 1437]
  29 (Inline Function) --------`--------     combase!DefaultInvokeInApartment+0x8e [onecore\com\combase\dcomrem\callctrl.cxx @ 3245]
  2a 000000fa`a9afb0e0 00007ffd`dc881686     combase!ReentrantSTAInvokeInApartment+0x1f1 [onecore\com\combase\dcomrem\reentrantsta.cpp @ 110]
  2b 000000fa`a9afb1d0 00007ffd`dc85ae94     combase!ComInvokeWithLockAndIPID+0x11c6 [onecore\com\combase\dcomrem\channelb.cxx @ 2153]
  2c 000000fa`a9afb4c0 00007ffd`dc9020b0     combase!ThreadDispatch+0x464 [onecore\com\combase\dcomrem\channelb.cxx @ 1634]
  2d 000000fa`a9afb640 00007ffd`dd9952e6     combase!ThreadWndProc+0xf0 [onecore\com\combase\dcomrem\chancont.cxx @ 733]
  2e 000000fa`a9afb6a0 00007ffd`dd995dbd     USER32!UserCallWinProcCheckWow+0x356
  2f 000000fa`a9afb800 00007ffd`dc9067c2     USER32!DispatchMessageWorker+0x1dd
  30 (Inline Function) --------`--------     combase!CCliModalLoop::MyDispatchMessage+0xc [onecore\com\combase\dcomrem\callctrl.cxx @ 2947]
  31 000000fa`a9afb880 00007ffd`dc9057b7     combase!CCliModalLoop::PeekRPCAndDDEMessage+0x7e [onecore\com\combase\dcomrem\callctrl.cxx @ 2559]
  32 000000fa`a9afb8f0 00007ffd`dc905395     combase!CCliModalLoop::BlockFn+0x1b7 [onecore\com\combase\dcomrem\callctrl.cxx @ 2055]
  33 000000fa`a9afb980 00007ffd`dc883a60     combase!ModalLoop+0xc9 [onecore\com\combase\dcomrem\chancont.cxx @ 172]
  34 (Inline Function) --------`--------     combase!ThreadSendReceive+0x1128 [onecore\com\combase\dcomrem\channelb.cxx @ 7268]
  35 (Inline Function) --------`--------     combase!CSyncClientCall::SwitchAptAndDispatchCall+0x11c4 [onecore\com\combase\dcomrem\channelb.cxx @ 5744]
  36 000000fa`a9afb9f0 00007ffd`dc9037f2     combase!CSyncClientCall::SendReceive2+0x1350 [onecore\com\combase\dcomrem\channelb.cxx @ 5306]
  37 (Inline Function) --------`--------     combase!SyncClientCallRetryContext::SendReceiveWithRetry+0x37 [onecore\com\combase\dcomrem\callctrl.cxx @ 1494]
  38 (Inline Function) --------`--------     combase!CSyncClientCall::SendReceiveInRetryContext+0x37 [onecore\com\combase\dcomrem\callctrl.cxx @ 581]
  39 000000fa`a9afbe10 00007ffd`dc8d6d2a     combase!ClassicSTAThreadSendReceive+0x1a2 [onecore\com\combase\dcomrem\callctrl.cxx @ 563]
  3a 000000fa`a9afbf30 00007ffd`dc8d65c3     combase!CSyncClientCall::SendReceive+0x52a [onecore\com\combase\dcomrem\ctxchnl.cxx @ 798]
  3b (Inline Function) --------`--------     combase!CClientChannel::SendReceive+0x59 [onecore\com\combase\dcomrem\ctxchnl.cxx @ 675]
  3c 000000fa`a9afc0a0 00007ffd`dd497a91     combase!NdrExtpProxySendReceive+0xc3 [onecore\com\combase\ndr\ndrole\proxy.cxx @ 1899]
  3d 000000fa`a9afc0f0 00007ffd`dc947e56     RPCRT4!NdrpClientCall3+0x431
  3e 000000fa`a9afc460 00007ffd`dcaa6642     combase!ObjectStublessClient+0x146 [onecore\com\combase\ndr\ndrole\amd64\stblsclt.cxx @ 366]
  3f 000000fa`a9afc7f0 00007ffd`d40dd72c     combase!ObjectStubless+0x42 [onecore\com\combase\ndr\ndrole\amd64\stubless.asm @ 176]
  40 000000fa`a9afc840 00007ffc`b99e5cd4     netprofm!PublicNetworkListManager::GetCost+0xdc
  41 000000fa`a9afc8b0 00007ffc`b99e5bfc     OLMAPI32!DllGetClassObject+0x644
  42 000000fa`a9afc8f0 00007ffc`b99e4519     OLMAPI32!DllGetClassObject+0x56c
  43 000000fa`a9afc950 00007ffc`b99e4355     OLMAPI32!ExchangePerf_DuplicateClientInfo+0x3f9
  44 000000fa`a9afcad0 00007ffc`b99d9277     OLMAPI32!ExchangePerf_DuplicateClientInfo+0x235
  45 000000fa`a9afcb10 00007ffc`b99d8fa2     OLMAPI32!HrGetMAPIMalloc2Spec+0x7d7
  46 000000fa`a9afce00 00007ffc`b99dade8     OLMAPI32!HrGetMAPIMalloc2Spec+0x502
  47 000000fa`a9afcec0 00007ffc`8bbcd7b3     OLMAPI32!HrCreateOfflineObj+0x358
  48 000000fa`a9afd5b0 00007ffc`8bbcec88     EMSMDB32!HrEmsuiInit+0x1a733
  49 000000fa`a9afd6d0 00007ffc`8bbcd9cb     EMSMDB32!HrEmsuiInit+0x1bc08
  4a 000000fa`a9afda80 00007ffc`b99ede5a     EMSMDB32!HrEmsuiInit+0x1a94b
  4b 000000fa`a9afdb00 00007ffc`b99edb6b     OLMAPI32!WrapStoreEntryID+0x39ea
  4c 000000fa`a9afdb90 00007ffc`b99edfa0     OLMAPI32!WrapStoreEntryID+0x36fb
  4d 000000fa`a9afdc80 00007ffc`b99ed184     OLMAPI32!WrapStoreEntryID+0x3b30
  4e 000000fa`a9afdcf0 00007ffc`b99ed90d     OLMAPI32!WrapStoreEntryID+0x2d14
  4f 000000fa`a9afde10 00007ffc`b9a88465     OLMAPI32!WrapStoreEntryID+0x349d
  50 000000fa`a9afdf30 00007ffc`b9a9ed60     OLMAPI32!PRProviderInit+0x1615
  51 000000fa`a9afe0e0 00007ff7`19dce766     OLMAPI32!HrCopyAdapterInfo+0x8a0
  52 000000fa`a9afe140 00007ff7`19f84810     OUTLOOK!FreeDelegatorInfo+0x34f86
  53 000000fa`a9afe3a0 00007ff7`19d5a33b     OUTLOOK!HrBgScheduleRepairApp+0x85cb0
  54 000000fa`a9afe400 00007ff7`19d5a25f     OUTLOOK+0x14a33b
  55 000000fa`a9afe460 00007ff7`1a04ef51     OUTLOOK+0x14a25f
  56 000000fa`a9afe510 00007ff7`1a09b2a7     OUTLOOK!OutlookSyncEventOccurredEx+0x2961
  57 000000fa`a9afe540 00007ff7`1a09b1e3     OUTLOOK!OutlookSyncEventOccurredEx+0x4ecb7
  58 000000fa`a9afe590 00007ff7`19d6a518     OUTLOOK!OutlookSyncEventOccurredEx+0x4ebf3
  59 000000fa`a9afe670 00007ff7`19d6a050     OUTLOOK!FFolderSupportsUnicode+0x63b8
  5a 000000fa`a9afe7d0 00007ff7`19cebb8c     OUTLOOK!FFolderSupportsUnicode+0x5ef0
  5b 000000fa`a9afe8d0 00007ffd`dc759c38     OUTLOOK+0xdbb8c
  5c 000000fa`a9afea10 00007ffd`dc759b50     OLEAUT32!IDispatch_Invoke_Stub+0xd8
  5d 000000fa`a9afeaa0 00007ffd`dd40cbd0     OLEAUT32!IDispatch_RemoteInvoke_Thunk+0x60
  5e 000000fa`a9afeb10 00007ffd`dc8d449f     RPCRT4!NdrStubCall2+0x8d0
  5f 000000fa`a9afedc0 00007ffd`dc757254     combase!CStdStubBuffer_Invoke+0xdf [onecore\com\combase\ndr\ndrole\stub.cxx @ 1479]
  60 000000fa`a9afee00 00007ffd`dc8d3eee     OLEAUT32!CStubWrapper::Invoke+0x44
  61 (Inline Function) --------`--------     combase!InvokeStubWithExceptionPolicyAndTracing::__l6::<lambda_c9f3956a20c9da92a64affc24fdd69ec>::operator()+0x2d [onecore\com\combase\dcomrem\channelb.cxx @ 1152]
  62 000000fa`a9afee40 00007ffd`dc8d36b8     combase!ObjectMethodExceptionHandlingAction<<lambda_c9f3956a20c9da92a64affc24fdd69ec> >+0x4e [onecore\com\combase\dcomrem\excepn.hxx @ 94]
  63 (Inline Function) --------`--------     combase!InvokeStubWithExceptionPolicyAndTracing+0x199 [onecore\com\combase\dcomrem\channelb.cxx @ 1150]
  64 000000fa`a9afeea0 00007ffd`dc95ad00     combase!DefaultStubInvoke+0x3c8 [onecore\com\combase\dcomrem\channelb.cxx @ 1219]
  65 000000fa`a9aff070 00007ffd`dc8d5d42     combase!SyncServerCall::StubInvoke+0x40 [onecore\com\combase\dcomrem\ServerCall.hpp @ 791]
  66 (Inline Function) --------`--------     combase!StubInvoke+0x201 [onecore\com\combase\dcomrem\channelb.cxx @ 1484]
  67 000000fa`a9aff0b0 00007ffd`dc8d4951     combase!ServerCall::ContextInvoke+0x3c2 [onecore\com\combase\dcomrem\ctxchnl.cxx @ 1437]
  68 (Inline Function) --------`--------     combase!DefaultInvokeInApartment+0x8e [onecore\com\combase\dcomrem\callctrl.cxx @ 3245]
  69 000000fa`a9aff250 00007ffd`dc881686     combase!ReentrantSTAInvokeInApartment+0x1f1 [onecore\com\combase\dcomrem\reentrantsta.cpp @ 110]
  6a 000000fa`a9aff340 00007ffd`dc85ae94     combase!ComInvokeWithLockAndIPID+0x11c6 [onecore\com\combase\dcomrem\channelb.cxx @ 2153]
  6b 000000fa`a9aff630 00007ffd`dc9020b0     combase!ThreadDispatch+0x464 [onecore\com\combase\dcomrem\channelb.cxx @ 1634]
  6c 000000fa`a9aff7b0 00007ffd`dd9952e6     combase!ThreadWndProc+0xf0 [onecore\com\combase\dcomrem\chancont.cxx @ 733]
  6d 000000fa`a9aff810 00007ffd`dd995dbd     USER32!UserCallWinProcCheckWow+0x356
  6e 000000fa`a9aff970 00007ff7`19cf68a7     USER32!DispatchMessageWorker+0x1dd
  6f 000000fa`a9aff9f0 00007ff7`19fff475     OUTLOOK+0xe68a7
  70 000000fa`a9affbc0 00007ff7`19e25c7f     OUTLOOK!RefreshOutlookETWLoggingState+0x1acc5
  71 000000fa`a9affc30 00007ff7`1a00a632     OUTLOOK!GetOutlookSafeModeState+0x4400f
  72 000000fa`a9affc80 00007ffd`dd75dda7     OUTLOOK!OlkGetResourceHandle+0x6af2
  73 000000fa`a9affcc0 00007ffd`ddfb9c5c     KERNEL32!BaseThreadInitThunk+0x17
  74 000000fa`a9affcf0 00000000`00000000     ntdll!RtlUserThreadStart+0x2c
  ```
  Thread 0x7440 (TID 29760, created 16:27:41), `kn 60`, verbatim. It is a CLR thread-pool thread in
  the Teams Meeting Add-in's telemetry setup:
  `<SetupTelemetryAsync>` -> `TelemetryManager.CompleteInitialize` ->
  `MsoAriaTelemetryProvider.GetCurrentSmtpAddress` -> `OutlookApplicationModel.GetCurrentUserPrimarySmtpAddress`,
  making a cross-apartment COM call (`MTAThreadDispatchCrossApartmentCall`) into the UI thread's STA:
  ```text
  === THREAD 0x7440 ===
  *** WARNING: Unable to verify checksum for C:\WINDOWS\assembly\NativeImages_v4.0.30319_64\System.Data\436b0645a91cd17436cf6460cc7670b2\System.Data.ni.dll
   # Child-SP          RetAddr               Call Site
  00 000000fa`ac86cda8 00007ffd`db50ab53     ntdll!NtWaitForMultipleObjects+0x14
  01 000000fa`ac86cdb0 00007ffc`87f2d61b     KERNELBASE!WaitForMultipleObjectsEx+0x123
  02 000000fa`ac86d0a0 00007ffd`dc905054     mso!Ordinal9147+0xeb
  03 000000fa`ac86d110 00007ffd`dc883393     combase!MTAThreadWaitForCall+0x104 [onecore\com\combase\dcomrem\channelb.cxx @ 7148]
  04 (Inline Function) --------`--------     combase!MTAThreadDispatchCrossApartmentCall+0x6c4 [onecore\com\combase\dcomrem\chancont.cxx @ 235]
  05 (Inline Function) --------`--------     combase!CSyncClientCall::SwitchAptAndDispatchCall+0xaf7 [onecore\com\combase\dcomrem\channelb.cxx @ 5700]
  06 000000fa`ac86d1c0 00007ffd`dc8d700d     combase!CSyncClientCall::SendReceive2+0xc83 [onecore\com\combase\dcomrem\channelb.cxx @ 5306]
  07 (Inline Function) --------`--------     combase!SyncClientCallRetryContext::SendReceiveWithRetry+0x42 [onecore\com\combase\dcomrem\callctrl.cxx @ 1494]
  08 000000fa`ac86d5e0 00007ffd`dc8d6a83     combase!CSyncClientCall::SendReceiveInRetryContext+0x5d [onecore\com\combase\dcomrem\callctrl.cxx @ 581]
  09 (Inline Function) --------`--------     combase!DefaultSendReceive+0x3c [onecore\com\combase\dcomrem\callctrl.cxx @ 539]
  0a 000000fa`ac86d620 00007ffd`dc8d65c3     combase!CSyncClientCall::SendReceive+0x283 [onecore\com\combase\dcomrem\ctxchnl.cxx @ 803]
  0b (Inline Function) --------`--------     combase!CClientChannel::SendReceive+0x59 [onecore\com\combase\dcomrem\ctxchnl.cxx @ 675]
  0c 000000fa`ac86d790 00007ffd`dd3c1526     combase!NdrExtpProxySendReceive+0xc3 [onecore\com\combase\ndr\ndrole\proxy.cxx @ 1899]
  0d 000000fa`ac86d7e0 00007ffd`dc947f5c     RPCRT4!NdrpClientCall2+0x4e6
  0e 000000fa`ac86de40 00007ffd`dcaa6642     combase!ObjectStublessClient+0x24c [onecore\com\combase\ndr\ndrole\amd64\stblsclt.cxx @ 365]
  0f 000000fa`ac86e1d0 00007ffc`3c6a62f8     combase!ObjectStubless+0x42 [onecore\com\combase\ndr\ndrole\amd64\stubless.asm @ 176]
  10 000000fa`ac86e220 00007ffc`3c6a5b22     Microsoft_Teams_MeetingAddin!DomainBoundILStubClass.IL_STUB_CLRtoCOM()+0xc8
  11 000000fa`ac86e310 00007ffc`3c6a591c     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Model.ApplicationModel.OutlookApplicationModel.GetCurrentUserPrimarySmtpAddress+0x172
  12 000000fa`ac86e370 00007ffc`3c6a56ae     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Telemetry.MsoAriaTelemetry.Provider.MsoAriaTelemetryProvider.GetCurrentSmtpAddress+0x4c
  13 000000fa`ac86e3b0 00007ffc`3c6a5596     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Telemetry.MsoAriaTelemetry.Provider.MsoAriaTelemetryProvider.SetupCommonProperties+0x3e
  14 000000fa`ac86e3e0 00007ffc`3c6a54de     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Telemetry.TelemetryFlowController.SetupCommonProperties+0x66
  15 000000fa`ac86e430 00007ffc`3c6a2b2a     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Telemetry.TelemetryManager.CompleteInitialize+0x3e
  16 000000fa`ac86e460 00007ffc`90ca66a0     Microsoft_Teams_MeetingAddin!Microsoft.Teams.MeetingAddin.Application.<SetupTelemetryAsync>d__75.MoveNext+0x4ca
  17 000000fa`ac86e4e0 00007ffc`90ca6525     mscorlib_ni!System.Threading.ExecutionContext.RunInternal+0x170
  18 000000fa`ac86e5b0 00007ffc`90d43e7f     mscorlib_ni!System.Threading.ExecutionContext.Run+0x15
  19 000000fa`ac86e5e0 00007ffc`90d493d3     mscorlib_ni!System.Runtime.CompilerServices.AsyncMethodBuilderCore.MoveNextRunner.Run+0x6f
  1a 000000fa`ac86e630 00007ffc`90d0fcf0     mscorlib_ni!System.Threading.Tasks.AwaitTaskContinuation.RunOrScheduleAction+0x63
  1b 000000fa`ac86e680 00007ffc`90d402fe     mscorlib_ni!System.Threading.Tasks.Task.FinishContinuations+0x400
  1c 000000fa`ac86e710 00007ffc`61f9bca3     mscorlib_ni!System.Threading.Tasks.Task<bool>.TrySetResult+0x5e
  1d 000000fa`ac86e750 00007ffc`61f9baa4     System_Data_ni+0x79bca3
  1e 000000fa`ac86e7a0 00007ffc`61f9bd51     System_Data_ni+0x79baa4
  1f 000000fa`ac86e7e0 00007ffc`61f9beac     System_Data_ni+0x79bd51
  20 000000fa`ac86e820 00007ffc`90d0fd37     System_Data_ni+0x79beac
  21 000000fa`ac86e870 00007ffc`90d412a4     mscorlib_ni!System.Threading.Tasks.Task.FinishContinuations+0x447
  22 000000fa`ac86e900 00007ffc`90d42077     mscorlib_ni!System.Threading.Tasks.Task.Finish+0x34
  23 000000fa`ac86e960 00007ffc`90d41751     mscorlib_ni!System.Threading.Tasks.Task.ExecuteWithThreadLocal+0x257
  24 000000fa`ac86ea10 00007ffc`90ccc966     mscorlib_ni!System.Threading.Tasks.Task.ExecuteEntry+0xa1
  25 000000fa`ac86ea50 00007ffc`9b921b63     mscorlib_ni!System.Threading.ThreadPoolWorkQueue.Dispatch+0x156
  26 000000fa`ac86eaf0 00007ffc`9b7de34b     clr!CallDescrWorkerInternal+0x83
  27 000000fa`ac86eb30 00007ffc`9b79682a     clr!CallDescrWorkerWithHandler+0x47
  28 000000fa`ac86eb70 00007ffc`9b79654a     clr!MethodDescCallSite::CallTargetWorker+0xfa
  29 000000fa`ac86ec70 00007ffc`9b7f702b     clr!QueueUserWorkItemManagedCallback+0x2a
  2a 000000fa`ac86ed60 00007ffc`9b7f6f4f     clr!ManagedThreadBase_DispatchInner+0x33
  2b 000000fa`ac86eda0 00007ffc`9b7f6e1b     clr!ManagedThreadBase_DispatchMiddle+0x83
  2c 000000fa`ac86ee90 00007ffc`9b91a3e5     clr!ManagedThreadBase_DispatchOuter+0x87
  2d 000000fa`ac86ef20 00007ffc`9bb86688     clr!ManagedThreadBase_DispatchInCorrectAD+0x15
  2e 000000fa`ac86ef50 00007ffc`9baa3b2f     clr!Thread::DoADCallBack+0x3a4
  2f 000000fa`ac86f0e0 00007ffc`9b7f6f4f     clr!ManagedThreadBase_DispatchInner+0x2acb37
  30 000000fa`ac86f120 00007ffc`9b7f6e1b     clr!ManagedThreadBase_DispatchMiddle+0x83
  31 000000fa`ac86f210 00007ffc`9b7f6d8f     clr!ManagedThreadBase_DispatchOuter+0x87
  32 000000fa`ac86f2a0 00007ffc`9b8cdc5a     clr!ManagedThreadBase_FullTransitionWithAD+0x2f
  33 000000fa`ac86f300 00007ffc`9b87e5fd     clr!ManagedPerAppDomainTPCount::DispatchWorkItem+0x9a
  34 000000fa`ac86f470 00007ffc`9b87e3b9     clr!ThreadpoolMgr::ExecuteWorkRequest+0x51
  35 000000fa`ac86f4a0 00007ffc`9b90643a     clr!ThreadpoolMgr::WorkerThreadStart+0xe9
  36 000000fa`ac86f530 00007ffd`dd75dda7     clr!Thread::intermediateThreadProc+0x8a
  37 000000fa`ac86fa70 00007ffd`ddfb9c5c     KERNEL32!BaseThreadInitThunk+0x17
  38 000000fa`ac86faa0 00000000`00000000     ntdll!RtlUserThreadStart+0x2c
  ```
  Counted over `~*kn 8` (top 8 frames of all 321 thread blocks):
  - 41 threads were in `ntdll!RtlpWaitOnCriticalSection`: 31 under `OLMAPI32!HrCreateOfflineObj+0x28b`,
    5 under `OLMAPI32!ReadRegStringValue+0xfbdc`, and 5 elsewhere.
  - Exactly one thread (0x7440) was in `combase!MTAThreadWaitForCall`.
  Representative waiters, verbatim (top 8 frames as captured):
  - 0xF5E8, 0x10C30 and 0xE940: `HrCreateOfflineObj`.
  - 0x12290: an EMSMDB32 lock.
  - 0xA74: `StartSyncLogging` -> `HrCreateServerRequestManager`.
  - 0x13264: `HrCreateOlkAsyncTask` -> `HrCreateConnectionManager`.
  - 0x7038: MSPST32.
  - 0x5C28 and 0x128CC: `ReadRegStringValue`.
  - 0x109F0: a pumping wait.
  - 0x12EF8, 0x55AC, 0xDE6C and 0x1133C: `HrCreateOfflineObj` via EMSMDB32, safe-mode code, free/busy
    and `HrGetGlobalOffline`.
  ```text
     7  Id: 12c58.f5e8 Suspend: 1 Teb: 000000fa`a987f000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`aa5fea48 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`aa5fea50 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`aa5feb50 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`aa5febd0 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`aa5fec10 00007ff7`19dafa98     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`aa5ff300 00007ff7`19daf812     OUTLOOK!FreeDelegatorInfo+0x162b8
  06 000000fa`aa5ff390 00007ff7`19e7823b     OUTLOOK!FreeDelegatorInfo+0x16032
  07 000000fa`aa5ff3c0 00007ff7`19e5efa1     OUTLOOK!HrGetDelegatorInfoSync+0x186eb
    11  Id: 12c58.109f0 Suspend: 1 Teb: 000000fa`a9891000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`aaeff2d8 00007ffc`87f62dc4     win32u!NtUserMsgWaitForMultipleObjectsEx+0x14
  01 000000fa`aaeff2e0 00007ffc`b99787a2     mso!Ordinal7677+0xd4
  02 000000fa`aaeff340 00007ffc`b9a3d311     OLMAPI32!ScRelocNotifications+0x572
  03 000000fa`aaeff660 00007ffc`b9a3d1b2     OLMAPI32!HrGetMessageClassFromContentClassA+0x611
  04 000000fa`aaeff6c0 00007ffc`b9a3ba8a     OLMAPI32!HrGetMessageClassFromContentClassA+0x4b2
  05 000000fa`aaeff700 00007ffc`b9a53488     OLMAPI32!HrCreateServerRequestManager+0x7d2a
  06 000000fa`aaeff780 00007ffd`0f6413ed     OLMAPI32!UlAddRef+0x9b8
  07 000000fa`aaeff7b0 00007ffd`0f5754ee     mso20win32client!Ordinal499+0x21d
    14  Id: 12c58.10c30 Suspend: 1 Teb: 000000fa`a989b000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`ab3fefd8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`ab3fefe0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`ab3ff0e0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`ab3ff160 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`ab3ff1a0 00007ff7`19dafa98     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`ab3ff890 00007ff7`19daf812     OUTLOOK!FreeDelegatorInfo+0x162b8
  06 000000fa`ab3ff920 00007ff7`19f0020b     OUTLOOK!FreeDelegatorInfo+0x16032
  07 000000fa`ab3ff950 00007ff7`19fe7bc0     OUTLOOK!HrBgScheduleRepairApp+0x16ab
    16  Id: 12c58.5c28 Suspend: 1 Teb: 000000fa`a989f000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`ab5fdb98 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`ab5fdba0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`ab5fdca0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`ab5fdd20 00007ffc`b9a320bc     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`ab5fdd60 00007ffd`0f733307     OLMAPI32!ReadRegStringValue+0xfbdc
  05 000000fa`ab5fdda0 00007ffc`b998a075     mso20win32client!Ordinal996+0x37
  06 000000fa`ab5fddd0 00007ffc`b998abfd     OLMAPI32!RPCCanceller_ExtAddRef+0x1315
  07 000000fa`ab5fde30 00007ffc`8bb7a114     OLMAPI32!ExchangePerf_GetCacheRecWrap+0x50d
    60  Id: 12c58.e940 Suspend: 1 Teb: 000000fa`a9907000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`ac66bab8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`ac66bac0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`ac66bbc0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`ac66bc40 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`ac66bc80 00007ffc`b99daa51     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`ac66c370 00007ffc`b9a3f585     OLMAPI32!HrOpenOfflineObj+0xb1
  06 000000fa`ac66c440 00007ffc`b9a3f9ae     OLMAPI32!HrSendHttpRequestWithHeaders+0x85
  07 000000fa`ac66c480 00007ffc`b9a3f537     OLMAPI32!HrSendHttpRequestWithHeadersAndId+0xde
    66  Id: 12c58.7038 Suspend: 1 Teb: 000000fa`a991d000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`ad16f768 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`ad16f770 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`ad16f870 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`ad16f8f0 00007ffc`9b5712bd     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`ad16f930 00007ffc`9b5ba1ba     MSPST32!OSTServiceEntry+0x6afad
  05 000000fa`ad16f960 00007ffc`9b56e16b     MSPST32!MSProviderInit+0xf23a
  06 000000fa`ad16f990 00007ffc`b9a5f239     MSPST32!OSTServiceEntry+0x67e5b
  07 000000fa`ad16fad0 00007ffc`b9a53488     OLMAPI32!HrConvertMAPIErrorEx2+0x3919
   253  Id: 12c58.13264 Suspend: 1 Teb: 000000fa`af972000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`afebf378 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`afebf380 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`afebf480 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`afebf500 00007ffc`b998420b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`afebf540 00007ffc`b99839ec     OLMAPI32!UnregisterOutlookEtwCallback+0x17b
  05 000000fa`afebf570 00007ffc`b9981291     OLMAPI32!HrCreateConnectionManager+0x1d0c
  06 000000fa`afebf5b0 00007ffc`b99817e3     OLMAPI32!HrCreateOlkAsyncTask+0x1bc1
  07 000000fa`afebf640 00007ffc`b9a48c8a     OLMAPI32!HrCreateOlkAsyncTask+0x2113
   254  Id: 12c58.55ac Suspend: 1 Teb: 000000fa`af976000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`b00bed48 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`b00bed50 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`b00bee50 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`b00beed0 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`b00bef10 00007ff7`19dafa98     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`b00bf600 00007ff7`19daf8bf     OUTLOOK!FreeDelegatorInfo+0x162b8
  06 000000fa`b00bf690 00007ff7`19de2393     OUTLOOK!FreeDelegatorInfo+0x160df
  07 000000fa`b00bf710 00007ff7`19de22f0     OUTLOOK!GetOutlookSafeModeState+0x723
   259  Id: 12c58.128cc Suspend: 1 Teb: 000000fa`af988000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`b07bf6f8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`b07bf700 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`b07bf800 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`b07bf880 00007ffc`b9a320bc     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`b07bf8c0 00007ffd`0f733307     OLMAPI32!ReadRegStringValue+0xfbdc
  05 000000fa`b07bf900 00007ffc`b9982b48     mso20win32client!Ordinal996+0x37
  06 000000fa`b07bf930 00007ffc`b9a1c4a1     OLMAPI32!HrCreateConnectionManager+0xe68
  07 000000fa`b07bf980 00007ffc`b9a3ba8a     OLMAPI32!FPropContainsProp+0x4d91
   263  Id: 12c58.12290 Suspend: 1 Teb: 000000fa`af990000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`b0bbd5e8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`b0bbd5f0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`b0bbd6f0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`b0bbd770 00007ffc`8bba66bf     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`b0bbd7b0 00007ffc`8bbf57c3     EMSMDB32+0xc66bf
  05 000000fa`b0bbd7e0 00007ffc`8beb288d     EMSMDB32!MSProviderInit+0x17063
  06 000000fa`b0bbd810 00007ffc`8bdaf7f9     EMSMDB32!ScStatOpen+0xc635d
  07 000000fa`b0bbd850 00007ffc`8bbf250b     EMSMDB32!EMSWizardEntry+0x501e9
   282  Id: 12c58.a74 Suspend: 1 Teb: 000000fa`af9ba000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`aa7ff288 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`aa7ff290 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`aa7ff390 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`aa7ff410 00007ffc`b9a3b9d9     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`aa7ff450 00007ffc`b9c80fe8     OLMAPI32!HrCreateServerRequestManager+0x7c79
  05 000000fa`aa7ff480 00007ffc`b9c87344     OLMAPI32!StartSyncLogging+0x7958
  06 000000fa`aa7ff570 00007ffc`b9a1c4a1     OLMAPI32!StartSyncLogging+0xdcb4
  07 000000fa`aa7ff600 00007ffc`b9a5f239     OLMAPI32!FPropContainsProp+0x4d91
   289  Id: 12c58.12ef8 Suspend: 1 Teb: 000000fa`af9d6000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`b27be8d8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`b27be8e0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`b27be9e0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`b27bea60 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`b27beaa0 00007ffc`b99daa51     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`b27bf190 00007ffc`8bd4b93e     OLMAPI32!HrOpenOfflineObj+0xb1
  06 000000fa`b27bf260 00007ffc`8bd4cae8     EMSMDB32!DllCanUnloadNow+0x1ba6e
  07 000000fa`b27bf2f0 00007ffc`8bd4c6d9     EMSMDB32!DllCanUnloadNow+0x1cc18
   290  Id: 12c58.de6c Suspend: 1 Teb: 000000fa`af9de000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`a9bfeac8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`a9bfead0 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`a9bfebd0 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`a9bfec50 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`a9bfec90 00007ff7`19dafa98     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`a9bff380 00007ff7`19daf812     OUTLOOK!FreeDelegatorInfo+0x162b8
  06 000000fa`a9bff410 00007ff7`1a1c0d10     OUTLOOK!FreeDelegatorInfo+0x16032
  07 000000fa`a9bff440 00007ff7`19f3f606     OUTLOOK!GetFBPublishingInterval+0xf660
   293  Id: 12c58.1133c Suspend: 1 Teb: 000000fa`af9e8000 Unfrozen
   # Child-SP          RetAddr               Call Site
  00 000000fa`aacfecf8 00007ffd`ddf0a61d     ntdll!NtWaitForAlertByThreadId+0x14
  01 000000fa`aacfed00 00007ffd`ddf0bd1f     ntdll!RtlpWaitOnCriticalSection+0x5ad
  02 000000fa`aacfee00 00007ffd`ddf0cf92     ntdll!RtlpEnterCriticalSectionContended+0x1ef
  03 000000fa`aacfee80 00007ffc`b99dad1b     ntdll!RtlEnterCriticalSection+0xf2
  04 000000fa`aacfeec0 00007ffc`b99daa51     OLMAPI32!HrCreateOfflineObj+0x28b
  05 000000fa`aacff5b0 00007ffc`b9a3f585     OLMAPI32!HrOpenOfflineObj+0xb1
  06 000000fa`aacff680 00007ffc`b9acde26     OLMAPI32!HrSendHttpRequestWithHeaders+0x85
  07 000000fa`aacff6c0 00007ffc`b9ace326     OLMAPI32!HrGetGlobalOffline+0x63a6
  ```
  No stack was captured on the MCP side, i.e. thread 0x5D64 in PID 68964. Only its wait-chain line in
  1.6 exists.
  **1.8 The triggering call.** Source: Claude Code transcripts (`.jsonl`, paths in PART 6), searched by
  the diagnosing agent. Lines verbatim from its search output. The first two lines are transcript
  records that only name outlookai tools, presumably when their schemas were loaded. The third is the
  actual `tool_use`.
  ```text
  == C:\Users\<user>\.claude\projects\C--Users-<user>-Downloads\579e5f93-028a-4ea2-a918-38697644e3ad\subagents\agent-a26ab951d6306c1a7.jsonl (modified 16:50:00, 10 outlookai mentions)
     2026-11-15T15:27:33.609Z  tool_use: mcp__outlookai__list_accounts,mcp__outlookai__list_folders,mcp__outlookai__read,mcp__outlookai__save_attachment,mcp__outlookai__search,mcp__outlookai__thread  input:
     2026-11-15T15:27:39.465Z  tool_use: mcp__outlookai__list_accounts,mcp__outlookai__list_folders,mcp__outlookai__read,mcp__outlookai__save_attachment,mcp__outlookai__search,mcp__outlookai__thread  input:
     2026-11-15T15:27:39.467Z  tool_use: mcp__outlookai__list_accounts  input: {
  == 579e5f93-028a-4ea2-a918-38697644e3ad.jsonl  first=2026-11-15T15:22:35.568Z  modified=16:56:36
  == agent-a26ab951d6306c1a7.jsonl  first=2026-11-15T15:27:21.454Z  modified=16:50:00
     [2026-11-15T15:50:00.482Z] [Request interrupted by user for tool use]
  == agent-a4de6896a59e9e3b0.jsonl  first=2026-11-15T15:50:14.382Z  modified=16:56:43
  ```
  - The first search covered every `.jsonl` under `C:\Users\<user>\.claude\projects` modified after
    2026-11-15 16:15. On 2026-11-15 it found exactly one outlookai `tool_use` before the diagnosing
    agent's own: the `list_accounts` above.
  - MCP 68964 serves claude.exe 76924, i.e. session 579e5f93 and its sub-agents, all of which share it.
    Up to the last search, at about 17:06, neither that session's main transcript nor its other
    sub-agents (a4de6896a59e9e3b0, a9750354f40169b80, ad7de4071de9a0e95) held any outlookai
    `tool_use` record.
  - A second search covered every `.jsonl` under `C:\Users\<user>\.claude\projects\C--Users-<user>-Downloads`
    modified after 16:45, for outlookai `tool_use` records at or after 15:45Z. All of them, verbatim
    below, are the diagnosing agent's (a921ba5ff51ee3377), after the fix.
  - **So, as far as these transcripts show, no client of MCP 68964 had a call outstanding at 17:02:02.**
    The one it had was interrupted at 15:50:00Z. That conclusion is an inference from transcripts;
    nothing inside the MCP server was observed.
  ```text
  2026-11-15T16:04:27.029Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__outlook_health
  2026-11-15T16:04:27.038Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__list_accounts
  2026-11-15T16:04:41.325Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:04:41.332Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:04:41.333Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:04:41.975Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:04:42.877Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:05:42.393Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  2026-11-15T16:05:43.148Z  agent-a921ba5ff51ee3377.jsonl  mcp__outlookai__search
  ```
  - The system32 session, from the transcript
    `C:\Users\<user>\.claude\projects\C--WINDOWS-system32\6c9c2c07-aff8-4d3a-b38c-fe8a18e27fd6.jsonl`
    (379,518 bytes):
    - Records: `queue-operation` at 15:27:44.832Z and 15:27:44.834Z; then a `user` record at
      15:27:45.643Z with `"content":"-"`, `"permissionMode":"bypassPermissions"`, `"promptSource":"sdk"`,
      `"entrypoint":"sdk-cli"`, `"cwd":"C:\\WINDOWS\\system32"`, `"version":"2.1.325"`.
    - A `stop_hook_summary` at 15:27:48.087Z.
    - What started it and what it did were NOT investigated.
  - Older outlookai use: the transcript of session 19c2d4d8-e339-4d1b-a4c4-0d75b56f8e30 mentions
    outlookai tools many times on 2026-11-12 and 2026-11-13. For example a `search` at
    2026-11-13T10:33:32.703Z (11:33:32 local, one minute after the first 13 Nov safe-mode prompt).
    Not analysed.
  **1.9 Network List Service check, 16:52:15.** A separate process created the NetworkListManager
  (CLSID DCB00C01-570F-4A9B-8D69-199FDBA5723B) and asked `IsConnected`, `IsConnectedToInternet`,
  `GetConnectivity` and `GetNetworks(3)`, with a 20 s self-timeout. Verbatim:
  ```text
  start 16:52:15.577
  activated in 60 ms
  IsConnected=True IsConnectedToInternet=True Connectivity=66 (83 ms)
    network 'Network' category=1 connected=True internet=True connectivity=66 domainType=0
    network 'Network 3' category=0 connected=False internet=False connectivity=0 domainType=0
    network 'Network 2' category=0 connected=False internet=False connectivity=0 domainType=0
  done in 108 ms
  ```
  **1.10 What was done, with raw output.** All of it was by the diagnosing agent, with the user's
  go-ahead relayed by the coordinating session. Nothing else on the machine was changed: no registry,
  add-in, profile, data file or Office setting, and no `/safe`. Registry keys were only exported,
  as a precaution.
  - 17:02:02, the kill, verbatim:
  ```text
  before: 76888 ["C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE" -Embedding] 11/15/2026 16:27:39
  Stop-Process issued at 17:02:02.374
  76888 gone: True at 17:02:02.402
  OUTLOOK.EXE processes 8s later:
    60368 parent=2228 ["C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE" -Embedding] 11/15/2026 17:02:02
  MCP 68964 alive: True
  ```
  - About 17:02:15, the new PID 60368 checked. Its windows are hidden and none is hung. The second
    `rctrl_renwnd32` with no title is unexplained. No multi-node wait chain existed in 60368 or any MCP
    server. Verbatim, filtered to these lines by the script:
  ```text
  60368 threads=285 cpu=28.328125 ws=389MB mainHwnd=0 title=''
  Token: Elevated=0 ElevationType=3 (1=default,2=full,3=limited) IntegrityRID=0x2000 (0x2000=Medium,0x3000=High,0x4000=System)
  Top-level windows: 32
    hwnd=0x2F1004 vis=False hung=False min=False owner=0x0 rect=(156,156,3036,1719) class='rctrl_renwnd32' title='Inbox - <mailbox-B> - Outlook'
    hwnd=0x670ECE vis=False hung=False min=False owner=0x0 rect=(0,0,136,39) class='rctrl_renwnd32' title=''
  --- wait chains (non-trivial) for Outlook main + MCP servers:
  ```
  - 17:02:54, the shell launch. `Start-Process explorer.exe '"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"'`
    was used so that the launch is the same as a Start-menu click and runs outside the agent's job
    object. Verbatim:
  ```text
  shell launch issued at 17:02:54.368
  OUTLOOK.EXE processes now:
    60368 parent=2228 ["C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE" -Embedding] 11/15/2026 17:02:02
  visible windows of 60368 at 17:03:01:
    hwnd=0x5A0EE6 vis=True hung=False min=False owner=0x0 rect=(-8,-8,3848,2120) class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook'
  ```
  **1.11 The second cold start, 67 ms after the kill.** Thread start times of every OutlookAI MCP
  server inside the window 17:01:55 to 17:02:10, verbatim. The format is
  `<pid> (claude <parent pid>) started <day hh:mm:ss>: <count> <tid>@<time>`.
  ```text
   26812 (claude  35756) started 13 17:22:19: threads started 17:01:55-17:02:10 = 0
   31340 (claude  49344) started 13 23:10:24: threads started 17:01:55-17:02:10 = 0
   32724 (claude   5912) started 13 17:22:08: threads started 17:01:55-17:02:10 = 0
   37340 (claude  52100) started 14 18:51:20: threads started 17:01:55-17:02:10 = 0
   65532 (claude  64472) started 13 17:46:34: threads started 17:01:55-17:02:10 = 0
   67204 (claude  78308) started 15 16:44:09: threads started 17:01:55-17:02:10 = 0
   68964 (claude  76924) started 15 16:21:10: threads started 17:01:55-17:02:10 = 3 75728@17:02:02.469 27600@17:02:09.738 73948@17:02:09.778
   73176 (claude  73536) started 14 18:56:58: threads started 17:01:55-17:02:10 = 0
  ```
  - Kill issued at 17:02:02.374 and the process gone by 17:02:02.402. MCP 68964's thread 75728 started
    at 17:02:02.469, 67 ms after the process was gone. OUTLOOK.EXE 60368 was created by DcomLaunch in the
    same second.
  - OutlookAI's reconcile timestamps from 1.13 are 16:02:03.197Z and 16:02:03.251Z.
  - The client of the original call had been interrupted at 16:50:00 (1.8), and no tool call was in
    flight (1.8).
  - Which code path in 68964 did this was NOT determined. `C:\Users\<user>\AppData\Local\OutlookAI\audit.log`
    was last written 2026-11-13 20:21:47, and no `*.log`, `*.txt`, `*.json` or `*.jsonl` under
    `%LOCALAPPDATA%\OutlookAI` was written after 16:20 on 2026-11-15.
  **1.12 Verification.** Source: the diagnosing agent's script. It sends `WM_NULL` with
  `SendMessageTimeout(SMTO_ABORTIFHUNG, 5000)`, reads the status bar through UI Automation, and lists
  the OST/NST `LastWriteTime` values. Sizes print with the Dutch thousands separator, so
  `24.273 MB` is 24,273 MB. Round 1, verbatim:
  ```text
  --- round 1 at 17:03:39
  process: threads=272 cpu=49.4s ws=411MB
    hwnd=0x5A0EE6 class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook' hung=False ; WM_NULL replied in 0 ms
      status bar: Items in View 288 | Unread Items in View 12 | Progress All folders are up to date. | Connectivity to your server. Connected to: Microsoft Exchange | Normal View. Show All Pinned Panes. | Reading View. Hide All Pinned Panes. | Zoom Out | Zoom | Zoom In | Zoom 10%
  --- round 2 at 17:03:49
  process: threads=274 cpu=49.6s ws=412MB
    hwnd=0x5A0EE6 class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook' hung=False ; WM_NULL replied in 0 ms
      status bar: Items in View 288 | Unread Items in View 12 | Progress All folders are up to date. | Connectivity to your server. Connected to: Microsoft Exchange | Normal View. Show All Pinned Panes. | Reading View. Hide All Pinned Panes. | Zoom Out | Zoom | Zoom In | Zoom 10%
  --- round 3 at 17:03:59
  process: threads=273 cpu=49.6s ws=412MB
    hwnd=0x5A0EE6 class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook' hung=False ; WM_NULL replied in 0 ms
      status bar: Items in View 288 | Unread Items in View 12 | Progress All folders are up to date. | Connectivity to your server. Connected to: Microsoft Exchange | Normal View. Show All Pinned Panes. | Reading View. Hide All Pinned Panes. | Zoom Out | Zoom | Zoom In | Zoom 10%
  --- OST/NST files (LastWriteTime shows whether Outlook is writing synced data):
    17:03:59    24.273 MB  <mailbox-B>.ost
    17:03:35     9.433 MB  <mailbox-A>.ost
    17:03:32        16 MB  <mailbox-C>.ost
    17:02:32        16 MB  <mailbox-C>.nst
    17:02:31        16 MB  <mailbox-A>.nst
    17:02:31        16 MB  <mailbox-B>(2).nst
    16:46:41        16 MB  <mailbox-D>.ost
    16:30:49        16 MB  <mailbox-D>.nst
    01:45:37        16 MB  <mailbox-B>.nst
  ```
  Round 2, after the outlookai tool calls and the COM reads, verbatim:
  ```text
  --- round 1 at 17:08:19
  process: threads=272 cpu=52.5s ws=415MB
    hwnd=0x5A0EE6 class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook' hung=False ; WM_NULL replied in 0 ms
      status bar: Items in View 288 | Unread Items in View 12 | Progress All folders are up to date. | Connectivity to your server. Connected to: Microsoft Exchange | Normal View. Show All Pinned Panes. | Reading View. Hide All Pinned Panes. | Zoom Out | Zoom | Zoom In | Zoom 10%
  --- round 2 at 17:08:34
  process: threads=273 cpu=52.6s ws=415MB
    hwnd=0x5A0EE6 class='rctrl_renwnd32' title='Inbox - <mailbox-A> - Outlook' hung=False ; WM_NULL replied in 0 ms
      status bar: Items in View 288 | Unread Items in View 12 | Progress All folders are up to date. | Connectivity to your server. Connected to: Microsoft Exchange | Normal View. Show All Pinned Panes. | Reading View. Hide All Pinned Panes. | Zoom Out | Zoom | Zoom In | Zoom 10%
  --- OST/NST files (LastWriteTime shows whether Outlook is writing synced data):
    17:07:41     9.433 MB  <mailbox-A>.ost
    17:07:41    24.273 MB  <mailbox-B>.ost
    17:07:41        16 MB  <mailbox-C>.ost
    17:05:22        16 MB  <mailbox-B>(2).nst
    17:05:22        16 MB  <mailbox-C>.nst
    17:04:51        16 MB  <mailbox-A>.nst
    16:46:41        16 MB  <mailbox-D>.ost
    16:30:49        16 MB  <mailbox-D>.nst
    01:45:37        16 MB  <mailbox-B>.nst
  ```
  **1.13 outlook_health and list_accounts after the fix, 17:04:27.** Tool output verbatim, except for
  the placeholder substitutions.
  ```text
  {"status":"ok","outlook":{"running":true,"headless":false,"version":"16.0.14334.20974","installerMutexHeld":false,"comConnected":true,"storesReachable":7,"stores":["<mailbox-C>","<store-D-online>","<delegate-mailbox-1>","<delegate-mailbox-2>","<mailbox-B>","<mailbox-D>","<mailbox-A>"]},"index":{"provider":"OleDb","newestIndexedUtc":"2026-11-15T16:02:55Z","ageMinutes":1.5396681216666666,"perStore":[{"store":"<mailbox-A>","newestIndexedUtc":"2026-11-15T15:49:02Z"},{"store":"<mailbox-B>","newestIndexedUtc":"2026-11-15T16:02:55Z"},{"store":"<mailbox-C>","newestIndexedUtc":"2026-11-15T16:02:25Z"},{"store":"<mailbox-A>","newestIndexedUtc":"2026-11-13T10:06:09Z"}],"wSearchStartMode":"automatic","indexerProcessRunning":true},"audit":{"path":"C:\\Users\\<user>\\AppData\\Local\\OutlookAI\\audit.log","writable":true},"tuning":{"managed":true,"enabled":true,"searchEnabled":true,"cachingEnabled":true,"ostEnabled":true,"restartNeeded":false,"lastReconcileUtc":"2026-11-15T16:02:03.1967660Z","uiSearchBackend":"local"},"registration":{"status":"ok","runningFrom":"C:\\Users\\<user>\\AppData\\Local\\OutlookAI\\Setup\\McpServer\\OutlookAI.McpServer.exe","registeredCommand":"${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe","addInStatus":"ok","addInHealed":false,"addInLastReconcileUtc":"2026-11-15T16:02:03.2510478Z","addInResolvedServerPath":"C:\\Users\\<user>\\AppData\\Local\\OutlookAI\\Setup\\McpServer\\OutlookAI.McpServer.exe"},"advice":["Index is current (newest mail 1.5 min ago); searches run at index speed."]}
  {"accounts":[{"smtpAddress":"<mailbox-A>","displayName":"<mailbox-A>","deliveryStore":"<mailbox-A>"},{"smtpAddress":"<mailbox-B>","displayName":"<mailbox-B>","deliveryStore":"<mailbox-B>"},{"smtpAddress":"<mailbox-C>","displayName":"<mailbox-C>","deliveryStore":"<mailbox-C>"},{"smtpAddress":"<mailbox-D>","displayName":"<mailbox-D>","deliveryStore":"<store-D-online>"}],"stores":[{"displayName":"<mailbox-C>","isDelegate":false,"isCachedExchange":true,"exchangeStoreType":4,"onlineOnly":false,"locallySearchable":true,"inLocalIndex":true},{"displayName":"<store-D-online>","isDelegate":false,"isCachedExchange":false,"exchangeStoreType":1,"onlineOnly":true,"locallySearchable":false,"inLocalIndex":false},{"displayName":"<delegate-mailbox-1>","isDelegate":true,"isCachedExchange":false,"exchangeStoreType":1,"onlineOnly":false,"locallySearchable":true,"inLocalIndex":true},{"displayName":"<delegate-mailbox-2>","isDelegate":true,"isCachedExchange":false,"exchangeStoreType":1,"onlineOnly":false,"locallySearchable":true,"inLocalIndex":true},{"displayName":"<mailbox-B>","isDelegate":false,"isCachedExchange":true,"exchangeStoreType":0,"onlineOnly":false,"locallySearchable":true,"inLocalIndex":true},{"displayName":"<mailbox-D>","isDelegate":false,"isCachedExchange":true,"exchangeStoreType":4,"onlineOnly":true,"locallySearchable":false,"inLocalIndex":false},{"displayName":"<mailbox-A>","isDelegate":false,"isCachedExchange":true,"exchangeStoreType":4,"onlineOnly":false,"locallySearchable":true,"inLocalIndex":true}]}
  ```
  The outlookai `search` tool (seven calls, 17:04:41 to 17:05:43) also worked against the healthy
  instance. No `read` call was made. The search results are left out (mailbox content). One `search` with a folder scope reported in
  its `advice`: "Freshness sweep covered no folder: 'Calendar' could not be opened in Outlook, so
  results for it are index-only". That message is quoted as seen and was not investigated.
  **1.14 Other observations.**
  - **ROT.** Against the healthy PID 60368, between 17:05 and 17:08,
    `[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')` in Windows PowerShell 5.1
    failed. `New-Object -ComObject Outlook.Application` then attached to 60368 without starting a second
    process. The script only did this while PID 60368 was the only OUTLOOK.EXE, and checked that before
    and after. Verbatim (the lines listing calendars are omitted):
  ```text
  OUTLOOK before: 60368
  GetActiveObject: Exception calling "GetActiveObject" with "1" argument(s): "Operation unavailable (Exception from HRESULT: 0x800401E3 (MK_E_UNAVAILABLE))"
  attached via CreateObject; OUTLOOK processes after attach: 60368
  OUTLOOK after: 60368
  ```
  - **The coordinating session's observation**, in its own words: at 17:11:24 on 2026-11-15 it confirmed
    that PID 60368 (`OUTLOOK.EXE -Embedding`, started 17:02:02) was the only Outlook process, with
    Responding=True and MainWindowTitle "Inbox - <mailbox-A> - Outlook".
  - **The user's symptom.** The user only said "why is my outlook not starting?" (15:44:15Z). No error
    text was given. No second OUTLOOK.EXE was seen during the diagnosis, and clicks are not logged, so
    it is not observed what happened on the user's clicks (see PART 2).
  - **Identity.** The diagnosing agent ran as the same user at Medium integrity (1.1), like Outlook and
    every MCP server whose token was checked (1.3).
  **1.15 What OutlookAI's source showed.** This is the working tree `C:\Source\SixFive7\OutlookAI` at
  HEAD `b5da70f` (clean). It was read only to find the COM activation, so nothing else was read.
  `McpServer/OutlookAI.Core/Com/OutlookComSession.cs`, verbatim excerpts:
  ```csharp
  336        /// <summary>
  337        /// Connects to Outlook via the COM object model. Outlook is a single-instance COM
  338        /// server: creating Outlook.Application attaches to the running instance, or starts
  339        /// OUTLOOK.EXE when none is running (allowed per S7/D17;
  340        /// <paramref name="allowStartingOutlook"/>=false refuses instead). Never restarts
  341        /// or closes an existing instance.
  342        /// </summary>
  343        public static OutlookComSession Connect(bool allowStartingOutlook = true, Action<OutlookComSession>? onOutlookGone = null)
  360            PumpedStaRunner runner = new PumpedStaRunner("OutlookAI.ComGateway.Sta");
  361            OutlookComSession session = new OutlookComSession(runner, startedOutlook: !wasRunning);
  365                runner.Run(() =>
  366                {
  367                    Type progIdType = Type.GetTypeFromProgID("Outlook.Application")
  368                        ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
  369                    session._application = Activator.CreateInstance(progIdType)
  370                        ?? throw new InvalidOperationException("Failed to create Outlook.Application.");
  371                    dynamic app = session._application;
  372                    dynamic ns = app.GetNamespace("MAPI");
  373                    session._namespace = (object)ns;
  374                    try
  375                    {
  376                        // Binds the default profile when Outlook was cold-started by us; a
  377                        // no-op/benign failure when a session is already active.
  378                        ns.Logon(Type.Missing, Type.Missing, false, false);
  379                    }
  401                    // D49 THE PROCESS PIN. A window-less Outlook EXITS the moment the only
  402                    // Inspector closes (Phase-1 row 1) - which is the compose path's own
  403                    // last step, so every headless draft was killing the instance it had
  404                    // just used, taking update_draft (com_failure) and three live-suite
  405                    // collections (RPC_S_SERVER_UNAVAILABLE) with it. A NON-DISPLAYED
  406                    // Explorer owns an invisible window that keeps the process alive; it
  407                    // shows nothing, leaves Process.MainWindowHandle zero (so outlook_health
  408                    // still reports headless), and does not cost the user promotability -
  409                    // launching outlook.exe with the pin held still opens a normal window.
  410                    session.EnsureComposeSurfacePin();
  416                session.WireProcessExitWatch();
  459        private void OnWatchedProcessExited(object? sender, EventArgs e)
  460        {
  461            SignalOutlookGone(GoneByProcessExit);
  462        }
  ```
  `McpServer/OutlookAI.Core/Com/ComGateway.cs`, verbatim excerpts:
  ```csharp
  271            catch (Exception ex) when (IsDisconnectException(ex))
  272            {
  273                // Same rule as the no-recovery overload: the dead session is dropped, the
  274                // work is never re-run. A budgeted operation is a multi-call lambda by
  275                // assumption, and replaying one replays the steps that already succeeded.
  276                Invalidate(session);
  277                throw;
  278            }
  290        /// <summary>
  291        /// Returns the live session, pinging a held one first and reconnecting when the
  292        /// ping fails. Serialized: concurrent callers queue here while a cold Outlook
  293        /// start is in progress.
  294        /// </summary>
  302            lock (_lock)
  303            {
  304                if (_session != null)
  305                {
  306                    try
  307                    {
  308                        _session.GetProfileName();
  309                        return _session;
  310                    }
  311                    catch (Exception ex) when (IsSessionUnusable(ex))
  312                    {
  313                        // Outlook went away (it is never stopped by us - user action or
  314                        // crash). Drop the dead session and reconnect below.
  ```
  Other lines a `rg` for `GetActiveObject|Outlook.Application|CoCreateInstance|Activator.CreateInstance|...`
  over `McpServer\*.cs` showed. These are pointers only and were not read further:
  ```text
  McpServer\OutlookAI.ComHost\Supervision\ComHostPolicy.cs:140:        /// Budget for establishing the COM session (which may cold-start OUTLOOK.EXE).
  McpServer\OutlookAI.ComHost\Supervision\ComHostPolicy.cs:345:        /// (<c>OUTLOOK.EXE -Embedding</c>, parent svchost - i.e. COM activation), and the
  McpServer\OutlookAI.ComHost\Supervision\ComHostSupervisor.cs:326:            // session, which may cold-start OUTLOOK.EXE, so it gets a wider floor than the
  McpServer\OutlookAI.ComHost\Supervision\ComHostSupervisor.cs:876:        /// WHAT THE KILL DOES NOT REACH: OUTLOOK.EXE. It is not in the tree, established
  McpServer\OutlookAI.ComHost\Supervision\ComHostSupervisor.cs:880:        /// (<c>OUTLOOK.EXE -Embedding</c>, parent <c>svchost.exe</c>, see
  McpServer\OutlookAI.Core\Com\ComOperationBudgets.cs:83:        /// Deadline for establishing the COM session, which may cold-start OUTLOOK.EXE.
  McpServer\OutlookAI.Core\Services\HealthReporting.cs:278:        /// Whether the running OUTLOOK.EXE is HEADLESS - no main window, tray icon only
  McpServer\OutlookAI.RemediationTools\ComStaRunner.cs:16:    /// <c>Outlook.Application</c> can COLD-START OUTLOOK.EXE, which is slow rather than
  McpServer\OutlookAI.McpServer.Tests\T2\LiveDisconnectRecoveryTests.cs:469:        // most likely to block without ever returning: it cold-starts OUTLOOK.EXE through
  McpServer\OutlookAI.McpServer.Tests\T2\LiveDisconnectRecoveryTests.cs:470:        // Activator.CreateInstance on the Outlook.Application ProgID, which has no timeout
  ```
  The `search` tool's own description says: "the server sweeps it live through Outlook (autostarting it
  headless when needed)". The `read` tool's says: "Needs Outlook (starts it if allowed)". The
  `outlook_health` tool's says: "attaches to Outlook only when it is already running, NEVER starts it".
  **PART 2 - INFERENCES (the diagnosing agent's reading of the causal chain; guesses, not proven).**
  - **2.1 Launcher.** MCP 68964 cold-started PID 76888 through `OutlookComSession.Connect`'s
    `Activator.CreateInstance` (line 369) while serving `list_accounts`. The evidence is the same-second
    timing, the `-Embedding` command line with a DcomLaunch parent, MCP thread 0x5D64 created in that
    second (probably the `PumpedStaRunner` STA "OutlookAI.ComGateway.Sta") and the code. Whether the
    installed build used this in-process path or the ComHost child (`OutlookAI.ComHost`) was not
    determined.
  - **2.2 The outer call.** The late-bound `IDispatch::Invoke` at the bottom of the UI thread's stack
    came from MCP 68964's connect sequence, most likely `ns.Logon(...)` (line 378) or the hidden-Explorer
    pin (line 410). It was probably cross-process. The only in-process cross-apartment caller found
    (0x7440) is the inner call, although that count used top-8 frames, so a deeper in-process caller
    would be missed. MCP 68964 was the only process found waiting on Outlook. The hidden Explorer
    `Inbox - <mailbox-B> - Outlook` existing in 76888 suggests the pin had at least started.
    Which of the two it was cannot be told from public symbols.
  - **2.3 The mechanism.**
    1. Serving that call, Outlook opened the Exchange offline store on its UI thread and took a lock in
       `OLMAPI32!HrCreateOfflineObj` (31 threads later queued there).
    2. Inside that, it made a cross-process COM call to netprofm (`GetCost`, plausibly for metered-network
       handling).
    3. An STA waiting on an outgoing COM call keeps dispatching incoming calls, so the Teams add-in's
       cross-apartment call from 0x7440 ran re-entrantly on top of the outer frame.
    4. That inner call needed the MAPI async-task singleton (`StartSyncLogging` ->
       `HrCreateSingletonQueueOlkAsyncTask`). It waited for it in a non-pumping
       `WaitForMultipleObjectsEx`, and whatever it waited for needed the lock or the UI thread the outer
       frame holds.
    5. Neither frame can finish, so Outlook deadlocked itself.
    6. The netprofm call itself had probably completed: netprofm answered a fresh client in 108 ms. WCT
       kept naming PID 3240 only because the outer frame was still registered as an outgoing call.
    The exact object of the inner wait was not identified.
  - **2.4 Why the cold start matters.**
    - In this `-Embedding` start the add-ins were loaded (16:27:41) before anything opened the stores.
      The first object-model call from outside then made Outlook open them while the Teams add-in's
      background telemetry was already calling into the object model.
    - In a normal start from the shell, Outlook opens the profile and stores before it loads add-ins.
      This is general knowledge and was NOT verified here.
    - The deadlock needs three things at once: an external call that opens stores on the UI thread, a
      cross-process call inside it (GetCost) that opens a re-entrancy window, and an in-process
      object-model call (Teams) landing in that window. That makes it timing-dependent. The second cold
      start by the same MCP at 17:02:02 did NOT deadlock.
  - **2.5 The user-visible symptom.** OUTLOOK.EXE is single-instance. A click on the shortcut starts a
    process that hands its request to the running instance and exits. Here the running instance was
    window-less and deadlocked, so nothing appeared. It does not matter that the user's clicks were not
    observed. The same hand-off is how the shell launch at 17:02:54 turned the healthy 60368 into a
    visible Outlook.
  - **2.6 The second cold start.** When 76888 died, MCP 68964 still had a COM call outstanding, from the
    request its client abandoned at 16:50:00. That call failed, and something in 68964 reconnected and
    cold-started Outlook again 67 ms later without any client asking. Candidates:
    - the abandoned request's retry;
    - `GetOrConnect` after `Invalidate`;
    - the process-exit watch (`OnWatchedProcessExited` -> `SignalOutlookGone`) and its consequences.
    The reconcile timestamps at 17:02:03 fit a fresh connect.
  - **2.7 Earlier days.** The 13 Nov hang (10:38:01), the two "couldn't start last time" prompts (11:32
    and 11:41) and 13 Outlook boots that day may be the same mechanism, given outlookai use that day. No
    dump exists, so this is unproven.
  - **2.8 Not involved:** elevation, new Outlook, an Office update, add-in disabling, network outage, and
    netprofm being down. See PART 3.
  **PART 3 - RULED OUT, AND HOW.**
  - **Elevation or integrity mismatch.** OUTLOOK.EXE 76888 and 60368 and MCP servers 68964, 67204, 32724
    and 65532 were all Medium, not elevated (1.3).
  - **New Outlook, or the classic/new toggle redirecting.** olk.exe is not registered for the user, the
    toggle is hidden, and all shortcuts target classic OUTLOOK.EXE (1.1).
  - **An Office update installing or pending.** LastScenario UPDATE Success, UpdatesReadyToApply 0, no
    UpdateToVersion, ClickToRunSvc running, OUTLOOK.EXE dated 21-10-2026 (1.1).
  - **A crash.** No Application Error 1000 or WER crash for OUTLOOK.EXE or olk.exe in 7 days (1.4). The
    process was alive.
  - **A busy loop.** 0.000 s CPU in 6 s, and no thread used CPU (1.3).
  - **netprofm hung for everyone.** A new client got answers in 108 ms (1.9).
  - **Outlook's resiliency having disabled or flagged an add-in.** No `DisabledItems`, and
    `CrashingAddinList` is empty (1.1).
  - **No network.** NCSI shows Internet capability since 13 Nov 11:14:23 (1.4), and the probe says
    IsConnectedToInternet=True (1.9).
  - **Another external process calling into Outlook.** WCT over all eight MCP servers and explorer.exe
    found only 68964 (1.6). Other processes were not scanned.
  **PART 4 - NOT CHECKED, AND WHY.**
  - **Which object-model call the outer frame was** (Logon, the pin, or another). Office has only public
    symbols, so there are no real function names.
  - **The MCP side's stack** (68964 thread 0x5D64). No debugger was attached to the MCP process.
  - **netprofm's own threads, and whether GetCost had returned.** As a non-admin, WCT got `PidOnly` for
    the LocalService svchost.
  - **The object the inner `mso!Ordinal9147` wait waited for,** and which thread would have signalled
    it. That needs a full dump and more analysis than was done.
  - **The integrity of MCP servers 26812, 31340, 37340 and 73176,** and the session ids of claude.exe
    5912, 35756, 64472, 49344, 52100 and 73536.
  - **The installed OutlookAI version,** its source revision, and which connect path (in-process
    `ComGateway` or `ComHost`) it runs.
  - **Why no OutlookAI or <custom-addin> assembly appears in the module list** while event 45
    says both loaded.
  - **What inside 68964 cold-started Outlook at 17:02:02.469.** No log covers it (1.11).
  - **Whether an MCP cancellation from the client** (the interrupt at 16:50:00) reaches the COM work.
  - **The origin of the `C:\WINDOWS\system32` Claude SDK session** that started 0.1 s after the UI thread
    blocked.
  - **Whether the 13 Nov hang had the same stack.** WER kept no dump: `Report Status: 97`, `Cab Id: 0`.
  - **The Windows Update and Setup event logs.** They were exported but not analysed, because the cause
    was found elsewhere.
  - **Windhawk's active mods.**
  - **What the user actually clicked,** and when.
  - **How 60368 behaves later,** e.g. after the user closes its window while the hidden-Explorer pin
    holds it.
  - **A reproduction.** None was attempted.
  **PART 5 - DIRECTIONS.** These are guesses, made WITHOUT knowing how OutlookAI works or why it does
  what it does, and none is recommended. The OutlookAI agent makes the trade-offs, and may well find
  that some of these contradict decisions already taken (S7, D17 and D49 appear in the code comments
  above).
  - **(A) No COM cold start from the MCP server.** When no OUTLOOK.EXE runs, either refuse with a message
    asking for Outlook to be started (the code already has `allowStartingOutlook = false` and the
    message "Outlook is not running and starting it was not allowed."), or start Outlook as the shell
    does. That means `OUTLOOK.EXE`, not COM, and attaching only once its main window exists and
    answers, e.g. `SendMessageTimeout(WM_NULL)`. Guessed trade-offs:
    - the first call after Outlook was closed gets slower, or fails;
    - a visible Outlook window appears on the user's desktop;
    - it reverses what the code calls "allowed per S7/D17" and the tools' "autostarting it headless";
    - the test tiers, which seem built around cold starts, may need changes.
  - **(B) Keep the cold start, but wait for Outlook to finish starting before any MAPI work.** Do not call
    `Logon` or create the pin straight after `CreateInstance`, and keep every call on the one STA.
    Also make work stop when its MCP client cancelled or went away, so an abandoned request cannot
    reconnect and cold-start Outlook again (17:02:02). Guessed trade-offs:
    - there may be no reliable "finished starting" signal for an `-Embedding` start;
    - add-ins' own background threads (Teams) still run during the start, so the race may only get
      narrower;
    - more waiting, and the cancellation plumbing is unknown.
  - **(C) Use the Outlook side.** OutlookAI's own VSTO add-in runs inside Outlook; it loaded at
    16:27:41 in 656 ms. It might tell the server when Outlook has finished starting and its stores are
    open, so the server never calls earlier. Or the user disables or defers the Teams Meeting Add-in,
    which removes this particular re-entrant caller. Guessed trade-offs:
    - add-in-to-server signalling may not exist and would be new machinery;
    - disabling Teams costs the user the Teams meeting button;
    - other add-ins could re-enter the same way;
    - another vendor's add-in is the user's configuration, not the product.
  - **(D) Make sure Outlook is already running,** e.g. started at login, so the server never cold-starts
    it. OutlookAI could document this, or tell the agent to ask the user to start Outlook when it is
    not running. Guessed trade-offs:
    - no code fix;
    - it depends on the user;
    - a closed Outlook brings the risk back.
  - **(E) Detect and recover.** The server knows when it started Outlook itself
    (`startedOutlook: !wasRunning`). It could watch that instance, and if its windows stay hung or the
    first calls never return within N seconds, end only that instance and start Outlook from the shell
    instead. Guessed trade-offs:
    - it contradicts the code's rule that Outlook "is never stopped by us";
    - a slow but healthy start could be killed;
    - a kill during store open may cost a resync and an "Outlook couldn't start last time... safe mode?"
      prompt, as on 13 Nov;
    - more machinery.
  - **(F, additional) Report it.** `outlook_health` could say "Outlook is running but hung", e.g. a
    window-less instance whose windows report hung or whose COM ping times out, with advice. Agents
    could then tell the user instead of blocking. Guessed trade-off: diagnosis only, no prevention.
  Questions that may matter before choosing (also guesses):
  - Which connect path does the installed build run?
  - What did S7, D17 and D49 decide, and why?
  - Does an MCP cancellation reach the COM work?
  - What exactly reconnected at 17:02:02.469?
  - Can this be reproduced on a test guest with the Teams Meeting Add-in installed, by cold-starting
    through COM and calling `Logon` and creating the pin immediately, several times?
  - Do the eight long-lived MCP servers (one per open Claude session, some two days old) each hold a
    session to Outlook, and does that matter?
  **PART 6 - POINTERS** (optional; they may no longer exist).
  - The real value behind every placeholder: `.work/outlook-hang-2026-11-15-identifiers.md` in this
    clone. It is local only and never committed; ask the user if it is missing.
  - Evidence folder: `C:\Users\<user>\Downloads\tmp-outlook-diag\`. It holds:
    - `cdb-stack-76888.txt` (all 3,700 lines of the stacks);
    - `waitchain-outlook-76888.txt` (all 317 threads' chains);
    - `waitchain-mcp-explorer-netprofm.txt`;
    - `outlook-76888-modules.txt` (all 255 modules);
    - `events-*.txt` (the exported logs);
    - `nlmprobe-out.txt`;
    - `verify-1.txt` and `verify-2-after-calendar-reads.txt`;
    - `changes.md`;
    - `reg-backup\*.reg`;
    - `scripts\*.ps1` (`procinfo`, `waitchain`, `nlmprobe`, `verify`, `events`).
    The `sym\` subfolder holds the downloaded public symbols.
  - Transcripts, under `C:\Users\<user>\.claude\projects\`:
    - the trigger: `C--Users-<user>-Downloads\579e5f93-028a-4ea2-a918-38697644e3ad.jsonl` and
      `...\579e5f93-028a-4ea2-a918-38697644e3ad\subagents\agent-a26ab951d6306c1a7.jsonl`;
    - the diagnosis: `C--Users-<user>-Downloads\fd2ffacb-6643-4a2c-afb0-8ead426fbd94.jsonl` and
      `...\fd2ffacb-6643-4a2c-afb0-8ead426fbd94\subagents\agent-a921ba5ff51ee3377.jsonl`;
    - the system32 session: `C--WINDOWS-system32\6c9c2c07-aff8-4d3a-b38c-fe8a18e27fd6.jsonl`;
    - the earlier outlookai use: `C--Users-<user>-Downloads\19c2d4d8-e339-4d1b-a4c4-0d75b56f8e30.jsonl`.

- [ ] **Fold every duplicated piece of code into one copy, and let the add-in use OutlookAI.Core
  (Q154, decided 2026-11-17).** The maintainer: *"Q154 b - refactor everything to bring this pattern
  everywhere relevant. I want as much code shared as possible. And as a rule no code duplicated."*
  (`AGENTS.md`, Shared code; `QUESTIONS.md` decision log, 2026-10-11 and 2026-11-17). The add-in
  references Core's net48 build since stage S1 (Q167 (a)), which ended the "Phase-6 no-coupling
  choice" (the gitignored v3.MD, lines 64, 79-80, 144, 184 and 264): both add-in build scripts
  restore Core's packages, the solution holds Core so that F5 builds it, and the add-in may use one
  part of Core, the namespace `OutlookAI.Core.AddIn`, which nothing in it may use to attach to
  Outlook, start a COM apartment thread or open the index (Q164 (a); `check-pinned-constants.ps1`
  check 15). It uses none of Core yet. Q162's signature choice is built on that reference, and Q145's
  writer of an account's offline period moves into that namespace for the same reason. Then the
  copies go. A read-only
  survey of master 9f4646d, made 2026-11-17, found 73 groups: the add-in against Core (A01-A12, the
  signature catalog and the signature and draft-region bookmark code among them), the add-in's
  linked source files (L01-L13), the server's own repetition (S01-S14), test doubles and helpers
  (T01-T16), PowerShell helpers restated across scripts (P01-P13, of which P13 lists same-named
  functions that are NOT duplicates) and values mirrored across languages (X01-X05). Each group has
  its copies, how they diverged, a target home, its blockers and a stage of the survey's refactor
  plan. It is in `.work\dup-survey-2026-11\` of the maintainer's checkout - README.txt,
  inventory.tsv (one row a group), members.tsv (every copy's file and lines) and the detectors' raw
  output - which is gitignored, so not in a clone. How far, and in what order, was decided on
  2026-11-18 as Q163 to Q171.
  Folded so far, each with its CHANGELOG entry, into `McpServer/OutlookAI.McpServer.Tests/Support/`:
  the survey's T01, T02 and T04 (the stand-ins for Outlook's connection and for the search index),
  T07 (repository paths), T08 (readers of compiled code and source text; ComposeOnReopenedDraftTests and
  ReadOnlyFolderLookupTests keep member cutters with declaration rules of their own), T09, T10 to
  T12 (the live tier's helpers, audit lines and tool-surface readers, with the test tags and the
  mailer's tag checks; streams adding new inline tag checks fold theirs when they land) and T05
  (the sweep and scan suites' one-mailbox profile, `AliceMailbox`, and the thread suites' walks,
  `ThreadWalks`; six more suites - the thread, unindexed-store and folder-walk ones - still spell
  that mailbox's SID, prefix and frontier under names of their own, beside stores of their own).
  PowerShell, into `Testbed/guest/ScriptHelpers.ps1` (decision 215 (a)), held there by
  check-powershell-51 check 4: P01's `Invoke-NativeCommand` for 13 of its 17 copies, P09's
  `Format-Seconds`, the build-VM runner's copy of which rounded 119.6 s to "1 m 60 s", P02's
  `Invoke-Logged`, `Show-LogTail` and `Format-Argument`, P12's `Test-IsElevated`, and P03's `Say` for
  the 18 copies that print the time and append to the script's log (the SDK installer, which travels
  alone, keeps an identical copy), beside `Write-ReportLine` for
  the five scripts that collect a report (the mail sink's installer and the build-VM runner keep
  their different ones as `Write-SinkLog` and `Write-RunLog`); P11's module
  import is `Testbed/host/OwnEditionModules.ps1`, which the `Tools/` scripts dot-source too.
  Invoke-LiveTierOnGuest.ps1 and Invoke-TrustRemovalProbe.ps1 fold with the rest of S7d,
  Build-Corpus.ps1 and Reset-HubPopulation.ps1 with S7c, and the check lists them until they do.
  What remains, one stage at a time: S3, the linked files and the shared logic (L01-L07, A06, A07,
  A10) into Core, the logic only the add-in runs into `OutlookAI.Core.AddIn` (Q163 (a)); S4, the
  signature and draft code (A01-A05, A08, A09, A12) with Q162's checkbox, Core winning where the
  copies drifted apart and the pane's visible differences measured on a test machine first
  (Q165 (c)), and the pane driven on a test machine (Q168 (b)); S6, the server's own repetition
  (S01-S14). In parallel: the rest of the shared test code (S5a), the PowerShell helpers (S7), and
  the rules restated in C# and PowerShell as shared C# source (Q171).

- [ ] **Let the add-in's AI choose the signature by itself, once the add-in uses OutlookAI.Core (Q158
  and Q162, decided 2026-11-17).** Signatures are managed through the MCP server only; the add-in
  gets no editor, list or dropdown, because Outlook has its own (Q158). It gets one checkbox
  instead, "Allow the AI to automatically choose the correct signature", on by default. With it on,
  the AI picks the signature when it writes or rewrites the body, and when a reply or forward opens,
  matching the language of the mail being answered - the maintainer's example is his "kind regards"
  and "vriendelijke groeten" signatures - from the signatures the sending account can use. It never
  replaces a signature someone chose by hand, and never changes a mail as it is sent (Q162 (b)). It
  uses Core's signature code, not the add-in's own copy, so it waits for the Core reference above
  (Q154). Today the task pane's "Select the best signature" does this by hand, through that copy.
  The Core code it is to call is in place, and the server's drafts use it:
  `SignatureEnvironment.ReadCurrent()`, then `SignatureSelector.Candidates(...)` for the signatures
  the sending account can use, with excerpts to judge the language by, and `SignatureSelector.Decide(...)`
  for the default and why (`QUESTIONS.md` decision log, 2026-11-17, Q147 carried out, item 10).

- [ ] **Remove the Cached Mode policy values OutlookAI 3.0.0.319 to 3.1.0.325 wrote, when the
  maintainer says so (Q151 (a), decided 2026-11-17).** Those versions wrote the five values as
  Outlook policies under `HKCU\Software\Policies` wherever the user's token could - the maintainer's
  workstation since 2026-07-23. Under Q145 (a) OutlookAI writes no policies any more; these stay
  where they are, labelled in OutlookAI Settings as "Set by an earlier OutlookAI". In his words:
  *"Q151 a - So I can see how it looks in our gui. Then after I will tell you to remove them."* When
  he does: first set each account's own offline period to All - in Account Settings, or through the
  account's own profile value once OutlookAI writes it (Q145) - and only then remove the values. A
  policy overrides an account's own period without changing it: one account on his workstation
  reads 12 months under the policy's All, and removing the policy alone would make Outlook drop that
  account's offline mail older than 12 months (inferred from Microsoft's KB 2733062).
  `Tuning\Applied` records which values OutlookAI wrote itself. The removal is built and tested, and
  nothing calls it (`Services/EarlierPolicyRemoval.cs`, `OutlookTuningService.RemoveEarlierPolicyValues`,
  T1 `EarlierPolicyRemovalTests`): it sets every account's own period to All first, through
  `IAccountOfflinePeriods`, and refuses whole until the add-in can (the item on the account's own
  period, below). Wiring it to a button or to startup is the change his word will be. **Open before
  then:** the policy's `DownloadSharedFolders` = 1 also overrides each account's own "Download shared
  folders" box. If an account has that box off, removing the policy drops its shared folders from the
  offline copy. Directions: (1) the removal first turns every account's box on, through the same seam
  (the box is bit 0x800 of the profile section's config flags, read on all four of the workstation's
  accounts); (2) it leaves that one policy value in place; (3) it removes it anyway, since the box is
  on by default. Recommended: (1).

- [ ] **Read the next `wwlib.dll` crash's dump, if there is one, and say whether it was the compose crash**
  (runbook 4.6, item 8). OUTLOOK.EXE died once in Word on `OutlookAI-Indexed` (4.2f run 4, 2026-10-03), in
  `NewDraft_Hub_SignatureOverride`, before the guests kept dumps. Every crash caught since was
  `OLMAPI32.DLL+0x2E411` in the compose inspector's close that `c4595b1` removed, and the Word crash struck in
  that same compose - so it is most likely the same fault, but that is inferred. Nothing to do unless it comes
  back: the runner now fetches and reads every dump. Directions: (1) wait for one; (2) run the compose
  classes in a loop on the indexed guest until one appears or 30 runs pass clean. Recommended: (1).

- [ ] **Find out why `LiveMoveArchiveTests.MoveChain` could not resolve the hub's new Archive folder
  on a guest that was not restarted** (seen once, 2026-10-03, `OutlookAI-Indexed`; `QUESTIONS.md`
  decision log, the indexed guest's first live runs, item 6). `archive_mail` created the hub's
  Archive folder and the test's own verify session then read "NoDesignatedArchiveFolder"; the item
  left in Archive failed five later tests' hub check. The run had staged the suite onto
  `CP-17C`'s running Outlook - no graceful restart, no step 9a. Every run since through
  `Testbed/host/Invoke-LiveTierOnGuest.ps1`, which always restarts, passed it (six of six). If a
  user's long-running Outlook can hit the same, a second `archive_mail` could miss the folder the
  first one made. Directions: (1) reproduce with a run that skips the restart; (2) re-read the
  designation from a freshly opened store in the verify session; (3) read it in the product the way
  the verify session does, after creating. Recommended: (1) first.

- [ ] **Let the add-in set an existing account's own offline period, once it can call Core
  (Q145 (a)).** OutlookAI writes the offline period as the user's own `SyncWindowSetting`, which only
  accounts added later take (KB 3115009). An account that already exists keeps its own period in its
  Outlook profile, so OutlookAI Settings tells the user how to set it in Account Settings. Measured on
  `OutlookAI-Exchange` on 2026-11-17, test account only: the period is the pair 0x6649 (months) and
  0x665A (days) in the account's own Exchange provider section, Outlook's eleven slider positions
  are the eleven pairs of Q143, and a pair written with Outlook closed or while it runs is honoured
  at its next start. The writer is built, in Core: `AccountOfflinePeriodWriter`, with T1
  `AccountOfflinePeriodWriterTests`. The add-in references Core since the code-sharing refactor's
  first stage (Q167 (a)), but may call only Core's namespace `OutlookAI.Core.AddIn` (Q164 (a)), so
  what waits is moving the writer, and the profile reader it uses, into it. Then implement
  `IAccountOfflinePeriods` (`Services/AccountOfflinePeriods.cs`, which says what it must do), set it
  in `OutlookTuningService.AccountPeriods`, and tell the user the new period takes effect when
  Outlook next starts. The dialog and the removal below both use it. Not measured: whether a running
  Outlook applies a written pair before it restarts, and whether Outlook then really downloads that
  much more mail.

- [ ] **Decide whether CodeQL should follow an agent's arguments beyond the MCP tools themselves
  (Q146 follow-up, 2026-11-17).** Since Q146 the release's CodeQL scan treats every MCP tool's
  parameters as untrusted input (`Tools/CodeQL`), and its census proves the model reaches all 22
  tools. But scratch queries over 9f4646d show the taint stopping early in two places. None of it
  reaches the COM host's session (0 of `OutlookComSession`'s parameters): `MailService` hands the
  work to that process through a gateway lambda and a pipe, which CodeQL does not follow, so what
  the COM host does with an agent's text - the folders it creates, the files it writes, the
  filters it builds - is outside the scan. And the eight list parameters (`ids`, `attachments`,
  `remove_attachments`) lose it at the call into `MailService`, where each becomes an
  `IReadOnlyList`. A scratch run that also made the parameters of the COM session's 27 operations
  sources found three more results, all of them `save_attachment`'s target folder (the next item).
  Directions: (1) one models-as-data row per COM-session operation, as Q146 (d) did for the tools,
  with the census counting them too - the same mechanism, but the rows must follow the contract
  (`IOutlookSession`) as it grows, and the list parameters stay uncovered; (2) a small CodeQL
  library in `Tools/CodeQL` that defines the sources in code - every tool parameter with its list
  elements, and every parameter of an `IOutlookSession` implementation - complete, with nothing
  to keep in step, but query code where (d) chose data; (3) keep the scan as it is, with the gaps
  written down in `Tools/Invoke-CodeQL.ps1` - free, but its silence about those two places then
  means nothing. Recommended: (1), once the next item is decided, because its three results would
  stop every release until then; (2) only if the list parameters turn out to matter.

- [ ] **Decide how wide the time-zone check's line is (Q144 follow-up, 2026-11-17).**
  `T1/TimeZoneMarkerTests` fails a test that can read the machine's zone and is not marked
  `[Trait("TimeZone", "Machine")]`, and the build VM runs the marked ones again in W. Europe. It reads
  compiled code, which cannot show a value's kind, so it counts a test that CAN reach a read, not
  only one that does: 145 of the 323 marked tests reach the zone only through five product
  conversions that convert a value already marked Local and nothing else - `WsSqlBuilder.FormatUtc`,
  `DaslDateLiteral.FormatUtc`, `IndexRowMapper.AsUtcDateTime`, `IndexSearchService.GetStaleness`,
  `ComDateValue.FromTableValue`. Under the rule that a marked test is rewritten to set its own zone
  when next touched, those 145 can lose their marks only if the product changes. Measured
  2026-11-17 with a scratch copy of the walk. Directions: (1) keep the line as it is -
  nothing to maintain, 323 marks, and the 383 marked cases take about a second of test time; (2)
  trust the five conversions through a reviewed list and count making a Local value
  (`DateTime.SpecifyKind(x, DateTimeKind.Local)`, `new DateTime(..., DateTimeKind.Local)`) as a read
  instead - 178 marks, and a list kept true by hand, as `ReadOnlyProductApi` is; (3) give the five a
  zone parameter, as `ComDateValue.FromItemValue` got in Q95, so their tests name a zone of their
  own - product work in five functions; (4) count only the tests' own code - 28 marks, and it misses
  every test whose product call converts in the machine's zone, `MailService.ToUtc` (122 tests)
  among them, the Q95 kind. Recommended: (1) now, and (3) for each function as it is next changed.

- [ ] **Decide whether the draft tools refuse an attachment path Outlook cannot take (found
  2026-11-18, by save_attachment's live tests).** On the unindexed test guest, `Attachments.Add` on a
  file called `CON.txt`, given as a `\\?\` path, crashed OUTLOOK.EXE with an access violation inside
  the call (live run 20261118-050200). The draft tools' `attachments` take any rooted path the server
  can read (`DraftAttachments.Validate`), and such a path passes every check there, so an agent could
  crash the user's Outlook through `new_draft`, `reply_draft`, `forward_draft` or `update_draft`. Not
  measured: whether every `\\?\` path crashes it, or only one that names a device. Directions: (1)
  refuse `\\?\` and `\\.\` paths, and file names Windows reads as a device, in
  `DraftAttachments.Validate`, saying why - cheap, and nothing an agent needs; (2) measure first which
  paths crash it and refuse exactly those - narrower, at the cost of a guest run that crashes Outlook
  on purpose; (3) leave it: an agent needs its own file tools to make such a file. Recommended: (1).

- [ ] **Decide whether the VM pools keep CPU free for the maintainer, as they keep memory free
  (2026-11-17).** The pools (`Testbed/host/TestbedPool.ps1`) start a member other than the first only
  while the host has its memory and 16 GB more available, and nothing looks at CPU. From 17:20 to 17:40
  that day three build runs and four Outlook guests ran at once: the host's 32 logical processors read
  82 to 100 % busy for the whole minute sampled, each build run took 17 to 19 minutes instead of 9 to
  10, and one timing-sensitive test (`T1.AuditLogStressTests`, eight writers checked for time order)
  failed on both build clones and passed alone minutes later (`Docs/live-tier-on-the-vm.md` section
  4.7). Directions: (1) leave it - the most runs at once, a busy PC while they overlap; (2) a CPU gate
  beside the memory one: a member other than the first starts only while the host's CPU, sampled over
  a few seconds, is under a threshold, else the run waits; (3) cap each testbed VM with Hyper-V
  (`Set-VMProcessor -Maximum`), so no VM can take all of its virtual processors' share; (4) fewer or
  smaller clones - one build clone, or clones with 2 virtual processors. Recommended: (2), at 75 %:
  it keeps every run the pools allow today when the PC is quiet, and stops adding load when it is not.


- [ ] **Waiting for the shared test mailbox (Q109).** The maintainer requested a free shared
  mailbox in his Microsoft 365 tenant on 2026-10-03, with full access for the Exchange VM's
  mailbox. On 2026-10-08 it did not exist yet; creating it lies with a colleague. Until it exists,
  the six `Requires=DelegateStore` Exchange tests stay disabled on the Exchange test VM. Once it
  does: enable them there, and move test writes from the Exchange VM's mailbox into the shared
  mailbox wherever a test allows it (Q110).
  Then also read which signature, if any, Outlook itself puts in a reply from that mailbox, and give
  OutlookAI's reply drafts there the same (Q160, decided 2026-11-17). Until then a reply from a
  delegate or shared mailbox gets no default signature, and its result says so; a signature the
  agent names still goes in (`QUESTIONS.md` decision log, 2026-11-17, Q147 carried out, item 7).

- [ ] **Match what Outlook does with a named signature on a reply from a shared or delegate mailbox,
  once the shared test mailbox exists (Q203 (d), decided 2026-11-18).** Today a signature the agent
  names goes into such a reply, and no default does (Q160; `QUESTIONS.md` decision log, 2026-11-17,
  Q147 carried out, item 7). The maintainer kept that for now, Q203 (a), and decided what follows,
  (d): once the shared test mailbox of Q109 exists, measure what Outlook itself does with a named
  signature on a reply from a shared or delegate mailbox, and make OutlookAI's reply drafts there do
  the same (`QUESTIONS.md` decision log, 2026-11-18). The item above holds the default signature's
  half (Q160).

- [ ] **Measure whether Outlook lets one message use another signature, or none, under a policy-forced
  default (Q156, G2's policy half).** Until it is measured, OutlookAI refuses a different signature and
  "none" while Group Policy or Cloud Policy forces a default (`SignatureSelector.ForcedConflict`). The
  live tier cannot set a policy: `HKCU\Software\Policies` is read-only for the user (its ACL grants
  ReadKey, checked on the workstation 2026-11-17), and the suite runs at RunLevel Limited. It needs an
  elevated step on a PST guest before the suite - a parameter of `Invoke-LiveTierOnGuest.ps1` that
  writes the policy values over PowerShell Direct, undone by the checkpoint restore - then a live test
  that reads Outlook's own view (`LiveOutlookTestMailer.ReadSignatureEngine`) and its compose window's
  Signature menu. The maintainer decides whether the runner gets that step. The same step settles G2's
  other half, a default deployed for every account in `Common\MailSettings` (QUESTIONS.md, 2026-11-17,
  item 5): its first run, on OAI-UNINDEXED on 2026-11-17, could not tell what Outlook does with it -
  Outlook's signature settings, read through a draft's editor, reported no default with it or without
  it - and the value had been written while Outlook ran, which Outlook may read only when it starts.

- [ ] **Decide how Office on the Exchange VM gets back the ability to create items (found 2026-11-17).**
  Its Office LTSC 2024 volume edition was never activated, its grace ran out in late October, and since
  then Outlook creates no item through its object model: 41 of a run's 87 tests failed at their first
  draft or mail, every read passed, and the write tests cannot pass there (runbook §4.4a, "The live runs
  of 2026-11-17"). The other guests do not meet it because their clocks are frozen inside the grace
  (Q130). Directions: (1) re-arm Office's grace with its own `ospprearm.exe`, a limited number of times,
  each run or monthly - free, but the licence terms need a look; (2) install Microsoft 365 Apps there,
  activated by the test account's own licence if the tenant gives it one - a licensed component, so
  a decision on AGENTS.md's Dependencies rule, and it would also let Q147 measure real roaming
  signatures; (3) give the VM a real activation (a MAK or KMS key) - a licence key; (4) keep the VM to
  `Writes=Nothing` tests until one of these lands. Proving the cause first means one item created by
  tested helper code after a re-arm.

- [ ] **Delete the one tagged leftover in the Exchange VM mailbox's Sent Items once the maintainer has
  confirmed its id (Q138 (c), decided 2026-11-17).** Found read-only on 2026-11-17: subject
  `[OutlookAI-McpTest]
  d3994a2ad25309442 d39 move seed`, sent 2026-07-26 03:46:54 (W. Europe), EntryID beginning
  `000000009D47E954`, in the real mailbox's Sent Items; nothing else tagged in either mailbox. After his
  confirmation: delete exactly that EntryID - EntryID allowlist AND ordinal tag match - through tested
  helper code and the suite, then remove it from Deleted Items too. The only write the real mailbox
  ever gets.

- [ ] **Decide what to do about the test account's automapping into the Exchange VM's mailbox.** Since
  the test account joined the profile, Outlook shows a THIRD store, named by the test account's display
  name: the test account's mailbox once more, as an online delegate store under the real mailbox's
  account - Exchange automaps a mailbox to a user with Full Access to it, so the real mailbox holds Full
  Access to the test account (a tenant setting, not the VM's). The settings leave it undeclared: it
  is the hub seen through another account, so a census of it
  would read the hub's own test traffic as a loss, and the write guards refuse it by name. Directions:
  (1) remove that Full Access, or re-grant it with automapping off (a colleague, in the tenant); (2)
  keep it and let it stand in for the shared test mailbox of Q109 for the read-only `DelegateStore`
  tests - it is a real delegate store, but its contents are the hub's churn, so the census would need to
  treat it as the hub; (3) keep it undeclared, as now. Recommended: (1) if the shared mailbox of Q109 is
  still coming, else (2) after a run shows what those tests make of it. The product side is settled by
  decision 207 (b), 2026-11-18: the maintainer's own Outlook has the same kind of pair, so OutlookAI
  treats the entry as a second view of the indexed mailbox - one store ID, listed and searched once,
  through the cached entry, and named so in `outlook_health` and `list_accounts` (`McpServer/README.md`
  load-bearing fact 18).

- [ ] **Two things the Exchange VM's first runs left open.** (1) The sign-in's verification-code step has
  never met a code page: Microsoft asked for no MFA code on either sign-in of 2026-10-03, nor for the test
  account's on 2026-10-08, nor in Outlook on the web on 2026-11-17, so only the RFC test vectors stand
  behind the TOTP generator - the first code page will be its first real proof. (2) The five
  `T2/LiveSearchInTests` throw NullReference when the settings carry no `subjectOnlyProbe` instead of
  refusing with the remedy.

- [ ] **Stop D49's lifetime pins being left behind, once Q149's measurement has said whether pins
  stay (Q118 (a), approved 2026-10-11, on hold for Q149).** Approved: a per-user list of live
  sessions and the pins they made, by owner PID and pin window handle; the last session out closes
  the pin; the pins of dead sessions are closed; `outlook_health` lists the sessions; and a
  machine-wide lock, so sessions that start Outlook at once make one pin. The window handles also let
  the show-me path recognise a pin (S6, below). On hold for the maintainer's choice between keeping
  Word with pins made only while drafting - then build (a) for those pins - and writing drafts as HTML
  with no compose window and no pins - then this item goes (`QUESTIONS.md` decision log, 2026-10-11
  and 2026-11-17). Q149 (d) measured on 2026-11-17 that pins stay for as long as drafts go through
  Word: without one, the compose window's close ended a window-less Outlook in 40 of 40 drafts on
  Office LTSC 2024 (`Docs/research/q149-draft-text`). **Measured** 2026-10-03 on
  `OAI-UNINDEXED` (`McpServer/Docs/com-host.md`, "Pins left behind, measured"): a session pins only
  when it finds no Explorer, so against a user's open Outlook no
  pin is ever made (11 sessions, 0 pins). A pin is LEFT when its maker did not start Outlook - sessions
  that connect while Outlook is starting or running window-less all pin, and none of them closes its
  own (2 at once left 1, 3 at once left 3) - or when the maker's COM host is killed. A pin left keeps a
  window-less OUTLOOK.EXE up: `Application.Quit()` did not end it (3 of 3), nor did the user's close of
  the window they had opened over it (2 of 2); File > Exit did. **S6 is true in one state:** after
  that close, `ActiveExplorer()` returns the hidden pin, the show-me path's `Activate()` makes it
  visible, and the user's close of it ends Outlook - harmless there, because the close raised Quit
  and dropped every attached session, so the pin had no live owner. **What the research of
  2026-10-11 found (a) also needs** (`.work\research-2026-10-11\q118-tray\report.md` of the
  maintainer's checkout, gitignored): record who started this OUTLOOK.EXE once per process, not per
  session, because a starter that leaves first ends Outlook under the others today; close only pins
  recorded by window handle, never a visible window or another program's hidden Explorer; while
  another session lives, adopt a dead session's pin rather than close it, since closing the only
  Explorer ends Outlook under everyone; let any session that hears Quit close the shared pin, so the
  user's exit wins for every session; mark "Outlook is exiting" across processes, so a new server's
  first call waits for that process to end instead of starting another; and have `outlook_health`
  also say who started Outlook and what will end it. Never measured: the tray icon in any of these
  states, and what its Exit Now does to a pinned Outlook.

- [ ] **Put a reply's or forward's text above the quoted original's header block, not inside it.**
  Measured on 2026-11-17 by Q149 (`Docs/research/q149-draft-text`, "Found on the way"): all 12
  replies and forwards the product wrote with a signature, on both Office LTSC 2024 test guests, had
  the body and the signature under the line that heads the quoted original. Word lays out their first
  eleven paragraphs with that block's top border, and the first of them in Times New Roman. A reply
  made with no inspector - every reply since the re-opened-draft compose (`c4595b1`) - has no
  `_MailAutoSig`, so `ApplySignatureToDocument` places the signature at the start of `_MailOriginal`,
  which sits inside the header block's bordered `div`, and `InsertBodyAboveSignature` then writes the
  body above the signature, inside the same block. A reply without a signature was not measured. Cell
  C4 of `T2/LiveQ149HtmlOnlyParityTests` shows it: today's reply against one laid out as `Reply()`
  leaves it.

- [ ] **Stop the compose window showing on screen in an Outlook with no window.** Measured on
  2026-11-17 by Q149 (`Docs/research/q149-draft-text`, cells L1 and L1c): in 50 of 50 drafts on the
  two Office LTSC 2024 test guests, the editor promotion showed the compose window at (25,25),
  768 x 519, for 130-646 ms, and gave Outlook the foreground. `ComposeSurface.PromoteForWordEditor`
  parks the window off-screen before `Activate()`, and `Activate()` puts it back where Outlook keeps
  it; the sweep after it hides it again. With an Outlook window open it stays off-screen (40 of 40).
  How to fix it depends on Q149's decision: writing drafts as HTML shows nothing, and one promotion
  per Outlook process (cells W6 and W7) would show it once per process instead of once per draft.

- [ ] **Let the count tripwire's census read a table date by its column spelling too.**
  `CensusTableRow.ReadUtc` still calls the one-argument `ComDateValue.FromTableValue`, which takes every
  value as UTC; Q11's measurement (2026-10-03) showed the explicit `ReceivedTime` column - the census's
  first spelling - reports LOCAL time. The census only compares its own readings with each other, so
  the fingerprints stay consistent and nothing is mis-judged; only the instant it would print beside a
  departed item is off by the UTC offset. Pass the spelling (`CensusColumnMap` knows the index, the
  names list the spelling) when the census is next touched.

- [ ] **Measure which store-hash input - and which name - a cached Exchange store is filed under
  (Q113 (b)), then remove the one open exception.** Every store that is not Exchange is now matched to
  the search index by its own name AND hash, never by a guess (Q113 (a), `McpServer/README.md`
  load-bearing fact 16). An Exchange store still keeps the pre-Q113 match - its hash alone when one
  documented input fits (`storeHash`), else the old name rule (`displayName`) - because Microsoft
  documents its input (the profile's `PR_MAPPING_SIGNATURE`; the product also tries the store's own
  signature and the entry-ID-plus-`.ost` variant) but nothing has measured it, nor the name the index
  files it under. MEASURED 2026-10-03 on the Exchange test VM (`OutlookAI-Exchange`, one cached Exchange account): `matchedBy` `storeHash`, `matchedInput` `profileMappingSignature` - the profile's `PR_MAPPING_SIGNATURE`, as Microsoft documents. Still open: the name the index files it under, and a shared mailbox's row (Q109). Read both from each Exchange row's `matchedBy`,
  `matchedInput` and `indexStore` in `outlook_health` (or the maintainer's own, read-only, on a build
  he asks for). Then make `StoreIndexIdentity.InExchangeException` false for the measured store
  types - the one switch - read the own name for them in `GetStoreDetails`, and delete
  `StoreIndexMatchKind.StoreHash` and `NameRule` with the tests that pin them
  (`T1/StoreIndexMatcherTests`, the "open exception" cases). Delegate rows stay `delegateFolder`.

- [ ] **Decide what an UNSCOPED search does with hits from another Outlook profile's stores (Q99
  finding).** One Windows user has one search index for all of their Outlook profiles, so a search
  without `store` returns index hits from every profile - measured on `OutlookAI-Indexed`, where the
  tier profile's unscoped search returned the corpus store of `CorpusProfile`, under the name
  `Outlook Data File` that the tier profile's own store also had (`Docs/live-tier-on-the-vm.md`
  section 8 item 24). Such a hit cannot be opened from the open profile. Unchanged by Q99, and harmless
  on a one-profile machine. Directions: (1) leave it, and document it; (2) FLAG such hits from the
  store map (a hit whose store root no store of this profile claimed); (3) DROP them; (4) scope an
  unscoped search to this profile's roots. Recommended: (2). The Exchange half it waited for is
  measured (2026-10-03, Q113 (b), on the Exchange test VM): a cached Exchange mailbox is tied by its
  store hash, input `profileMappingSignature`, so a root the hash did not tie is not this profile's own
  cached Exchange store. Delegate roots stay `delegateFolder`; a delegate's row is not measured yet. Since
  2026-11-18 (the stale-copy decision) an unscoped search already leaves out another profile's store
  when it carries the name of a store of this profile that the index ties by its hash - the decoy of
  runbook section 8 item 24 - and reports it (`index.storeCopiesLeftOut`); another profile's stores of
  other names are unchanged.

- [ ] **Measure `/ \ * ?` in an ATTACHMENT's file name in an `/at=` URL (Q99 follow-up, 2026-10-03).**
  Measured: the index does not encode `%` there (`OutlookAI 50% off.txt`, `OutlookAI %2A look-alike.txt`
  appear as they are - `Docs/live-tier-on-the-vm.md` section 8 item 29), so the product decodes nothing.
  The other four cannot be measured on a guest as it stands: no route the object model leaves open puts
  them into a file name (no file on disk can carry one, Outlook refuses a PropertyAccessor write of
  `PR_ATTACH_LONG_FILENAME`, the URL carries the file name and never the display name, and an embedded
  item's file name is its subject with exactly those four removed). Only mail from another system can
  carry them. If the index writes such a `/` raw, `MapiItemUrl.TryParse` takes the attachment suffix
  as the LAST `/`-separated segment, so it would misread that row - latent, unmeasured. Directions:
  (1) leave it - no route, and probably rare; (2) make the parser find `/at=` after the item segment
  whatever the name holds, pinned in T1 with a fabricated URL; (3) a tested route that submits raw MIME
  carrying such a file name to a guest's mail sink, so the index's spelling can be measured.
  Recommended: (3), then (2) if `/` comes back raw - a parser change for a spelling nobody has seen
  would be a guess.

- [ ] **Measure what reading an online-only Exchange store's top folder costs cold, and while its server is unreachable (stream K, 2026-11-18).**
  Every store's top folder is read now, once per store-list refresh, so an Exchange store Outlook does
  not cache gets the 28-character S ID like the rest (the store-ID follow-up's option (b),
  `QUESTIONS.md` decision log, 2026-11-18). Measured on the Exchange VM: 2.2-3.0 ms for the automapped
  online entry - but warm, because that Outlook has a window and its folder pane holds every top folder
  open. Not measured: the first read in an Outlook the MCP server started headless, where nothing holds
  the folder open, and a read while the server cannot be reached. The one wait seen was 26 s, for
  `GetFolderFromID` of an ID no folder has, while Outlook was reconnecting after the VM resumed and
  still reported itself connected; the online entry's top folder read in 2.3 ms in that same window. So
  the read is skipped only while Outlook says it is offline or disconnected
  (`OutlookComSession.ReadsTopFolder`), and a server lost without Outlook noticing is not covered -
  while `outlook_health` reads the stores inside its 5 s budget. Directions: (1) measure both,
  read-only, on the Exchange VM: a first read after a graceful restart with Outlook started headless,
  and a read with the VM's network adapter disconnected on the host and connected again after; (2) leave
  it - the rule covers what Outlook knows; (3) read no online-only store's top folder, as before
  2026-11-18, so such a store keeps its T ID; (4) read an online store's top folder only when a tool
  asks for its ID, not on every refresh. Recommended: (1), at the next Exchange VM session that restarts
  Outlook anyway - the network half only if the cold read waits.

- [ ] **Put a release candidate's MCP server in front of one session on the workstation, without installing it.**
  The manual pre-release checks (`Docs/release-manual-checks.md`, Q74 D2) exercise the Exchange-only
  write paths through the MCP server, by hand, on the maintainer's own profile - and nothing can put a
  candidate's server there today: `Tools/Switch-AddInBuild.ps1` copies and registers the ADD-IN only and
  never touches the server. Until this exists the list runs right after a release is installed, with the
  previous installer kept for rollback. Recommended shape: the same copy-then-register discipline as
  Q81 - a server build copied out of any build folder and registered for one Claude Code session, never
  globally - so it can never become the server every other session starts.

- [ ] **What still stops a rebuilder rebuilding the test VM from this repository alone.**
  `Testbed/` is the entry point and holds the runnable half - parameter set, host and guest
  scripts, the settings template, the credential contract - and `Tools/Checks/check-testbed-references.ps1`
  fails the build when a document names something the repository does not contain. The corpus
  parameters are now recorded and verified: **`vm2` / seed `7777` / anchor `2026-08-19` /
  20,000 items, default shape**, recovered from the manifest header on the guest and confirmed by
  re-running `corpus-plan`, which reproduced the per-folder and seven-day counts the measurement
  docs quote. Every doc that used `vm1 / 4242 / 2026-08-01 / 40000` was quoting an EXAMPLE.

  What is left, in the order it blocks a rebuild:

  - [ ] **Answer the open questions in `Testbed/README.md` section 6.** They are the facts that are
        genuinely not recorded anywhere - Hyper-V spec, Windows edition, Office version and bitness,
        which Outlook profile is default and how the switch is automated, whether the three-store
        layout or the mail sink exist at all. Each needs a guest or the maintainer; none can be
        derived from the repository. Since 2026-10-03 only the questions that apply to the
        script-built guests are left to answer: those section 6 asks only about the original
        guest, `OutlookAI-TestVM`, went unanswerable when it was deleted that day.

        **Two corrections, 2026-09-15, because this line had drifted from the section it points at.**
        (a) It said **eleven** questions; section 6 now numbers **16** items - 13 questions plus 3
        things nobody can put in a repository - and several carry their own "half-answered" notes.
        Quote the section, not this count. (b) **"The second Windows account" is no longer one of
        them.** Section 1.1 of the runbook used to need two Windows logons per machine, because
        Windows Search indexes a `mapi16://{SID}/` scope per ACCOUNT; the build is now two GUESTS
        (`OutlookAI-Indexed` / `OutlookAI-Unindexed`), so index state is a property of the machine
        and Indexing Options controls it unambiguously (`49e563e`). Stated carefully, because these
        are different facts and only the second happened: the per-account index assumption was NOT
        disproved - the design stopped depending on it, which also retired the riskiest unverified
        assumption in the whole layout.
  - [ ] **Fold the recovered facts into `Docs/live-tier-on-the-vm.md`.** Its section 8 lists ~20
        open items; the corpus parameters (item 15), the PST path and display name (item 11), the
        scheduled-task recipe (item 10) and how results leave the guest (item 13) are now answered
        elsewhere in the repository, and its worked examples still use the `vm1` corpus.
  - [ ] **The history rewrite for the leaked guest password is still outstanding.** The value is
        dead - rotated 2026-08-24, and the current one is provably absent from HEAD and from every
        commit - but the old username-and-password line remains reachable in history from
        `d499bf1` to `54ecd26^`. Tracked in `Docs/autonomous-session-log.md`; noted here because a
        rebuilder reading the testbed docs will find the reference to it.

- [ ] **Residual questions from the 2026-08-19/20 atomicity-claims sweep.** All 31 claims of
  non-effect were enumerated, 16 were wrong and all 16 are fixed (`Docs/completeness-gaps.md`
  section 7b, T1 `AtomicityClaimsTests`). These are the things reading could not settle, and the
  two wrinkles beside the send path that are NOT the audited defect.

  - [ ] **Which HRESULTs Outlook actually raises from `MailItem.Delete()`, `Move()` and
        `Display()` when the RPC channel breaks mid-call.** The may-or-may-not reading of
        `RPC_S_CALL_FAILED` is documented and is this repo's own stated basis for the retry
        classification, but whether these three calls produce it in practice is unmeasured. It
        decides whether the unknown-outcome wording on those three paths describes a rare event or
        a theoretical one. **The wording was deliberately fixed WITHOUT waiting for it**, on the
        maintainer's reasoning: a claim about what did not happen should not rest on an unmeasured
        probability. Measurable on the VM with `OUTLOOKAI_COMHOST_FAULT` plus a read afterwards.
  - [ ] **Whether a MUTATING response can actually exceed the 64 MB frame.** Row J8's verdict is
        "false by construction": `ComDraftUpdateResult` and `ComDraftCreateResult` carry no body,
        only recipients and attachment metadata, and no realistic 64 MB case could be constructed
        by reading - a recipient count in the hundreds of thousands would be needed. The serialised
        size of a `ComDraftUpdateResult` is measured nowhere. The fix is one `if` either way, so
        this is a curiosity rather than a blocker.
  - [ ] **Whether `TryDiscardDraft` can reach its catch-all through a NON-`IsComCallFailure`
        exception after `Delete()`.** A `NullReferenceException` or `InvalidOperationException`
        from `collection.Count` or the indexer inside `TryFindDiscardedCopy` would escape
        `_runner.Run` entirely rather than becoming `com_failure`, and what `PumpedStaRunner` does
        with such an escape was not established. It changes which message the caller gets, not
        whether the delete happened.
  - [ ] **Whether soft-deleted drafts actually survive on the maintainer's profile.** The discard
        fix leans on recoverability from Deleted Items. Outlook's "empty Deleted Items on exit"
        option and Exchange retention tags can both remove the item, and neither is visible from
        this codebase. The message says to look in Deleted Items, which is right in either case;
        what is unknown is how often looking will find anything.
  - [ ] **`send` catches `TimeoutException` only** - noted by the audit and NOT part of the defect
        it was auditing, so deliberately left alone in that pass. A child that dies for any other
        reason raises `ComHostUnavailableException`, which is not a `TimeoutException`, so that
        path is saved by `ComHostSupervisor.DescribeInterruption` instead. The outcome is right by
        a different route rather than by this `catch`, which means the send path's own
        `send_outcome_unknown` audit line is NOT written for it. Worth widening the filter, or at
        least writing down that the audit line is conditional on which way the child died.
  - [ ] **`SendUsingAccount` is written to the user's draft before the identity readback and is
        never restored.** The "Nothing was sent" claims stay true, and the messages now SAY the pin
        may have been rewritten (rows 9 and 10) - but saying so is a smaller fix than restoring it.
        Restoring needs the previous value captured before the putref and put back on every failure
        path, which is one more mutating call on a path that is currently refusing to mutate
        anything, so it wants a decision rather than a quiet fix.
  - [ ] **Seven decision lines from this pass are provably unguarded, established by mutation
        rather than assumed.** Each was reverted, built, run against the whole non-live suite and
        restored; 30 of 37 were caught, and eight of those thirty only after the gap they exposed
        was closed with a new test. The full table is in `tmp-aitrace/mutation-table.md`.
        - **Five live in `OutlookComSession`, behind a COM call no non-live test can execute**,
          which is the same class as `sortApplied` and the attachment-plan execution already
          recorded above: the saved-draft id being published on the failure path and re-read after
          the relocate (row 5's COM half), `TryMoveItemToPath` reporting created folders on the
          failure path (row 6's COM half), `TrySaveAttachment` reporting the attempted path, and
          the size read being best-effort (row 11's two halves). Every T1 test substitutes the
          session, so a test can only prove that the SERVICE layer uses what it is given. What
          each is worth is not in doubt - the shapes are read off the code - but they are
          unexercised until a live run. The cheap substitute is the one already used elsewhere: a
          temporary build that forces the branch.
        - **`CompleteMove`'s audit-failure branch** (`Ok=false` over an item that MOVED, reported
          as `outcome: applied`) needs `AuditLog.Append` to FAIL, and it writes to
          `%LOCALAPPDATA%\OutlookAI` through a path that is not injectable. `AppendTo` takes a
          directory and is used by `AuditLogTests`; wiring the service layer to it would make this
          reachable, and is a bigger change than the row it guards. (2026-10-03: no wiring is
          needed any more. Since Q86 every test process appends to its own throwaway log,
          `AuditLog.EffectiveLogPath`, so a T1 test can make `Append` fail by holding that file
          open without sharing for the duration of the call - `T1/AuditLogToolTests` already
          locks it that way.)
        - **The supervisor's own wiring for the interrupted-request outcome** needs a child that
          dies while holding a request. Both ends are pinned separately - the value
          (`MutationOutcome.ForInterrupted`) and the carrier
          (`ComHostUnavailableException.Outcome`) - so what is unguarded is the line that joins
          them.

  - [ ] **The `outcome` field is not consumed by anything yet.** It is additive and null-omitted,
        the tool descriptions teach it, and T1 asserts a `com_failure` never carries `unchanged` -
        but no agent behaviour depends on it, so its value is currently "the claim is testable"
        rather than "the claim is acted on". If it turns out nothing ever branches on it, the
        honest conclusion is that the prose was the whole fix and the field is cost.

- [ ] **Residual gaps left by the 2026-08-19 timeout pass.** The values, the three inventory
  defects, the graceful sweep expiry and the kill work all landed; these did not, and each one
  is written down because a mutation check proved it unguarded rather than because it was
  guessed at.

  - [ ] **Nothing in the non-live tier notices if `ComGateway`'s budget overload goes back to
        being a pass-through.** Reverting `BudgetedSessionProxy.Wrap(session, budget)` to
        `session` leaves all 1,887 tests green. The MECHANISM is pinned (T1
        `InProcessBudgetTests` drives the proxy directly, and removing its dispatch check
        fails); the WIRING is not, because exercising it needs a real COM session. The live
        tier is where it is exercised, which is the tier that had no budget at all until this
        pass. Options: accept and rely on T2; add `InternalsVisibleTo` to `OutlookAI.Core` and
        pin the wiring through an internal seam; or a structural IL assertion, which is
        fragile and unlike anything else here. (2026-10-03: the `InternalsVisibleTo` now
        exists - Q86 added it for the audit-log redirect - so the second option costs only
        the seam.)
  - [ ] **The grace values themselves are unmeasured.** `CleanExitGraceMilliseconds` (250)
        and `ShutdownExitGraceMilliseconds` (2000) are judgements: nobody has timed how long
        `OutlookComSession.Dispose` takes against a real Outlook, so nobody knows whether
        2000 ms is generous or short. Measuring it needs a live-tier run that times the
        child's exit with a session attached. A wrong value degrades to today's behaviour
        (the child is terminated and releases nothing), which is why it shipped unmeasured.
  - [ ] **`MailService.SearchIndexTimeoutSeconds` is pinned only from above.** T1 asserts it
        never exceeds `OleDbIndexClient.DefaultCommandTimeoutSeconds`, so reverting 60 to 15
        fails nothing. A lower bound would need a measurement constant for "how long a
        statement on a saturated indexer legitimately takes", and no such measurement exists.
  - [ ] **`T3/McpStdioClient` still has one budget for a whole session.** One
        `CancellationTokenSource` bounds every read and write for the client's lifetime, so it
        is a session budget masquerading as a per-call one. The exhaustive-scan live test now
        passes an explicitly derived budget, which is the case that would have broken first;
        the general split (session budget plus a per-`RoundTripAsync` budget, both named)
        is still open. Raising the DEFAULT is deliberately not the fix - it is the non-live
        run's only safety net against a hung stdio test, and the build VM gives a whole run 60
        minutes (Testbed/host/Invoke-TestsOnBuildVm.ps1 -RunTimeoutMinutes).
  - [ ] **Claude Code's 30-minute stdio idle abort is now the nearest client-side limit, and
        nobody owns it.** A 600 s exhaustive scan is 600 s of complete silence on the pipe -
        this server sends no progress notifications. It fits (600 s < 1800 s idle < the
        ~27.8 h per-call hard ceiling), but the idle limit is a client default nobody here
        chose, no test watches it, and a user who sets a per-server `timeout` in `.mcp.json`
        for a sensible reason will land far below 600 s. Re-measure instruction: `grep -a` the
        shipped `~/.local/share/claude/versions/<ver>` binary for
        `CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT` and resolve the stdio branch's constant.
  - [ ] **The graceful sweep expiry has never fired against real mail.** It is pinned at both
        ends in T1 - the pure boundary (`OutlookComSession.SweepBudgetSpent`) and the whole
        reporting chain from `ComSweepResult` to the advice sentence - but the walk that stops
        exists only over live COM folders. With the budget now at 165 s inner / 180 s outer and
        a measured whole-profile sweep of ~60 s, it should never fire in ordinary use, which is
        the point and also the reason it will stay unexercised.

- [ ] **Residual gaps left by the 2026-08-19 re-entrant `update_draft` pass.** The intent record,
  the idempotence key, the attachment reconciler and the add-before-remove reorder all landed and
  are pinned in T1 `DraftUpdateReentrancyTests`. These did not, and each is here because a mutation
  check proved it unguarded rather than because it was guessed at.

  - [ ] **The attachment enumeration guard is provably unguarded** (mutation M15, 2026-08-19):
        making `TryUpdateDraft` enumerate the draft's attachments even when the request touches
        none leaves all 1,921 tests green. It is a cost guard - one COM call per attachment on
        every body-only revision - with no observable payload, so nothing outside a live profile
        can see it. Options: accept and rely on T2; or count contract-level COM calls in a live
        fixture, which nothing does today.
  - [ ] **The COM-side EXECUTION of the plan is unguarded by any non-live test.** `DraftAttachmentPlan`
        is pinned state by state, and the two decisions the COM sequence makes were lifted out so they
        could be reverted (`BuildForAttempt`, `ComDraftUpdateResume.ThreadIndexFor`). What no T1 can
        reach is `RemoveAttachments` deleting the N lowest-indexed instances of a name, and the fact
        that additions now run before removals - both are `dynamic` COM against a live
        `Attachments` collection. Reversing the order back leaves the suite green. T2
        `LiveUpdateDiscardTests` is where it would be exercised; extending it to assert the ORDER
        needs a way to observe a mid-sequence state, which nothing has today.
  - [ ] **Nothing verifies that an interrupted attempt's partial writes are DURABLE.** The whole
        design is deliberately indifferent to it - it converges on the end state from whatever the
        draft is observed to hold - but the question is still unanswered and worth answering, because
        it decides how often the resume does anything at all. Whether Outlook keeps an unsaved
        `Attachments.Add` / `Attachment.Delete` when the automation client dies is undocumented by
        Microsoft (`tmp-aitrace/kill-safety.md` section 2.3). Measurable on the dev machine with
        `OUTLOOKAI_COMHOST_FAULT=hang:TryUpdateDraft` plus a read of the draft afterwards.
  - [ ] **The gap map's remaining line-number citations are stale.** Checked 2026-08-19: all nine
        `OutlookComSession.cs` references had drifted (converted to symbols in this pass), and the
        `MailService.cs`, `MailModels.cs` and `OutlookTools.cs` ones are stale too - `MailService.cs:220`
        lands in an unrelated comment, `:612` and `:3995` on bare `</summary>` lines. Convert each as
        its row is next touched; a stale number reads as evidence and points the next reader at
        unrelated code.

- [ ] **PARTLY FIXED 2026-09-15 - option (2) shipped; option (1) is the COM half and still needs a
  live run.** Fix the census's identity-to-count degradation, which is why the tripwire needs a
  re-census at all on a stable mailbox. Found by reading during the 2026-08-24 investigation.

  **What shipped (option 2, "carry the reason").** `FolderCensus.CountOnly` now REQUIRES a
  `CensusCountReason`, and refuses `Walked` - so a count-only reading whose reason says it was
  walked cannot be constructed, and "reason not recorded" cannot quietly become the commonest
  answer. `CensusIdentityPlan.TryIdentify` reports which of the six refusals applied, the walk
  failure path reports `TableUnusable`, and `T2/CensusReadingStrength` - pure - turns a PAIR of
  readings into the clause the verdict carries. Three things follow:
  - **The wrong sentence is gone.** `EvaluateByCount` said `(folder above the identity budget)`
    whatever the cause. That was right in one case of six and pointed the reader at the wrong
    number in the other five - in an emergency, in a message read once, by somebody who believes
    mail has just been deleted.
  - **A degraded pair says it is degraded.** An `ITEMS LOST` over a folder the baseline identified
    and this pass could only count now says the POST-RUN reading was WEAKER, names why, and says
    the same mailbox state can read as a loss on one pass and as filed on another for that reason
    alone - so "this folder lost items" and "this reading was weaker than the one it is compared
    against" stop printing identically.
  - **The silent case is noted.** A degraded pair whose count did NOT move is where the guard
    quietly loses its teeth: the baseline could tell a filing from a deletion there and this pass
    cannot, so an item removed while another arrived is invisible. Noted, never failed - nothing
    was observed to leave.

  Pinned by 22 new T1 tests (`T1/CensusReadingStrengthTests`), including a source read of the live
  census's call site, which is the half no CI test can execute. **Nothing about which runs FAIL
  changed**: the same deltas fail, with a verdict that now says what it is made of.

  **Still open: option (1), re-walk the degraded folder before concluding anything.** That is the
  only one that removes the false failure at its SOURCE rather than explaining it, and it is one
  COM table read inside `CaptureFolder` - so it cannot land without a live run to check it.
  Option (3), refusing to compare a degraded pair at all, stays rejected: noisier by a lot, and it
  would fail runs for a property of the census rather than of the mailbox.

  <details><summary>Original entry</summary>

  **The defect.** Whether a folder is compared BY IDENTITY or BY COUNT is decided independently
  on each pass, and the post-run decision is timing-dependent. `CensusIdentityPlan.Repeating`
  refuses a folder that grew past `500 x 4`; `ShouldIdentify` refuses one once the 120 s
  identity budget has expired for that store; and `WalkFolderItems` returns null - meaning
  "count this folder" - on any COM call failure, a missing column, a row with no EntryID, a
  duplicate EntryID, `EndOfTable` false, or a count that moved under the read. Any of those on
  the post-run pass but not the baseline silently switches the folder from
  `EvaluateByIdentity` to `EvaluateByCount`.

  **Why that matters.** The count rule cannot exonerate a filing and cannot see a departure
  masked by an arrival. So the SAME mailbox state yields `note: filed (not loss)` on one reading
  and `ITEMS LOST` on the next, purely because the second reading was weaker. That is a census
  that reports differently twice with nothing having changed - the thing the original "treating
  as enumeration noise" message blamed on COM - and it is exactly the shape a re-census 30 s
  later would clear, which is how it has been hiding.

  **It also lies about itself.** `EvaluateByCount`'s failure text says
  `(folder above the identity budget)`, which is one of at least five possible reasons and is
  the wrong one whenever the cause was the clock, a transient COM failure or an unusable table.

  Three ways to fix it, cheapest first:

  1. **Re-walk the folder, not the run.** When a folder the baseline identified degrades on the
     post-run pass, walk that one folder again immediately - one table read - before concluding
     anything. Removes the timing dependence at its source and is the only option that makes the
     re-census rung unnecessary for this class.
  2. **Carry the reason.** `FolderCensus.CountOnly(count, reason)`, and a failure that names it:
     "the post-run reading of this folder was WEAKER than the baseline's (identity time budget
     expired)". Does not fix the false failure, but makes it self-explaining and removes the
     wrong sentence.
  3. **Refuse to compare a degraded pair at all** - treat baseline-identified plus
     post-run-counted as an unmeasured folder and fail the run for that reason instead. Strictly
     honest, and noisier than (1) by a lot.

  (1) and (2) compose and are the recommendation. Nothing here was implemented: it changes what
  the guard proves, which is the maintainer's call.

  </details>

- [ ] **Two things `Docs/live-tier-on-the-vm.md` should say, found while answering "can the
  tripwire be always correct?" (2026-08-24).** Not edited here - another agent owns that file.
  - **AutoArchive must be OFF in the VM's Outlook profile**, and it belongs in section 2 beside
    the other profile settings. It is a client-side actor that moves items out of a PST on a
    schedule, which is indistinguishable from mail loss to a before/after census, and it is the
    one such actor a test machine can have.
  - **Section 7's prediction is now wrong in one respect:** on a `Portable` profile the retry
    ladder does not run at all (`TripwireRetryPolicy.None`), so a suspected loss on the VM fails
    on the first reading with `NO RE-CENSUS IS CONFIGURED` rather than after two re-censuses and
    a re-run. Open item 18 ("the re-census-then-re-run policy is decided and not built") is now
    closed in both halves.

  **Four more, from the bystander work (2026-08-24).** Same reason - not edited here.
  - **Section 1.3 and section 2.6 must say the bystander is DECLARED, not merely listed.**
    "The bystander is listed in `expectedStoreDisplayNames` and is never the hub" is now half
    the instruction: it also has to be named in the new `bystanderStoreDisplayNames`, and it
    must stay in `expectedStoreDisplayNames` (that is what censuses it and what
    `list_accounts` exactness counts). Listing it in only one of the two is a refusal, not a
    warning.
    **CORRECTED AND DONE 2026-09-24 (Q77): that last sentence was wrong.** The census adds every
    declared bystander back in, so one left out of `expectedStoreDisplayNames` is still watched and
    is not refused, and a store only in `expectedStoreDisplayNames` is the identity account's
    legitimate shape; the tier refuses only when no watched store is both non-hub and write-denied.
    The maintainer chose to correct the documents rather than the loader, and sections 1.3, 2.6 and
    2.10 now say what the code does. The 2.10 example and field table below are done too - with the
    split into a watched and an indexed list (Q70) on top.
  - **The section 2.10 settings example needs the new key**, matching
    `Testbed/live-test-settings.example.json`: `"bystanderStoreDisplayNames": [ "OutlookAI
    Bystander" ]`. Without it a rebuilder follows the runbook and gets a bystander the suite
    may still write to, which is the defect this closed.
  - **Section 6's "what to read in the output" gains a line**, printed straight after the
    settings line: `[tripwire] watch soundness: N declared bystander(s), M store(s) this census
    can fail on, K watched store(s) the suite may still write to`. `M=0` is the machine-readable
    form of section 7's "the guard runs and proves nothing", and it says so in the same words.
  - **Open behaviour item 20 is closed** and should say how: not by narrowing the grant - two
    live tests use it - but by a declaration the allowlist honours and the tripwire verifies.

  **Four more, from the refuse-a-vacuous-census work (2026-08-24).** Same reason.
  - **Section 1.3 must say the bystander is MANDATORY, not advisable.** "The bystander is the one
    people leave out, and the tripwire is useless without it" now understates it: without one the
    tier refuses to start, and no setting turns that off. Same for the sentence in section 2.6.
  - **Section 7's "with one PST that is also the hub" prediction no longer describes a run.**
    That configuration never reaches a census - `EnsureBaseline` throws `NO STORE THIS CENSUS
    WATCHES CAN PRODUCE A FAILURE` before the health gate and before any COM call. The paragraph
    should say so, because as written it tells a rebuilder what to expect from a run that cannot
    happen.
  - **Section 6's `[tripwire] watch soundness:` line can no longer report `M=0`** on a run that
    got as far as printing anything, so the "M=0 means the guard proves nothing" reading belongs
    with the refusal instead.
  - **The section 2.10 settings example and field table need the corpus store in
    `bystanderStoreDisplayNames`**, matching `Testbed/live-test-settings.example.json` and
    `Testbed/README.md` §3b: `[ "OutlookAI Bystander", "Corpus A" ]`. The table's
    `expectedStoreDisplayNames` row should also say that a declared bystander the profile does
    not mount refuses the tier - which is why `Corpus B`, ~~in the other Windows account's
    profile~~ **on the other GUEST** (corrected 2026-09-15: the two-Windows-accounts layout was
    replaced by two guests, `OutlookAI-Indexed` with `Corpus A` and `OutlookAI-Unindexed` with
    `Corpus B` - `49e563e`), goes in that machine's settings file and not this one. The reasoning
    is unchanged and now easier to state: it is a different MACHINE, so it is a different settings
    file.

- [ ] **PENDING TASK - process `C:/Source/SixFive7/BrowserAI/.work/truncation-prompt-for-sibling-project.md`.**
  *(Path corrected 2026-09-15: it was written with a Windows backslash before `truncation`, which this
  file stored as a literal TAB - so the path as printed named a file that cannot exist and the `t` was
  missing from the name. Forward slashes here so it cannot happen again. The file DOES still exist at
  the corrected path, checked 2026-09-15, so this item is genuinely outstanding rather than moot.)*
  The maintainer asked for this at 09:00 on 2026-08-18. It is expected to be the portable
  description-budget prompt written for another project; read it and act on what it asks for. Recorded
  here because auto-compaction was imminent when it was requested.

- [ ] **Re-run the unindexed-store probes on a MIXED profile - the one shape no machine here has.**
  Group A and E of `Docs/completeness-gaps.md` are now all closed (A1-A5, E1). Everything about
  them has been verified on two profile shapes: the fully-indexed developer profile, and the
  Hyper-V VM whose only store is an unindexed PST. **The shape that carries A1's residue is
  neither of those** - one INDEXED mailbox plus one UNINDEXED data file, so that the profile-wide
  frontier probe succeeds while one store is still absent from the index catalog. That is the
  ordinary "Exchange account plus archive.pst" desktop, it is what T1
  `UnindexedStoreReportingTests` models with a stand-in index client, and no machine to hand can
  produce it live. What T1 cannot exercise is the real chain: a live `DiscoverStoreScopes` whose
  sample returns one store's prefixes and not the other's, and `StoreHasIndexRows` answering
  false against a real SystemIndex for a mounted PST.

  Mount a PST on a profile that already has an indexed account (or add an account to the VM), do
  not add it to Indexing Options, then confirm READ-ONLY, with `search` only:
  - a plain unscoped search names it in `sweep.storesWithoutIndex` with `no_index_frontier`
    (this already worked - it is the regression guard);
  - an unscoped search with `before` older than 7 days reports `sweep.notNeeded:true` **and**
    `indexFrontierMissing:true`, `freshness:"partial"`, `degraded:true` - it used to say
    `freshness:"live"` with nothing else;
  - an attachment-only search reports `freshness:"index-only"` **and** `no_index_frontier`;
  - `outlook_health` shows the PST as `index.perStore[].inLocalIndex:false` with a problem line,
    while the indexed account still reports a frontier.

  Also worth measuring on that profile: the cost of the one COM store-list read the three
  no-sweep paths now pay (5-minute cached, so expected to be a pipe round trip), and whether
  `StoreIndexProbeBudgetMs` (1.5 s) is ever the thing that cuts the probe short.

- [ ] **An answer too big to frame: refused (done), but not yet prevented (open).** Found by the
  boundary audit on 2026-08-18. **The refusal and the measurement landed 2026-08-18; the four
  responses below are still the maintainer's to choose between, and nothing here pre-empts them.**

  **Done - (a) the refusal.** `ComHostServer.ServeAsync` now catches the framing refusal around
  its write and answers with a `ComHostResponseTooLargeException` frame naming the operation, the
  size and the limit; the serve loop keeps serving, so one oversized answer costs one call instead
  of the session. `ComHostErrorMapper` rebuilds the type, and `OutlookTools.GuardAsync` returns it
  as `ResponseTooLarge` with advice to NARROW the request rather than retry it. The testing
  objection below is answered by the fourth option it did not list: `ComHostServer` takes its
  ceiling as a constructor parameter defaulting to `ComHostProtocol.MaxFrameBytes`, so T1 reaches
  the branch with a 64 KB answer against a 4 KB ceiling - real serve loop, real framing, real
  mapper, no 64 MB allocation. Measured 2026-08-18: the two refusal tests run in 73 ms and all six
  new tests in 99 ms, against a T1 tier of 1 m 55 s.

  **Done - the measurement.** `ComHostFrameMeter` keeps a per-process high-water mark of the
  largest payload seen (both directions: encode and read, because the big frames are built in the
  child whose counters die with it) plus a refusal count. `outlook_health` reports both beside the
  limit, as `comHost.largestFrameBytes`, `frameLimitBytes` and `framesRefusedTooLarge`, and a
  refusal also raises a `problems` line. Lifetime is one SERVER process and survives child
  restarts, deliberately: the child is restartable and the question is about the product. So the
  number that says whether 64 MB is right is now collectable from any running install - it was
  previously unmeasured, which is why the entry below argues from the caps instead.

  **Done - (c), chosen by the maintainer 2026-08-18: cap bodies at the COM layer, so a frame that
  cannot be sent cannot be built.** (b) and (d) were the alternatives and were NOT chosen -
  `MaxFrameBytes` is unchanged and the sweep result is still one frame. Two bounds, both in
  `OutlookComSession`, both applied in `SnapshotBrief` where the body is read:
  `SweepBodyCharsCap` (500 000 chars, = `MailService.BodyCharsCap`, the largest body window `read`
  will ever return in one call) per ITEM, and `SweepBodyBytesBudget` (32 MiB, = `MaxFrameBytes / 2`)
  across the whole sweep. The per-item cap alone provably cannot do the job: the item count is
  4 folders x `SweepPerFolderCap` x every store, and the store count is unbounded, so 800 items at
  500 000 chars is already ~400 MB of theoretical worst case on ONE store. What it does instead is
  keep the budget FAIR - one enormous mail cannot spend it all and blind the sweep to everything
  behind it. The budget is counted in encoded BYTES rather than characters because the pipe escapes
  every non-ASCII character to six bytes: a character budget would have to be sized for that case
  and would then bite at ~5.5 M characters, which an ordinary unindexed-PST sweep really reaches.

  **The cut is never silent, and it is not a display truncation.** These bodies are matched against
  (`FreshMerge.MatchesTerms`) and shown to nobody, so unlike `read`'s windowing - which pages and
  loses nothing - a cut here can make a search MISS a real match. `ComMailBrief.BodyTruncated`
  carries the fact per item; `sweep.itemsBodyCapped` and `sweep.itemsBodyCappedUnmatched` reach the
  payload; the `body_cap` coverage code is raised on the INTERSECTION alone (cut AND unmatched),
  since a cut body on an item that matched anyway cost nothing; `sweep.bodyBudgetExhausted` says
  which bound cut, because the remedies differ. The two facts CAN be told apart and are; what
  cannot be settled is whether the term really sat past the cut, since that needs the text the
  bound refused to carry, so the sentence says "may be" and never "is". T1 `SweepBodyCapTests`
  (18 tests) plus two wire round-trips in `ComHostProtocolTests`, mutation-checked with 13 separate
  reverts, each disabling one decision.

  **What remains reachable, stated rather than implied.** (i) The NON-body half of a frame is still
  unbounded in the store count: per-item EntryIDs, StoreIDs, subjects and folder names are ~1-2 KB
  each, so ~40 mounted stores each holding 200 items in all four arrival-path folders inside the
  window would build an oversized frame carrying no body text at all. Never observed; the typed
  refusal is the backstop. (ii) The byte budget is enforced against an OVER-estimate of the encoded
  size (`EncodedBodyByteCeiling`), so it errs toward cutting early, never toward a frame that will
  not send. (iii) The bound has only ever been exercised in T1 - see the live-profile item below.
  (iv) `read` is DELIBERATELY not capped at the COM layer: `TryReadItem` still returns the whole
  body, plus the whole `HTMLBody` when `include_html` is set, so one pathological mail could in
  principle build an oversized `read` frame. Capping it there would break the one contract that
  makes `read` lossless - `bodyTotalChars` is measured from the full body and `body_offset` pages
  the whole of it, so a COM-side cut would make the total a lie and the tail unreachable. The
  measured `read` payload is ~0.5 MB, 0.8% of the limit. The sweep was the case worth closing
  because its size is driven by MAIL VOLUME rather than by one mail, and because its bodies are
  never shown, so a cut there costs matching rather than reading.

  **The other two frames that carry mail were checked and need nothing.** `ExhaustiveScan` and the
  `thread` walk both snapshot briefs with `includeBody: false` - the exhaustive tier matches
  server-side through DASL and the thread walk needs no body at all - so the sweep is the only
  frame in this server that ever carried body text in bulk.

  **MEASURED 2026-08-18 on the real 5-store profile, with that high-water mark** (read-only:
  `outlook_health`, `list_accounts`, four searches; nothing created, moved or deleted). Largest
  frame **441,930 bytes - 432 KB, 0.66% of the limit, about 152x headroom**; zero refusals.
  **This corrects the derived worst case below:** "reachable by ordinary use" is too strong. The
  filter-only search, the one that should have swept hardest, **timed out** - `Outlook did not
  respond to 'SweepFoldersNewerThan' within 30000 ms` - the supervisor replaced the COM host, and
  the search degraded to `index-only` and still answered with 100 hits. So **the 30-second sweep
  budget bites long before the 200-items-per-folder cap does**, and the cap arithmetic is the wrong
  worst case for an Exchange store. **SUPERSEDED IN PART, 2026-08-19: the 30 s budget is now 180 s,
  and with the bound lifted the frame measurement changes completely.** On a purpose-built corpus
  (one unindexed PST, 20,000 items across the four arrival-path folders, the 200-per-folder cap
  engaged) a single store's sweep produced a frame high-water of **10,734,599 bytes - 10.2 MB over
  758 items, ~13.5 KB per item**, and five such stores extrapolates to **~54 MB against the 64 MB
  limit**. So the 432 KB measured on the real profile was bounded by the TIMEOUT, not by the item
  caps, and `SweepBodyBytesBudget` (32 MiB) is load-bearing rather than insurance: it bites before
  the frame limit does, which is exactly its design intent. The residual PST case below is no longer
  the one this machine cannot produce - it has now been produced, on the test VM, and the bounds
  held. The residual risk narrows, and stays real: a fast LOCAL store
  absent from the index, where the window falls back to seven days, holding a lot of recent large
  mail - the archive/PST shape, and the one case this machine cannot produce, because the only
  unindexed store to hand is the test VM's and it is empty. **Bearing on the options:** (b) was not
  urgent on this evidence, and (c) was the one that would close the residual case outright - which
  is why (c) is what the maintainer chose (see above). The residual PST case is still the one this
  machine cannot produce, so the new bounds have never been exercised against real mail.
  Incidentally, the timeout path was observed working on a real profile for the first time: no
  hang, host replaced, honest degraded answer naming the reason.

  **The limit is reachable by ordinary use - derived from the caps 2026-08-18, not measured.** One
  `SweepFoldersNewerThan` answer is a single frame and `MailService` calls it with
  `includeBodies: true`, so a frame carries 4 arrival-path folders x `SweepPerFolderCap` (200)
  items per store, times every store in the profile. **The bodies were not capped at the COM
  layer** - `SnapshotBrief` took `item.Body` whole, and `BodyCharsDefault`/`BodyCharsCap` are
  applied in `MailService`, on the FAR side of the frame: they bound what the agent sees, not what
  crosses the pipe. **(They are capped there now - `SweepBodyCharsCap` / `SweepBodyBytesBudget`,
  2026-08-18 - so the arithmetic below is the worst case as it WAS.)** That puts 64 MB at ~80 KB average body on a one-store profile, ~27 KB on three
  stores, ~16 KB on five. An 80 KB body is an ordinary long quoted thread. **And the path there is
  the unindexed-store case**: the sweep window is normally minutes wide, so 200-per-folder never
  fills, EXCEPT when a store is missing from the index and the window falls back to seven days. So
  the more degraded the index, the larger the frame - and before the refusal landed, the frame
  bursting killed the subsystem that was compensating for the degraded index. It now refuses that
  one call instead, which is why (c) and (d) are still worth choosing between: the sweep on a
  degraded profile is exactly the caller with nothing left to narrow. This is where options (c)
  and (d) come from. Write-up in the session trace folder under Downloads
  (`tmp-aitrace/frame-size-analysis.md`).

  **The defect, as found (kept for the record; the refusal half is fixed).**
  `ComHostProtocol.EncodeFrame` refuses a payload over `MaxFrameBytes` (64 MB) by throwing
  `ComHostProtocolException` - a deliberate, specific, actionable failure. But it was thrown from
  `ComHostServer.WriteAsync`, which `ServeAsync` guarded with `catch (IOException)` only. So the
  exception left the serve loop, `Program.Main` printed it to stderr and the child exited with 1.
  The caller learned "the COM host went away", which is the one fact that says nothing about what
  to do next; the sentence naming the size and the limit reached only the child's stderr. This is
  the same species as the wrapper defect above - a good message that does not survive the process
  boundary - and it is the only other instance the audit found.

  **How likely is it?** Low - `McpServer/Docs/com-host.md` calls 64 MB "far above any real payload" and a
  `read` returns ~0.5 MB. The candidates are `SweepFoldersNewerThan(includeBodies: true)` and
  `ExhaustiveScan` over a large window. That "low" is still a derivation rather than an
  observation; the high-water mark now accumulating in `outlook_health` is what will replace it,
  and until an install has been read it says nothing on its own.

- [ ] **Verify the sweep's new body bounds against a live profile - T1 owns every decision above the COM call and none of the COM half.**
  `SweepBodyCharsCap` / `SweepBodyBytesBudget` landed 2026-08-18 and are pinned by T1
  `SweepBodyCapTests` (18 tests, mutation-checked). What T1 owns is the pure cut, the byte ceiling
  against the real serializer, the frame-half invariant, the payload fields, the coverage code, the
  advice split and the body-cache guarantee. What it cannot produce is a real `item.Body` big enough
  to cut, so **no swept body has ever actually been truncated on this machine.**

  - **What to confirm read-only, and it needs the shape the whole frame-size item is about:** a
    LOCAL store (PST/archive) absent from the index, so the sweep window falls back to
    `EmptyIndexSweepWindow`, holding enough recent large mail to move real body volume. Then an
    ordinary `search` naming that store should report `comHost.largestFrameBytes` (in
    `outlook_health`) well under `frameLimitBytes` with `framesRefusedTooLarge: 0`, and - only if
    something really was cut - `sweep.itemsBodyCapped` with its advice sentence.
  - **The two numbers worth measuring while that store is mounted**, because both are predictions
    rather than observations: the largest frame such a sweep actually produces (to see how much of
    the 32 MiB budget real mail uses), and the per-item cost of the body pass, which is one linear
    scan of at most `SweepBodyCharsCap` characters per item on top of the COM `.Body` read that
    produced it. The scan is expected to be noise next to the read; nothing has timed it.
  - **Unverifiable without a mailbox that has one:** whether any real mail body exceeds 500 000
    characters at all. If none ever does, the per-item cap is pure insurance and only the budget
    can ever bite - which would be the good outcome and should be recorded as such.

- [ ] **Measure the sweep and scan budgets against a known corpus - the tooling exists, the corpus does not yet.**
  The store shape the item above asks for cannot be borrowed from anywhere: every store on the real
  profile is indexed, so `EmptyIndexSweepWindow` never engages, and the Hyper-V VM's PST is the only
  unindexed store to hand and it is empty. So the corpus is built rather than found.
  `McpServer/OutlookAI.RemediationTools` gained five commands for it - `corpus-plan`, `corpus-probe`,
  `corpus-build`, `corpus-teardown`, `corpus-reindex` - and `Docs/corpus-measurement-plan.md` is the
  plan for what to run against the result and what each number would settle. T1
  `CorpusGeneratorTests` pins the size distribution, the date spread, the seeding, the store
  refusals, the teardown rule, the manifest format and the date-fidelity verdicts; only the COM
  calls are outside that tier, and they carry no decisions.

  - **Built once, 2026-08-19, and the measurement still has not been taken.** 40 000 items into
    the VM's PST in 12m27s at 50.9 items/sec, zero failures; resumability and determinism both
    demonstrated (a re-run skipped the 2 000 items an earlier timing run had made). Three faults
    came out of it, all now guarded in code rather than written down as cautions:
    - **Items were queued for delivery.** 5 532 landed in the target store's **Outbox** - inert
      on that VM only because its profile has no mail account. The store guard could not catch
      it: "local .pst" and "an account's delivery store" are not mutually exclusive.
      `CorpusSafety.EvaluateProfile` now refuses unless `Session.Accounts` is EMPTY, with no
      override. Stricter than "no account delivers here" because the object model cannot express
      the narrower rule - `SendUsingAccount` is per item, so any account may send a message that
      lives anywhere.

      **The population is identified, 2026-08-24.** The plan for that shape marks **5,532 items
      unread**, exactly - `corpus-plan` prints it now. So the queued items are precisely the ones
      the plan wanted left unread, and the read state was the only thing the builder treated
      differently: it set `MailItem.UnRead` and then wrote `PR_MESSAGE_FLAGS` WHOLESALE, as
      `MSGFLAG_READ` for a read item and as **0** for an unread one. Both are gone. One
      read-modify-write now carries the read state, clears `MSGFLAG_SUBMIT` on every item -
      that is the bit meaning "queued for delivery" - clears `MSGFLAG_UNSENT` only when the
      placement rung calls for it, and preserves every bit it does not own. **The mechanism
      inside Outlook is still unproved and only a build can prove it**, so `corpus-census`
      reports Outbox strays split by the plan's intended read state: a small build either
      confirms the identity or kills it.
    - **Every item was filed as a draft**, so the sweep saw 6 of 40 000 in 234-367 ms.
      `Items.Add` + `Save` produces an UNSENT item and Outlook files those in Drafts whatever
      folder they were added to; the sweep covers Inbox/Sent/Deleted/Junk and not Drafts. The
      root cause was a design fault of mine: the MSGFLAG_UNSENT write lived as a rung of the
      DATE ladder, so `--allow-undated` silently disabled placement too. `CorpusPlacement` is now
      its own probed ladder, and a rung passes only when the item's `Parent` is the target folder
      AND that folder's `GetTable` returns it.
    - **`corpus-reindex` and the post-teardown count looked only at Inbox/Sent/Junk/Deleted**, so
      with all 40 000 items in Drafts the recovery path would have reported ZERO and teardown
      would have claimed a clean store. Drafts (16) and Outbox (4) are in the scan set now. Same
      lesson as the Outbox omission the live tier's sweep already records.
    - **Nothing in the tool noticed any of the three.** A person looking at Outlook found them.
      `corpus-census` closes that: a read-only scan compared against the plan - right count,
      right folders, one copy each, nothing stranded in Drafts or the Outbox - run by every
      build on itself and settable as its exit code.

  - **The placement probe's folder-table check is fixed, 2026-08-24.** Against an empty Inbox the
    move rungs verified; with ~22,000 items present the same rungs reported the item correctly
    parented and ABSENT from the folder's table, so the build refused a placement that works.
    The refusal was right given what it observed and the observation was wrong: the filter asked
    for every corpus subject in the folder and the walk stopped at its 2,000-row cap. It now
    filters on the probe's own reserved ordinal (`CorpusPlan.DaslSubjectFragment`, bracket-free
    so a `[` can never open a DASL character class), so it selects roughly one row; and reaching
    the cap is reported as INCONCLUSIVE rather than as "not there", with a refusal that blames
    the measurement instead of the store. The date probe's exclusion half had the same defect
    and is fixed the same way.

  - [ ] **Move `T2/CorpusFreshnessTests.cs` to T1.** It is pure - no Outlook, no COM, no settings
    file, no `Category=Live` - and belongs beside `CorpusGeneratorTests`. It sits in T2 only
    because T1 was owned by a parallel worktree while it was written. A rename, nothing else.
  - **Still to do: tear down `CP-07-CORPUS-40K` and re-run.** The 40 000 drafts are still in the
    PST; the manifest is 40 002 lines and a copy is outside the guest. Teardown deletes by
    EntryID allowlist AND tag, and now reaches Drafts.
  - **The date verdict from that run proves nothing and must not be quoted.** The probe reported
    `readBack` correct with `daslIn=False`, which reads as "the date does not drive selection" -
    but the item was in Drafts while the probe queried the Inbox's table, so "not in this folder"
    explains it equally well. The two failures were indistinguishable in the output. The probe now
    settles placement first and builds its date probe with the placement that verified, so a
    re-run isolates the question.
  - **The one thing that must be settled first, before any of it is worth doing:** whether that PST
    accepts back-dated mail at all. `MailItem.SentOn` is read-only in the object model, and an item
    created straight into a folder is UNSENT, which some stores date themselves. `corpus-probe`
    settles it empirically - it writes one throwaway item per method, re-opens it by EntryID, reads
    `ReceivedTime` back, and then asks a DASL date restriction on either side of the instant whether
    it selects the item - and `corpus-build` refuses to build an undated corpus unless
    `--allow-undated` says so in as many words. **An undated corpus would make both windows select
    the same population while looking exactly like a good corpus**, which is why the refusal is the
    default rather than a warning.
  - **A product gap fell out of it, recorded as H3 in `Docs/completeness-gaps.md` and OPEN.** The
    sweep restricts on `(datereceived >= X) OR (date >= X)`, so mail carrying neither property is
    selected by NO window - absent from the freshness tier rather than mis-dated, and absent
    however wide the window is opened - while `sweep.foldersSwept` still counts the folder and
    `freshness` still says `live`. Real users hit this with imported, copied or restored mail.
    The filter is code and is not in doubt; the supporting observation is confounded by the
    placement fault above, and the row says so.
  - **The blocker for the 180 s proposal, found while writing the plan and not yet acted on:**
    `SearchBudgetMs` is `SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs`, and T1
    `BudgetCompositionTests.SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline` asserts
    that sum fits inside `ComOperationBudgets.OperationDeadlineMs` (120 s). (That test is now
    `...FitsTheFreshnessDeadline` and names the sweep's own class - see the RESOLVED row below.) At 180 s the sum is
    195 s and that test fails before anything reaches a mailbox. Anything above roughly 105 s moves
    the operation deadline too, and with it the child work budget and `ExhaustiveTimeBudgetMs`. Decide
    the shape of that change before measuring, so the measurement is aimed at the right question.

- [ ] **Verify the three folder-walk reporting fixes against a live profile - the COM half none of them can reach from T1.**
  G2, G3 and G4 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `FolderWalkReportingTests` (21 tests, driving the real `MailService` through a stand-in session
  and index client; mutation-checked - removing any one of the three fix lines fails 4 of them).
  What T1 owns is everything ABOVE the COM call, which is where the whole defect lived, since all
  three drops were decisions taken there. What it cannot produce is the COM failure itself:

  - **G2 needs a store whose `DisplayName` read throws.** Never observed on any machine here; it
    was found by reading the `catch` that swallowed it. Nothing in the repo knows what actually
    provokes it - a damaged profile entry, a data file that will not open, a store mid-removal are
    guesses. Worth trying: mount a PST, then rename or delete the file underneath Outlook while a
    session holds it. What to confirm read-only: the store appears in `list_folders` under
    `(unnamed store N)` with `nameUnreadable: true`, an unscoped `search` reports
    `sweep.storesUnnamed`, hits from it carry the label as their `store`, `outlook_health` lists it
    (via the same COM store list, so it should now reach `index.perStore[].inLocalIndex: false`),
    and `search(store: "(unnamed store N)")` is refused with the placeholder message rather than
    the typo one. **Also unverified in the real world:** whether a store that will not name itself
    will still answer `GetDefaultFolder` and `GetTable`, i.e. whether the sweep of it actually
    returns mail rather than four skips. The code handles both; only a live case can say which
    happens.
  - **G3 needs 10 000 folders in one profile, or a 65-level-deep tree.** Both are constructible in
    a test PST and neither exists here. The cheap partial check is a temporary build with the cap
    lowered (it is a `public const` on `MailService`, read at the call site) against the real
    profile: confirm `truncated: true`, `walkCapReached: true`, NO `nextOffset`, and the advice
    sentence - and that a store-by-store listing then returns the tree the capped call could not.
  - **G4 needs a delegate mailbox whose folder walk hits a bound**, i.e. G3's condition inside a
    shared mailbox. The dev profile has two delegate mailboxes indexing 11 and 23 folder paths, so
    the honest statement is that this flag has never fired on real data and cannot until the walk
    cap is reachable. The lowered-cap build covers it in the same pass: a delegate folder search
    should then report `scope.folderNamesTruncated: true`, `degraded: true`, and the INCOMPLETE
    SCOPE sentence.

  Read-only throughout - `list_folders`, `search`, `outlook_health`, `list_accounts`. No mailbox
  writes are needed for any of it.

- [ ] **Verify the sweep's sort-failure detection against a live profile - the one half of H2 that T1 cannot reach.**
  H2, G5, B2 and F3 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `SearchCoverageClaimTests` (20 tests driving the real `MailService` through a stand-in session
  and index client, mutation-checked: five separate revert-one-line runs each fail exactly the
  tests that own that line). Three of the four are settled by that, because the whole defect lived
  above the COM call. H2 is not, and the gap is narrow and specific:

  - **`SweepFolder`'s new `out bool sortApplied` is set inside the `catch` around
    `Table.Sort`, and that catch has never been observed to fire on any machine here** - the
    defect was found by reading the swallowed exception, not by hitting it. So the value the whole
    row turns on is produced by a line no test executes. What IS pinned: the flag's journey across
    the process boundary (`ComHostProtocolTests.ComSweepResult_CarriesTheUnsortedCappedFoldersAcrossTheWire`),
    its per-store filtering in `ApplySweepCounters`, the code split, and both advice sentences.
  - **What would provoke it is unknown.** `Table.Sort` needs the property present as a column, and
    the code adds `urn:schemas:httpmail:datereceived` to `Columns` before sorting, so the ordinary
    path cannot fail. Guesses worth trying, none verified: a folder whose `DefaultItemType` is mail
    but whose contents are not, a folder on a store mid-reconnect, or a search folder. A cheaper
    substitute is a temporary build that forces `sortApplied = false` and confirms the payload
    end-to-end on the real profile.
  - **What to confirm read-only**, once a folder can be made to both refuse the sort and exceed
    `SweepPerFolderCap` (200 items in the freshness window - see the `EmptyIndexSweepWindow` row in
    `Docs/magic-numbers.md` for how wide the window has to get): `sweep.itemCappedFolders` names the
    folder, `sweep.itemCappedFoldersUnsorted` names it too, `sweep.coverageGaps` carries
    `item_cap_unsorted` and NOT `item_cap`, and the advice contains the word ARBITRARY and neither
    "newest-first" nor "OLDEST".

  Also unmeasured, and cheap to settle on the dev profile while the above is open: how often the
  G5 probes now run. The trigger widened from "the merged answer was empty" to "the index tier
  returned no rows", which costs two TOP-1 statements per folder-scoped search that the index did
  not answer. It was accepted on the standing rule that completeness outranks speed, so the number
  is worth having rather than worth acting on.

- [ ] **Verify the exhaustive scan's depth guard against a live profile - the half of F4 that T1 cannot reach.**
  F4 was closed on 2026-08-18 and is pinned by T1 `ScanDepthAndSweepScopeTests` end to end from
  `ComExhaustiveResult` to the payload, plus the process-boundary round trip in
  `ComHostProtocolTests`. What no test executes is the guard itself: `if (depth > FolderWalkDepthGuard)`
  needs a folder tree more than 64 levels deep, which no CI runner and no real mailbox has. The same
  shape as H2's `sortApplied`, and the same cheap substitute applies - a temporary build with the
  guard lowered to 2 or 3, then one read-only `exhaustive: true` search of a store with any nesting
  at all, confirming `exhaustive.depthLimitReached: true`, `depth_limit` in
  `exhaustive.coverageGaps`, `freshness: "partial"`, `degraded: true`, and an advice sentence
  naming the guard's value. Worth pairing with the lowered-cap build the G3/G4 item above already
  asks for; both are read-only and neither needs a mailbox write.

- [ ] **Add the second PST to the test VM - it is what makes the count tripwire mean anything
      there.** The tripwire exempts the hub store, so a machine whose only store IS the hub gives
      it nothing to watch: it will census, report zero failures, and be structurally incapable of
      reporting anything else. Recommended layout and the reason are in
      `Docs/live-tier-on-the-vm.md` section 2.3. Add it through Outlook's own UI (File > Account
      Settings > Data Files > Add) rather than a script - creating stores is not something the
      tested helpers do, and mailbox mutation from ad-hoc shell code is the thing that once
      destroyed real mail.
      **2026-09-24:** the store is named as an `.invalid` address now (`bystander@vm.invalid` in the
      example), because a small store is found in the index only through mail addressed to it, and it
      gets the generator's bystander population (next item). Adding it by script has been permitted
      on a guest since 2026-09-15, and attaching it to the second profile while still empty since
      2026-09-24 - `Docs/live-tier-on-the-vm.md` §2.6 draws both lines.

- [ ] **Make `corpus-teardown` drain the folders it created, as it drains the items: in a PST its
      `Folder.Delete()` MOVES them into Deleted Items, so every hub rebuild leaves two more empty
      ones there.** Measured on `OutlookAI-Unindexed`, 2026-10-03 (runbook §3b item 7, §4.1d):
      after the first `Reset-HubPopulation.ps1` run the hub's Deleted Items held
      `OutlookAI-Corpus-Folder-Projects` and `-Notices`, empty; after the second, also `... (2)` of
      each - Outlook renames on collision, so they pile up - and both teardowns printed
      `folders removed 2`. Harmless to the tests (the live tier's test folders are a different
      string), but the hub grows by two folders a run and the count says something that did not
      happen. Direction (a) of §3b item 7 was chosen overnight on the maintainer's behalf, for his
      review: first a read-only corpus-tool option that resolves a manifest's recorded folder
      EntryIDs and says where each sits now - `CP-12B-POPULATIONS-V2` holds two torn-down manifests
      and their four folders to run it on - then, if a folder keeps its EntryID across the move, a
      second `Delete()` from Deleted Items by EntryID AND prefix, and a count that says which
      happened.

- [ ] **UNTESTED: what an ADVISED EVENT SINK leaves inside Outlook when its COM host is killed.
      Two of the three nominated mechanisms were measured on 2026-09-15 and both came back
      negative; the third cannot be created from PowerShell, so it is the one gap left.**

      **Measured on the test guest, 2026-09-15.** A holder process was killed with
      `TerminateProcess` - the same call `ComHostSupervisor` makes - and then a fresh client was
      timed against the same Outlook instance:

      | Holder held before being killed | Fresh client afterwards |
      | --- | --- |
      | `Application` + `NameSpace` + `Folder` | `CreateObject` 0.37 s, `GetDefaultFolder` 0.05 s |
      | the same **plus a pin `Explorer`** (never displayed, never `Close()`d) | `CreateObject` 0.37 s, `GetDefaultFolder` 0.05 s, `explorers=1` |

      **Neither reproduced anything.** A killed holder's references do not block, delay or
      otherwise inconvenience the next client, with or without a pin.

      **THE NINE-MINUTE COM BLOCK THAT STARTED THIS WAS NOT REPRODUCED AND REMAINS UNEXPLAINED.**
      It was produced by throwaway probes, not by the product, and it **must not be cited as
      evidence about the kill path** in any direction - not as a reason to change the kill, and
      not as a reason to trust it. `Docs/autonomous-session-log.md` records "COM references
      orphaned by two earlier probes" as the only surviving explanation; the trials above are the
      direct test of exactly that, and it failed, so that paragraph is superseded and the cause is
      open. Nothing about `ComHostSupervisor`'s kill was ever implicated.

      **The pin leak is real, and it is a nuisance rather than a hang.** The orphaned `Explorer`
      **persists and is visible to the next client** (`explorers=1`), and `OUTLOOK.EXE` stayed up
      throughout. So a host killed while holding a pin leaves an invisible window behind: a slow
      accumulation over many kills, and an irritation to a human who then tries to exit Outlook
      and finds it will not go. It is not a wedge and it is not why anything blocked. Note also
      that **in the first trial `OUTLOOK.EXE` survived the holder's death with NO pin at all.** That
      does not weaken the pin's justification, as this note used to say (corrected 2026-11-17, Q149):
      a holder's death releases its references, which only starts Outlook's reference-counted wait -
      about 11.5 minutes on Office LTSC 2021 - while what the pin guards against is a WINDOW closing:
      the hidden compose inspector's close, which ends a window-less Outlook at once even with
      references held (D49 row 1, 2 of 2 on LTSC 2021; measured on LTSC 2024 by Q149's cell L1,
      `Docs/research/q149-draft-text/`).

      **DECIDED 2026-09-15: do not chase the sink now.** It is the only one of the three where a
      pointer is handed **into** Outlook rather than held by us, which is why it is worth testing
      at all - but it cannot be created from PowerShell, so testing it means deploying something
      that can, and the new guests are about to have exactly that deployed anyway. **Fold it into
      commissioning**: the real COM host lands on each guest as a build step, and
      `McpServer/OutlookAI.McpServer.Tests/T3/ComHostSupervisionLiveTests.cs` already drives the
      timeout / kill / respawn path there.

      **The experiment, concretely, so nobody has to redesign it:**

      1. **Start from a freshly restarted guest.** Not a convenience: the two trials above only
         mean anything because no earlier probe's orphaned state was left on the machine, and the
         nine-minute block itself is now believed to be orphaned-state contamination of some kind.
         Assert `OUTLOOK.EXE` is not running before anything binds COM.
      2. **Connect the real COM host**, so a genuine advised event sink is registered - the thing
         PowerShell cannot produce.
      3. **Kill it with `TerminateProcess`** - `Process.Kill`, the supervisor's own path, not a
         graceful shutdown. A graceful exit unadvises and proves nothing.
      4. **Time a fresh client's `GetDefaultFolder` under a watchdog.** The watchdog is the point:
         the failure mode being looked for is a call that never returns, so a bare stopwatch
         cannot record it. Bound it, and record the bound as the result when it expires.
      5. Record `explorers=` and whether `OUTLOOK.EXE` survived, the same two observations the
         trials above took, so the three mechanisms stay comparable.

- [ ] **Live-only, and unguarded by any non-live test: three decisions inside the count tripwire's
      verification.** Established by construction rather than by mutation, because they sit behind
      a COM census that no CI test can execute: `CollectionFinished`'s early return on `NotLast`;
      the keep-alive release and the `_verified` latch in `Verify(final)`; and the key-based
      intersection in the confirmation census. The KEY itself is pinned in
      `T1/StoreCountTripwireTests`; the code that uses it is not. The cheap substitute is the same
      one this file already records for `sortApplied`: a temporary build that forces the branch.
      **Narrowed 2026-08-24:** the intersection moved out of `Verify` into `T2\TripwireRetryLadder`
      and is now pinned by `T1/TripwireRetryLadderTests` against a fake census source, bounds and
      reporting included. What is left live-only is the LiveRetrySource wiring around it - the
      three censuses it takes, the 30 s waits, and the collections it names as implicated.

- [ ] **Watch for the count tripwire firing on a PST's Junk folder.** The census marks self-pruning
      folders by asking the store for its default Deleted Items, Junk and sync-issue folders. A PST
      may refuse `GetDefaultFolder` for Junk, in which case a generator-made "Junk Email" folder is
      an ordinary folder and a decrease in it FAILS rather than being noted. Nothing prunes it on a
      machine with no accounts, so this should stay theoretical - but if the tripwire ever fires on
      Junk, this is why, and the fix is to mark volatility by folder NAME as well as by default-folder
      identity.

- [ ] **Two live tests still degrade silently, and were left that way deliberately.**
      `LiveDraftTests.ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain` and its `LiveSendTests` twin
      loop over `expectedStoreDisplayNames`, so on a machine with fewer stores they sweep fewer and
      still pass. **This got sharper on 2026-08-24 and is now the more urgent half of this entry:**
      both used to be `LiveTier=ProfileBound` and therefore never selected on a test machine, and
      both are now in the VM bucket (`Requires=MailAccount,MultipleStores,Transport`), so a VM run
      WILL schedule them and they will pass having swept whatever the machine happens to have.
      Changing a sweep assertion without a live run to check it was not a trade worth making
      unsupervised, so this is left for the first real VM run. The other two of the four
      (`LiveStaleIndexRowTests`, `LiveManageSignatureTests.DefaultAssignment`) now refuse on a
      Production profile.

- [ ] **Residual gaps left by the 2026-08-23 tier-3 classification pass.** `Category!=Live`
      no longer reaches a mailbox: eleven T3 tests that called `outlook_health`,
      `list_accounts` or `search` moved into `ComHostSupervisionLiveTests`,
      `OutlookAvailabilityLiveTests` and `OutlookHealthLiveToolShapeTests`
      (`Category=Live` + `Requires=OutlookInstance`), `McpStdioClient` now refuses those three
      tools unless a test declares them, and
      `T1/LiveTierInventoryTests.EveryStdioTestReachingOutlook_DeclaresIt` reads the
      declaration back out of the IL. **Confirmed by measurement 2026-08-24:** the full
      `--filter "Category!=Live"` runs 2,226 cases green on a machine with Outlook up, moving
      Outlook's own CPU by 0.19 s in 107 s and lowering its handle count - so the standing local
      `FullyQualifiedName!~Tests.T3.` half of the filter is no longer needed and CI's plain
      `Category!=Live` is safe for the whole suite. What that pass found and did NOT fix:

  - [ ] **`list_accounts` starts Outlook, and nothing in the tool layer stops it.** The
        supervisor's liveness verdict for `NotRunning` is `MayStart`, which calls
        `BeginWarmUp` and connects to `Outlook.Application` - so on a machine with Outlook
        installed but closed, a bare `list_accounts` launches it. `outlook_health` guards its
        own probe with `if (outlookRunning)`; `list_accounts`, `list_folders`, `read`,
        `search`'s sweep and every draft path do not. Correct for a shipped tool
        (S7/D17 permits the cold start); a hazard for a test tier, and it was reachable from
        the default run until this change. Decide whether the live tier should force
        `allowStartingOutlook: false` for the T3 stdio tests, which would need a server-side
        switch it does not have.
  - [ ] **Four T3 tests pass while asserting almost nothing on a machine with no Outlook**,
        each by an early `return` that is documented where it sits. They are not new and none
        is wrong, but together they are why "the CI tier is green" said less than it looked:
        `OutlookAvailabilityLiveTests.SearchAlwaysAnswers_AndSaysWhetherItIsComplete` (returns
        the moment `search` reports an error, which on an indexless machine is every run),
        `...ATransientOutlookState_AnswersFastAndCarriesRetryGuidance` (returns when Outlook
        is healthy, keeping only the timing assertion),
        `ComHostSupervisionLiveTests.NoComHostSurvivesTheServer` (returns when no COM host was
        spawned, keeping only the stdin-close assertion) and
        `OutlookHealthLiveToolShapeTests.OutlookHealth_CarriesTheFreshnessBlock_WithOrWithoutAnIndex`
        (skips the advice assertion when the index provider is unavailable). All four are now
        `Category=Live`, so the question is what the VM run should assert INSTEAD of returning.
        **The first three were decided and converted on 2026-10-03 (Q101)**: each return goes
        through `T2/LivePopulationCoverage` - a refusal on `Production`, a `PROVED NOTHING:` line on
        `Portable` - and the search test now fails on any error but an unreachable index, which it
        recognises by `outlook_health`'s own `index.provider` verdict. **The fourth is still as
        described**, and is the one left in this entry. Worth knowing before the next Production run:
        the transient-state test can only exercise its check while Outlook is starting, hung or
        unavailable, so on a workstation with a healthy Outlook it now FAILS by design; the server's
        `OUTLOOKAI_COMHOST_LIVENESS` override could make it force such a state instead, if that is
        ever preferred to the Q57 answer.
  - [ ] **The pin reads tool NAMES, not arguments.** `search`, `read`, `thread`, the draft
        tools, `move_mail` and the show-me tools all have a refusal that fires before any COM
        work, which is what the protocol-only half of T3 is built on - so they cannot be
        blanket-guarded, and a future test that calls one with arguments that DO reach Outlook
        would not be caught. A bounded exhaustive `search` is the realistic case.

- [ ] **`corpus-reindex` → `corpus-teardown` silently loses rows whenever two scan rows share an
      ordinal.** Found 2026-09-16 while establishing what the census's duplicate-ordinal fault is
      actually for. `CorpusCommands.RunReindex` writes one manifest line per scan row, taking
      `row.Ordinal` straight through
      (`McpServer/OutlookAI.RemediationTools/CorpusCommands.cs:1073-1077` since the population work of
      2026-09-24, which also writes the created folders first; the item loop is unchanged), and `CorpusManifest.Parse`
      keys items by ordinal - `manifest._items[item.Ordinal] = item;`
      (`McpServer/OutlookAI.RemediationTools/CorpusManifest.cs:226`), and `Add` does the same at
      `:248`. So the second line for an ordinal **overwrites** the first, in memory, with no
      unparseable-line record and no count anywhere that a reader could compare against the
      "Wrote N entries" the command prints.
      **What it costs.** Reindex is the RECOVERY path for a lost manifest, and a duplicated ordinal
      is one of the two things it exists to recover from: a build interrupted between the COM create
      and the manifest flush re-creates that ordinal on the next run, leaving an orphan copy that
      only a scan can find. Reindex would find both copies, write both lines, and then hand
      `corpus-teardown` a manifest naming only one of them - dropping the orphan from the very
      recovery manifest that exists to catch it. Teardown's phase 2 re-scan would in practice still
      reach it, so the item is not permanently stranded; what is lost is the EVIDENCE, at the one
      moment somebody is looking for it. Bounded by how rarely a build is interrupted, which is why
      this is recorded rather than fixed.
      **What the fix would be.** Key the manifest by EntryID rather than by ordinal (the EntryID is
      what every delete is addressed with; the ordinal is only ever a label), or - smaller - have
      `Parse` and `Add` refuse to overwrite and route the loser to `UnparseableLines`, and have
      `RunReindex` print the duplicate count beside the total it already prints.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)

- [ ] **Remove the 13 stale Agent-tool worktrees and their `worktree-agent-*` branches.**
  All 13 under .claude/worktrees are clean, unlocked and 0 commits ahead of master (0 to 107
  behind), so every commit they hold is on master already. Each still has a checked-out CLAUDE.md,
  which a session started or resumed there would load. Confirm no agent is running in one (the last
  writes were 2026-09-24 to 2026-09-27), then for each: `git worktree remove .claude/worktrees/agent-<id>`
  and `git branch -d worktree-agent-<id>`; `-d` refuses a branch that is not merged, which is the check.
  Added in 2.1.277: https://code.claude.com/docs/en/changelog#2-1-277
  Extended in 2.1.281: https://code.claude.com/docs/en/changelog#2-1-281
  Remaining differences: https://github.com/anthropics/claude-code/tree/main/mods/agents-md#where-it-still-differs-from-claudemd

- [ ] **Clean up old `.work` folders (`.work/g2-cp11b`, `.work/q81-addin-registration`, `.work/testbed-livetier-payload`, `.work/worktree-archive`).**
  Their copies and clones carry CLAUDE.md/AGENTS.md files. `.claude/settings.json` keeps them out of
  Claude's context (`claudeMdExcludes`), but delete what is no longer needed.
  Added in 2.1.277: https://code.claude.com/docs/en/changelog#2-1-277
  Extended in 2.1.281: https://code.claude.com/docs/en/changelog#2-1-281
  Remaining differences: https://github.com/anthropics/claude-code/tree/main/mods/agents-md#where-it-still-differs-from-claudemd

- [ ] **Move OutlookAI from MIT to the FSL-based licence BrowserAI carries.**
  Decided by the maintainer on 2026-10-01, in his words: *"Move both OutlookAI and the new library
  to the FSL based license BrowserAI has."* The new library is the shared MCP registration library
  that is being extracted from BrowserAI. Nothing in this repository has been changed for it: this
  entry is the whole of the change so far, added from the BrowserAI session on his instruction.
  **The licence to copy** is `LICENSE` in the BrowserAI repository: the Functional Source License
  1.1 (MIT Future License), modified to a five-year term. It is a bespoke variant, and its own text
  forbids the plain `FSL-1.1-MIT` identifier. BrowserAI names it
  `LicenseRef-BrowserAI-FSL-1.1-MIT-5yr` in its SPDX headers, so OutlookAI needs a name of its own
  in the same form.
  **What the move touches, to be checked when it is done:** `LICENSE`; the licence line in
  `README.md`; any licence field in the project files and in `Installer.iss`; source-file headers,
  if this repository wants them; and `CHANGELOG.md` under Unreleased, because the licence a user
  receives changes.
  **One thing to settle first.** Every release published so far went out under MIT. Decide whether
  the change applies from the next release or from a stated date, and say so in the release notes.
