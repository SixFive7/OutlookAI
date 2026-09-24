#Requires -Version 5.1
<#
    ============================================================================================
    THE EXTENDED MAPI HALF OF THIS FILE IS GONE. IT WAS MEASURED BROKEN, AND NOTHING CALLS IT.
    ============================================================================================

    WHAT WAS HERE UNTIL NOW. A C# Extended MAPI interop - `IProfAdmin`, `IMsgServiceAdmin`,
    `IMsgServiceAdmin2`, the `SRowSet` readers, `Invoke-WithProfAdmin`, `Get-MapiProfile`,
    `Get-MapiService`, `Add-MapiPstService` - roughly 550 lines, compiled by `Add-Type` on every
    dot-source. Three scripts were built on it and none of them worked.

    WHY IT WENT. The first time any of it executed anywhere, on a guest, it threw:

        Exception calling "MAPIAdminProfiles" with "2" argument(s): "Unable to cast COM object of
        type 'System.__ComObject' to interface type 'OutlookAI.Testbed.IProfAdmin'. This operation
        failed because the QueryInterface call on the COM component for the interface with IID
        '{00020379-0000-0000-C000-000000000046}' failed due to the following error: No such
        interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."

    WHERE. Guest OAI-UNINDEXED, user vmadmin, 64-bit elevated Windows PowerShell 5.1, Office LTSC
    2024 (ProPlus2024Volume / PerpetualVL2024), build 16.0.17932.20884, 2026-09-16, called from
    `Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`.

    `MAPIInitialize` SUCCEEDED and `MAPIAdminProfiles` SUCCEEDED. What failed is the
    QueryInterface the CLR performs when it marshals that call's out-parameter into a
    `[ComImport]` interface. THE CAUSE WAS NOT ESTABLISHED and is not established now - see
    'IF SOMEBODY WANTS TO REOPEN IT' at the bottom, which names the one cheap experiment that
    would settle it and says plainly that nobody has run it.

    WHAT REPLACED IT, per capability. Every one of these is either documented by Microsoft or
    measured on a guest; none of them needs `IProfAdmin`:

      create a profile ............ a .prf import (`ImportPRF`), `New-OutlookProfile.ps1`
      name a PST at creation ...... the same .prf's `[ServiceN] Name=`, measured to carry `@`
      add a PST to a live profile.. `NameSpace.AddStoreEx`, `Add-OutlookPstStore.ps1`
      name a store afterwards ..... rename the store's root folder, `Rename-OutlookStore.ps1`
      switch the default profile... the HKCU `DefaultProfile` value, `Set-DefaultOutlookProfile.ps1`
      DELETE a profile ............ NOTHING FREE. No caller needs it any more; see below.

    DELETING A PROFILE HAS NO ROUTE, AND THAT IS A FINDING RATHER THAN AN OMISSION.
    `IProfAdmin::DeleteProfile` is dead with the rest of the interface. Removing the
    `...\Outlook\Profiles\<name>` key by hand is reverse engineering with no published recipe, it
    leaves the PST files and the `DefaultProfile` value pointing at nothing, and on the
    maintainer's workstation a bug in it destroys a real profile carrying real delegate
    mailboxes. The only thing that ever needed it was `Add-OutlookPstStore.ps1 -NameProbe`, which
    created and deleted a throwaway profile to find out whether Outlook accepts `@` in a store
    display name - and that question is now ANSWERED by measurement, twice, so the probe is gone
    and the capability is not needed. The remaining route for a human is the Mail control panel.

    WHY NOT KEEP THE MAPI CODE AS A FALLBACK. Because that is the bug this file was bitten by,
    twice over: three scripts carried "never been executed" banners while depending on an interop
    that had also never been executed, and the whole set failed on first contact with a guest.
    `Set-DefaultOutlookProfile.ps1` says it in its own banner - one route, exercised on every run,
    or none. Keeping ~550 lines of unreachable C# also meant every script that wants nothing but
    the guest guard paid an `Add-Type` compile for it, and a compile error in code nobody calls
    would have taken down scripts that make no MAPI call at all.

    IT IS IN GIT. Commit 8b610c2 is the last one that contains the interop in full, with its five
    constraints (C# 5 only; a type cannot be redefined in a live session; every vtable slot
    declared in order; ULONG_PTR is IntPtr; ANSI throughout) and the vtable layouts. Nothing is
    lost - it is one `git show` away - and the file's history is where a dead route belongs.

    THE NAME OF THIS FILE IS NOW HISTORICAL. It is dot-sourced by six scripts under this
    directory and named in `Testbed/README.md`, so renaming it would be a wider change than this
    one; what it actually holds now is the SHARED GUEST LAYER - the guard that keeps these
    scripts off the maintainer's workstation, the PST path normaliser, and one Outlook COM
    session helper. No MAPI call is made anywhere in this repository any more.

    IT IS DOT-SOURCED, NEVER RUN:

        . "$PSScriptRoot\OutlookMapiInterop.ps1"

    Same shape as Testbed/host/TestbedLeasePath.ps1: one definition, dot-sourced by every side, so
    they cannot disagree about it.

    WINDOWS POWERSHELL 5.1. No ternary, no `??`.

    ---------------------------------------------------------------------------------------------
    IF SOMEBODY WANTS TO REOPEN THE E_NOINTERFACE: IT IS SETTLED. RUN 2026-09-24. REFUTED.
    THE E_NOINTERFACE WAS THIS PROJECT'S DECLARATION, NOT OFFICE.

    The hypothesis this section carried was that MAPI's own profile-administration object does
    not answer QueryInterface for its own IID. The experiment below was run, exactly as this
    section specified it, and the object DOES answer - for the right IID. The IID the deleted
    [ComImport] declaration carried was not IID_IProfAdmin at all.

    THE IIDs, from Microsoft's own headers (github.com/microsoft/MAPIStubLibrary, the MAPI header
    set Microsoft publishes):

        include/MAPIGuid.h   DEFINE_OLEGUID(IID_IProfAdmin,        0x0002031C, 0, 0);
        include/MAPIGuid.h   DEFINE_OLEGUID(IID_IMsgServiceAdmin,  0x0002031D, 0, 0);
        include/MAPIAux.h    DEFINE_OLEGUID(IID_IMsgServiceAdmin2, 0x00020387, 0, 0);
        include/MAPIAux.h    DEFINE_OLEGUID(IID_IMessageRaw,       0x0002038A, 0, 0);

    and what the interop deleted at commit 8b610c2 declared (`git show 8b610c2:<this file>`):

        IProfAdmin          [Guid("00020379-0000-0000-C000-000000000046")]   WRONG - is 0002031C
        IMsgServiceAdmin    [Guid("0002037A-0000-0000-C000-000000000046")]   WRONG - is 0002031D
        IMsgServiceAdmin2   [Guid("0002038A-0000-0000-C000-000000000046")]   WRONG - is 00020387;
                                                                             0002038A is IMessageRaw

    None of 00020379 or 0002037A appears in either header. All three were wrong.

    THE MEASUREMENT. Guest OAI-INDEXED, user vmadmin, 64-bit Windows PowerShell 5.1.26100.7920,
    Office LTSC 2024 (Click-to-Run VersionToReport 16.0.17932.20996), stub routing
    DLLPathEx = ...\root\VFS\ProgramFilesCommonX64\system\msmapi\1033\msmapi32.dll (exists).
    DllImport of the stub: MAPIInitialize(IntPtr), MAPIAdminProfiles(uint, out IntPtr),
    MAPIUninitialize(); then Marshal.QueryInterface on the returned pointer. No method of the
    object was called, nothing was written, every pointer was released, MAPIUninitialize ran.
    Run twice - session 0 over PowerShell Direct (MTA, the conditions of the 2026-09-16
    failure) and session 1 through Register-InteractiveTask.ps1 (STA) - identical results:

        MAPIInitialize(NULL)             hr=0x00000000
        MAPIAdminProfiles(0, out ptr)    hr=0x00000000  ptr=0x0000021218C08AC0   (session 0)
                                                        ptr=0x0000021FA9EB1890   (session 1)
        QI {00020379-...} (the IID the deleted interop declared)
                                         hr=0x80004002 (E_NOINTERFACE)  ppv=NULL
        QI {0002031C-...} (IID_IProfAdmin per MAPIGuid.h)
                                         hr=0x00000000  ppv=the same pointer
        QI {00000000-...} (IID_IUnknown, control)
                                         hr=0x00000000  ppv=the same pointer
        Release(object)                  refcount now 0
        MAPIUninitialize()               done

    So the 2026-09-16 failure reproduces exactly - E_NOINTERFACE on {00020379-...} - and is
    fully explained: the CLR QueryInterfaces for the GUID on the [ComImport] declaration, and
    that GUID named no interface. The object answers IID_IProfAdmin and IUnknown with itself.

    WHAT THIS DOES AND DOES NOT ESTABLISH.
      * It DOES establish that "IProfAdmin is unreachable from PowerShell on Office LTSC 2024"
        was never measured: what was measured is that a declaration with the wrong IID fails.
        The same is true of IMsgServiceAdmin and IMsgServiceAdmin2, whose IIDs were also wrong.
      * It does NOT establish that the rest of the deleted interop would have worked. No call
        ever got past the cast, so its vtable slot order, its ANSI marshalling and its SRowSet
        readers were never executed anywhere. A corrected IID is the first unknown, not the last.
      * It changes no route in this repository. By decision (Q67, 2026-09-24: knowledge only), no
        vtable route and no revived interop was built on this result. The .prf / AddStoreEx /
        registry routes above remain the ones in use, and they are measured.
      * DELETING A PROFILE, listed above as "NOTHING FREE", has a documented free route after all
        (IProfAdmin::DeleteProfile, behind the correct IID). It is still unbuilt and unexercised,
        and nothing currently needs it.

    NOTE FOR WHOEVER EDITS THIS FILE'S TOP BANNER: its "THE CAUSE WAS NOT ESTABLISHED and is not
    established now" was written before this run and is superseded by this section.
#>

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# The guest guard. FIRST CALL IN EVERY SCRIPT THAT WRITES.
#
# These scripts create and reconfigure Outlook profiles and stores. Run one on the maintainer's
# workstation by mistake and it operates on a real profile with real delegate mailboxes. The
# cheapest reliable difference between that machine and a guest is WHO IS LOGGED ON: the guests
# autologon as vmadmin (Testbed/README.md section 2), and no other machine in this project does.
#
# It fails CLOSED and it is not silenceable by a flag - the only way past it is to name the
# account you mean, which is a thing you cannot do by accident.
# ---------------------------------------------------------------------------------------------
function Assert-TestbedGuest {
    param([string[]] $ExpectedUser = @('vmadmin'))

    $who = $env:USERNAME
    foreach ($candidate in $ExpectedUser) {
        if ($who -eq $candidate) { return }
    }

    throw @"
REFUSING TO RUN. This session is logged on as '$who', which is not one of: $($ExpectedUser -join ', ').

These scripts create and reconfigure Outlook profiles and stores. On the maintainer's workstation
that would operate on a real profile carrying real mail and delegate mailboxes.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2). If you are building the
SECOND Windows account - the unindexed one in Docs/live-tier-on-the-vm.md section 2.4 - it has a
different name and nothing here can guess it, so pass it explicitly:

    -ExpectedUser <that account's username>

Do not 'fix' this by widening the default. The default is the guard.
"@
}

<#
    Refuses while Outlook is running.

    It does NOT kill it. Mailbox-safety rule 7 forbids taskkill on OUTLOOK.EXE outright, and a
    running Outlook writes its own view of the profile back at shutdown - so killing it here would
    trade a clean refusal for a silent revert an hour later.

    NOTE WHICH SCRIPTS THIS APPLIES TO NOW. It guards the registry-and-file writers, where a
    running Outlook would overwrite what was just written: New-OutlookProfile.ps1's .prf import
    (a .prf is read at startup) and Set-DefaultOutlookProfile.ps1. It deliberately does NOT guard
    Add-OutlookPstStore.ps1, which drives the object model and therefore REQUIRES a live Outlook.
#>
function Assert-OutlookNotRunning {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw @"
REFUSING: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

What this script writes is read by Outlook at STARTUP, and a running Outlook writes its own view
of the profile back when it closes - which silently reverts whatever was just written.

Close Outlook properly and run this again. DO NOT taskkill it - mailbox-safety rule 7 forbids
that outright, and a forced kill can leave the profile mid-write.
"@
    }
}

<#
    Normalises a PST path ONCE, and everything downstream reuses the exact string this returns.

    Why it matters: Outlook normalises the path it reports back, so comparing a store's FilePath
    against the literal a caller typed can miss. Every comparison in these scripts is against the
    string this function returned. GetFullPathName settles the spelling; it does not require the
    file to exist, which is correct here because both remaining routes create it - the PST
    provider does when a .prf names it, and AddStoreEx does when it is handed a path with no file.
#>
function Resolve-PstPath {
    param([Parameter(Mandatory = $true)] [string] $Path)

    $full = [System.IO.Path]::GetFullPath($Path)
    $directory = [System.IO.Path]::GetDirectoryName($full)
    if (-not (Test-Path -LiteralPath $directory)) {
        throw "The directory for the PST does not exist: $directory. Create it first; Outlook creates the .pst, not the folder."
    }
    return $full
}

<#
    Binds Outlook over COM, hands the MAPI NameSpace to a script block, and RELEASES EVERYTHING
    in a finally. Shared rather than copied because the release half is safety-critical and a
    copied one is one that can drift.

    WHY THE RELEASE MATTERS MORE THAN IT LOOKS, measured on this project's guest 2026-09-15: a
    read-only GetDefaultFolder blocked for NINE MINUTES against an Outlook that was otherwise
    healthy, and the only surviving explanation was COM references orphaned by two earlier probes
    that exited while still holding Application and NameSpace - one of them via `exit 2` mid-flight.
    So: DO NOT CALL `exit` INSIDE -Body. Return a value, or throw; both run the finally, and
    `exit` from inside a try can skip it. That is the exact mistake that poisoned the instance.

    IT DOES NOT QUIT OUTLOOK, EVER. Mailbox-safety rule 7: never taskkill OUTLOOK.EXE, release COM
    references BEFORE any Quit (quitting while refs are held zombifies the process), and prefer
    leaving Outlook headless. Leaving it up is also what the corpus build wants - Build-Corpus.ps1
    refuses unless Outlook is already running and warm.

    WHY GetDefaultFolder AND NOT NameSpace.Logon. Rename-OutlookStore.ps1 - the one script in this
    directory that is MEASURED WORKING - initialises MAPI with GetDefaultFolder against the
    DEFAULT profile, and says why: Microsoft documents that Logon can raise the profile picker
    even when a default is set, and a dialog on an unattended guest is a hang rather than a
    prompt. This helper follows it exactly. The consequence is the important part and every caller
    states it: THESE SCRIPTS OPERATE ON THE DEFAULT PROFILE. Point the default at the profile you
    mean first (Set-DefaultOutlookProfile.ps1), then run them.

    -Body receives the NameSpace. $ns.CurrentProfileName is what says which profile you actually
    got, and callers check it rather than assuming.
#>
function Invoke-WithOutlookSession {
    param([Parameter(Mandatory = $true)] [scriptblock] $Body)

    $outlook = $null
    $ns = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace('MAPI')

        # olFolderInbox = 6. This is the call that initialises MAPI against the default profile.
        # Inbox deliberately: the same project measured GetDefaultFolder(Inbox) at 0.05 s and once
        # saw GetDefaultFolder(DeletedItems) block for nine minutes, unexplained to this day.
        #
        # It is wrapped only to say WHICH call failed. A raw COM HRESULT here reads as "Outlook is
        # broken" when the usual cause is much duller and much more actionable: the default profile
        # has no store to have an Inbox in.
        try { $null = $ns.GetDefaultFolder(6) }
        catch {
            throw @"
Could not open the default profile's Inbox, so MAPI was never initialised and nothing below it can
be trusted. Underlying error: $($_.Exception.Message)

The usual cause is not a broken Outlook. It is that the DEFAULT profile has no store - a profile
built with no PST has nowhere for an Inbox to be. Check which profile is default, and what it
holds:

    .\Set-DefaultOutlookProfile.ps1 -ListOnly
"@
        }

        & $Body $ns
    }
    finally {
        foreach ($reference in @($ns, $outlook)) {
            if ($null -ne $reference) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference) }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
