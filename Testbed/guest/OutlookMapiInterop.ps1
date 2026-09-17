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
    IF SOMEBODY WANTS TO REOPEN THE E_NOINTERFACE, here is the state of it, stated as what is
    known and what is guessed, because the next person will otherwise redo the guessing.

    KNOWN: the stub resolves (`DLLPathEx` points at a real msmapi32.dll under the Click-to-Run
    root\VFS tree), the process is 64-bit and Office is x64, `MAPIInitialize` returns S_OK, and
    `MAPIAdminProfiles` returns S_OK. The failure is in the CLR's marshalling of the out-parameter.

    NOT KNOWN, AND NOT GUESSED AT IN CODE: why. The leading hypothesis is that MAPI's own
    `IProfAdmin` object does not answer `QueryInterface` for `IID_IProfAdmin` - a C++ caller
    receives the pointer directly from `MAPIAdminProfiles` and never asks, so an incomplete
    `QueryInterface` there would have gone unnoticed for decades, while a `[ComImport]`
    out-parameter forces exactly that question. That is a HYPOTHESIS. It has not been tested and
    nothing here depends on it.

    THE EXPERIMENT THAT WOULD SETTLE IT, ~10 lines, guest-only, writes nothing: declare
    `MAPIAdminProfiles(uint, out IntPtr)`, call it, and then call `Marshal.QueryInterface` on the
    returned pointer for `IID_IProfAdmin`. A non-null pointer plus an E_NOINTERFACE from that QI
    confirms the hypothesis; anything else refutes it. IT HAS NOT BEEN RUN.

    AND IF IT IS CONFIRMED, THE FIX IS STILL NOT FREE. It would mean hand-building the vtable
    calls with `Marshal.GetDelegateForFunctionPointer` - guessing at slot offsets, which is the
    exact failure mode this rewrite exists to remove, and which fails as an access violation with
    no error text rather than as an exception. Do not ship one without a guest run behind it.
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
