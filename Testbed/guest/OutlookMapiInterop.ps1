<#
    ============================================================================================
    RUN 2026-09-16 ON A GUEST, AND `MAPIAdminProfiles` FAILED. READ THIS BEFORE USING IT.
    ============================================================================================

    This banner replaces the "never been executed" one, as that banner asked - and the answer is
    not the happy one. The first time anything in this file was executed anywhere, it threw:

        Exception calling "MAPIAdminProfiles" with "2" argument(s): "Unable to cast COM object of
        type 'System.__ComObject' to interface type 'OutlookAI.Testbed.IProfAdmin'. This operation
        failed because the QueryInterface call on the COM component for the interface with IID
        '{00020379-0000-0000-C000-000000000046}' failed due to the following error: No such
        interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."

    Measured on `OutlookAI-Unindexed` / OAI-UNINDEXED, Office LTSC 2024 (`ProPlus2024Volume` on
    `PerpetualVL2024`, build 16.0.17932.20884), 64-bit elevated PowerShell, user `vmadmin`,
    called from `Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`.

    So the MAPI stub RESOLVES and `MAPIAdminProfiles` is CALLABLE, but the object it hands back
    does not answer to this file's `IProfAdmin` declaration. Whether that is the declaration, the
    IID, the vtable order or an initialisation ordering problem has NOT been established. Do not
    assume any other entry point here works because this one was reached: nothing else in this
    file has been executed either, and one proven-broken export is a reason to treat the rest as
    unproven rather than as merely unrun.

    WHAT WAS DONE INSTEAD. `Set-DefaultOutlookProfile.ps1` no longer depends on this path; the
    default profile is set through the documented HKCU `DefaultProfile` value, which is MEASURED
    working on the same guest on the same day. See that script's header.

    THE GENERAL LESSON, recorded because it cost five failed corpus builds to learn: a script
    verified by PARSING is a script whose syntax is verified. This one parsed perfectly and its
    central call does not work.

    WHAT THIS FILE IS. The shared Extended MAPI layer for the profile scripts beside it. It is
    DOT-SOURCED, never run:

        . "$PSScriptRoot\OutlookMapiInterop.ps1"

    Same shape as Testbed/host/TestbedLeasePath.ps1: one definition, dot-sourced by every side, so
    they cannot disagree about it.

    WHY MAPI AND NOT THE REGISTRY. IProfAdmin and IMsgServiceAdmin are the documented, supported
    interface for profile administration. The registry layout underneath is reverse engineering
    with no published end-to-end recipe on Office 16.x. There is a second reason and it is the
    stronger one for this project: the PST provider's configure call takes PR_DISPLAY_NAME and
    applies it AT CREATION TIME, which is the only route anybody found to a store whose display
    name is exactly what we asked for. The object model's Store.DisplayName is read-only, and
    PropertyAccessor.SetProperty on 0x3001001F at store level is reported blocked.

    THE PREFLIGHT THAT SAYS THIS IS WORTH TRYING, measured on a guest running Office LTSC 2024
    build 16.0.17932.20996:

        DLLPathEx = C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\system\
                    msmapi\1033\msmapi32.dll        [exists: TRUE]
        Office platform x64;  PowerShell Is64BitProcess True

    DLLPath being a bare unresolvable 'mapi32.dll' next to a DLLPathEx that resolves is the
    HEALTHY shape on Click-to-Run, not a fault - Click-to-Run installs into a virtualised root\VFS
    tree the old value could never name.

    FIVE CONSTRAINTS THE C# BELOW OBEYS. Each produces a distinctive failure if broken, and the
    third one produces the worst failure in this whole set:

    1. C# 5 ONLY. The guest has a .NET runtime and NO SDK (Testbed/README.md section 2), so
       Add-Type compiles with the csc.exe that ships in the .NET Framework. No string
       interpolation, no expression-bodied members, no nameof. Breaking this is a compile error at
       dot-source time: loud, immediate, harmless.
    2. A TYPE CANNOT BE REDEFINED IN A LIVE SESSION. Re-dot-sourcing after editing the C# throws
       'type already exists'. This file guards on the type being present and skips Add-Type, so
       while you are iterating on the interop you must start a FRESH PowerShell each time.
    3. EVERY VTABLE SLOT IS DECLARED, IN ORDER, INCLUDING THE ONES NOBODY CALLS. A COM interface
       declaration is positional. Omit one method - even a deprecated one - and every method after
       it calls the wrong function pointer. That does not look like a mistake: it looks like MAPI
       returning nonsense, or an access violation that takes the whole PowerShell process down
       with no error text. If a first run dies with no output, THIS IS WHERE TO LOOK. The unused
       slots below are declared with IntPtr arguments and named, deliberately, rather than omitted.
    4. ULONG_PTR IS IntPtr, NOT uint. ulUIParam is pointer-sized; declaring it uint on x64
       corrupts the stack for every argument after it.
    5. ANSI THROUGHOUT. MAPI_UNICODE is documented as unsupported on the service-admin calls, so
       every string here is CharSet.Ansi / UnmanagedType.LPStr. The exception is deliberate and
       marked: the SPropValue payloads use PT_UNICODE tags and real wide strings, because those go
       to the PST provider rather than through the admin API.

    THREADING. Windows PowerShell 5.1 hosts in STA, which is what MAPI wants; -MTA gives
    RPC_E_CHANGED_MODE. Every MAPI call must stay on ONE thread - no Start-Job, no runspaces. The
    scripts using this file are straight-line single-threaded on purpose.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1), Windows PowerShell 5.1. Profile
    administration itself does not need Outlook running, but the verify steps do, and Outlook can
    never finish starting in session 0.
#>

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# The guest guard. FIRST CALL IN EVERY SCRIPT THAT WRITES.
#
# These scripts destroy Outlook profiles. Run one on the maintainer's workstation by mistake and
# it takes a real profile with real delegate mailboxes with it. The cheapest reliable difference
# between that machine and a guest is WHO IS LOGGED ON: the guests autologon as vmadmin
# (Testbed/README.md section 2), and no other machine in this project does.
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

These scripts create, reconfigure and DELETE Outlook profiles. On the maintainer's workstation
that would destroy a real profile carrying real mail and delegate mailboxes.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2). If you are building the
SECOND Windows account - the unindexed one in Docs/live-tier-on-the-vm.md section 2.4 - it has a
different name and nothing here can guess it, so pass it explicitly:

    -ExpectedUser <that account's username>

Do not 'fix' this by widening the default. The default is the guard.
"@
}

# ---------------------------------------------------------------------------------------------
# The interop. Defined once per PowerShell process; see constraint 2 above.
# ---------------------------------------------------------------------------------------------
if (-not ('OutlookAI.Testbed.Mapi' -as [type])) {

    $mapiSource = @'
using System;
using System.Runtime.InteropServices;

namespace OutlookAI.Testbed
{
    // IProfAdmin. MAPIX.H order, every slot, IUnknown excluded (the CLR supplies those three).
    [ComImport]
    [Guid("00020379-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IProfAdmin
    {
        [PreserveSig] int GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        [PreserveSig] int GetProfileTable(uint ulFlags, out IntPtr lppTable);
        [PreserveSig] int CreateProfile(
            [MarshalAs(UnmanagedType.LPStr)] string lpszProfileName,
            [MarshalAs(UnmanagedType.LPStr)] string lpszPassword,
            IntPtr ulUIParam,
            uint ulFlags);
        [PreserveSig] int DeleteProfile([MarshalAs(UnmanagedType.LPStr)] string lpszProfileName, uint ulFlags);
        // Declared but never called. Removing it would shift every slot below by one.
        [PreserveSig] int ChangeProfilePassword(IntPtr a, IntPtr b, IntPtr c, uint ulFlags);
        [PreserveSig] int CopyProfile(IntPtr a, IntPtr b, IntPtr c, IntPtr ulUIParam, uint ulFlags);
        [PreserveSig] int RenameProfile(IntPtr a, IntPtr b, IntPtr c, IntPtr ulUIParam, uint ulFlags);
        [PreserveSig] int SetDefaultProfile([MarshalAs(UnmanagedType.LPStr)] string lpszProfileName, uint ulFlags);
        [PreserveSig] int AdminServices(
            [MarshalAs(UnmanagedType.LPStr)] string lpszProfileName,
            [MarshalAs(UnmanagedType.LPStr)] string lpszPassword,
            IntPtr ulUIParam,
            uint ulFlags,
            out IMsgServiceAdmin lppServiceAdmin);
    }

    // IMsgServiceAdmin. MAPIX.H order, every slot.
    [ComImport]
    [Guid("0002037A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsgServiceAdmin
    {
        [PreserveSig] int GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        [PreserveSig] int GetMsgServiceTable(uint ulFlags, out IntPtr lppTable);
        [PreserveSig] int CreateMsgService(
            [MarshalAs(UnmanagedType.LPStr)] string lpszService,
            [MarshalAs(UnmanagedType.LPStr)] string lpszDisplayName,
            IntPtr ulUIParam,
            uint ulFlags);
        [PreserveSig] int DeleteMsgService(byte[] lpuid);
        [PreserveSig] int CopyMsgService(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, uint g);
        [PreserveSig] int RenameMsgService(IntPtr a, uint b, IntPtr c);
        [PreserveSig] int ConfigureMsgService(
            byte[] lpuid,
            IntPtr ulUIParam,
            uint ulFlags,
            uint cValues,
            IntPtr lpProps);
        [PreserveSig] int OpenProfileSection(IntPtr a, IntPtr b, uint c, out IntPtr d);
        [PreserveSig] int MsgServiceTransportOrder(uint a, IntPtr b, uint c);
        [PreserveSig] int AdminProviders(IntPtr a, uint b, out IntPtr c);
        [PreserveSig] int SetPrimaryIdentity(IntPtr a, uint b);
        [PreserveSig] int GetProviderTable(uint ulFlags, out IntPtr lppTable);
    }

    // IMsgServiceAdmin2 = IMsgServiceAdmin + CreateMsgServiceEx. All twelve inherited slots are
    // repeated first because a derived COM interface's vtable starts with its base's.
    //
    // THE IID BELOW IS THE ONE THING IN THIS FILE I COULD NOT VERIFY AGAINST A PRIMARY SOURCE.
    // If it is wrong, QueryInterface returns E_NOINTERFACE and the callers fall back to
    // create-then-diff, which works on plain IMsgServiceAdmin. So a wrong GUID here costs a code
    // path, not a run - which is exactly why the fallback exists.
    [ComImport]
    [Guid("0002038A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsgServiceAdmin2
    {
        [PreserveSig] int GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        [PreserveSig] int GetMsgServiceTable(uint ulFlags, out IntPtr lppTable);
        [PreserveSig] int CreateMsgService(IntPtr a, IntPtr b, IntPtr c, uint d);
        [PreserveSig] int DeleteMsgService(byte[] lpuid);
        [PreserveSig] int CopyMsgService(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, uint g);
        [PreserveSig] int RenameMsgService(IntPtr a, uint b, IntPtr c);
        [PreserveSig] int ConfigureMsgService(byte[] lpuid, IntPtr ulUIParam, uint ulFlags, uint cValues, IntPtr lpProps);
        [PreserveSig] int OpenProfileSection(IntPtr a, IntPtr b, uint c, out IntPtr d);
        [PreserveSig] int MsgServiceTransportOrder(uint a, IntPtr b, uint c);
        [PreserveSig] int AdminProviders(IntPtr a, uint b, out IntPtr c);
        [PreserveSig] int SetPrimaryIdentity(IntPtr a, uint b);
        [PreserveSig] int GetProviderTable(uint ulFlags, out IntPtr lppTable);
        [PreserveSig] int CreateMsgServiceEx(
            [MarshalAs(UnmanagedType.LPStr)] string lpszService,
            [MarshalAs(UnmanagedType.LPStr)] string lpszDisplayName,
            IntPtr ulUIParam,
            uint ulFlags,
            byte[] lpuidService);
    }

    public static class Mapi
    {
        // ---- flags and property tags ------------------------------------------------------
        public const uint MAPI_DEFAULT_SERVICES = 0x00000001;
        public const uint MAPI_DIALOG           = 0x00000008;

        public const uint PR_DISPLAY_NAME_A  = 0x3001001E;
        public const uint PR_DISPLAY_NAME_W  = 0x3001001F;
        public const uint PR_SERVICE_NAME_A  = 0x3D09001E;
        public const uint PR_SERVICE_UID     = 0x3D0D0102;
        public const uint PR_DEFAULT_PROFILE = 0x3D04000B;
        public const uint PR_PST_PATH_W      = 0x6700001F;

        public const int MAPI_E_NO_ACCESS        = unchecked((int)0x80070005);
        public const int MAPI_E_NOT_FOUND        = unchecked((int)0x8004010F);
        public const int MAPI_E_NOT_INITIALIZED  = unchecked((int)0x80040605);
        public const int MAPI_E_FAILONEPROVIDER  = unchecked((int)0x8004011D);
        public const int MAPI_E_BAD_CHARWIDTH    = unchecked((int)0x80040103);
        public const int MAPI_E_NO_SUPPORT       = unchecked((int)0x80040102);
        public const int MAPI_E_UNCONFIGURED     = unchecked((int)0x8004011C);
        public const int MAPI_E_EXTENDED_ERROR   = unchecked((int)0x80040119);

        // ---- exports ----------------------------------------------------------------------
        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern int MAPIInitialize(IntPtr lpMapiInit);

        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern void MAPIUninitialize();

        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern int MAPIAdminProfiles(uint ulFlags, out IProfAdmin lppProfAdmin);

        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern void MAPIFreeBuffer(IntPtr lpBuffer);

        // HrQueryAllRows and FreeProws are MAPI utility exports. If either comes back as an
        // EntryPointNotFoundException the stub did not forward it by name, and the repair is to
        // declare IMAPITable's twenty-three slots and call SetColumns/QueryRows by hand. The
        // callers report that case by name rather than as a generic failure.
        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern int HrQueryAllRows(
            IntPtr lpTable, IntPtr lpPropTags, IntPtr lpRestriction,
            IntPtr lpSortOrderSet, int crowsMax, out IntPtr lppRows);

        [DllImport("mapi32.dll", CharSet = CharSet.Ansi)]
        public static extern void FreeProws(IntPtr lpRows);

        // ---- struct geometry --------------------------------------------------------------
        // SPropValue { ULONG ulPropTag; ULONG dwAlignPad; union _PV Value; }
        // The union's largest member is SBinary { ULONG cb; LPBYTE lpb; }, which is two
        // pointer-slots wide once padded. So: 8 header bytes, then two pointer slots.
        //   x64 -> 8 + 16 = 24     x86 -> 8 + 8 = 16
        // Computed rather than hardcoded so a 32-bit run fails honestly instead of silently.
        public static int SPropValueSize { get { return 8 + (2 * IntPtr.Size); } }

        // SRow { ULONG ulAdrEntryPad; ULONG cValues; LPSPropValue lpProps; }
        public static int SRowSize { get { return 8 + IntPtr.Size; } }

        // SRowSet { ULONG cRows; SRow aRow[]; } - aRow is pointer-aligned, so it starts at 8.
        public static int SRowSetHeaderSize { get { return 8; } }

        public const int SPropValueValueOffset = 8;

        /// <summary>Builds an SPropTagArray { ULONG cValues; ULONG aulPropTag[n] }.</summary>
        public static IntPtr AllocPropTagArray(uint[] tags)
        {
            IntPtr block = Marshal.AllocHGlobal(4 + (4 * tags.Length));
            Marshal.WriteInt32(block, 0, tags.Length);
            for (int i = 0; i < tags.Length; i++)
            {
                Marshal.WriteInt32(block, 4 + (4 * i), unchecked((int)tags[i]));
            }
            return block;
        }

        /// <summary>
        /// Builds an SPropValue array of PT_UNICODE strings. The wide strings are separate
        /// allocations; FreePropValueArray releases both halves.
        /// </summary>
        public static IntPtr AllocUnicodeStringProps(uint[] tags, string[] values, out IntPtr[] strings)
        {
            int size = SPropValueSize;
            IntPtr block = Marshal.AllocHGlobal(size * tags.Length);
            strings = new IntPtr[tags.Length];
            for (int i = 0; i < tags.Length; i++)
            {
                IntPtr at = new IntPtr(block.ToInt64() + (size * i));
                Marshal.WriteInt32(at, 0, unchecked((int)tags[i]));
                Marshal.WriteInt32(at, 4, 0);                       // dwAlignPad
                strings[i] = Marshal.StringToHGlobalUni(values[i]);
                Marshal.WriteIntPtr(at, SPropValueValueOffset, strings[i]);
            }
            return block;
        }

        public static void FreePropValueArray(IntPtr block, IntPtr[] strings)
        {
            if (strings != null)
            {
                for (int i = 0; i < strings.Length; i++)
                {
                    if (strings[i] != IntPtr.Zero) { Marshal.FreeHGlobal(strings[i]); }
                }
            }
            if (block != IntPtr.Zero) { Marshal.FreeHGlobal(block); }
        }

        // ---- SRowSet readers ---------------------------------------------------------------
        public static int RowCount(IntPtr rowSet)
        {
            if (rowSet == IntPtr.Zero) { return 0; }
            return Marshal.ReadInt32(rowSet, 0);
        }

        private static IntPtr RowAt(IntPtr rowSet, int index)
        {
            return new IntPtr(rowSet.ToInt64() + SRowSetHeaderSize + ((long)SRowSize * index));
        }

        private static IntPtr PropAt(IntPtr row, int index)
        {
            IntPtr props = Marshal.ReadIntPtr(row, 8);
            return new IntPtr(props.ToInt64() + ((long)SPropValueSize * index));
        }

        public static int PropCount(IntPtr rowSet, int rowIndex)
        {
            return Marshal.ReadInt32(RowAt(rowSet, rowIndex), 4);
        }

        public static uint PropTag(IntPtr rowSet, int rowIndex, int propIndex)
        {
            return unchecked((uint)Marshal.ReadInt32(PropAt(RowAt(rowSet, rowIndex), propIndex), 0));
        }

        /// <summary>PT_STRING8 payload, or null when the slot holds an error/absent value.</summary>
        public static string PropAnsiString(IntPtr rowSet, int rowIndex, int propIndex)
        {
            IntPtr prop = PropAt(RowAt(rowSet, rowIndex), propIndex);
            uint tag = unchecked((uint)Marshal.ReadInt32(prop, 0));
            if ((tag & 0xFFFF) != 0x001E) { return null; }
            IntPtr text = Marshal.ReadIntPtr(prop, SPropValueValueOffset);
            if (text == IntPtr.Zero) { return null; }
            return Marshal.PtrToStringAnsi(text);
        }

        /// <summary>PT_BOOLEAN payload as a nullable-ish int: -1 when the slot is not PT_BOOLEAN.</summary>
        public static int PropBoolean(IntPtr rowSet, int rowIndex, int propIndex)
        {
            IntPtr prop = PropAt(RowAt(rowSet, rowIndex), propIndex);
            uint tag = unchecked((uint)Marshal.ReadInt32(prop, 0));
            if ((tag & 0xFFFF) != 0x000B) { return -1; }
            return Marshal.ReadInt16(prop, SPropValueValueOffset) == 0 ? 0 : 1;
        }

        /// <summary>PT_BINARY payload (SBinary { ULONG cb; LPBYTE lpb; }), or null.</summary>
        public static byte[] PropBinary(IntPtr rowSet, int rowIndex, int propIndex)
        {
            IntPtr prop = PropAt(RowAt(rowSet, rowIndex), propIndex);
            uint tag = unchecked((uint)Marshal.ReadInt32(prop, 0));
            if ((tag & 0xFFFF) != 0x0102) { return null; }
            int cb = Marshal.ReadInt32(prop, SPropValueValueOffset);
            IntPtr lpb = Marshal.ReadIntPtr(prop, SPropValueValueOffset + IntPtr.Size);
            if (lpb == IntPtr.Zero || cb <= 0) { return null; }
            byte[] bytes = new byte[cb];
            Marshal.Copy(lpb, bytes, 0, cb);
            return bytes;
        }

        /// <summary>QueryInterface for IMsgServiceAdmin2; null when the interface is absent.</summary>
        public static IMsgServiceAdmin2 AsServiceAdmin2(IMsgServiceAdmin admin)
        {
            try { return (IMsgServiceAdmin2)admin; }
            catch (InvalidCastException) { return null; }
        }
    }
}
'@

    Add-Type -TypeDefinition $mapiSource -Language CSharp
}

# ---------------------------------------------------------------------------------------------
# PowerShell wrappers. Every one of them checks its own result; none of them trusts an HRESULT
# alone where a table read can settle the same question.
# ---------------------------------------------------------------------------------------------

function ConvertTo-MapiErrorText {
    param([int] $HResult)

    $known = @{
        ([OutlookAI.Testbed.Mapi]::MAPI_E_NO_ACCESS)       = 'MAPI_E_NO_ACCESS - the name is already taken, or the profile is in use'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_NOT_FOUND)       = 'MAPI_E_NOT_FOUND - no such profile or service'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_NOT_INITIALIZED) = 'MAPI_E_NOT_INITIALIZED - MAPIInitialize was not called on this thread'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_FAILONEPROVIDER) = 'MAPI_E_FAILONEPROVIDER - a provider refused; on a PST this is usually the path (casing included)'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_BAD_CHARWIDTH)   = 'MAPI_E_BAD_CHARWIDTH - ANSI/Unicode mismatch; MAPI_UNICODE is not supported on the admin calls'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_NO_SUPPORT)      = 'MAPI_E_NO_SUPPORT - this provider does not implement the call'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_UNCONFIGURED)    = 'MAPI_E_UNCONFIGURED - the service exists but was never configured'
        ([OutlookAI.Testbed.Mapi]::MAPI_E_EXTENDED_ERROR)  = 'MAPI_E_EXTENDED_ERROR - call GetLastError for detail'
    }

    if ($known.ContainsKey($HResult)) { return $known[$HResult] }
    return ('0x{0:X8}' -f $HResult)
}

function Assert-MapiOk {
    param([int] $HResult, [string] $What)

    if ($HResult -ne 0) {
        throw "$What failed: $(ConvertTo-MapiErrorText -HResult $HResult)"
    }
}

<#
    Opens a MAPI admin session and hands it to a script block. MAPIInitialize and
    MAPIUninitialize are paired in a finally, which is not optional: an unpaired Initialize
    leaves the process holding MAPI for its lifetime.
#>
function Invoke-WithProfAdmin {
    param([Parameter(Mandatory = $true)] [scriptblock] $Body)

    $initialised = $false
    $admin = $null
    try {
        Assert-MapiOk -HResult ([OutlookAI.Testbed.Mapi]::MAPIInitialize([IntPtr]::Zero)) -What 'MAPIInitialize'
        $initialised = $true

        $admin = $null
        Assert-MapiOk -HResult ([OutlookAI.Testbed.Mapi]::MAPIAdminProfiles(0, [ref] $admin)) -What 'MAPIAdminProfiles'

        & $Body $admin
    }
    finally {
        if ($null -ne $admin) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($admin)
        }
        if ($initialised) {
            [OutlookAI.Testbed.Mapi]::MAPIUninitialize()
        }
    }
}

<#
    Reads a MAPI table into PowerShell objects. Columns are requested by tag, so the property
    order in each row matches the order asked for.

    -Kind is only used to make the failure message say which table could not be read.
#>
function Read-MapiTable {
    param(
        [Parameter(Mandatory = $true)] [IntPtr] $Table,
        [Parameter(Mandatory = $true)] [uint32[]] $Tags,
        [string] $Kind = 'table'
    )

    $tagArray = [IntPtr]::Zero
    $rows = [IntPtr]::Zero
    try {
        $tagArray = [OutlookAI.Testbed.Mapi]::AllocPropTagArray($Tags)

        $hr = 0
        try {
            $hr = [OutlookAI.Testbed.Mapi]::HrQueryAllRows($Table, $tagArray, [IntPtr]::Zero, [IntPtr]::Zero, 0, [ref] $rows)
        }
        catch [EntryPointNotFoundException] {
            throw @"
mapi32.dll did not export HrQueryAllRows by name, so the $Kind cannot be read this way.

This is a KNOWN possible outcome, not a mystery: the MAPI stub forwards some utility functions by
ordinal only. The repair is to declare IMAPITable's twenty-three vtable slots in
OutlookMapiInterop.ps1 and call SetColumns/QueryRows by hand. Nothing was written before this
point, so the profile is unchanged.
"@
        }
        Assert-MapiOk -HResult $hr -What "HrQueryAllRows over the $Kind"

        $result = @()
        $count = [OutlookAI.Testbed.Mapi]::RowCount($rows)
        for ($r = 0; $r -lt $count; $r++) {
            $record = [ordered]@{}
            $propCount = [OutlookAI.Testbed.Mapi]::PropCount($rows, $r)
            for ($p = 0; $p -lt $propCount; $p++) {
                $tag = [OutlookAI.Testbed.Mapi]::PropTag($rows, $r, $p)
                $type = $tag -band 0xFFFF
                $value = $null
                if ($type -eq 0x001E) { $value = [OutlookAI.Testbed.Mapi]::PropAnsiString($rows, $r, $p) }
                elseif ($type -eq 0x000B) { $value = [OutlookAI.Testbed.Mapi]::PropBoolean($rows, $r, $p) }
                elseif ($type -eq 0x0102) { $value = [OutlookAI.Testbed.Mapi]::PropBinary($rows, $r, $p) }
                $record[('0x{0:X8}' -f $tag)] = $value
            }
            $result += [pscustomobject] $record
        }
        return , $result
    }
    finally {
        if ($rows -ne [IntPtr]::Zero) { [OutlookAI.Testbed.Mapi]::FreeProws($rows) }
        if ($tagArray -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($tagArray) }
    }
}

<#
    The profile table, as objects with Name and IsDefault.

    THIS IS THE ONLY TRUSTWORTHY ANSWER TO 'does that profile exist'. DeleteProfile returns S_OK
    without deleting when the profile is in use, and CreateProfile's HRESULT says nothing about
    what the table ended up holding.
#>
function Get-MapiProfile {
    param([Parameter(Mandatory = $true)] $ProfAdmin)

    $table = [IntPtr]::Zero
    Assert-MapiOk -HResult $ProfAdmin.GetProfileTable(0, [ref] $table) -What 'IProfAdmin::GetProfileTable'
    try {
        $nameTag = [OutlookAI.Testbed.Mapi]::PR_DISPLAY_NAME_A
        $defaultTag = [OutlookAI.Testbed.Mapi]::PR_DEFAULT_PROFILE
        $rows = Read-MapiTable -Table $table -Tags @($nameTag, $defaultTag) -Kind 'profile table'

        $nameKey = '0x{0:X8}' -f $nameTag
        $defaultKey = '0x{0:X8}' -f $defaultTag
        $out = @()
        foreach ($row in $rows) {
            $out += [pscustomobject]@{
                Name      = $row.$nameKey
                IsDefault = ($row.$defaultKey -eq 1)
            }
        }
        return , $out
    }
    finally {
        if ($table -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($table) }
    }
}

<#
    The message-service table for one profile, as objects with ServiceName, DisplayName and Uid.
    Used both to verify a store was added under the name asked for, and to find the UID of a
    service CreateMsgService just made without telling us.
#>
function Get-MapiService {
    param([Parameter(Mandatory = $true)] $ServiceAdmin)

    $table = [IntPtr]::Zero
    Assert-MapiOk -HResult $ServiceAdmin.GetMsgServiceTable(0, [ref] $table) -What 'IMsgServiceAdmin::GetMsgServiceTable'
    try {
        $uidTag = [OutlookAI.Testbed.Mapi]::PR_SERVICE_UID
        $svcTag = [OutlookAI.Testbed.Mapi]::PR_SERVICE_NAME_A
        $nameTag = [OutlookAI.Testbed.Mapi]::PR_DISPLAY_NAME_A
        $rows = Read-MapiTable -Table $table -Tags @($uidTag, $svcTag, $nameTag) -Kind 'message service table'

        $uidKey = '0x{0:X8}' -f $uidTag
        $svcKey = '0x{0:X8}' -f $svcTag
        $nameKey = '0x{0:X8}' -f $nameTag
        $out = @()
        foreach ($row in $rows) {
            $out += [pscustomobject]@{
                Uid         = $row.$uidKey
                ServiceName = $row.$svcKey
                DisplayName = $row.$nameKey
            }
        }
        return , $out
    }
    finally {
        if ($table -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($table) }
    }
}

<#
    Normalises a PST path ONCE, and everything downstream reuses the exact string this returns.

    Why it matters: the same file spelled with different casing in two profiles is reported to
    give MAPI_E_FAILONEPROVIDER. GetFullPathName is what settles the spelling; it does not require
    the file to exist, which is correct here because the provider creates it.
#>
function Resolve-PstPath {
    param([Parameter(Mandatory = $true)] [string] $Path)

    $full = [System.IO.Path]::GetFullPath($Path)
    $directory = [System.IO.Path]::GetDirectoryName($full)
    if (-not (Test-Path -LiteralPath $directory)) {
        throw "The directory for the PST does not exist: $directory. Create it first; the provider creates the .pst, not the folder."
    }
    return $full
}

<#
    Adds a PST message service to a profile and configures it with an EXACT display name.

    The display name is set AT CREATION TIME, through ConfigureMsgService's PR_DISPLAY_NAME. That
    is the whole reason this file exists rather than a two-line AddStoreEx call: the object model
    cannot name a store, and this can.

    Returns the service UID. Verification is the CALLER's job and every caller does it - this
    function deliberately does not decide whether the name it read back is acceptable.
#>
function Add-MapiPstService {
    param(
        [Parameter(Mandatory = $true)] $ServiceAdmin,
        [Parameter(Mandatory = $true)] [string] $PstPath,
        [Parameter(Mandatory = $true)] [string] $DisplayName,
        [string] $ServiceName = 'MSUPST MS'
    )

    $before = Get-MapiService -ServiceAdmin $ServiceAdmin
    $beforeUids = @()
    foreach ($service in $before) {
        if ($null -ne $service.Uid) { $beforeUids += [System.BitConverter]::ToString($service.Uid) }
    }

    $uid = New-Object 'byte[]' 16
    $usedEx = $false
    $admin2 = [OutlookAI.Testbed.Mapi]::AsServiceAdmin2($ServiceAdmin)

    if ($null -ne $admin2) {
        $hr = $admin2.CreateMsgServiceEx($ServiceName, $DisplayName, [IntPtr]::Zero, 0, $uid)
        Assert-MapiOk -HResult $hr -What "IMsgServiceAdmin2::CreateMsgServiceEx('$ServiceName')"
        $usedEx = $true
    }
    else {
        $hr = $ServiceAdmin.CreateMsgService($ServiceName, $DisplayName, [IntPtr]::Zero, 0)
        Assert-MapiOk -HResult $hr -What "IMsgServiceAdmin::CreateMsgService('$ServiceName')"

        # CreateMsgService does not hand back the UID it made, so diff the table.
        $after = Get-MapiService -ServiceAdmin $ServiceAdmin
        $new = @()
        foreach ($service in $after) {
            if ($null -eq $service.Uid) { continue }
            if ($beforeUids -notcontains [System.BitConverter]::ToString($service.Uid)) { $new += $service }
        }
        if ($new.Count -ne 1) {
            throw "Expected exactly one new message service after CreateMsgService; the table gained $($new.Count). The profile is now in a state this script did not intend - revert the checkpoint."
        }
        $uid = $new[0].Uid
    }

    Write-Host ("  created via {0}" -f $(if ($usedEx) { 'IMsgServiceAdmin2::CreateMsgServiceEx' } else { 'IMsgServiceAdmin::CreateMsgService + table diff' }))

    # PT_UNICODE payloads on purpose - see constraint 5 in the header. These go to the PST
    # provider, not through the ANSI-only admin API.
    $strings = $null
    $props = [IntPtr]::Zero
    try {
        $props = [OutlookAI.Testbed.Mapi]::AllocUnicodeStringProps(
            @([OutlookAI.Testbed.Mapi]::PR_PST_PATH_W, [OutlookAI.Testbed.Mapi]::PR_DISPLAY_NAME_W),
            @($PstPath, $DisplayName),
            [ref] $strings)

        $hr = $ServiceAdmin.ConfigureMsgService($uid, [IntPtr]::Zero, 0, 2, $props)
        Assert-MapiOk -HResult $hr -What "IMsgServiceAdmin::ConfigureMsgService('$DisplayName')"
    }
    finally {
        [OutlookAI.Testbed.Mapi]::FreePropValueArray($props, $strings)
    }

    return $uid
}

<#
    Refuses while Outlook is running.

    It does NOT kill it. Mailbox-safety rule 7 forbids taskkill on OUTLOOK.EXE outright, and a
    running Outlook writes its own view of the profile back at shutdown - so killing it here would
    trade a clean refusal for a silent revert an hour later.
#>
function Assert-OutlookNotRunning {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw @"
REFUSING: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

Profile administration while Outlook is running is not reliable: the profile is read at logon and
a running Outlook writes its own view back when it closes, which silently reverts whatever this
script did.

Close Outlook properly and run this again. DO NOT taskkill it - mailbox-safety rule 7 forbids
that outright, and a forced kill can leave the profile mid-write.
"@
    }
}
