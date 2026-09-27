#Requires -Version 5.1
<#
.SYNOPSIS
    Imports the running PowerShell's OWN Security and Utility modules, by path. Dot-sourced near
    the top of every host script that uses a cmdlet from either, before it uses one.

.DESCRIPTION
    MEASURED 2026-09-27 on the maintainer's workstation (Windows 11 10.0.26200, PowerShell 7.6.6,
    Windows PowerShell 5.1.26100): Windows PowerShell 5.1 started by Start-Process from PowerShell 7
    inherits 7's PSModulePath, which lists 7's own module folder ahead of 5.1's, so 5.1 resolves
    these two modules to 7's copies:

      * Microsoft.PowerShell.Security does not load at all ("The member AuditToString is already
        present"): no Cert: drive, no ConvertTo-SecureString, no Get-AuthenticodeSignature.
      * Microsoft.PowerShell.Utility loads without the script half of 5.1's, which is where 5.1
        keeps Get-FileHash, New-TemporaryFile, New-Guid, Format-Hex, Import-PowerShellDataFile and
        ConvertFrom-SddlString.

    Nothing fails until the first such cmdlet runs, and then the script stops on "not recognized".
    5.1 started any other way - directly from a PowerShell 7 prompt, which resets the path, or from
    cmd, Explorer or a scheduled task - has both modules, and so does PowerShell 7. Importing the
    running edition's own copy by path, from $PSHOME, is right in every one of those cases: where
    that copy is already loaded, the import changes nothing.

    Tools/Switch-AddInBuild.ps1 found the Security half first (Q81) and imports that module the
    same way, inline, because it lives outside Testbed/. Nothing here prints or reads anything.

    Dot-source it AFTER $ErrorActionPreference = 'Stop', so a module that will not load stops the
    script here, by name, instead of at a cmdlet much later:

        . (Join-Path $PSScriptRoot 'OwnEditionModules.ps1')
#>

foreach ($ownEditionModuleName in @('Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility')) {
    Import-Module (Join-Path $PSHOME "Modules\$ownEditionModuleName\$ownEditionModuleName.psd1")
}
