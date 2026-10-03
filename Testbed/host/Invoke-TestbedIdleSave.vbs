' Runs Invoke-TestbedIdleSave.ps1 with no window at all, for the OutlookAI-TestbedIdleSave
' scheduled task (Register-IdleSaveTask.ps1 registers this file as the task's action).
'
' WHY THIS EXISTS. The task used to start powershell.exe directly. powershell.exe is a
' console-subsystem program, so Windows creates a console window for it first - on the
' maintainer's workstation a Windows Terminal window - and creating it activates it: every
' 15 minutes it took the foreground and keyboard focus for 1-2 seconds, once 141 ms after a
' keystroke (measured 2026-10-03 by another session with SpawnSpotter, 9 runs of 9).
' -WindowStyle Hidden cannot prevent that: it is applied after the console already exists.
'
' wscript.exe is a GUI-subsystem host, so it never allocates a console. Run(..., 0, True)
' starts PowerShell with SW_HIDE supplied at creation, so no window is ever created, and
' waits for it, so the task's last result is the script's own exit code. The same pattern
' runs the maintainer's "Patch Claude Code Extension" task. A task with LogonType S4U would
' need no wrapper, but registering one needs administrator rights.
Option Explicit
Dim shell, fso, folder, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ _
    & folder & "\Invoke-TestbedIdleSave.ps1"""
WScript.Quit shell.Run(command, 0, True)
