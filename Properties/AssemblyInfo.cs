using System.Reflection;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("OutlookAI")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("OutlookAI")]
[assembly: AssemblyCopyright("Copyright ©  2026")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("78af2871-0ceb-4451-b80d-455552e37c91")]

// Every [DllImport] in this assembly loads its DLL from System32 and nowhere else (CA5392, Q125).
// Each one names a Windows DLL - kernel32, user32, gdi32, shell32, advapi32, wintrust, and
// wtsapi32 in the helper - and without this the search starts in the program's own folder. The
// add-in is installed per user, under %LOCALAPPDATA%, where the user's own processes can write,
// and this file is compiled into OutlookAI.PolicyWriter.exe as well: the helper that runs ELEVATED
// from that same folder. wtsapi32.dll is not one of Windows' KnownDLLs, so a copy planted beside
// the helper would have been loaded into it, as administrator.
[assembly: DefaultDllImportSearchPaths(DllImportSearchPath.System32)]

[assembly: AssemblyVersion("99.99.99.0")]
[assembly: AssemblyFileVersion("99.99.99.0")]

