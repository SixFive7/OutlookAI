using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using OutlookAI.Core.Audit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Sends every audit line this test process writes to a throwaway directory of its own, before
/// any test runs (Q86, decided 2026-10-03).
/// <para>
/// <b>Why it exists.</b> The non-live suite drives the product's real write paths through fakes
/// - <c>T1/AtomicityClaimsTests</c>, <c>T1/DraftUpdateReentrancyTests</c> and others - and every one
/// of them ends in <see cref="AuditLog.Append"/>, which wrote to the maintainer's REAL
/// <c>%LOCALAPPDATA%\OutlookAI\audit.log</c>. Re-counted 2026-09-28: at least 15,739 of its 24,413
/// lines carried fake EntryIDs. Every agent runs this suite on that workstation, and his own
/// OutlookAI server appends to the same file at any moment, including during a run.
/// </para>
/// <para>
/// <b>How.</b> A module initializer, because it is the one hook the runtime guarantees to run
/// before ANY code in this assembly - every test class, fixture, data source and the collection
/// orderer included - and xunit 2 has no assembly fixture. It calls
/// <see cref="AuditLog.RedirectThisProcess"/>, reached through <c>InternalsVisibleTo</c>; nothing
/// outside a test assembly can, so this is not a product setting and the product's default path
/// is unchanged. The directory is under <c>%TEMP%</c>, unique to this process (its id plus a GUID,
/// so two worktrees testing at once never share one), and deleted when the process exits. A
/// process killed before that leaves a few kilobytes in <c>%TEMP%\OutlookAI-TestAudit</c>, which
/// nothing reads.
/// </para>
/// <para>
/// <b>What it cannot reach</b>, and does not pretend to: a process this one STARTS. The T3 tier
/// runs the real server executable, which appends to the real log because nothing may redirect a
/// shipped process (no setting, no environment variable). That tier stays off the log by never
/// making the server write one - see <c>McpStdioClient</c>'s audit-log refusal.
/// </para>
/// <para>
/// Proven by <c>T1/AuditLogIsolationTests</c>, which never opens the real file: it checks the
/// effective directory, appends a sentinel and finds it here, and reads the compiled IL of every
/// test for anything that names the real log.
/// </para>
/// </summary>
internal static class AuditIsolation
{
    /// <summary>The folder under <c>%TEMP%</c> that holds one throwaway audit directory per test process.</summary>
    internal const string ParentFolderName = "OutlookAI-TestAudit";

    private static string? _throwawayDirectory;

    /// <summary>The throwaway directory THIS test process's audit lines go to.</summary>
    internal static string ThrowawayDirectory => _throwawayDirectory
        ?? throw new InvalidOperationException(
            "The test assembly's module initializer never redirected the audit log, so audit lines from this run "
            + "would reach the real log. Nothing in this assembly may run before it - find out what did.");

    [ModuleInitializer]
    [SuppressMessage(
        "Usage",
        "CA2255:The 'ModuleInitializer' attribute should not be used in libraries",
        Justification = "A test assembly is the one library where running first is the point: no test may run before the redirect.")]
    internal static void RedirectThisProcessesAuditLog()
    {
        string directory = Path.Combine(
            Path.GetTempPath(),
            ParentFolderName,
            Environment.ProcessId.ToString(CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));

        // Throws, and so fails the assembly load and with it every test, if the directory were
        // ever the real one or not throwaway. That is the intended failure mode: a run that
        // cannot be isolated must not run at all.
        AuditLog.RedirectThisProcess(directory);
        _throwawayDirectory = directory;

        AppDomain.CurrentDomain.ProcessExit += (_, _) => TryDelete(directory);
    }

    private static void TryDelete(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
            // Best effort at exit: a leftover directory in %TEMP% costs a few kilobytes and is
            // read by nothing, which is not worth failing a finished run over.
        }
    }
}
