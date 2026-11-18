using System.Diagnostics;
using System.Reflection;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Only an OUTLOOK.EXE in the caller's own Windows session is this user's Outlook (decided 2026-11-18). Another
/// user signed in to the same machine runs their own, which this server can neither attach to nor see the windows
/// of: counted as this user's, it read as an Outlook starting with no windows, and after two minutes as a stuck
/// one, with the other user's process ID for this user to end.
/// </summary>
public sealed class OutlookProcessesTests
{
    [Fact]
    public void AnotherSessionsOutlook_IsNotThisUsers()
    {
        List<int> mine = OutlookProcesses.InSession(Processes((4100, 1), (5200, 2), (6300, 1)), 1);

        Assert.Equal(new[] { 4100, 6300 }, mine);
    }

    [Fact]
    public void OnlyAnotherUsersOutlook_MeansThisUserHasNone()
    {
        Assert.Empty(OutlookProcesses.InSession(Processes((5200, 2)), 1));
    }

    [Fact]
    public void AnOutlookWhoseSessionCouldNotBeRead_IsNotCounted()
    {
        // This user's own Outlook can always be read; one that cannot is another user's or already gone.
        List<int> mine = OutlookProcesses.InSession(
            new[] { new KeyValuePair<int, int?>(4100, null), new KeyValuePair<int, int?>(6300, 1) }, 1);

        Assert.Equal(new[] { 6300 }, mine);
    }

    [Fact]
    public void WithTheCallersOwnSessionUnknown_EveryOutlookCounts_AsBeforeTheRule()
    {
        // Rather than report no Outlook at all - which would have the server start one beside it.
        List<int> all = OutlookProcesses.InSession(Processes((4100, 1), (5200, 2)), null);

        Assert.Equal(new[] { 4100, 5200 }, all);
    }

    [Fact]
    public void TheCallersSession_IsTheSessionThisProcessRunsIn()
    {
        using Process self = Process.GetCurrentProcess();

        Assert.Equal(self.SessionId, OutlookProcesses.CallerSessionId);
    }

    [Fact]
    public void EveryOutlookThisMachineLists_InThisSession_IsInThisSession()
    {
        Process[] mine = OutlookProcesses.InThisSession();
        try
        {
            Assert.All(mine, process => Assert.Equal(OutlookProcesses.CallerSessionId, process.SessionId));
        }
        finally
        {
            foreach (Process process in mine)
            {
                process.Dispose();
            }
        }
    }

    [Fact]
    public void TheLivenessCheck_AndIsOutlookRunning_AskOnlyForThisSessionsOutlook()
    {
        Assert.Contains(IlReader.CalleesOf(Method(typeof(OutlookLiveness), nameof(OutlookLiveness.Snapshot))), IsOutlookProcesses);
        Assert.Contains(IlReader.CalleesOf(Method(typeof(OutlookComSession), nameof(OutlookComSession.IsOutlookProcessRunning))), IsOutlookProcesses);
    }

    [Fact]
    public void NoServerCode_ListsOutlookProcesses_ButThroughTheSessionFilter()
    {
        // The rule holds by construction: a lookup of OUTLOOK.EXE anywhere else would count another user's.
        // Process.GetProcessesByName stays allowed in the one generic helper, which health uses for the search
        // indexer - a service, in session 0 for every user - and the image name itself only here.
        Assembly[] server = { typeof(OutlookProcesses).Assembly, typeof(ComHostSupervisor).Assembly, typeof(OutlookTools).Assembly };
        List<string> listers = new();
        List<string> namers = new();
        foreach (Type type in server.SelectMany(a => a.GetTypes()))
        {
            foreach (MethodBase body in IlReader.BodiesOf(type))
            {
                if (IlReader.CalleesOf(body).Any(m => m.DeclaringType == typeof(Process) && m.Name == nameof(Process.GetProcessesByName)))
                {
                    listers.Add(type.FullName + "." + body.Name);
                }
            }

            if (IlReader.StringsOf(type).Contains("OUTLOOK"))
            {
                namers.Add(type.FullName!);
            }
        }

        Assert.Equal(
            new[] { typeof(OutlookProcesses).FullName + "." + nameof(OutlookProcesses.InThisSession), typeof(HealthReporting).FullName + "." + nameof(HealthReporting.TryIsProcessRunning) }.OrderBy(n => n, StringComparer.Ordinal),
            listers.Distinct().OrderBy(n => n, StringComparer.Ordinal));
        Assert.Equal(new[] { typeof(OutlookProcesses).FullName }, namers.Distinct());
    }

    private static bool IsOutlookProcesses(MethodBase method) => method.DeclaringType == typeof(OutlookProcesses);

    private static MethodBase Method(Type type, string name)
    {
        return type.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static, Type.EmptyTypes)
            ?? throw new InvalidOperationException(type.Name + "." + name + " was not found");
    }

    private static IEnumerable<KeyValuePair<int, int?>> Processes(params (int Id, int Session)[] processes)
    {
        return processes.Select(p => new KeyValuePair<int, int?>(p.Id, p.Session));
    }
}
