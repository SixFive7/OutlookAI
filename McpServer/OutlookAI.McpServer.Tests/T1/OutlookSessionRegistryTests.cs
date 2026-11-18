using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The per-user session registry (Q118 (a), approved 2026-10-11; Q208 B and Q175 (b), decided 2026-11-18): every
/// OutlookAI COM session joins it, an Outlook with no window has ONE lifetime pin that every session shares, the last
/// session out closes it, a session that hears Quit closes it, and a compose window is promoted once per Outlook
/// process. The rules are pure (<see cref="SessionRegistryRules"/>) and the store is a file under a named mutex
/// (<see cref="OutlookSessionRegistry"/>); both are pinned here without an Outlook.
/// </summary>
public sealed class OutlookSessionRegistryTests
{
    private static readonly ProcessStamp Outlook = new(4100, 638_000_000_000_000_000);
    private static readonly ProcessStamp OtherOutlook = new(4200, 638_100_000_000_000_000);
    private static readonly ProcessStamp HostA = new(5100, 638_000_000_100_000_000);
    private static readonly ProcessStamp HostB = new(5200, 638_000_000_200_000_000);
    private static readonly ProcessStamp ThisTestProcess = new(5300, 638_000_000_300_000_000);

    private static RegisteredSession Session(ProcessStamp host, int sequence, ProcessStamp outlook, bool startedOutlook = false)
    {
        return new RegisteredSession(
            host.Pid.ToString(CultureInfo.InvariantCulture) + "." + host.StartedUtcTicks.ToString(CultureInfo.InvariantCulture) + "." + sequence.ToString(CultureInfo.InvariantCulture),
            host,
            outlook,
            sinceUtcTicks: 638_000_000_400_000_000 + (host.Pid * 1000L) + sequence,
            startedOutlook);
    }

    // ------------------------------------------------------------------ the file's text

    [Fact]
    public void FormatAndParse_RoundTripEveryField()
    {
        SessionRegistryState state = new();
        RegisteredSession a = Session(HostA, 1, Outlook, startedOutlook: true);
        SessionRegistryRules.Join(state, a);
        SessionRegistryRules.Join(state, Session(HostB, 2, Outlook));
        SessionRegistryRules.PinMade(state, Outlook, 0x000A_0B2C, a.Key, 638_000_000_500_000_000);
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, editorObtained: true, 638_000_000_600_000_000);
        state.FindOrAdd(OtherOutlook).DirectRefused = true;

        SessionRegistryState back = SessionRegistryRules.Parse(SessionRegistryRules.Format(state));

        RegisteredOutlook o = Assert.IsType<RegisteredOutlook>(back.Find(Outlook));
        Assert.Equal(a.Key, o.StartedBy);
        Assert.Equal(0x000A_0B2C, o.PinWindow);
        Assert.Equal(a.Key, o.PinMadeBy);
        Assert.Equal(638_000_000_500_000_000, o.PinSinceUtcTicks);
        Assert.Equal(638_000_000_600_000_000, o.PromotedUtcTicks);
        Assert.False(o.DirectRefused);
        Assert.True(back.Find(OtherOutlook)!.DirectRefused);
        Assert.Equal(new[] { a.Key, Session(HostB, 2, Outlook).Key }, back.Sessions.Select(s => s.Key).ToArray());
        RegisteredSession readA = back.Sessions[0];
        Assert.Equal(HostA, readA.Process);
        Assert.Equal(Outlook, readA.Outlook);
        Assert.True(readA.StartedOutlook);
        Assert.False(back.Sessions[1].StartedOutlook);
        Assert.Equal(a.SinceUtcTicks, readA.SinceUtcTicks);
    }

    [Fact]
    public void Parse_DropsWhatItCannotRead_AndKeepsTheRest()
    {
        string text = string.Join("\n", new[]
        {
            SessionRegistryRules.Header,
            "garbage line",
            "outlook stamp=notastamp pin=12",
            "outlook stamp=4100@638000000000000000 pin=77 pinBy=5100.1.1 pinSince=9",
            "outlook stamp=4100@638000000000000000 pin=88",
            "session key=5100.638000000100000000.1 process=5100@638000000100000000 outlook=4100@638000000000000000 since=5 startedOutlook=1",
            "session key=5100.638000000100000000.1 process=5100@638000000100000000 outlook=4100@638000000000000000 since=6 startedOutlook=0",
            "session key= process=5100@1 outlook=4100@1",
            "session key=x process=bad outlook=4100@1",
            "outlook stamp=4200@638100000000000000 pin=abc promoted=xyz",
            string.Empty,
        });

        SessionRegistryState state = SessionRegistryRules.Parse(text);

        Assert.Equal(2, state.Outlooks.Count);
        Assert.Equal(77, state.Find(Outlook)!.PinWindow);
        Assert.Equal(0, state.Find(OtherOutlook)!.PinWindow);
        Assert.Equal(0, state.Find(OtherOutlook)!.PromotedUtcTicks);
        RegisteredSession only = Assert.Single(state.Sessions);
        Assert.Equal(5, only.SinceUtcTicks);
        Assert.Empty(SessionRegistryRules.Parse(null).Sessions);
        Assert.Empty(SessionRegistryRules.Parse(string.Empty).Outlooks);
    }

    [Fact]
    public void ProcessStamp_ReadsBackItsOwnForm_AndRefusesAnythingElse()
    {
        ProcessStamp stamp = new(1234, 638_000_000_000_000_001);
        Assert.Equal("1234@638000000000000001", stamp.ToString());
        Assert.Equal(stamp, ProcessStamp.TryParse(stamp.ToString()));
        Assert.Null(ProcessStamp.TryParse(null));
        Assert.Null(ProcessStamp.TryParse("1234"));
        Assert.Null(ProcessStamp.TryParse("@5"));
        Assert.Null(ProcessStamp.TryParse("0@5"));
        Assert.Null(ProcessStamp.TryParse("-3@5"));
        Assert.Null(ProcessStamp.TryParse("12@x"));
    }

    [Fact]
    public void ProcessStamp_FindsThisProcessRunning_AndAStampOfAnEndedOneNot()
    {
        ProcessStamp self = ProcessStamp.Current();
        Assert.True(ProcessStamp.IsRunning(self));

        // The same id with another start time is a later process that reused it: not this one.
        Assert.False(ProcessStamp.IsRunning(new ProcessStamp(self.Pid, self.StartedUtcTicks + 1)));
        Assert.False(ProcessStamp.IsRunning(new ProcessStamp(0, 1)));
    }

    // ------------------------------------------------------------------ prune and join

    [Fact]
    public void Prune_DropsSessionsWhoseProcessIsGone_AndOutlooksThatAreGone()
    {
        SessionRegistryState state = new();
        SessionRegistryRules.Join(state, Session(HostA, 1, Outlook));
        SessionRegistryRules.Join(state, Session(HostB, 2, Outlook));
        SessionRegistryRules.Join(state, Session(HostA, 3, OtherOutlook));
        HashSet<ProcessStamp> running = new() { HostA, Outlook };

        int pruned = SessionRegistryRules.Prune(state, running.Contains, ThisTestProcess, liveInThisProcess: null);

        Assert.Equal(2, pruned);
        Assert.Equal(new[] { Session(HostA, 1, Outlook).Key }, state.Sessions.Select(s => s.Key).ToArray());
        Assert.Null(state.Find(OtherOutlook));
        Assert.NotNull(state.Find(Outlook));
    }

    [Fact]
    public void Prune_DropsASessionOfThisProcessThatThisProcessNoLongerHolds()
    {
        SessionRegistryState state = new();
        RegisteredSession open = Session(ThisTestProcess, 1, Outlook);
        RegisteredSession leaked = Session(ThisTestProcess, 2, Outlook);
        SessionRegistryRules.Join(state, open);
        SessionRegistryRules.Join(state, leaked);

        SessionRegistryRules.Prune(state, _ => true, ThisTestProcess, key => key == open.Key);

        Assert.Equal(new[] { open.Key }, state.Sessions.Select(s => s.Key).ToArray());
    }

    [Fact]
    public void Join_RecordsWhoStartedOutlook_Once()
    {
        SessionRegistryState state = new();
        RegisteredSession starter = Session(HostA, 1, Outlook, startedOutlook: true);
        SessionRegistryRules.Join(state, Session(HostB, 1, Outlook));
        Assert.Null(state.Find(Outlook)!.StartedBy);

        SessionRegistryRules.Join(state, starter);
        SessionRegistryRules.Join(state, Session(HostB, 2, Outlook, startedOutlook: true));

        Assert.Equal(starter.Key, state.Find(Outlook)!.StartedBy);
        Assert.Equal(3, state.Sessions.Count);

        // The same key again replaces, never duplicates.
        SessionRegistryRules.Join(state, starter);
        Assert.Equal(3, state.Sessions.Count);
    }

    // ------------------------------------------------------------------ leave: who closes the pin

    [Fact]
    public void Leave_TheLastSessionOut_ClosesThePin_WhoeverMadeIt_AndWhoeverStartedOutlook()
    {
        SessionRegistryState state = new();
        RegisteredSession starter = Session(HostA, 1, Outlook, startedOutlook: true);
        RegisteredSession drafter = Session(HostB, 1, Outlook);
        SessionRegistryRules.Join(state, starter);
        SessionRegistryRules.Join(state, drafter);
        SessionRegistryRules.PinMade(state, Outlook, 99, drafter.Key, 1);

        // The starter leaves first: before Q118 (a) its exit closed its pin and ended Outlook under the other
        // session (measured 2026-10-03). Now the pin stays while anyone is attached.
        LeaveVerdict first = SessionRegistryRules.Leave(state, starter.Key, Outlook, heardQuit: false);
        Assert.Equal(0, first.ClosePinWindow);
        Assert.Equal(1, first.OthersAttached);

        // The drafter leaves last, and did not start Outlook: before Q118 (a) its pin was left behind, keeping a
        // window-less Outlook up past Application.Quit() (3 of 3). Now the last session out closes it.
        LeaveVerdict last = SessionRegistryRules.Leave(state, drafter.Key, Outlook, heardQuit: false);
        Assert.Equal(99, last.ClosePinWindow);
        Assert.Equal(0, last.OthersAttached);
        Assert.Empty(state.Sessions);

        // The record goes only once the close happened.
        Assert.Equal(99, state.Find(Outlook)!.PinWindow);
        SessionRegistryRules.PinClosed(state, Outlook, last.ClosePinWindow);
        Assert.Equal(0, state.Find(Outlook)!.PinWindow);
        Assert.Null(state.Find(Outlook)!.PinMadeBy);
    }

    [Fact]
    public void Leave_AfterQuit_ClosesThePin_EvenUnderOtherSessions()
    {
        SessionRegistryState state = new();
        RegisteredSession a = Session(HostA, 1, Outlook);
        RegisteredSession b = Session(HostB, 1, Outlook);
        SessionRegistryRules.Join(state, a);
        SessionRegistryRules.Join(state, b);
        SessionRegistryRules.PinMade(state, Outlook, 42, a.Key, 1);

        LeaveVerdict verdict = SessionRegistryRules.Leave(state, b.Key, Outlook, heardQuit: true);

        Assert.Equal(42, verdict.ClosePinWindow);
        Assert.Equal(1, verdict.OthersAttached);
    }

    [Fact]
    public void Leave_WithNoPin_ClosesNothing_AndSessionsOfAnotherOutlookDoNotCount()
    {
        SessionRegistryState state = new();
        RegisteredSession a = Session(HostA, 1, Outlook);
        SessionRegistryRules.Join(state, a);
        SessionRegistryRules.Join(state, Session(HostB, 1, OtherOutlook));

        Assert.Equal(0, SessionRegistryRules.Leave(state, a.Key, Outlook, heardQuit: false).ClosePinWindow);

        SessionRegistryRules.Join(state, a);
        SessionRegistryRules.PinMade(state, Outlook, 7, a.Key, 1);
        LeaveVerdict verdict = SessionRegistryRules.Leave(state, a.Key, Outlook, heardQuit: false);
        Assert.Equal(7, verdict.ClosePinWindow);
        Assert.Equal(0, verdict.OthersAttached);
    }

    [Fact]
    public void ADeadSessionsPin_IsAdoptedWhileAnotherLives_AndClosedByTheLastOneOut()
    {
        SessionRegistryState state = new();
        RegisteredSession killed = Session(HostA, 1, Outlook, startedOutlook: true);
        RegisteredSession survivor = Session(HostB, 1, Outlook);
        SessionRegistryRules.Join(state, killed);
        SessionRegistryRules.Join(state, survivor);
        SessionRegistryRules.PinMade(state, Outlook, 31, killed.Key, 1);

        // HostA's COM host was killed (a deadline expiry): the next registry change prunes it, and its pin stays.
        SessionRegistryRules.Prune(state, s => s != HostA, ThisTestProcess, liveInThisProcess: null);
        Assert.Equal(31, state.Find(Outlook)!.PinWindow);

        LeaveVerdict verdict = SessionRegistryRules.Leave(state, survivor.Key, Outlook, heardQuit: false);
        Assert.Equal(31, verdict.ClosePinWindow);
    }

    [Fact]
    public void PinClosed_ClearsOnlyTheWindowItNames()
    {
        SessionRegistryState state = new();
        SessionRegistryRules.PinMade(state, Outlook, 5, "k", 1);
        SessionRegistryRules.PinClosed(state, Outlook, 6);
        Assert.Equal(5, state.Find(Outlook)!.PinWindow);
        SessionRegistryRules.PinClosed(state, Outlook, 0);
        Assert.Equal(0, state.Find(Outlook)!.PinWindow);
        SessionRegistryRules.PinClosed(state, OtherOutlook, 0);
        Assert.Null(state.Find(OtherOutlook));
    }

    // ------------------------------------------------------------------ the editor route (Q208 B)

    [Theory]
    [InlineData(true, false, false, DraftEditorRoute.Direct)]
    [InlineData(true, true, false, DraftEditorRoute.Direct)]
    [InlineData(false, false, false, DraftEditorRoute.Promote)]
    [InlineData(false, true, false, DraftEditorRoute.Direct)]
    [InlineData(false, true, true, DraftEditorRoute.Promote)]
    [InlineData(true, true, true, DraftEditorRoute.Promote)]
    public void TheEditorRoute_IsDirectWheneverAWindowShowsOrThisOutlookWasPromoted(bool showsWindow, bool promoted, bool directRefused, DraftEditorRoute expected)
    {
        RegisteredOutlook record = new(Outlook) { PromotedUtcTicks = promoted ? 1 : 0, DirectRefused = directRefused };
        Assert.Equal(expected, SessionRegistryRules.DecideEditorRoute(showsWindow, record));
    }

    [Fact]
    public void TheEditorRoute_WithNoRecord_PromotesOnlyAWindowlessOutlook()
    {
        Assert.Equal(DraftEditorRoute.Promote, SessionRegistryRules.DecideEditorRoute(false, null));
        Assert.Equal(DraftEditorRoute.Direct, SessionRegistryRules.DecideEditorRoute(true, null));
    }

    [Fact]
    public void OnePromotionPerOutlook_ThenDirect_AndARefusedDirectReadSendsTheNextDraftBack()
    {
        SessionRegistryState state = new();

        // The first draft in a window-less Outlook promotes - the one moment a window can show.
        Assert.Equal(DraftEditorRoute.Promote, SessionRegistryRules.DecideEditorRoute(false, state.Find(Outlook)));
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, editorObtained: true, 100);

        // Every later draft in that process reads directly (W6, W7: 120 of 120).
        for (int draft = 2; draft <= 10; draft++)
        {
            Assert.Equal(DraftEditorRoute.Direct, SessionRegistryRules.DecideEditorRoute(false, state.Find(Outlook)));
            SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Direct, editorObtained: true, 100 + draft);
        }

        Assert.Equal(100, state.Find(Outlook)!.PromotedUtcTicks);

        // A refused direct read is never left to refuse every draft after it.
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Direct, editorObtained: false, 200);
        Assert.Equal(DraftEditorRoute.Promote, SessionRegistryRules.DecideEditorRoute(false, state.Find(Outlook)));
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, editorObtained: true, 300);
        Assert.Equal(DraftEditorRoute.Direct, SessionRegistryRules.DecideEditorRoute(false, state.Find(Outlook)));
        Assert.Equal(100, state.Find(Outlook)!.PromotedUtcTicks);

        // A promotion that did not produce the editor records nothing: the next draft tries it again.
        SessionRegistryState other = new();
        SessionRegistryRules.EditorOutcome(other, OtherOutlook, DraftEditorRoute.Promote, editorObtained: false, 1);
        Assert.Equal(DraftEditorRoute.Promote, SessionRegistryRules.DecideEditorRoute(false, other.Find(OtherOutlook)));
    }

    [Fact]
    public void ANewOutlookProcess_StartsUnpromoted()
    {
        SessionRegistryState state = new();
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, editorObtained: true, 1);

        // Outlook restarted: a new process stamp, so its first window-less draft promotes again - once per start.
        Assert.Equal(DraftEditorRoute.Promote, SessionRegistryRules.DecideEditorRoute(false, state.Find(OtherOutlook)));
        SessionRegistryRules.Prune(state, s => s != Outlook, ThisTestProcess, null);
        Assert.Null(state.Find(Outlook));
    }

    // ------------------------------------------------------------------ what outlook_health says

    [Fact]
    public void WhatEndsOutlook_SaysWhatHoldsItAndWhatWillEndIt()
    {
        Assert.Equal("Outlook is not running.", SessionRegistryRules.DescribeWhatEndsOutlook(false, false, null, 0));
        Assert.Contains("closes its last window", SessionRegistryRules.DescribeWhatEndsOutlook(true, true, null, 3), StringComparison.Ordinal);

        RegisteredOutlook pinned = new(Outlook) { PinWindow = 9 };
        Assert.Contains("the last OutlookAI session to end closes it", SessionRegistryRules.DescribeWhatEndsOutlook(true, false, pinned, 2), StringComparison.Ordinal);
        Assert.Contains("no session attached", SessionRegistryRules.DescribeWhatEndsOutlook(true, false, pinned, 0), StringComparison.Ordinal);
        Assert.Contains("ends on its own", SessionRegistryRules.DescribeWhatEndsOutlook(true, false, new RegisteredOutlook(Outlook), 1), StringComparison.Ordinal);
        Assert.Contains("ends on its own", SessionRegistryRules.DescribeWhatEndsOutlook(true, false, null, 0), StringComparison.Ordinal);
    }

    [Fact]
    public void TheHealthView_ListsTheLiveSessionsOfThisOutlook_AndNamesALeftPin()
    {
        SessionRegistryState state = new();
        RegisteredSession maker = Session(HostA, 1, Outlook, startedOutlook: true);
        RegisteredSession reader = Session(HostB, 1, Outlook);
        SessionRegistryRules.Join(state, maker);
        SessionRegistryRules.Join(state, reader);
        SessionRegistryRules.Join(state, Session(HostB, 2, OtherOutlook));
        SessionRegistryRules.PinMade(state, Outlook, 0x1F00, maker.Key, new DateTime(2026, 11, 18, 9, 0, 0, DateTimeKind.Utc).Ticks);
        SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, true, new DateTime(2026, 11, 18, 9, 1, 0, DateTimeKind.Utc).Ticks);

        OutlookSessionsHealthView? view = HealthReporting.DescribeOutlookSessions(state, Outlook, false, s => true, HostB.Pid, out string? problem);

        Assert.Null(problem);
        Assert.NotNull(view);
        Assert.Equal(new[] { HostA.Pid, HostB.Pid }, view!.Attached.Select(r => r.Pid).ToArray());
        Assert.Equal(new[] { false, true }, view.Attached.Select(r => r.ThisServer).ToArray());
        Assert.True(view.Attached[0].StartedOutlook);
        Assert.True(view.StartedByOutlookAI);
        Assert.Equal("0x1F00", view.PinWindow);
        Assert.Equal(HostA.Pid, view.PinMadeByPid);
        Assert.True(view.PinMakerAttached);
        Assert.Equal(new DateTime(2026, 11, 18, 9, 1, 0, DateTimeKind.Utc), view.PromotedUtc);
        Assert.False(view.NextDraftPromotes);

        // Every session gone, the pin still there: a problem, with what ends it.
        OutlookSessionsHealthView? orphan = HealthReporting.DescribeOutlookSessions(state, Outlook, false, s => s != HostA && s != HostB, null, out string? orphanProblem);
        Assert.Empty(orphan!.Attached);
        Assert.False(orphan.PinMakerAttached);
        Assert.NotNull(orphanProblem);
        Assert.Contains("File > Exit", orphanProblem!, StringComparison.Ordinal);

        // A window the user has open is no problem, pin or not; and no Outlook, no block.
        Assert.NotNull(HealthReporting.DescribeOutlookSessions(state, Outlook, true, s => false, null, out string? windowed));
        Assert.Null(windowed);
        Assert.Null(HealthReporting.DescribeOutlookSessions(state, null, false, s => true, null, out _));
        OutlookSessionsHealthView? unknown = HealthReporting.DescribeOutlookSessions(null, Outlook, false, s => true, null, out _);
        Assert.Empty(unknown!.Attached);
        Assert.Null(unknown.StartedByOutlookAI);
        Assert.Null(unknown.PinWindow);
    }

    // ------------------------------------------------------------------ the store: a file under a named mutex

    [Fact]
    public void TheStore_KeepsEveryChangeInItsFile_AndReadsItBack()
    {
        using TempRegistry temp = new();
        OutlookSessionRegistry registry = temp.Registry(s => true);
        RegisteredSession a = Session(HostA, 1, Outlook);

        int answer = registry.Update(state =>
        {
            SessionRegistryRules.Join(state, a);
            SessionRegistryRules.PinMade(state, Outlook, 12, a.Key, 3);
            return 17;
        }, out RegistryAccess access);

        Assert.Equal(17, answer);
        Assert.Equal(RegistryAccess.Locked, access);
        Assert.True(File.Exists(temp.Path));
        Assert.StartsWith(SessionRegistryRules.Header, File.ReadAllText(temp.Path), StringComparison.Ordinal);
        SessionRegistryState? read = registry.TryRead();
        Assert.Equal(12, read!.Find(Outlook)!.PinWindow);
        Assert.Equal(a.Key, Assert.Single(read.Sessions).Key);
    }

    [Fact]
    public void TheStore_PrunesDeadSessionsOnEveryChange()
    {
        using TempRegistry temp = new();
        RegisteredSession dead = Session(HostA, 1, Outlook);
        temp.Registry(s => true).Update(state =>
        {
            SessionRegistryRules.Join(state, dead);
            return 0;
        }, out _);

        temp.Registry(s => s != HostA).Update(state => 0, out _);

        Assert.Empty(temp.Registry(s => true).TryRead()!.Sessions);
    }

    [Fact]
    public void TheStore_SerialisesChangesFromManyThreads()
    {
        using TempRegistry temp = new();
        OutlookSessionRegistry registry = temp.Registry(s => true);
        System.Collections.Concurrent.ConcurrentBag<RegistryAccess> accesses = new();
        Thread[] threads = Enumerable.Range(1, 8).Select(i => new Thread(() =>
        {
            for (int j = 0; j < 5; j++)
            {
                registry.Update(state =>
                {
                    SessionRegistryRules.Join(state, Session(HostA, (i * 100) + j, Outlook));
                    return 0;
                }, out RegistryAccess access);
                accesses.Add(access);
            }
        })).ToArray();

        foreach (Thread thread in threads)
        {
            thread.Start();
        }

        foreach (Thread thread in threads)
        {
            thread.Join();
        }

        Assert.Equal(40, accesses.Count(a => a == RegistryAccess.Locked));
        Assert.Equal(40, registry.TryRead()!.Sessions.Count);
    }

    [Fact]
    public void TheStore_ReadsAroundADamagedFile()
    {
        using TempRegistry temp = new();
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(temp.Path)!);
        File.WriteAllText(temp.Path, "\u0000\u0001 not a registry\nsession key=broken\n");
        OutlookSessionRegistry registry = temp.Registry(s => true);

        registry.Update(state =>
        {
            SessionRegistryRules.Join(state, Session(HostA, 1, Outlook));
            return 0;
        }, out RegistryAccess access);

        Assert.Equal(RegistryAccess.Locked, access);
        Assert.Single(registry.TryRead()!.Sessions);
    }

    [Fact]
    public void TheStore_WhenTheLockIsNotHad_ChangesThisProcesssCopy_AndSaysSo()
    {
        using TempRegistry temp = new();
        OutlookSessionRegistry registry = temp.Registry(s => true, lockWaitMilliseconds: 200);
        using ManualResetEventSlim held = new();
        using ManualResetEventSlim release = new();
        Thread holder = new(() =>
        {
            using Mutex mutex = new(false, temp.MutexName);
            mutex.WaitOne();
            held.Set();
            release.Wait();
            mutex.ReleaseMutex();
        });
        holder.Start();
        held.Wait();
        try
        {
            int ran = 0;
            registry.Update(state =>
            {
                ran++;
                SessionRegistryRules.EditorOutcome(state, Outlook, DraftEditorRoute.Promote, true, 5);
                return 0;
            }, out RegistryAccess access);

            Assert.Equal(1, ran);
            Assert.Equal(RegistryAccess.LockTimedOut, access);
            Assert.False(File.Exists(temp.Path), "nothing may be written without the lock");

            // The copy keeps the promotion once per process even so.
            DraftEditorRoute route = registry.Update(state => SessionRegistryRules.DecideEditorRoute(false, state.Find(Outlook)), out _);
            Assert.Equal(DraftEditorRoute.Direct, route);
        }
        finally
        {
            release.Set();
            holder.Join();
        }
    }

    [Fact]
    public void TheStore_TakesTheLockOverFromAHolderThatDied()
    {
        using TempRegistry temp = new();
        Mutex? abandoned = null;
        Thread dies = new(() =>
        {
            abandoned = new Mutex(false, temp.MutexName);
            abandoned.WaitOne();

            // Ends holding it, never released: a COM host killed in the middle of a change.
        });
        dies.Start();
        dies.Join();
        try
        {
            // The dead thread's handle stays open, so the mutex is the abandoned one, not a fresh one of the same name.
            temp.Registry(s => true).Update(state =>
            {
                SessionRegistryRules.Join(state, Session(HostA, 1, Outlook));
                return 0;
            }, out RegistryAccess access);

            Assert.Equal(RegistryAccess.Locked, access);
            Assert.Single(temp.Registry(s => true).TryRead()!.Sessions);
        }
        finally
        {
            abandoned?.Dispose();
        }
    }

    [Fact]
    public void SessionKeys_AreUniqueAndCarryTheirProcess()
    {
        ProcessStamp self = ProcessStamp.Current();
        string first = OutlookSessionRegistry.NewSessionKey(self);
        string second = OutlookSessionRegistry.NewSessionKey(self);
        Assert.NotEqual(first, second);
        Assert.StartsWith(self.Pid.ToString(CultureInfo.InvariantCulture) + "." + self.StartedUtcTicks.ToString(CultureInfo.InvariantCulture) + ".", first, StringComparison.Ordinal);
    }

    [Fact]
    public void TheDefaultRegistry_IsAFileUnderTheOutlookAIStateFolder()
    {
        Assert.Equal(
            System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", OutlookSessionRegistry.FileName),
            OutlookSessionRegistry.Default.FilePath);
        Assert.StartsWith(@"Local\", OutlookSessionRegistry.DefaultMutexName, StringComparison.Ordinal);
    }

    /// <summary>A registry under %TEMP% with a mutex name of its own, deleted afterwards - never the user's.</summary>
    private sealed class TempRegistry : IDisposable
    {
        private readonly string _directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OutlookAI-T1-Registry", Guid.NewGuid().ToString("N"));

        public TempRegistry()
        {
            MutexName = @"Local\OutlookAI.T1Registry." + Guid.NewGuid().ToString("N");
        }

        public string Path => System.IO.Path.Combine(_directory, OutlookSessionRegistry.FileName);

        public string MutexName { get; }

        public OutlookSessionRegistry Registry(Func<ProcessStamp, bool> isRunning, int lockWaitMilliseconds = 5_000)
        {
            return new OutlookSessionRegistry(Path, MutexName, isRunning, lockWaitMilliseconds);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_directory))
                {
                    Directory.Delete(_directory, recursive: true);
                }
            }
            catch (IOException)
            {
            }
        }
    }
}
