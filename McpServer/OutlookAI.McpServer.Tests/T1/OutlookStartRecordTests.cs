using System.Globalization;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what OutlookAI keeps about the Outlook it starts (Q176 (c)), the add-in's start-up stamp it waits
/// for (Q175 (a)), the start log (Q178 (e)) and the start lock (Q175 (e)) - each decided by the maintainer
/// on 2026-11-18 after the hang of 2026-11-15, where nothing OutlookAI wrote could say which call was in
/// progress inside the Outlook it had started, nor what started Outlook again after it was ended.
/// <para>
/// Formats, parsing and the stamp rule are pure. The log is written to this test process's throwaway
/// directory (<c>AuditIsolation</c>), and the lock uses mutex names of its own, so nothing here touches the
/// real record, log or lock.
/// </para>
/// </summary>
public sealed class OutlookStartRecordTests
{
    private static readonly DateTime Started = new(2026, 11, 15, 15, 27, 39, 467, DateTimeKind.Utc);

    // ------------------------------------------------------------------ the connection record

    [Fact]
    public void AConnectionRecord_ReadsBackAsWritten()
    {
        OutlookConnectRecord written = new(76888, Started, startedOutlook: true, serverProcessId: 68964, hostProcessId: 1234,
            OutlookStartSteps.Logon, Started.AddSeconds(21), Started.AddSeconds(-1));

        OutlookConnectRecord? read = OutlookConnectRecord.Parse(written.Format());

        Assert.NotNull(read);
        Assert.Equal(76888, read!.OutlookProcessId);
        Assert.Equal(Started, read.OutlookStartUtc);
        Assert.Equal(DateTimeKind.Utc, read.OutlookStartUtc!.Value.Kind);
        Assert.True(read.StartedOutlook);
        Assert.Equal(68964, read.ServerProcessId);
        Assert.Equal(1234, read.HostProcessId);
        Assert.Equal(OutlookStartSteps.Logon, read.Step);
        Assert.Equal(Started.AddSeconds(21), read.StepSinceUtc);
        Assert.Equal(Started.AddSeconds(-1), read.BeganUtc);
        Assert.True(read.InProgress);
    }

    [Fact]
    public void AConnectionRecord_BeforeItKnowsItsOutlook_ReadsBackWithout()
    {
        OutlookConnectRecord written = new(null, null, false, 10, 11, OutlookStartSteps.CreateApplication, Started, Started);
        OutlookConnectRecord? read = OutlookConnectRecord.Parse(written.Format());

        Assert.NotNull(read);
        Assert.Null(read!.OutlookProcessId);
        Assert.Null(read.OutlookStartUtc);
        Assert.False(read.StartedOutlook);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not a record")]
    [InlineData("server=1;host=2;stepSinceUtc=2026-11-15T15:27:39.0000000Z;beganUtc=2026-11-15T15:27:39.0000000Z")]
    [InlineData("server=x;host=2;step=logon;stepSinceUtc=2026-11-15T15:27:39.0000000Z;beganUtc=2026-11-15T15:27:39.0000000Z")]
    public void AnythingThatIsNotARecord_ReadsAsNone(string? text)
    {
        Assert.Null(OutlookConnectRecord.Parse(text));
    }

    [Theory]
    [InlineData(OutlookStartSteps.Connected, false)]
    [InlineData(OutlookStartSteps.Failed, false)]
    [InlineData(OutlookStartSteps.CreateApplication, true)]
    [InlineData(OutlookStartSteps.Settle, true)]
    [InlineData(OutlookStartSteps.Logon, true)]
    [InlineData(OutlookStartSteps.Pin, true)]
    public void OnlyConnectedAndFailed_EndAConnection(string step, bool inProgress)
    {
        OutlookConnectRecord record = new(1, Started, false, 2, 3, step, Started, Started);
        Assert.Equal(inProgress, record.InProgress);
        Assert.Equal(!inProgress, OutlookStartSteps.IsFinal(step));
    }

    [Fact]
    public void AValueThatCouldCarryAnotherField_IsRefused_NotEscaped()
    {
        // Every value is a number, a time or a step token; one with a separator in it is a defect.
        OutlookConnectRecord record = new(1, Started, false, 2, 3, "logon;server=9", Started, Started);
        Assert.Throws<ArgumentException>(() => record.Format());
    }

    // ------------------------------------------------------------------ notes about one Outlook

    [Fact]
    public void ANote_ReadsBackAsWritten()
    {
        OutlookProcessNote written = new(76888, Started, Started.AddSeconds(5), 68964);
        OutlookProcessNote? read = OutlookProcessNote.Parse(written.Format());

        Assert.NotNull(read);
        Assert.Equal(76888, read!.ProcessId);
        Assert.Equal(Started, read.ProcessStartUtc);
        Assert.Equal(Started.AddSeconds(5), read.AtUtc);
        Assert.Equal(68964, read.ServerProcessId);
    }

    [Fact]
    public void ANote_IsAboutOneOutlook_ByItsIdAndItsStart()
    {
        OutlookProcessNote note = new(76888, Started, Started.AddSeconds(5), null);

        Assert.True(note.IsAbout(76888, Started));
        Assert.True(note.IsAbout(76888, Started.AddMilliseconds(400)));

        // Another Outlook - and the same id reused by a later process, which is the case the start time is
        // there for: a note about the Outlook that hung must never describe the next one.
        Assert.False(note.IsAbout(60368, Started));
        Assert.False(note.IsAbout(76888, Started.AddMinutes(35)));
        Assert.False(note.IsAbout(null, null));
    }

    // ------------------------------------------------------------------ the add-in's start-up stamp

    [Fact]
    public void TheAddInsOwnStamp_CountsForTheOutlookItNames_AndNoOther()
    {
        string after = Started.AddSeconds(0.7).ToString("o", CultureInfo.InvariantCulture);

        AddInStartupStamp own = AddInStartupStamp.Evaluate(76888, Started, 3, 76888, after, null);
        Assert.True(own.Expected);
        Assert.True(own.Present);
        Assert.Equal(AddInStartupStamp.StartupStampSource, own.Source);
        Assert.True(own.AsFact);

        Assert.False(AddInStartupStamp.Evaluate(76888, Started, 3, 60368, after, null).Present);
    }

    [Fact]
    public void AStampOlderThanTheOutlook_DoesNotCount_SoAReusedIdOrAHalfWrittenPairSeesNoStamp()
    {
        // The add-in writes the id, then the time. A reader between the two sees the new id with the old
        // time, and a reused id carries an old time too; both must read as no stamp.
        string before = Started.AddSeconds(-30).ToString("o", CultureInfo.InvariantCulture);
        AddInStartupStamp torn = AddInStartupStamp.Evaluate(76888, Started, 3, 76888, before, before);

        Assert.False(torn.Present);
        Assert.False(torn.AsFact);
    }

    [Fact]
    public void AStampWithinASecondBeforeTheStart_Counts_ForTheClocksDifferentResolutions()
    {
        string justBefore = Started.AddMilliseconds(-500).ToString("o", CultureInfo.InvariantCulture);
        Assert.True(AddInStartupStamp.Evaluate(76888, Started, 3, 76888, justBefore, null).Present);
    }

    [Fact]
    public void AnOlderAddIn_IsRecognisedByItsTuningReconcile()
    {
        // The live-tier guests run the add-in their frozen checkpoint holds, which writes no start-up stamp
        // of its own - but every start of it writes its tuning reconcile's time.
        string reconcile = Started.AddSeconds(3.5).ToString("o", CultureInfo.InvariantCulture);
        AddInStartupStamp stamp = AddInStartupStamp.Evaluate(76888, Started, 3, null, null, reconcile);

        Assert.True(stamp.Present);
        Assert.Equal(AddInStartupStamp.TuningReconcileSource, stamp.Source);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    [InlineData(2)]
    public void AnAddInNotSetToLoad_IsNotWaitedFor(int? loadBehavior)
    {
        AddInStartupStamp stamp = AddInStartupStamp.Evaluate(76888, Started, loadBehavior, null, null, null);
        Assert.False(stamp.Expected);
        Assert.Null(stamp.AsFact);
    }

    [Fact]
    public void AnOutlookWhoseStartCannotBeRead_HasNoCheckableStamp_AndNoneIsWaitedFor()
    {
        AddInStartupStamp stamp = AddInStartupStamp.Evaluate(76888, null, 3, 76888, Started.ToString("o", CultureInfo.InvariantCulture), null);
        Assert.Null(stamp.AsFact);
    }

    // ------------------------------------------------------------------ the start log

    [Fact]
    public void ALogLine_IsATime_AnEvent_TheProcesses_AndTokens()
    {
        string line = OutlookStartLog.FormatLine(
            Started,
            68964,
            1234,
            "settle.end",
            new[]
            {
                OutlookStartLog.Field("outcome", "settled"),
                OutlookStartLog.Field("waitedMs", 17_250L),
                OutlookStartLog.Field("started", true),
                OutlookStartLog.Field("outlook", 76888),
                OutlookStartLog.Field("missing", null),
            });

        Assert.Equal(
            "2026-11-15T15:27:39.4670000Z settle.end server=68964 host=1234 outcome=settled waitedMs=17250 started=true outlook=76888",
            line);
    }

    [Theory]
    [InlineData("Inbox - someone@example.com - Outlook")]
    [InlineData("Folder 'Nope/Missing' was not found in store \"InjectedStore\"")]
    [InlineData("line one\r\nline two")]
    public void NoFreeTextReachesTheLog_EvenByMistake(string text)
    {
        // Q178 (e): "with no mail content". A store name, a subject or an exception message handed to the
        // log comes out as a token: no spaces, quotes, @ or line breaks survive, and it is cut short.
        string token = OutlookStartLog.Token(text);

        Assert.DoesNotContain(" ", token, StringComparison.Ordinal);
        Assert.DoesNotContain("@", token, StringComparison.Ordinal);
        Assert.DoesNotContain("'", token, StringComparison.Ordinal);
        Assert.DoesNotContain("\"", token, StringComparison.Ordinal);
        Assert.DoesNotContain("\n", token, StringComparison.Ordinal);
        Assert.Equal(token.Length, Math.Min(text.Length, OutlookStartLog.MaxValueLength));
    }

    [Fact]
    public void ALongValue_IsCut()
    {
        Assert.Equal(OutlookStartLog.MaxValueLength, OutlookStartLog.Token(new string('a', 500)).Length);
    }

    [Fact]
    public void TheLog_IsWrittenBesideThisProcesssAuditLog_AndKeptShort()
    {
        // Into the test run's throwaway directory, never the real one (AuditIsolation).
        Assert.True(AuditLog.IsRedirected, "the audit log is not redirected, so this would write the real start log");
        string directory = AuditLog.EffectiveDirectory;
        string path = OutlookStartLog.EffectivePath;
        string previous = Path.Combine(directory, OutlookStartLog.PreviousFileName);
        Assert.Equal(Path.Combine(directory, OutlookStartLog.FileName), path);

        Directory.CreateDirectory(directory);
        File.WriteAllText(path, new string('x', (int)OutlookStartLog.RotateAtBytes) + "\r\n");
        if (File.Exists(previous))
        {
            File.Delete(previous);
        }

        OutlookStartLog.Write("test.line", OutlookStartLog.Field("n", 1));

        Assert.True(File.Exists(previous), "the full log was not moved aside");
        Assert.True(new FileInfo(previous).Length > OutlookStartLog.RotateAtBytes);
        string[] lines = File.ReadAllLines(path);
        string written = Assert.Single(lines);
        Assert.Contains(" test.line server=", written, StringComparison.Ordinal);
        Assert.EndsWith(" n=1", written, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the start lock

    private static string TestLockName() => @"Local\OutlookAI.Test.StartLock." + Guid.NewGuid().ToString("N");

    [Fact]
    public void TheLock_IsSessionWide_LikeOutlook()
    {
        Assert.StartsWith(@"Local\", OutlookStartLock.DefaultName, StringComparison.Ordinal);
    }

    [Fact]
    public void OneConnectionAtATime_TheSecondIsTurnedAway_UntilTheFirstLetsGo()
    {
        string name = TestLockName();
        Assert.False(OutlookStartLock.IsHeldElsewhere(out _, name), "a lock nobody ever made is not held");

        using (OutlookStartLock? first = OutlookStartLock.TryAcquire(name))
        {
            Assert.NotNull(first);
            Assert.True(first!.Held);

            // Another thread - another connection, as far as a mutex can tell.
            Assert.True(OnAnotherThread(() => OutlookStartLock.TryAcquire(name) == null));
            Assert.True(OnAnotherThread(() => OutlookStartLock.IsHeldElsewhere(out _, name)));

            // The holder's own look does not count itself.
            Assert.False(OutlookStartLock.IsHeldElsewhere(out _, name));
        }

        Assert.False(OnAnotherThread(() => OutlookStartLock.IsHeldElsewhere(out _, name)));
        Assert.True(OnAnotherThread(() =>
        {
            using OutlookStartLock? again = OutlookStartLock.TryAcquire(name);
            return again != null && !again.PreviousHolderDied;
        }));
    }

    [Fact]
    public void AHolderThatDies_HandsTheLockOn_AndTheNextConnectionIsTold()
    {
        string name = TestLockName();

        // A thread that takes the lock and ends without releasing it: what a COM host killed at its
        // deadline leaves behind.
        Thread dying = new(() => _ = OutlookStartLock.TryAcquire(name));
        dying.Start();
        dying.Join();

        using OutlookStartLock? next = OutlookStartLock.TryAcquire(name);
        Assert.NotNull(next);
        Assert.True(next!.PreviousHolderDied);
    }

    private static bool OnAnotherThread(Func<bool> body)
    {
        bool result = false;
        Thread thread = new(() => result = body());
        thread.Start();
        thread.Join();
        return result;
    }
}
