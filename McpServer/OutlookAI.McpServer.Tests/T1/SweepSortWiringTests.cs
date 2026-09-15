using System.Reflection;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the two lines that join the sweep's sort-refusal DECISION to the counter the field reads,
/// and the scan cursor's date fallback - the two mutants that survived the 2026-08-24 pass.
/// <para>
/// <b>What was already pinned, and why it was not enough.</b>
/// <c>OutlookComSession.SweepSortWasRefused</c> (the refusal test) and
/// <c>OutlookComSession.AddSortRefusal</c> (the counter) are pure, public and driven directly by
/// <c>SweepRefusalTelemetryTests</c>. Re-running the mutation pass one line at a time on
/// 2026-08-24 nevertheless left two mutants alive, and both live at a JOIN rather than in a
/// decision:
/// </para>
/// <list type="number">
/// <item>Inverting the argument at BOTH call sites -
/// <c>AddSortRefusal(tally.SortRefused, !sortRefused)</c> - builds clean and leaves the whole
/// suite green. The call sites are inside <c>SweepFoldersNewerThan</c> and
/// <c>SweepFolderTree</c>, whose walks are <c>dynamic</c> COM against live Outlook folders, so no
/// mailbox-free test can enter them.</item>
/// <item><c>ScanSingleFolder</c>'s cursor fallback dropped its conversion - exactly the defect the
/// line was added to fix, a local instant handed to a cursor named Utc - and the suite stayed
/// green. That one is fixed by MOVING the decision (<c>OutlookComSession.ScanCursorDate</c>), which
/// this file drives directly; what remains unreachable is the call to it.</item>
/// </list>
/// <para>
/// <b>Why this reads the compiled IL, and what that is worth.</b> A stand-in harness for the whole
/// of <c>SweepFolder</c> was considered and REFUSED on 2026-08-24, on the maintainer's reasoning
/// that the more faithful the fake, the more you end up testing your model of Outlook rather than
/// Outlook. The substitute this repository already uses for a join it cannot execute is to read
/// the join out of the compiled method - the same instrument
/// <c>TripwireReRunDriverTests</c> uses for the child-process launch. It proves ONE thing, and
/// narrowly: the flag reaches the counter without being negated on the way. It does not prove the
/// walk computes the flag correctly; <c>SweepRefusalTelemetryTests</c> owns that.
/// </para>
/// <para>
/// <b>The detector self-tests, because a detector that silently stops detecting is worse than no
/// detector.</b> <see cref="TheDetectorStillRecognisesTheExactMutationItExistsFor"/> compiles both
/// shapes into this test class and asserts the reader fires on the negated one and not on the
/// plain one. Same discipline as <c>LiveTierClockDriftTests</c>, and for the same reason.
/// </para>
/// <para>
/// No COM, no Outlook, no mailbox: reflection over the already-built assembly only.
/// </para>
/// </summary>
public sealed class SweepSortWiringTests
{
    /// <summary>
    /// The two COM walks that feed the counter. Named rather than discovered: a walk that stopped
    /// calling <c>AddSortRefusal</c> at all is the same defect as one that negates it, and a
    /// discovery-based list would simply stop looking at it.
    /// </summary>
    private static readonly string[] SweepWalks = { "SweepFoldersNewerThan", "SweepFolderTree" };

    // ------------------------------------------------------------- the sort-refusal join

    [Fact]
    public void BothSweepWalksStillFeedTheRefusalCounter()
    {
        foreach (string walk in SweepWalks)
        {
            Assert.Single(CallsTo(Walk(walk), nameof(OutlookComSession.AddSortRefusal)));
        }
    }

    [Fact]
    public void NeitherSweepWalkNegatesTheFlagOnItsWayToTheCounter()
    {
        // The surviving mutant, stated as an assertion. Inverted, a healthy profile's zero becomes
        // "every folder refused the sort" and a genuinely refusing profile's evidence becomes a
        // zero - and both readings would be believed, because sweep.sortRefusedFolders is the only
        // witness there is. This is the number that settled, in 2026-08-23, that Table.Sort had
        // never applied on any store for any user.
        foreach (string walk in SweepWalks)
        {
            Assert.All(
                CallsTo(Walk(walk), nameof(OutlookComSession.AddSortRefusal)),
                call => Assert.False(
                    NegatedImmediatelyBefore(Il(call.Method), call.Offset),
                    walk + " negates the sort-refusal flag before counting it."));
        }
    }

    [Fact]
    public void TheDetectorStillRecognisesTheExactMutationItExistsFor()
    {
        // The control run. A pin that reads bytes has to prove it can still see the shape it was
        // written for, or a compiler change turns it into a test that passes by looking at nothing.
        MethodInfo negated = typeof(SweepSortWiringTests).GetMethod(
            nameof(MutatedShape_CountsTheFlagNEGATED), BindingFlags.NonPublic | BindingFlags.Static)!;
        MethodInfo plain = typeof(SweepSortWiringTests).GetMethod(
            nameof(ShippedShape_CountsTheFlagAsItIs), BindingFlags.NonPublic | BindingFlags.Static)!;

        (MethodInfo _, int negatedOffset) = Assert.Single(
            CallsTo(new[] { negated }, nameof(OutlookComSession.AddSortRefusal)));
        Assert.True(
            NegatedImmediatelyBefore(Il(negated), negatedOffset),
            "the detector no longer recognises a negated flag, so the two assertions above prove nothing.");

        (MethodInfo _, int plainOffset) = Assert.Single(
            CallsTo(new[] { plain }, nameof(OutlookComSession.AddSortRefusal)));
        Assert.False(
            NegatedImmediatelyBefore(Il(plain), plainOffset),
            "the detector fires on the SHIPPED shape, so it would fail whatever the walks did.");
    }

    /// <summary>The shape both sweep walks carry. Compiled, never called.</summary>
    private static int ShippedShape_CountsTheFlagAsItIs(int refusedSoFar, bool sortRefused)
    {
        return OutlookComSession.AddSortRefusal(refusedSoFar, sortRefused);
    }

    /// <summary>The mutation that survived the whole suite on 2026-08-24. Compiled, never called.</summary>
    private static int MutatedShape_CountsTheFlagNEGATED(int refusedSoFar, bool sortRefused)
    {
        return OutlookComSession.AddSortRefusal(refusedSoFar, !sortRefused);
    }

    // ------------------------------------------------------------- the scan cursor's date

    [Fact]
    public void TheTableDateWinsWhenTheRowCarriedOne()
    {
        DateTime fromTable = new(2026, 8, 24, 9, 0, 0, DateTimeKind.Utc);
        DateTime fromItem = new(2001, 1, 1, 0, 0, 0, DateTimeKind.Local);

        Assert.Equal(fromTable, OutlookComSession.ScanCursorDate(fromTable, fromItem));
    }

    [Fact]
    public void AnAbsentTableDateFallsBackToTheItemValueCONVERTED()
    {
        // The defect, as a test. brief.ReceivedTime is local wall time carrying
        // DateTimeKind.Unspecified, and the cursor it feeds is named Utc. Using it raw is what put
        // a local instant into the next page's date bound, which on this machine is an hour or two
        // of a resumable scan either skipped or re-read.
        DateTime unspecifiedLocal = new(2026, 8, 24, 9, 0, 0, DateTimeKind.Unspecified);

        DateTime? cursor = OutlookComSession.ScanCursorDate(null, unspecifiedLocal);

        Assert.Equal(ComDateValue.FromItemValue(unspecifiedLocal), cursor);
        Assert.NotEqual(unspecifiedLocal, cursor);
        Assert.Equal(DateTimeKind.Utc, cursor!.Value.Kind);

        // And NOT the table conversion, which would leave an unspecified kind alone and call it
        // UTC - the two differ by exactly the machine's offset, which is why the kind cannot be
        // used to tell which conversion was meant.
        Assert.NotEqual(ComDateValue.FromTableValue(unspecifiedLocal), cursor);
    }

    [Fact]
    public void NeitherHalfPresentIsStillNull()
    {
        Assert.Null(OutlookComSession.ScanCursorDate(null, null));
    }

    [Fact]
    public void TheScanStillAsksForTheCursorDateRatherThanBuildingItInline()
    {
        // The one line CI cannot execute: the call inside ScanSingleFolder's dynamic row loop.
        // Inlining the fallback again is what the move was for, so the pin is that the call is
        // there at all.
        Assert.Single(CallsTo(Walk("ScanSingleFolder"), nameof(OutlookComSession.ScanCursorDate)));
    }

    // ------------------------------------------------------------- the IL reader

    /// <summary>
    /// Every compiled body that a source method's statements ended up in: the method itself, plus
    /// the compiler-generated ones named after it.
    /// <para>
    /// Load-bearing, and it was found by the pin failing rather than by reading. The sweep walks
    /// run inside <c>_runner.Run(() =&gt; { ... })</c> - a lambda on the STA pump - so the
    /// statements that feed the counter are compiled into a display class named
    /// <c>&lt;SweepFoldersNewerThan&gt;b__N</c> and the declared method contains no such call at
    /// all. A pin that looked only at the declared method would have found nothing and said so
    /// (<see cref="BothSweepWalksStillFeedTheRefusalCounter"/> is what catches that), but a pin
    /// that only checked for the ABSENCE of a negation would have passed on an empty set - which
    /// is the vacuous-guard shape this repository refuses elsewhere.
    /// </para>
    /// </summary>
    private static IReadOnlyList<MethodInfo> Walk(string name)
    {
        const BindingFlags All = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
        string lambdaMarker = "<" + name + ">";

        List<MethodInfo> bodies = new();
        foreach (Type type in new[] { typeof(OutlookComSession) }
                     .Concat(typeof(OutlookComSession).GetNestedTypes(All)))
        {
            bodies.AddRange(type.GetMethods(All).Where(
                m => string.Equals(m.Name, name, StringComparison.Ordinal)
                    || m.Name.Contains(lambdaMarker, StringComparison.Ordinal)));
        }

        if (bodies.Count == 0)
        {
            throw new InvalidOperationException(
                "OutlookComSession." + name + " is gone or renamed, and this pin no longer reaches it.");
        }

        return bodies;
    }

    private static byte[] Il(MethodInfo method)
    {
        return method.GetMethodBody()?.GetILAsByteArray()
            ?? throw new InvalidOperationException(method.Name + " has no IL body to read.");
    }

    /// <summary>
    /// Every <c>call</c> (0x28) across <paramref name="bodies"/> whose metadata token RESOLVES to
    /// a method named <paramref name="callee"/> on <see cref="OutlookComSession"/>, as the body it
    /// sits in and its byte offset. The token is resolved rather than pattern-matched, so a stray
    /// operand byte that happens to read as 0x28 cannot pass for an instruction - the same reader
    /// <c>TripwireReRunDriverTests.CalleesOf</c> uses.
    /// </summary>
    private static List<(MethodInfo Method, int Offset)> CallsTo(
        IReadOnlyList<MethodInfo> bodies, string callee)
    {
        List<(MethodInfo, int)> found = new();
        foreach (MethodInfo method in bodies)
        {
            byte[] il = Il(method);
            for (int i = 0; i + 4 < il.Length; i++)
            {
                if (il[i] != 0x28)
                {
                    continue;
                }

                int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
                try
                {
                    MethodBase? resolved = method.Module.ResolveMethod(token);
                    if (resolved != null
                        && resolved.DeclaringType == typeof(OutlookComSession)
                        && string.Equals(resolved.Name, callee, StringComparison.Ordinal))
                    {
                        found.Add((method, i));
                    }
                }
                catch (ArgumentException)
                {
                    // Not a real call - the bytes happened to look like one.
                }
            }
        }

        return found;
    }

    /// <summary>
    /// True when the three bytes before <paramref name="callOffset"/> are
    /// <c>ldc.i4.0; ceq</c> - which is how C# compiles <c>!flag</c> (and <c>flag == false</c>)
    /// on a <c>bool</c> local immediately before a call.
    /// <para>
    /// Deliberately narrow. It catches the mutation that was actually measured to survive and
    /// nothing else: a caller that computed the inversion further away, or through a helper, would
    /// pass this and is a different, far more visible edit.
    /// <see cref="TheDetectorStillRecognisesTheExactMutationItExistsFor"/> is what keeps the claim
    /// honest.
    /// </para>
    /// </summary>
    private static bool NegatedImmediatelyBefore(byte[] il, int callOffset)
    {
        return callOffset >= 3
            && il[callOffset - 3] == 0x16
            && il[callOffset - 2] == 0xFE
            && il[callOffset - 1] == 0x01;
    }
}
