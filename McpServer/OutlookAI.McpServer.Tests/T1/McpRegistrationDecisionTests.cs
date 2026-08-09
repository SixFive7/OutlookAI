using OutlookAI.Services;
using Xunit;

using Action = OutlookAI.Services.McpRegistrationDecision.RegistrationAction;
using Entry = OutlookAI.Services.McpRegistrationDecision.EntryState;
using Prompt = OutlookAI.Services.McpRegistrationDecision.PromptKind;
using Window = OutlookAI.Services.McpRegistrationDecision.OutlookWindow;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// When the add-in may change Claude Code's user-scope MCP registration by itself, and when
/// it has to stop and ask (<c>Services/McpRegistrationDecision.cs</c>, LINKED into this test
/// project — see the tests csproj).
///
/// This is the rule "never act on inferred intent" made mechanical, and it is worth pinning
/// exhaustively because both ways of getting it wrong are destructive: adopting an entry that
/// is not ours overwrites a command someone chose on purpose, and treating an absent
/// preference as an opt-out deletes an entry nobody asked us to delete. Every combination of
/// (stored intent × what the entry is) appears below exactly once.
///
/// The headless case is pinned just as hard. Outlook is deliberately autostarted with no
/// window for agent sessions, and a modal dialog there would be invisible to everyone and
/// would wedge the reconcile that raised it — so "nobody is looking" must produce a deferral
/// that writes nothing at all, never a guess.
/// </summary>
public sealed class McpRegistrationDecisionTests
{
    private static McpRegistrationDecision.Decision Decide(
        int? stored, Entry entry, bool canAskUser = true, bool intentJustDeclared = false)
    {
        return McpRegistrationDecision.Decide(stored, entry, canAskUser, intentJustDeclared);
    }

    // ===== The whole matrix =====
    //
    // stored: null = never decided, 1 = registration on, 0 = registration off.

    [Theory]
    // Never decided. Only an entry that is ALREADY ours may be resolved without asking:
    // upgrading a working install must not interrogate the user about it.
    [InlineData(null, Entry.Absent, Prompt.FirstRun, Action.None)]
    [InlineData(null, Entry.Ours, Prompt.None, Action.AdoptAndRegister)]
    [InlineData(null, Entry.Foreign, Prompt.ForeignEntry, Action.None)]
    [InlineData(null, Entry.Unreadable, Prompt.None, Action.None)]
    // On. A missing entry contradicts that and only the user can say why; anything else is
    // kept pointing at the installed server, which is the drift-healing the toggle promises.
    [InlineData(1, Entry.Absent, Prompt.EntryMissing, Action.None)]
    [InlineData(1, Entry.Ours, Prompt.None, Action.Register)]
    [InlineData(1, Entry.Foreign, Prompt.None, Action.Register)]
    [InlineData(1, Entry.Unreadable, Prompt.None, Action.None)]
    // Off. An entry of ours that is nevertheless there was added outside Outlook, so ask
    // before deleting anything; an entry that is not ours is simply none of our business.
    [InlineData(0, Entry.Absent, Prompt.None, Action.None)]
    [InlineData(0, Entry.Ours, Prompt.EntryUnexpected, Action.None)]
    [InlineData(0, Entry.Foreign, Prompt.None, Action.None)]
    [InlineData(0, Entry.Unreadable, Prompt.None, Action.None)]
    public void EveryCombinationOfIntentAndEntry(int? stored, Entry entry, Prompt prompt, Action action)
    {
        McpRegistrationDecision.Decision decision = Decide(stored, entry);

        Assert.Equal(prompt, decision.Prompt);
        Assert.Equal(action, decision.Action);
        Assert.False(decision.Deferred);
    }

    [Theory]
    [InlineData(null, Entry.Absent, Prompt.FirstRun)]
    [InlineData(null, Entry.Foreign, Prompt.ForeignEntry)]
    [InlineData(1, Entry.Absent, Prompt.EntryMissing)]
    [InlineData(0, Entry.Ours, Prompt.EntryUnexpected)]
    public void AQuestionIsNeverAskedAndActedOnAtTheSameTime(int? stored, Entry entry, Prompt prompt)
    {
        // THE invariant behind every safety claim here: while a question is outstanding,
        // nothing is written. There is no branch that half-registers and then asks.
        McpRegistrationDecision.Decision decision = Decide(stored, entry);

        Assert.Equal(prompt, decision.Prompt);
        Assert.Equal(Action.None, decision.Action);
        Assert.True(decision.ChangesNothing);
    }

    // ===== Decision 1: only an entry that is already ours counts as an implicit opt-in =====

    [Fact]
    public void AnEntryThatIsAlreadyOursIsAdoptedWithoutAsking()
    {
        // The migration rule, and the one silent resolution of an undecided machine: it was
        // registered by a version that did it unconditionally, so nothing the user can observe
        // changes. The opt-in is persisted with it.
        McpRegistrationDecision.Decision decision = Decide(null, Entry.Ours);

        Assert.Equal(Prompt.None, decision.Prompt);
        Assert.Equal(Action.AdoptAndRegister, decision.Action);
    }

    [Fact]
    public void AnEntryPointingSomewhereElseIsNeverAdoptedAndNeverOverwritten()
    {
        // `claude mcp add --scope user outlookai -- C:\my\wrapper.cmd` is a deliberate act.
        // Reading it as "you already opted in" and healing it would destroy it, and would
        // switch a setting on that the user never touched. So: ask, and change nothing.
        McpRegistrationDecision.Decision decision = Decide(null, Entry.Foreign);

        Assert.Equal(Prompt.ForeignEntry, decision.Prompt);
        Assert.Equal(Action.None, decision.Action);
    }

    [Fact]
    public void AnEntryPointingSomewhereElseIsNeverDeletedEither()
    {
        // The other half, and the one that would otherwise undo the user's own answer: having
        // said "leave it alone" (which stores an explicit off), the very next Outlook start
        // must not delete the entry they just protected.
        McpRegistrationDecision.Decision decision = Decide(0, Entry.Foreign);

        Assert.Equal(Prompt.None, decision.Prompt);
        Assert.Equal(Action.None, decision.Action);
        Assert.False(decision.Deferred);
    }

    [Fact]
    public void RemovalOnlyEverTargetsAnEntryOfOurs()
    {
        // Across the entire matrix, in both intent modes: nothing but Ours is ever removed.
        foreach (int? stored in new int?[] { null, 0, 1 })
        {
            foreach (Entry entry in new[] { Entry.Absent, Entry.Ours, Entry.Foreign, Entry.Unreadable })
            {
                foreach (bool declared in new[] { false, true })
                {
                    foreach (bool canAsk in new[] { false, true })
                    {
                        McpRegistrationDecision.Decision decision = Decide(stored, entry, canAsk, declared);
                        if (decision.Action == Action.Remove)
                        {
                            Assert.Equal(Entry.Ours, entry);
                        }
                    }
                }
            }
        }
    }

    // ===== A configuration that could not be read is never touched, and never asked about =====

    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    [InlineData(1)]
    public void AnUnreadableConfigurationProducesNoQuestionAndNoWrite(int? stored)
    {
        foreach (bool declared in new[] { false, true })
        {
            McpRegistrationDecision.Decision decision = Decide(stored, Entry.Unreadable, true, declared);

            Assert.Equal(Prompt.None, decision.Prompt);
            Assert.Equal(Action.None, decision.Action);
            Assert.False(decision.Deferred);
        }
    }

    // ===== Decision 2: a headless Outlook defers, and writes nothing =====

    [Theory]
    [InlineData(null, Entry.Absent)]
    [InlineData(null, Entry.Foreign)]
    [InlineData(1, Entry.Absent)]
    [InlineData(0, Entry.Ours)]
    public void WithNobodyToAskEveryQuestionDefersAndNothingIsWritten(int? stored, Entry entry)
    {
        McpRegistrationDecision.Decision decision = Decide(stored, entry, canAskUser: false);

        Assert.Equal(Prompt.None, decision.Prompt);
        Assert.Equal(Action.None, decision.Action);
        Assert.True(decision.Deferred);
        Assert.True(decision.ChangesNothing);
    }

    [Fact]
    public void AHeadlessOutlookIsNobodyToAsk_SoTheStateIsLeftExactlyAsFound()
    {
        // End to end, through the very predicate the add-in feeds with real window handles:
        // the background Outlook the mail server autostarts owns no window a human can see
        // (D17/D33), so every question defers and not one of them turns into a write. The
        // next interactive session asks instead.
        Window[] headless =
        {
            new Window(visible: false, minimized: false, left: 0, top: 0, right: 1200, bottom: 800),
            new Window(visible: true, minimized: false, left: -32000, top: -32000, right: -31000, bottom: -31200),
            new Window(visible: true, minimized: false, left: 0, top: 0, right: 0, bottom: 0),
        };

        bool canAskUser = McpRegistrationDecision.AnyWindowAHumanCanSee(headless);
        Assert.False(canAskUser);

        (int? Stored, Entry Entry)[] everyQuestion =
        {
            (null, Entry.Absent),
            (null, Entry.Foreign),
            (1, Entry.Absent),
            (0, Entry.Ours),
        };

        foreach ((int? stored, Entry entry) in everyQuestion)
        {
            McpRegistrationDecision.Decision decision = Decide(stored, entry, canAskUser);

            Assert.Equal(Prompt.None, decision.Prompt);
            Assert.Equal(Action.None, decision.Action);
            Assert.True(decision.Deferred);
        }
    }

    [Theory]
    // The states that are NOT questions carry on working headless — that is the whole point
    // of the background reconcile: an install whose intent is already known keeps itself
    // correct without anyone being there.
    [InlineData(null, Entry.Ours, Action.AdoptAndRegister)]
    [InlineData(1, Entry.Ours, Action.Register)]
    [InlineData(1, Entry.Foreign, Action.Register)]
    [InlineData(0, Entry.Absent, Action.None)]
    [InlineData(0, Entry.Foreign, Action.None)]
    public void HavingNobodyToAskChangesNothingThatWasNeverAQuestion(int? stored, Entry entry, Action action)
    {
        McpRegistrationDecision.Decision decision = Decide(stored, entry, canAskUser: false);

        Assert.Equal(Prompt.None, decision.Prompt);
        Assert.Equal(action, decision.Action);
        Assert.False(decision.Deferred);
    }

    // ===== An answer is acted on, not questioned again =====

    [Theory]
    // Ticking the box in OutlookAI Settings, or answering one of the prompts, IS the intent.
    // Asking again about the state that prompted the question would be absurd.
    [InlineData(1, Entry.Absent, Action.Register)]
    [InlineData(1, Entry.Ours, Action.Register)]
    [InlineData(1, Entry.Foreign, Action.Register)]
    [InlineData(0, Entry.Ours, Action.Remove)]
    [InlineData(0, Entry.Absent, Action.None)]
    [InlineData(0, Entry.Foreign, Action.None)]
    public void AFreshlyDeclaredIntentIsActedOnImmediately(int? stored, Entry entry, Action action)
    {
        McpRegistrationDecision.Decision decision = Decide(stored, entry, true, intentJustDeclared: true);

        Assert.Equal(Prompt.None, decision.Prompt);
        Assert.Equal(action, decision.Action);
        Assert.False(decision.Deferred);
    }

    [Fact]
    public void UntickingTheBoxRemovesTheEntryThereAndThen()
    {
        // The deregistration path: without this, unticking would ask "we found our entry,
        // remove it?" a second after the user asked for exactly that.
        Assert.Equal(Action.Remove, Decide(0, Entry.Ours, true, intentJustDeclared: true).Action);
        Assert.Equal(Prompt.EntryUnexpected, Decide(0, Entry.Ours).Prompt);
    }

    [Fact]
    public void EvenAFreshlyDeclaredIntentCannotTouchAnUnreadableConfiguration()
    {
        Assert.Equal(Action.None, Decide(1, Entry.Unreadable, true, intentJustDeclared: true).Action);
        Assert.Equal(Action.None, Decide(0, Entry.Unreadable, true, intentJustDeclared: true).Action);
    }

    // ===== What the settings dialog shows =====

    [Fact]
    public void AnEntryOfOursReadsAsOptedInWhileNothingIsStored()
    {
        // So the tick box and the reconcile agree about a machine that is about to be
        // adopted silently.
        Assert.True(McpRegistrationDecision.ResolveOptIn(null, ourEntryAlreadyPresent: true));
    }

    [Fact]
    public void AFreshInstallReadsAsNotOptedIn()
    {
        Assert.False(McpRegistrationDecision.ResolveOptIn(null, ourEntryAlreadyPresent: false));
    }

    [Fact]
    public void AnExplicitChoiceBeatsWhatIsOnDisk()
    {
        // Both directions: turning it off must not be undone by the entry still being there
        // on the very next start, and turning it on must survive the entry being deleted.
        Assert.False(McpRegistrationDecision.ResolveOptIn(0, ourEntryAlreadyPresent: true));
        Assert.True(McpRegistrationDecision.ResolveOptIn(1, ourEntryAlreadyPresent: false));
    }

    // ===== Is there a human here at all? =====

    [Fact]
    public void AnOrdinaryOutlookWindowIsSomebodyToAsk()
    {
        Assert.True(McpRegistrationDecision.AnyWindowAHumanCanSee(new[]
        {
            new Window(visible: true, minimized: false, left: 100, top: 100, right: 1300, bottom: 900),
        }));
    }

    [Fact]
    public void AMinimizedOutlookIsSomebodyToAsk()
    {
        // The case the rectangle test alone gets wrong: Windows parks a minimized window in
        // the same far-off-screen corner the invisible compose surface uses, so without the
        // minimized flag the user's own Outlook would be mistaken for a hidden one and they
        // would never be asked anything.
        Assert.True(McpRegistrationDecision.AnyWindowAHumanCanSee(new[]
        {
            new Window(visible: true, minimized: true, left: -32000, top: -32000, right: -31840, bottom: -31976),
        }));
    }

    [Fact]
    public void AParkedComposeWindowIsNobody()
    {
        // What the mail server's invisible compose surface leaves lying around while it drives
        // Word. Visible to Win32 for an instant, never to a human.
        Assert.False(McpRegistrationDecision.AnyWindowAHumanCanSee(new[]
        {
            new Window(
                visible: true,
                minimized: false,
                left: McpRegistrationDecision.ParkX,
                top: McpRegistrationDecision.ParkY,
                right: McpRegistrationDecision.ParkX + 900,
                bottom: McpRegistrationDecision.ParkY + 600),
        }));
    }

    [Theory]
    // A hidden window, a zero-size helper window, and nothing at all: all "nobody".
    [InlineData(false, false, 0, 0, 1200, 800)]
    [InlineData(true, false, 0, 0, 0, 0)]
    [InlineData(true, false, 50, 50, 50, 50)]
    public void NothingAHumanCouldSee(bool visible, bool minimized, int left, int top, int right, int bottom)
    {
        Assert.False(McpRegistrationDecision.AnyWindowAHumanCanSee(new[]
        {
            new Window(visible, minimized, left, top, right, bottom),
        }));
    }

    [Fact]
    public void OneVisibleWindowAmongHiddenOnesIsEnough()
    {
        Assert.True(McpRegistrationDecision.AnyWindowAHumanCanSee(new[]
        {
            new Window(visible: false, minimized: false, left: 0, top: 0, right: 800, bottom: 600),
            new Window(visible: true, minimized: false, left: -32000, top: -32000, right: -31000, bottom: -31400),
            new Window(visible: true, minimized: false, left: 10, top: 10, right: 900, bottom: 700),
        }));
    }

    [Fact]
    public void NoWindowsAtAllIsNobody()
    {
        Assert.False(McpRegistrationDecision.AnyWindowAHumanCanSee(new Window[0]));
        Assert.False(McpRegistrationDecision.AnyWindowAHumanCanSee(null!));
    }
}
