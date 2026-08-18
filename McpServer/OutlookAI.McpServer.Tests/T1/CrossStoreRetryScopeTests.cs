using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// When a bare EntryID does not open, the service tries the OTHER stores. This pins WHEN it
/// is allowed to, for the five loops that were still fanning out on any failure at all.
/// <para>
/// The rule is one sentence: retry only when the item could not be OPENED. Past the open,
/// a failure is an ANSWER about an item that was found - "the attachment index is out of
/// range", "this is not a mail item", "Outlook would not render the body" - and asking the
/// same question of every other store cannot change it. What it CAN do is repeat the work:
/// two of these five are classified MUTATING in <see cref="ComSessionOperations"/> because
/// <c>open_in_outlook</c> puts a window on the user's screen (and can mark the mail read)
/// and <c>save_attachment</c> writes a file. A blind fan-out on those is one side effect per
/// store, over an item that had already been found, ending in a failure anyway.
/// </para>
/// <para>
/// The three draft loops were tightened first (eee02f2). These five were deliberately left
/// alone at the time, because their COM layer did not set the "ItemNotFound" token either -
/// tightening them without that groundwork would have made every retry dead code, which is
/// exactly what had already happened to <c>update_draft</c> and <c>discard_draft</c>. The
/// token is now set at every <c>GetItemFromID</c> and nowhere else, so both halves are
/// pinned here: a non-open failure must stop at the first store, and a genuine not-found
/// must still reach every one of them.
/// </para>
/// <para>
/// No Outlook and no mailbox: the session is a stand-in that counts calls and answers with
/// the token it is told to.
/// </para>
/// </summary>
public sealed class CrossStoreRetryScopeTests
{
    /// <summary>Three stores, so "stopped at the first" and "tried them all" cannot be confused.</summary>
    private static readonly ComStoreDetail[] Stores =
    {
        new ComStoreDetail("Work", "store-work", 0, true),
        new ComStoreDetail("Archive.pst", "store-archive", 3, null),
        new ComStoreDetail("Delegate", "store-delegate", 1, true),
    };

    /// <summary>A plausible bare EntryID: hex, even length, and long enough to be accepted as one.</summary>
    private static readonly string BareEntryId = new string('A', 140);

    public static TheoryData<string> Loops()
    {
        return new TheoryData<string>
        {
            nameof(IOutlookSession.TryReadItem),
            nameof(IOutlookSession.TrySaveAttachment),
            nameof(IOutlookSession.TryDisplayItem),
            nameof(IOutlookSession.TryGetSendableDraftState),
            nameof(IOutlookSession.TryGetMailInfo),
        };
    }

    [Theory]
    [MemberData(nameof(Loops))]
    public void AFailureThatIsNotAMissingItem_StopsAtTheFirstStore(string operation)
    {
        // Anything other than the not-found token means a store opened the item and then
        // refused. That is the answer, and the other stores cannot improve on it.
        CountingSession session = new CountingSession("COMException 0x80040111");
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        Drive(service, operation);

        Assert.Equal(1, session.CallsTo(operation));
    }

    [Theory]
    [MemberData(nameof(Loops))]
    public void AnItemThatIsNowhere_IsLookedForInEveryStore(string operation)
    {
        // The other half. Tightening a loop is only safe if it still does the thing it was
        // written for, and the way that breaks is silent: the loop stays, the token never
        // arrives, and a draft in a non-default store answers with an opaque COM code.
        CountingSession session = new CountingSession(ComErrorTokens.ItemNotFound);
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        Drive(service, operation);

        // One attempt with no store, then one per store, all of them failing to find it.
        Assert.Equal(1 + Stores.Length, session.CallsTo(operation));
    }

    [Theory]
    [MemberData(nameof(Loops))]
    public void AnEmptyFailureReason_StopsAtTheFirstStore(string operation)
    {
        // Fail-closed. A COM path that reports failure without saying why is not evidence
        // that the item is elsewhere, and the old `result == null` guard treated it as if it
        // were.
        CountingSession session = new CountingSession(null);
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        Drive(service, operation);

        Assert.Equal(1, session.CallsTo(operation));
    }

    /// <summary>
    /// Calls the tool entry point that owns each loop. All five end in failure here, which
    /// is the point - the loop is what is being observed, not the result.
    /// </summary>
    private static void Drive(MailService service, string operation)
    {
        try
        {
            switch (operation)
            {
                case nameof(IOutlookSession.TryReadItem):
                    _ = service.Read(BareEntryId);
                    break;
                case nameof(IOutlookSession.TrySaveAttachment):
                    // Rooted, because that is all the tool validates, and never touched: the
                    // directory is created inside the COM session, which here is a stand-in.
                    _ = service.SaveAttachment(BareEntryId, 1, @"C:\OutlookAI-test-never-written");
                    break;
                case nameof(IOutlookSession.TryDisplayItem):
                    _ = service.OpenInOutlook(BareEntryId);
                    break;
                case nameof(IOutlookSession.TryGetSendableDraftState):
                    _ = service.Send(BareEntryId);
                    break;
                case nameof(IOutlookSession.TryGetMailInfo):
                    // archive_mail reports per-item outcomes instead of throwing, so this one
                    // returns rather than raising.
                    _ = service.ArchiveMail(new[] { BareEntryId });
                    break;
                default:
                    throw new ArgumentException("No entry point wired for '" + operation + "'.", nameof(operation));
            }
        }
        catch (InvalidOperationException)
        {
            // Expected: nothing opened the item, so the tool reports that it could not.
        }
    }

    /// <summary>Runs operations straight against the stand-in session.</summary>
    private sealed class DirectGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal DirectGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => operation(_session);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }

    /// <summary>
    /// A session that finds nothing, counts how often it was asked, and reports the failure
    /// reason the test chose. A <see cref="DispatchProxy"/> rather than 26 stubs, so a method
    /// added to the contract needs no change here.
    /// </summary>
    private sealed class CountingSession
    {
        private readonly Dictionary<string, int> _calls = new Dictionary<string, int>(StringComparer.Ordinal);

        internal CountingSession(string? failureReason)
        {
            AsSession = RecordingSession.Create(this, failureReason);
        }

        internal IOutlookSession AsSession { get; }

        internal int CallsTo(string operation)
        {
            lock (_calls)
            {
                return _calls.TryGetValue(operation, out int count) ? count : 0;
            }
        }

        private void Record(string operation)
        {
            lock (_calls)
            {
                _calls[operation] = (_calls.TryGetValue(operation, out int count) ? count : 0) + 1;
            }
        }

        internal class RecordingSession : DispatchProxy
        {
            private CountingSession _owner = null!;
            private string? _failureReason;

            internal static IOutlookSession Create(CountingSession owner, string? failureReason)
            {
                object proxy = Create<IOutlookSession, RecordingSession>()
                    ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
                ((RecordingSession)proxy)._owner = owner;
                ((RecordingSession)proxy)._failureReason = failureReason;
                return (IOutlookSession)proxy;
            }

            /// <inheritdoc />
            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                ArgumentNullException.ThrowIfNull(targetMethod);

                _owner.Record(targetMethod.Name);

                if (string.Equals(targetMethod.Name, nameof(IOutlookSession.GetStoreDetails), StringComparison.Ordinal))
                {
                    return Stores;
                }

                // Every out parameter is assigned: an unassigned slot is unboxed on the way
                // back and would fail the call for a reason that has nothing to do with the
                // loop under test.
                ParameterInfo[] parameters = targetMethod.GetParameters();
                for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
                {
                    ParameterInfo parameter = parameters[i];
                    if (!parameter.IsOut)
                    {
                        continue;
                    }

                    Type slot = parameter.ParameterType.GetElementType()!;
                    if (slot == typeof(string) && string.Equals(parameter.Name, "error", StringComparison.Ordinal))
                    {
                        args[i] = _failureReason;
                        continue;
                    }

                    args[i] = slot.IsValueType ? Activator.CreateInstance(slot) : null;
                }

                Type returnType = targetMethod.ReturnType;
                return returnType != typeof(void) && returnType.IsValueType
                    ? Activator.CreateInstance(returnType)
                    : null;
            }
        }
    }
}
