using System.Reflection;
using System.Text.RegularExpressions;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The write paths only an Exchange profile can take, pinned at the unit tier because no live
/// tier can take them (Q74, decision D1). The test VMs carry POP3/PST profiles only, and the
/// maintainer's workstation - the one machine with Exchange - is read-only for live tests, so
/// send-on-behalf and every write into a delegate or shared mailbox are exercised by no live
/// test anywhere. What T1 cannot reach - what Outlook and Exchange then DO with these values -
/// is the manual pre-release checklist (D2).
/// <para>
/// Pinned here. (1) The two decisions the COM layer makes on these paths, extracted as pure
/// seams the shipped code calls: which <c>SentOnBehalfOfName</c> a confirmed send writes
/// (<see cref="OutlookComSession.OnBehalfOfToApply"/>), and whether an account delivers into the
/// store a draft lives in (<see cref="OutlookComSession.IsDeliveryStoreFor"/>). A source-level
/// check holds the call sites to them: <c>TrySendDraft</c> writes the on-behalf name only through
/// the seam, only after the sending identity is verified and before <c>Send()</c>, and never
/// picks the sending account from it. (2) The service layer's handling of
/// <c>sent_on_behalf_of</c> and of a draft no account delivers into, driven through a stand-in
/// session. The Exchange branches of the special-folder lookups behind new_draft, the derived
/// drafts, update_draft, discard_draft, move_mail and archive_mail are pinned in
/// <see cref="ReadOnlyFolderLookupTests"/>.
/// </para>
/// <para>
/// No Outlook, no COM and no mailbox. Like <see cref="AtomicityClaimsTests"/>, the service-layer
/// tests reach the product's audit log, which is the real one under %LOCALAPPDATA%: they append
/// <c>send_token_issued</c>, <c>send_refused</c> and <c>reply_draft</c> lines for synthetic ids.
/// None of them reaches a send - the stand-in stops every confirmed send before it would answer.
/// </para>
/// </summary>
public sealed class ExchangeWritePathTests
{
    /// <summary>A primary mailbox's StoreID, as read from the draft.</summary>
    private const string OwnStoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000";

    /// <summary>The same mailbox's StoreID as a different retrieval path wraps it.</summary>
    private const string OwnStoreIdWrappedElsewhere = "00000000DCA740C8C042101AB4B908002B2FE18201000000454D534D44422E444C4C00000000";

    /// <summary>A delegate (shared) mailbox's StoreID: an Exchange additional mailbox no account delivers into.</summary>
    private const string DelegateStoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C0000DE1E";

    private const string OwnStore = "me@example.test";

    private const string DelegateStore = "shared@example.test";

    /// <summary>Plausible bare EntryIDs: hex, even length, long enough to be accepted as one.</summary>
    private static readonly string DraftId = string.Concat(Enumerable.Repeat("D1A0", 35));

    private static readonly string SourceMailId = string.Concat(Enumerable.Repeat("D1B0", 35));

    private static readonly string ReplyDraftId = string.Concat(Enumerable.Repeat("D1C0", 35));

    // ------------------------------------------------------------------ the on-behalf seam

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t\r\n")]
    public void ABlankOnBehalf_WritesNothingOnTheDraft(string? requested)
    {
        // Nothing to write and nothing to report: the send goes out as the account itself.
        Assert.Null(OutlookComSession.OnBehalfOfToApply(requested));
    }

    [Theory]
    [InlineData("boss@example.test")]
    [InlineData("Boss@Example.TEST")]
    [InlineData("Sales Team")]
    public void AnOnBehalf_IsWrittenExactlyAsConfirmed(string requested)
    {
        // The service layer trims it and the token's hash compares it case-insensitively, so
        // the spelling is the caller's - this layer neither drops nor rewrites it.
        Assert.Equal(requested, OutlookComSession.OnBehalfOfToApply(requested));
    }

    [Fact]
    public void TheConfirmedSend_WritesOnBehalf_OnlyThroughTheSeam_AfterTheIdentityCheck_AndBeforeSend()
    {
        string code = CodeOf(SessionSource, nameof(OutlookComSession.TrySendDraft));

        Assert.Contains("OnBehalfOfToApply(sentOnBehalfOfName)", code, StringComparison.Ordinal);

        // ONE write of the property, of the seam's value - not a second, inline copy of the rule.
        MatchCollection writes = Regex.Matches(code, @"\.SentOnBehalfOfName\s*=(?!=)");
        Assert.True(writes.Count == 1, "TrySendDraft writes SentOnBehalfOfName " + writes.Count + " times; it must write it once.");
        Assert.Contains(".SentOnBehalfOfName = onBehalfOf;", code, StringComparison.Ordinal);

        // ORDER. Written only once the sending identity is verified, so an identity abort never
        // leaves an on-behalf name on the user's draft - and before Send(), or it is not sent.
        int verified = code.IndexOf("\"SendIdentityVerificationFailed\"", StringComparison.Ordinal);
        int sent = code.IndexOf(".Send();", StringComparison.Ordinal);
        Assert.True(verified >= 0 && sent >= 0, "TrySendDraft no longer reads as the sequence this test was written for.");
        Assert.True(verified < writes[0].Index, "SentOnBehalfOfName is written before the sending identity is verified.");
        Assert.True(writes[0].Index < sent, "SentOnBehalfOfName is written after Send().");

        // The raw request feeds exactly two things: the content hash and the seam.
        string[] uses = code.Split('\n')
            .Where(line => Regex.IsMatch(line, @"\bsentOnBehalfOfName\b"))
            .Select(line => line.Trim())
            .ToArray();
        Assert.True(
            uses.Length == 3,
            "sentOnBehalfOfName is expected in the signature, the hash and the seam only; found: " + string.Join(" | ", uses));
    }

    [Fact]
    public void TheSendingAccount_ComesFromTheDraftsOwnStore_NeverFromTheOnBehalfAddress()
    {
        // The plausible regression this guards: "send as the on-behalf account when one is
        // given". The on-behalf address names whom the mail is sent for - Exchange decides
        // whether that is allowed - and is never the identity a send is verified against: that
        // is the account delivering into the draft's store.
        string code = CodeOf(SessionSource, nameof(OutlookComSession.TrySendDraft));

        Assert.Contains("FindAccountByDeliveryStore(info.StoreId, info.StoreDisplayName)", code, StringComparison.Ordinal);
        Assert.DoesNotContain("FindAccountBySmtp(", code, StringComparison.Ordinal);
        Assert.DoesNotMatch(@"FindAccount\w*\([^)]*sentOnBehalfOfName", code);
    }

    // ------------------------------------------------------------------ which account delivers into a store

    [Theory]
    [InlineData(nameof(OutlookComSession.TryGetSendableDraftState))]
    [InlineData(nameof(OutlookComSession.TrySendDraft))]
    [InlineData(nameof(OutlookComSession.TryCreateDerivedDraft))]
    public void EveryStoreDerivedIdentity_GoesThroughTheDeliveryStoreLookup(string member)
    {
        // The send's step 1 (no_sending_account), its step 2, and the identity a reply or
        // forward is pinned to - all three derive the account from the store the item lives in.
        Assert.Contains("FindAccountByDeliveryStore(", CodeOf(SessionSource, member), StringComparison.Ordinal);
    }

    [Fact]
    public void TheDeliveryStoreLookup_DecidesThroughThePredicate_AndNowhereElse()
    {
        string code = CodeOf(SessionSource, "FindAccountByDeliveryStore");

        Assert.Contains(
            "IsDeliveryStoreFor(storeId, storeDisplayName, deliveryStoreId, deliveryStoreName)", code, StringComparison.Ordinal);
        Assert.DoesNotContain("string.Equals(", code, StringComparison.Ordinal);
    }

    [Fact]
    public void TheAccountsOwnStore_IsMatchedByStoreId_InAnyCase()
    {
        // The StoreID alone settles it: a display name the user renamed does not unmatch it.
        Assert.True(OutlookComSession.IsDeliveryStoreFor(OwnStoreId, "Renamed by the user", OwnStoreId, OwnStore));

        // ...whatever case the hex was read in. The names differ here on purpose: with equal
        // names the display-name fallback matches on its own, and a case-sensitive StoreID
        // comparison would pass unnoticed.
        Assert.True(OutlookComSession.IsDeliveryStoreFor(OwnStoreId.ToLowerInvariant(), "Renamed by the user", OwnStoreId, OwnStore));
    }

    [Fact]
    public void AStoreIdWrappedDifferently_StillMatchesByDisplayName()
    {
        // The fallback exists because store EntryID wrappings can differ between retrieval
        // paths. Without it, a mailbox whose StoreID came back wrapped differently reads as one
        // no account delivers into, and a send from it is refused as no_sending_account.
        Assert.True(OutlookComSession.IsDeliveryStoreFor(OwnStoreIdWrappedElsewhere, OwnStore, OwnStoreId, "ME@example.test"));
    }

    [Fact]
    public void ADelegateMailbox_IsNoAccountsDeliveryStore()
    {
        // THE Exchange-only case: an additional (delegate or shared) mailbox that is not itself
        // an account of the profile. No account delivers into it, so a draft there is refused
        // by send and a reply to mail there is not pinned to any account.
        Assert.False(OutlookComSession.IsDeliveryStoreFor(DelegateStoreId, DelegateStore, OwnStoreId, OwnStore));
    }

    [Theory]
    [InlineData("sales@example.test", "sales@example.test (Archive)")]
    [InlineData("sales@example.test (Archive)", "sales@example.test")]
    [InlineData("sales@example.test", " sales@example.test")]
    [InlineData("Mailbox - Sales", "Sales")]
    public void ANearMissName_IsNotTheSameStore(string storeName, string deliveryStoreName)
    {
        // Whole-string equality only: a looser match would adopt a delegate mailbox whose name
        // merely contains an account's, and send from that account.
        Assert.False(OutlookComSession.IsDeliveryStoreFor(DelegateStoreId, storeName, OwnStoreId, deliveryStoreName));
    }

    [Theory]
    [InlineData(null, null, null, null)]
    [InlineData(null, null, OwnStoreId, OwnStore)]
    [InlineData(OwnStoreId, OwnStore, null, null)]
    [InlineData(null, OwnStore, OwnStoreId, null)]
    [InlineData(OwnStoreId, null, null, OwnStore)]
    public void AnIdentityThatWouldNotRead_MatchesNothing(
        string? storeId, string? storeName, string? deliveryStoreId, string? deliveryStoreName)
    {
        // Two unreadable values are not "the same store": that would hand a draft whose store
        // could not be read to the first account whose delivery store could not be read either.
        Assert.False(OutlookComSession.IsDeliveryStoreFor(storeId, storeName, deliveryStoreId, deliveryStoreName));
    }

    // ------------------------------------------------------------------ send: sent_on_behalf_of through the service

    [Theory]
    [InlineData("  boss@example.test  ", "boss@example.test")]
    [InlineData("Boss@Example.test", "Boss@Example.test")]
    [InlineData("   ", null)]
    [InlineData(null, null)]
    public void StepOne_ShowsTheOnBehalfTheTokenIsBoundTo_AndSendsNothing(string? requested, string? shown)
    {
        StandIn standIn = StandIn.WithDraft(OwnDraft());
        using MailService service = new MailService(new DirectGateway(standIn.AsSession));

        SendOutcome step1 = service.Send(DraftId, confirmToken: null, sentOnBehalfOf: requested);

        Assert.Equal("confirmation_required", step1.Status);
        Assert.False(step1.Sent);
        Assert.NotNull(step1.ConfirmToken);
        Assert.Equal(shown, step1.SentOnBehalfOf);
        Assert.Empty(standIn.SendRequests);
    }

    [Fact]
    public void TheConfirmedOnBehalf_ReachesOutlook_UnderTheHashTheTokenWasIssuedFor()
    {
        // Whitespace and case may differ between the two calls - the hash folds both - and what
        // reaches the COM layer is the trimmed value of the CONFIRMING call, bound into the hash
        // that layer re-verifies immediately before Send().
        ComSendableDraftState draft = OwnDraft();
        StandIn standIn = StandIn.WithDraft(draft);
        using MailService service = new MailService(new DirectGateway(standIn.AsSession));

        SendOutcome step1 = service.Send(DraftId, null, " Boss@Example.test ");
        Assert.Throws<StandInStoppedTheSend>(() => service.Send(DraftId, step1.ConfirmToken, "boss@example.test "));

        SendRequest request = Assert.Single(standIn.SendRequests);
        Assert.Equal(DraftId, request.EntryId);
        Assert.Equal(OwnStoreId, request.StoreId);
        Assert.Equal("boss@example.test", request.SentOnBehalfOfName);
        Assert.Equal(HashOf(draft, "boss@example.test"), request.ExpectedContentHash);
        Assert.NotEqual(HashOf(draft, null), request.ExpectedContentHash);
    }

    [Theory]
    [InlineData("boss@example.test", "other@example.test")]
    [InlineData("boss@example.test", null)]
    [InlineData(null, "boss@example.test")]
    public void AnOnBehalfThatChangedBetweenTheTwoCalls_IsRefused_AndOutlookIsNeverAskedToSend(string? confirmed, string? sentWith)
    {
        // The user confirmed one sender line; a different one - or none, or one added only on
        // the confirming call - is a different mail, so the token does not cover it.
        StandIn standIn = StandIn.WithDraft(OwnDraft());
        using MailService service = new MailService(new DirectGateway(standIn.AsSession));

        SendOutcome step1 = service.Send(DraftId, null, confirmed);
        SendRefusedException refusal = Assert.Throws<SendRefusedException>(
            () => service.Send(DraftId, step1.ConfirmToken, sentWith));

        Assert.Equal("draft_changed", refusal.Reason);
        Assert.Contains("Nothing was sent", refusal.Message, StringComparison.Ordinal);
        Assert.Empty(standIn.SendRequests);
    }

    // ------------------------------------------------------------------ send: a draft no account delivers into

    [Theory]
    [InlineData(null)]
    [InlineData("shared@example.test")]
    public void ADraftInAMailboxNoAccountDeliversInto_IsRefusedBeforeAnyToken_EvenWithAnOnBehalf(string? onBehalfOf)
    {
        // A draft in a delegate mailbox. Naming that mailbox as the on-behalf sender does not
        // make it sendable: on-behalf never chooses the account a send is verified against.
        StandIn standIn = StandIn.WithDraft(DelegateDraft());
        SendConfirmationTokens tokens = new SendConfirmationTokens();
        using MailService service = new MailService(new DirectGateway(standIn.AsSession), tokens);

        SendRefusedException refusal = Assert.Throws<SendRefusedException>(
            () => service.Send(DraftId, confirmToken: null, sentOnBehalfOf: onBehalfOf));

        Assert.Equal("no_sending_account", refusal.Reason);
        Assert.Equal(0, tokens.PendingCount);
        Assert.Empty(standIn.SendRequests);
    }

    [Fact]
    public void AStoreNoAccountDeliversInto_AtTheMomentOfSending_IsARefusal_NotAnUnknownOutcome()
    {
        // The COM layer's own check, made again inside the send: by then no account delivers
        // into the draft's store (or the one that does reports no SMTP address). Nothing was
        // sent - Send() was never called - so this must stay a named refusal and not fall
        // through to the catch-all.
        StandIn standIn = StandIn.WithDraft(OwnDraft());
        standIn.SendRefusal = "NoSendingAccountForStore";
        using MailService service = new MailService(new DirectGateway(standIn.AsSession));

        SendOutcome step1 = service.Send(DraftId);
        SendRefusedException refusal = Assert.Throws<SendRefusedException>(() => service.Send(DraftId, step1.ConfirmToken));

        Assert.Equal("no_sending_account", refusal.Reason);
        Assert.Contains("Nothing was sent", refusal.Message, StringComparison.Ordinal);
        Assert.Single(standIn.SendRequests);
    }

    // ------------------------------------------------------------------ a reply to mail in a delegate mailbox

    [Fact]
    public void AReplyToMailNoAccountDeliversInto_ReportsItsSenderAsUnpinned_EvenWhenTheDraftNamesAnAccount()
    {
        // accountResolved says whether the PRODUCT pinned SendUsingAccount from a matched
        // account - not whether the saved draft reports one. On delegate mail nothing matched,
        // so the From is whatever Outlook picks, and that is what the agent must be told.
        StandIn standIn = StandIn.WithDerivedDraft(new ComDraftCreateResult(
            new ComDraftInfo(
                ReplyDraftId,
                DelegateStore,
                DelegateStoreId,
                "Drafts",
                "folder-drafts",
                "RE: Rota",
                sendUsingAccountSmtp: OwnStore,
                conversationIndex: null,
                conversationId: "conversation-1",
                recipients: new[] { new ComRecipientInfo("to", "Rota", "rota@example.test") }),
            accountResolved: false,
            signatureInjected: true,
            bodyTextCharsBeforeSignature: 0,
            bodyTextCharsAfterSignature: 0,
            movedToDrafts: false,
            initialSaveFolderName: "Drafts",
            displayed: false,
            bodyPlacedViaWordEditor: true));
        using MailService service = new MailService(new DirectGateway(standIn.AsSession));

        DraftOutcome reply = service.ReplyDraft(SourceMailId, "Noted, thanks.", display: false);

        Assert.False(reply.AccountResolved);
        Assert.Equal(DelegateStore, reply.Store);
        Assert.Equal(ReplyDraftId, reply.EntryId);
    }

    // ------------------------------------------------------------------ support

    private const string SessionSource = "McpServer/OutlookAI.Core/Com/OutlookComSession.cs";

    /// <summary>
    /// A member declaration at class-member indentation - EXACTLY eight spaces in these sources -
    /// the grammar <see cref="ReadOnlyFolderLookupTests"/> reads the same file with.
    /// </summary>
    private static readonly Regex MemberDeclaration = new Regex(
        @"^ {8}(?![ /\[{}#])[^=;]*?\b(?<name>\w+)\s*(?:<[^>]*>)?\s*\(",
        RegexOptions.CultureInvariant);

    private static ComSendableDraftState OwnDraft() => Draft(OwnStoreId, OwnStore, resolvedAccount: OwnStore);

    private static ComSendableDraftState DelegateDraft() => Draft(DelegateStoreId, DelegateStore, resolvedAccount: null);

    private static ComSendableDraftState Draft(string storeId, string store, string? resolvedAccount)
    {
        return new ComSendableDraftState(
            DraftId,
            storeId,
            store,
            "Drafts",
            "Quarterly figures",
            isSent: false,
            "Please find the figures attached.",
            resolvedAccount,
            new[] { new ComRecipientInfo("to", "Finance", "finance@example.test") },
            attachments: null,
            bodyHtmlDigest: SendContentHash.DigestHtml("<p>Please find the figures attached.</p>"));
    }

    private static string HashOf(ComSendableDraftState draft, string? onBehalfOf)
    {
        return SendContentHash.Compute(
            draft.Subject, draft.Recipients, draft.BodyText, onBehalfOf, draft.Attachments, draft.BodyHtmlDigest);
    }

    /// <summary>
    /// The CODE of one member: its declaration line through the line before the next member
    /// declaration, with comment lines dropped - the doc comment of the member that follows
    /// sits inside that range, and prose must never satisfy or fail a check about code.
    /// </summary>
    private static string CodeOf(string relativePath, string member)
    {
        string[] lines = File.ReadAllLines(Path.Combine(RepoRoot(), relativePath));
        int start = Array.FindIndex(lines, line => DeclaredName(line) == member);
        Assert.True(start >= 0, member + " was not found in " + relativePath + " - this check has stopped proving anything.");

        List<string> code = new List<string> { lines[start] };
        for (int i = start + 1; i < lines.Length && DeclaredName(lines[i]) == null; i++)
        {
            if (!lines[i].TrimStart().StartsWith("//", StringComparison.Ordinal))
            {
                code.Add(lines[i]);
            }
        }

        return string.Join("\n", code);
    }

    private static string? DeclaredName(string line)
    {
        Match declaration = MemberDeclaration.Match(line);
        return declaration.Success ? declaration.Groups["name"].Value : null;
    }

    private static string RepoRoot()
    {
        string testProjectDir = typeof(ExchangeWritePathTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }

    /// <summary>What the service asked the COM layer to send.</summary>
    private sealed record SendRequest(string EntryId, string? StoreId, string ExpectedContentHash, string? SentOnBehalfOfName);

    /// <summary>
    /// Thrown by the stand-in at the point a real session would start sending. Nothing in the
    /// send path catches it, so the confirmed send ends there: no transport, no send audit line.
    /// </summary>
    private sealed class StandInStoppedTheSend : Exception
    {
        public StandInStoppedTheSend()
            : base("The stand-in session stops every confirmed send before it would answer.")
        {
        }
    }

    /// <summary>Runs operations straight against the stand-in session, with no budget layer.</summary>
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
    /// A session holding one draft (for send) or answering one reply (for the derived-draft
    /// path), recording what it is asked to send, and refusing every other call - so a flow
    /// that starts asking Outlook something new fails here instead of passing on a default.
    /// </summary>
    private sealed class StandIn
    {
        private StandIn()
        {
            AsSession = Proxy.Create(this);
        }

        internal IOutlookSession AsSession { get; }

        internal List<SendRequest> SendRequests { get; } = new List<SendRequest>();

        /// <summary>The COM refusal token a confirmed send answers with; null stops it by throwing.</summary>
        internal string? SendRefusal { get; set; }

        private ComSendableDraftState? Draft { get; set; }

        private ComDraftCreateResult? DerivedDraft { get; set; }

        internal static StandIn WithDraft(ComSendableDraftState draft) => new StandIn { Draft = draft };

        internal static StandIn WithDerivedDraft(ComDraftCreateResult derived) => new StandIn { DerivedDraft = derived };

        private object? Handle(MethodInfo method, object?[]? args)
        {
            switch (method.Name)
            {
                case nameof(IOutlookSession.TryGetSendableDraftState) when Draft != null:
                    SetOut(method, args, "error", null);
                    return Draft;

                case nameof(IOutlookSession.TrySendDraft) when Draft != null:
                    SendRequests.Add(new SendRequest(
                        (string)Arg(method, args, "entryIdHex")!,
                        (string?)Arg(method, args, "storeId"),
                        (string)Arg(method, args, "expectedContentHash")!,
                        (string?)Arg(method, args, "sentOnBehalfOfName")));
                    if (SendRefusal == null)
                    {
                        throw new StandInStoppedTheSend();
                    }

                    SetOut(method, args, "error", SendRefusal);
                    return null;

                case nameof(IOutlookSession.TryCreateDerivedDraft) when DerivedDraft != null:
                    SetOut(method, args, "savedDraftEntryId", DerivedDraft.Draft.EntryId);
                    SetOut(method, args, "createdFolder", null);
                    SetOut(method, args, "error", null);
                    return DerivedDraft;

                default:
                    throw new InvalidOperationException(
                        "The stand-in session was asked for '" + method.Name + "', which this flow is not expected to need.");
            }
        }

        private static object? Arg(MethodInfo method, object?[]? args, string name)
        {
            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
            {
                if (string.Equals(parameters[i].Name, name, StringComparison.Ordinal))
                {
                    return args[i];
                }
            }

            throw new InvalidOperationException(method.Name + " has no parameter named '" + name + "'.");
        }

        private static void SetOut(MethodInfo method, object?[]? args, string name, object? value)
        {
            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
            {
                if (parameters[i].IsOut && string.Equals(parameters[i].Name, name, StringComparison.Ordinal))
                {
                    args[i] = value;
                }
            }
        }

        /// <summary>Not sealed: <see cref="DispatchProxy"/> derives from it at runtime.</summary>
        internal class Proxy : DispatchProxy
        {
            private StandIn _owner = null!;

            internal static IOutlookSession Create(StandIn owner)
            {
                object proxy = Create<IOutlookSession, Proxy>()
                    ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
                ((Proxy)proxy)._owner = owner;
                return (IOutlookSession)proxy;
            }

            /// <inheritdoc />
            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                ArgumentNullException.ThrowIfNull(targetMethod);
                return _owner.Handle(targetMethod, args);
            }
        }
    }
}
