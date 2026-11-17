using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// What a live run on the Exchange VM's test account may do to ONE item, decided from the item itself
/// (2026-10-08, the Phase 2 write-safety design): it must sit in the hub, its subject must carry the tag
/// AND this run's marker (<see cref="LiveRunMarker"/>), and - where the operation addresses mail, a reply,
/// a reply-all or a send - its sender and every recipient must be the hub. An item that is not this run's
/// is refused, so no untagged item and no other run's item is ever touched, whatever a test passes.
/// <para>
/// Shared by the in-process guard (<see cref="LiveGuardedOutlookSession"/>) and the test-side MCP client's
/// gate (<see cref="StdioHubOnlyGate"/>), which both read the item through the product's own session.
/// </para>
/// </summary>
public sealed class LiveRunItemCheck
{
    /// <summary>Builds the check for one run.</summary>
    /// <param name="hubStoreDisplayName">The hub: the one store whose items may be touched.</param>
    /// <param name="runMarker">This run's marker; every item touched must carry it.</param>
    /// <param name="recipients">The addresses mail may go to - the hub's.</param>
    public LiveRunItemCheck(string hubStoreDisplayName, string runMarker, RecipientAllowlist recipients)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hubStoreDisplayName);
        ArgumentException.ThrowIfNullOrWhiteSpace(runMarker);
        HubStoreDisplayName = hubStoreDisplayName;
        RunMarker = runMarker;
        Recipients = recipients ?? throw new ArgumentNullException(nameof(recipients));
    }

    /// <summary>The hub.</summary>
    public string HubStoreDisplayName { get; }

    /// <summary>This run's marker.</summary>
    public string RunMarker { get; }

    /// <summary>Whom mail may go to.</summary>
    public RecipientAllowlist Recipients { get; }

    /// <summary>The phrase every refusal of this check carries.</summary>
    public const string Refused = "NOT THIS RUN'S ITEM";

    /// <summary>
    /// The check for this process, or null where none is armed (every machine but the Exchange VM's test
    /// account).
    /// </summary>
    public static LiveRunItemCheck? Current => CurrentCheck.Value;

    private static readonly Lazy<LiveRunItemCheck?> CurrentCheck = new(
        () => For(LiveTestSettings.Load()), LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>The check <paramref name="settings"/> arm, or null.</summary>
    public static LiveRunItemCheck? For(LiveTestSettings settings)
    {
        RecipientAllowlist? recipients = LiveRecipientGuard.For(settings);
        return recipients == null ? null : new LiveRunItemCheck(settings.TestHubStoreDisplayName, LiveRunMarker.Current, recipients);
    }

    /// <summary>True when <paramref name="subject"/> carries the tag (ordinal) and this run's marker.</summary>
    public bool CarriesTagAndMarker(string? subject)
    {
        return subject != null
            && subject.Contains(LiveOutlookTestMailer.SubjectTag, StringComparison.Ordinal)
            && subject.Contains(RunMarker, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The refusal for a subject a test is about to give an item, or null.</summary>
    public string? RefusalForNewSubject(string? subject, string operation)
    {
        return CarriesTagAndMarker(subject)
            ? null
            : Refused + ": '" + operation + "' would make an item whose subject does not carry both "
                + LiveOutlookTestMailer.SubjectTag + " and this run's marker " + RunMarker + " - and an item that does not "
                + "carry both is one no sweep of this run may take back. Put the tag and the fixture's marker in it.";
    }

    /// <summary>
    /// The refusal for touching the item <paramref name="entryId"/>, or null. Read through
    /// <paramref name="session"/> (the product's own reader); content-free - the subject is matched, never
    /// shown. An item that cannot be read at all is passed: the operation that follows cannot find it
    /// either, and fails on its own (the negative tests that hand a deleted or bogus id rely on that).
    /// </summary>
    /// <param name="session">The product's session.</param>
    /// <param name="entryId">The item.</param>
    /// <param name="storeId">Its store, when known.</param>
    /// <param name="operation">What was about to happen.</param>
    /// <param name="addressesMail">True for a reply, a reply-all and a send: sender and recipients must be the hub.</param>
    public string? RefusalForItem(IOutlookSession session, string entryId, string? storeId, string operation, bool addressesMail)
    {
        ArgumentNullException.ThrowIfNull(session);
        ComItemDetail? item = session.TryReadItem(entryId, storeId, false, false, out _);
        if (item == null && storeId == null && HubStoreId(session) is { } hubStoreId)
        {
            // An EntryID handed in without its store opens against the DEFAULT store, and the hub is not
            // that one here (the real mailbox is the profile's first account): ask the hub as well, so an item of
            // the hub is judged rather than passed for want of a store.
            item = session.TryReadItem(entryId, hubStoreId, false, false, out _);
        }

        if (item == null)
        {
            return null;
        }

        string? refusal = RefusalForReading(
            item.StoreDisplayName, item.Subject, item.SenderAddress, item.Recipients.Select(r => r.Address), operation, addressesMail);
        return refusal;
    }

    /// <summary>The hub's StoreID, read off the profile's stores; null when the hub is not among them.</summary>
    public string? HubStoreId(IOutlookSession session)
    {
        ArgumentNullException.ThrowIfNull(session);
        return session.GetStoreDetails()
            .FirstOrDefault(d => string.Equals(d.DisplayName, HubStoreDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId;
    }

    /// <summary>The decision of <see cref="RefusalForItem"/> over what was read. Pure: T1 drives it.</summary>
    public string? RefusalForReading(
        string? storeDisplayName,
        string? subject,
        string? senderAddress,
        IEnumerable<string?> recipientAddresses,
        string operation,
        bool addressesMail)
    {
        if (!string.Equals(storeDisplayName, HubStoreDisplayName, StringComparison.OrdinalIgnoreCase))
        {
            return Refused + ": '" + operation + "' was handed an item in a store that is not the hub (it is in '"
                + (storeDisplayName ?? "(unknown)") + "'). On the Exchange VM only the test account may be written.";
        }

        if (!CarriesTagAndMarker(subject))
        {
            return Refused + ": '" + operation + "' was handed a hub item whose subject does not carry both "
                + LiveOutlookTestMailer.SubjectTag + " and this run's marker " + RunMarker + " (the subject is not shown). "
                + "A run touches only what it made itself.";
        }

        if (!addressesMail)
        {
            return null;
        }

        List<string?> addresses = new() { senderAddress };
        addresses.AddRange(recipientAddresses ?? Array.Empty<string?>());
        return Recipients.RefusalForItem(addresses, operation);
    }
}

/// <summary>
/// The product's <see cref="IOutlookSession"/>, with every write it can make on the Exchange VM's test
/// account checked first (2026-10-08): a draft must be addressed to the hub and nobody else, carry the tag
/// and the run's marker, and come from the hub's account; a reply or a reply-all only to this run's own
/// hub-to-hub mail; an update, a discard, a move or a display only of this run's own items; a send only of
/// a draft whose every recipient - read off the draft itself, the moment before - is the hub. Every read
/// passes straight through. A refusal throws <see cref="InvalidOperationException"/> before the write.
/// <para>
/// Installed by <see cref="LiveMailServices"/>, the one place the live fixtures get their
/// <see cref="MailService"/> from, so no in-process test reaches the product without it on that machine.
/// </para>
/// </summary>
public sealed class LiveGuardedOutlookSession : IOutlookSession
{
    private readonly IOutlookSession _inner;
    private readonly LiveRunItemCheck _check;

    /// <summary>Wraps <paramref name="inner"/>.</summary>
    public LiveGuardedOutlookSession(IOutlookSession inner, LiveRunItemCheck check)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _check = check ?? throw new ArgumentNullException(nameof(check));
    }

    private static void Refuse(string? refusal)
    {
        if (refusal != null)
        {
            throw new InvalidOperationException(refusal);
        }
    }

    private void RefuseRecipients(string operation, params IEnumerable<string>?[] fields)
    {
        Refuse(_check.Recipients.RefusalFor(fields.Where(f => f != null).SelectMany(f => f!).Select(s => (string?)s), operation));
    }

    private void RefuseItem(string entryIdHex, string? storeId, string operation, bool addressesMail)
    {
        Refuse(_check.RefusalForItem(_inner, entryIdHex, storeId, operation, addressesMail));
    }

    // ---------------------------------------------------------------- reads: straight through

    /// <inheritdoc/>
    public string GetProfileName() => _inner.GetProfileName();

    /// <inheritdoc/>
    public IReadOnlyList<ComAccountInfo> GetAccounts() => _inner.GetAccounts();

    /// <inheritdoc/>
    public IReadOnlyList<ComStoreDetail> GetStoreDetails() => _inner.GetStoreDetails();

    /// <inheritdoc/>
    public ComFolderTree ListFolders(string? storeEntryId, int absoluteWalkCap) => _inner.ListFolders(storeEntryId, absoluteWalkCap);

    /// <inheritdoc/>
    public ComFolderDetail? TryDescribeFolder(string folderEntryId, out string? error) => _inner.TryDescribeFolder(folderEntryId, out error);

    /// <inheritdoc/>
    public ComFolderDetail? TryCreateFolder(string parentFolderEntryId, string name, out string? error)
    {
        // A folder is made in the hub and nowhere else: the parent is opened (read-only) first, and its
        // store must be the hub - by its display name, as every other hub check here is made.
        ComFolderDetail? parent = _inner.TryDescribeFolder(parentFolderEntryId, out string? describeError);
        if (parent == null || !string.Equals(parent.StoreDisplayName, _check.HubStoreDisplayName, StringComparison.OrdinalIgnoreCase))
        {
            Refuse(LiveRunItemCheck.Refused + ": create_folder may make a folder only in the hub, and the parent "
                + (parent == null ? "could not be opened (" + (describeError ?? "no error") + ")" : "is in '" + parent.StoreDisplayName + "'") + ".");
        }

        return _inner.TryCreateFolder(parentFolderEntryId, name, out error);
    }

    /// <inheritdoc/>
    public ComFolderPathList ListFolderPaths(string storeDisplayName, int absoluteWalkCap) => _inner.ListFolderPaths(storeDisplayName, absoluteWalkCap);

    /// <inheritdoc/>
    public IReadOnlyList<IReadOnlyList<string>> FindFolderPathsByLeafName(string storeDisplayName, string leafName, int absoluteWalkCap)
        => _inner.FindFolderPathsByLeafName(storeDisplayName, leafName, absoluteWalkCap);

    /// <inheritdoc/>
    public ComSweepResult SweepFoldersNewerThan(
        DateTime sinceUtc,
        int perFolderCap,
        bool includeBodies,
        string? onlyStoreDisplayName,
        IReadOnlyList<string>? folderPath = null,
        bool includeSubfolders = true,
        IReadOnlyDictionary<string, DateTime>? perStoreSinceUtc = null,
        int timeBudgetMs = 0,
        string? folderEntryId = null)
        => _inner.SweepFoldersNewerThan(sinceUtc, perFolderCap, includeBodies, onlyStoreDisplayName, folderPath, includeSubfolders, perStoreSinceUtc, timeBudgetMs, folderEntryId);

    /// <inheritdoc/>
    public ComExhaustiveResult ExhaustiveScan(
        string storeDisplayName,
        IReadOnlyList<string>? folderPath,
        IReadOnlyList<string>? terms,
        DateTime? sinceUtc,
        DateTime? beforeUtc,
        int maxItems,
        int timeBudgetMs,
        SearchIn searchIn = SearchInValues.Default,
        bool includeSubfolders = false,
        ComScanCursor? resumeFrom = null,
        string? folderEntryId = null)
        => _inner.ExhaustiveScan(storeDisplayName, folderPath, terms, sinceUtc, beforeUtc, maxItems, timeBudgetMs, searchIn, includeSubfolders, resumeFrom, folderEntryId);

    /// <inheritdoc/>
    public ComItemDetail? TryReadItem(string entryIdHex, string? storeId, bool includeHeaders, bool includeBody, out string? error, bool includeHtml = false)
        => _inner.TryReadItem(entryIdHex, storeId, includeHeaders, includeBody, out error, includeHtml);

    /// <inheritdoc/>
    public IReadOnlyList<ComMailBrief>? TryGetConversationItems(string entryIdHex, string? storeId, int maxItems, out string? error)
        => _inner.TryGetConversationItems(entryIdHex, storeId, maxItems, out error);

    /// <inheritdoc/>
    public IReadOnlyList<ComAttachmentInfo> SnapshotAttachmentsById(string entryId, string? storeId) => _inner.SnapshotAttachmentsById(entryId, storeId);

    /// <inheritdoc/>
    public string? TrySaveAttachment(
        string entryIdHex,
        string? storeId,
        int attachmentIndex,
        string targetDirectory,
        out long sizeBytes,
        out string? attemptedPath,
        out string? error)
        => _inner.TrySaveAttachment(entryIdHex, storeId, attachmentIndex, targetDirectory, out sizeBytes, out attemptedPath, out error);

    /// <inheritdoc/>
    public ComOpenResult? TryResolveByPath(
        string storeDisplayName,
        IReadOnlyList<string> folderPath,
        string itemSubject,
        DateTime? indexReceivedUtc,
        out string? error,
        int toleranceSeconds = 120)
        => _inner.TryResolveByPath(storeDisplayName, folderPath, itemSubject, indexReceivedUtc, out error, toleranceSeconds);

    /// <inheritdoc/>
    public ComExplorerState? TryGotoFolder(string? folderEntryId, string? storeEntryId, out string? error)
        => _inner.TryGotoFolder(folderEntryId, storeEntryId, out error);

    /// <inheritdoc/>
    public ComExplorerState? TryShowSearchResults(string query, int olSearchScope, string? folderEntryId, string? storeEntryId, out string? error)
        => _inner.TryShowSearchResults(query, olSearchScope, folderEntryId, storeEntryId, out error);

    /// <inheritdoc/>
    public ComDraftInfo? TryGetMailInfo(string entryIdHex, string? storeId, out string? error) => _inner.TryGetMailInfo(entryIdHex, storeId, out error);

    /// <inheritdoc/>
    public ComArchiveFolderInfo? TryResolveArchiveFolder(string storeDisplayName, out string? error) => _inner.TryResolveArchiveFolder(storeDisplayName, out error);

    /// <inheritdoc/>
    public ComSendableDraftState? TryGetSendableDraftState(string entryIdHex, string? storeId, out string? error)
        => _inner.TryGetSendableDraftState(entryIdHex, storeId, out error);

    // ---------------------------------------------------------------- writes: checked first

    /// <inheritdoc/>
    public ComOpenResult? TryDisplayItem(string entryIdHex, string? storeId, out string? error)
    {
        // Opening an item marks it read: a change to the item, so only this run's own.
        RefuseItem(entryIdHex, storeId, "open_in_outlook", addressesMail: false);
        return _inner.TryDisplayItem(entryIdHex, storeId, out error);
    }

    /// <inheritdoc/>
    public ComArchiveFolderInfo? TryResolveOrCreateArchiveFolder(string storeDisplayName, out string? createdFolderPath, out string? error)
    {
        if (!string.Equals(storeDisplayName, _check.HubStoreDisplayName, StringComparison.OrdinalIgnoreCase))
        {
            Refuse(LiveRunItemCheck.Refused + ": the archive folder may be resolved or created only in the hub, not in '" + storeDisplayName + "'.");
        }

        return _inner.TryResolveOrCreateArchiveFolder(storeDisplayName, out createdFolderPath, out error);
    }

    /// <inheritdoc/>
    public ComMoveItemResult? TryMoveItemToFolder(string entryIdHex, string? storeId, string targetFolderEntryId, out string? error)
    {
        RefuseItem(entryIdHex, storeId, "move_mail", addressesMail: false);
        return _inner.TryMoveItemToFolder(entryIdHex, storeId, targetFolderEntryId, out error);
    }

    /// <inheritdoc/>
    public ComMoveItemResult? TryMoveItemToFolderId(string entryIdHex, string? storeId, string targetFolderEntryId, string targetStoreId, out string? error)
    {
        RefuseItem(entryIdHex, storeId, "move_mail", addressesMail: false);
        return _inner.TryMoveItemToFolderId(entryIdHex, storeId, targetFolderEntryId, targetStoreId, out error);
    }

    /// <inheritdoc/>
    public ComDraftCreateResult? TryCreateNewDraft(
        string accountSmtpAddress,
        IReadOnlyList<string> toRecipients,
        string subject,
        ComDraftBody body,
        bool display,
        ComSignatureOverride? signatureOverride,
        ComDraftOptions? options,
        out string? savedDraftEntryId,
        out string? error)
    {
        if (!_check.Recipients.IsAllowed(accountSmtpAddress))
        {
            Refuse(RecipientAllowlist.Refused + ": new_draft from the account '" + accountSmtpAddress + "' - on this machine a draft is made "
                + "only in the hub's own account.");
        }

        RefuseRecipients("new_draft", toRecipients, options?.CcRecipients, options?.BccRecipients);
        Refuse(_check.RefusalForNewSubject(subject, "new_draft"));
        return _inner.TryCreateNewDraft(accountSmtpAddress, toRecipients, subject, body, display, signatureOverride, options, out savedDraftEntryId, out error);
    }

    /// <inheritdoc/>
    public ComDraftCreateResult? TryCreateDerivedDraft(
        string sourceEntryIdHex,
        string? sourceStoreId,
        ComDerivedDraftKind kind,
        IReadOnlyList<string> toRecipients,
        ComDraftBody body,
        bool display,
        ComSignatureOverride? signatureOverride,
        ComDraftOptions? options,
        out string? savedDraftEntryId,
        out string? error)
    {
        string operation = kind switch
        {
            ComDerivedDraftKind.Reply => "reply_draft",
            ComDerivedDraftKind.ReplyAll => "replyall_draft",
            _ => "forward_draft",
        };

        // A reply goes to the source's sender, a reply-all to its sender and recipients: both are mail
        // addressed by the SOURCE, so the source must be this run's own hub-to-hub mail.
        RefuseItem(sourceEntryIdHex, sourceStoreId, operation, addressesMail: kind != ComDerivedDraftKind.Forward);
        RefuseRecipients(operation, toRecipients, options?.CcRecipients, options?.BccRecipients);
        if (options?.SubjectOverride != null)
        {
            Refuse(_check.RefusalForNewSubject(options.SubjectOverride, operation));
        }

        ComDraftCreateResult? created = _inner.TryCreateDerivedDraft(
            sourceEntryIdHex, sourceStoreId, kind, toRecipients, body, display, signatureOverride, options, out savedDraftEntryId, out error);

        // What Outlook actually put on the draft, read back: the rule is about the draft, not the arguments.
        if (savedDraftEntryId != null)
        {
            ComDraftInfo? draft = _inner.TryGetMailInfo(savedDraftEntryId, null, out _);
            if (draft != null)
            {
                // An entry Outlook could not resolve has no address - its name is what the test gave, and a
                // name carrying the run's marker resolves to nobody. The draft is judged again before any send.
                Refuse(_check.Recipients.RefusalForItem(
                    draft.Recipients.Select(r => string.IsNullOrWhiteSpace(r.Address) ? r.Name : r.Address),
                    operation + " (the saved draft)",
                    allowMarkedNames: true));
            }
        }

        return created;
    }

    /// <inheritdoc/>
    public ComDraftUpdateResult? TryUpdateDraft(
        string entryIdHex,
        string? storeId,
        ComDraftBody? body,
        string? subject,
        IReadOnlyList<string>? toRecipients,
        IReadOnlyList<string>? ccRecipients,
        IReadOnlyList<string>? bccRecipients,
        int? importance,
        bool? requestReadReceipt,
        ComSignatureOverride? signatureOverride,
        IReadOnlyList<string> attachmentsToAdd,
        IReadOnlyList<string> attachmentsToRemove,
        ComDraftUpdateResume? resume,
        bool display,
        out string? error)
    {
        RefuseItem(entryIdHex, storeId, "update_draft", addressesMail: false);
        RefuseRecipients("update_draft", toRecipients, ccRecipients, bccRecipients);
        if (subject != null)
        {
            Refuse(_check.RefusalForNewSubject(subject, "update_draft"));
        }

        return _inner.TryUpdateDraft(
            entryIdHex, storeId, body, subject, toRecipients, ccRecipients, bccRecipients, importance, requestReadReceipt,
            signatureOverride, attachmentsToAdd, attachmentsToRemove, resume, display, out error);
    }

    /// <inheritdoc/>
    public ComDraftDiscardResult? TryDiscardDraft(string entryIdHex, string? storeId, out string? error)
    {
        RefuseItem(entryIdHex, storeId, "discard_draft", addressesMail: false);
        return _inner.TryDiscardDraft(entryIdHex, storeId, out error);
    }

    /// <inheritdoc/>
    public ComSendResult? TrySendDraft(string entryIdHex, string? storeId, string expectedContentHash, string? sentOnBehalfOfName, out string? error)
    {
        // The draft as it stands now, the moment before it goes: its own recipients decide, not what a
        // test once asked for.
        ComSendableDraftState? state = _inner.TryGetSendableDraftState(entryIdHex, storeId, out _);
        if (state == null && storeId == null && _check.HubStoreId(_inner) is { } hubStoreId)
        {
            state = _inner.TryGetSendableDraftState(entryIdHex, hubStoreId, out _);
        }

        if (state != null)
        {
            Refuse(_check.RefusalForReading(
                state.StoreDisplayName, state.Subject, state.ResolvedAccountSmtp, state.Recipients.Select(r => r.Address), "send", addressesMail: true));
        }

        if (sentOnBehalfOfName != null && !_check.Recipients.IsAllowed(sentOnBehalfOfName))
        {
            Refuse(RecipientAllowlist.Refused + ": send on behalf of '" + sentOnBehalfOfName + "' - on this machine mail goes from the hub alone.");
        }

        return _inner.TrySendDraft(entryIdHex, storeId, expectedContentHash, sentOnBehalfOfName, out error);
    }
}

/// <summary>
/// An <see cref="IComGateway"/> whose every operation runs on a <see cref="LiveGuardedOutlookSession"/>.
/// Everything else - the connection, its loss, the diagnostics - is the inner gateway's.
/// </summary>
public sealed class LiveGuardedGateway : IComGateway
{
    private readonly IComGateway _inner;
    private readonly LiveRunItemCheck _check;

    /// <summary>Wraps <paramref name="inner"/>.</summary>
    public LiveGuardedGateway(IComGateway inner, LiveRunItemCheck check)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _check = check ?? throw new ArgumentNullException(nameof(check));
    }

    /// <inheritdoc/>
    public event Action? OutlookGone
    {
        add => _inner.OutlookGone += value;
        remove => _inner.OutlookGone -= value;
    }

    /// <inheritdoc/>
    public bool IsConnected => _inner.IsConnected;

    /// <inheritdoc/>
    public bool? QuitSinkActive => _inner.QuitSinkActive;

    /// <inheritdoc/>
    public bool ProbeConnected() => _inner.ProbeConnected();

    /// <inheritdoc/>
    public T Run<T>(Func<IOutlookSession, T> operation)
    {
        ArgumentNullException.ThrowIfNull(operation);
        return _inner.Run(session => operation(new LiveGuardedOutlookSession(session, _check)));
    }

    /// <inheritdoc/>
    public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
    {
        ArgumentNullException.ThrowIfNull(operation);
        return _inner.Run(session => operation(new LiveGuardedOutlookSession(session, _check)), recovery);
    }

    /// <inheritdoc/>
    public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
    {
        ArgumentNullException.ThrowIfNull(operation);
        return _inner.Run(session => operation(new LiveGuardedOutlookSession(session, _check)), budgetMilliseconds, allowConnectFloor);
    }

    /// <inheritdoc/>
    public ComHostDiagnostics GetDiagnostics() => _inner.GetDiagnostics();

    /// <inheritdoc/>
    public void Dispose() => _inner.Dispose();
}

/// <summary>
/// Where every live fixture gets its <see cref="MailService"/>: the product's own, over a
/// <see cref="LiveGuardedGateway"/> wherever <see cref="LiveRunItemCheck.Current"/> is armed (the Exchange
/// VM's test account), and exactly <see cref="MailService.CreateDefault"/> everywhere else.
/// </summary>
public static class LiveMailServices
{
    /// <summary>The default service, guarded where this machine arms the check.</summary>
    public static MailService CreateDefault()
    {
        LiveRunItemCheck? check = LiveRunItemCheck.Current;
        return check == null
            ? MailService.CreateDefault()
            : new MailService(new LiveGuardedGateway(new ComGateway(allowStartingOutlook: true), check));
    }

    /// <summary>A service with its own send-confirmation tokens (the short-TTL test), guarded the same way.</summary>
    public static MailService Create(SendConfirmationTokens tokens)
    {
        IComGateway gateway = new ComGateway(allowStartingOutlook: true);
        LiveRunItemCheck? check = LiveRunItemCheck.Current;
        return new MailService(check == null ? gateway : new LiveGuardedGateway(gateway, check), tokens);
    }
}
