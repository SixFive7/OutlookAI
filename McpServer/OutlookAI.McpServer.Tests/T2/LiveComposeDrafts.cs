using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// One Q149 run's identity: the hub, the marker every subject carries, the EntryIDs this run created - the
/// allowlist every change and delete is checked against - and where the attempt lines go.
/// </summary>
public sealed class Q149Run
{
    private readonly ConcurrentDictionary<string, byte> _created = new(StringComparer.OrdinalIgnoreCase);
    private readonly Action<string> _write;

    public Q149Run(LiveTestSettings settings, string marker, Action<string> write)
    {
        Settings = settings ?? throw new ArgumentNullException(nameof(settings));
        Marker = string.IsNullOrWhiteSpace(marker) || marker.Length < 12
            ? throw new ArgumentException("A marker of 12 characters or more is required (S3).", nameof(marker))
            : marker;
        _write = write ?? throw new ArgumentNullException(nameof(write));
    }

    public LiveTestSettings Settings { get; }

    public string Hub => Settings.TestHubStoreDisplayName;

    /// <summary>The hub's own address: its store's display name, which on every test guest is the account's address.</summary>
    public string HubAddress => Settings.TestHubStoreDisplayName;

    public string Marker { get; }

    public IReadOnlyCollection<string> Created => _created.Keys.ToList();

    public string TaggedSubject(string label)
    {
        return LiveOutlookTestMailer.TaggedSubject(Marker, "q149 " + label);
    }

    /// <summary>Records an EntryID as one this run created, so it may later be changed or deleted.</summary>
    public void Allow(string? entryId)
    {
        if (!string.IsNullOrWhiteSpace(entryId))
        {
            _created[entryId!] = 0;
        }
    }

    /// <summary>Writes one attempt line, to the test output and the console capture both.</summary>
    public void Write(Q149Attempt attempt)
    {
        ArgumentNullException.ThrowIfNull(attempt);
        string line = attempt.ToLine();
        _write(line);
        Console.WriteLine(line);
    }

    /// <summary>One timestamped progress line, to the test output and the console capture both - so a run whose results
    /// cannot be fetched whole still says how far each step got, and when.</summary>
    public void Log(string line)
    {
        string stamped = "q149 " + DateTime.UtcNow.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + ": " + line;
        _write(stamped);
        Console.WriteLine(stamped);
    }
}

/// <summary>
/// Every write the Q149 measurement makes to the test mailbox, each refused before Outlook is asked
/// unless it is a draft in the hub carrying the tag and this run's marker - and, for a change to an item
/// that already exists, one this run created (<see cref="Q149ProbeGuards"/>).
/// </summary>
internal static class Q149Drafts
{
    /// <summary>
    /// Saves a new draft in the hub's Drafts with no inspector at all - subject first, so it is tagged from
    /// the moment it exists - optionally with an HTML body, and returns its EntryID and the hub's StoreID.
    /// The object <c>Items.Add</c> returned is released at once; nothing ever opens an inspector on it.
    /// </summary>
    internal static (string EntryId, string StoreId, Q149HtmlFacts BeforeBody) CreateTaggedDraft(
        Q149Run run, Q149OutlookClient client, string label, string? htmlBody, bool addressToHub)
    {
        string subject = run.TaggedSubject(label);
        Refuse(Q149ProbeGuards.RefuseSubject(subject, run.Marker));
        LiveStoreWriteGuard.Assert(run.Hub, StoreWriteKind.Draft, "q149 create draft");
        (string entryId, string storeId, Q149HtmlFacts before) = client.Run((app, ns) =>
        {
            dynamic? store = null;
            dynamic? drafts = null;
            dynamic? items = null;
            dynamic? mail = null;
            try
            {
                store = FindStore(ns, run.Hub) ?? throw new InvalidOperationException("Q149: the hub store is not in the profile.");
                drafts = store.GetDefaultFolder(LiveComposeProbe.Drafts);
                items = drafts.Items;
                mail = items.Add(0);
                mail.Subject = subject;
                SetSendUsingHub(ns, mail, run.HubAddress);
                if (addressToHub)
                {
                    AddRecipient(mail, run.HubAddress);
                }

                Q149HtmlFacts stub = Q149Html.Analyze((string?)mail.HTMLBody);
                if (htmlBody != null)
                {
                    mail.HTMLBody = htmlBody;
                }

                mail.Save();
                return ((string)mail.EntryID, (string)store.StoreID, stub);
            }
            finally
            {
                Q149Com.Release(mail);
                Q149Com.Release(items);
                Q149Com.Release(drafts);
                Q149Com.Release(store);
            }
        });
        run.Allow(entryId);
        return (entryId, storeId, before);
    }

    /// <summary>
    /// A mail this run sends the hub itself, tagged and in the run's allowlist once it has arrived - the source every
    /// reply and forward of the measurement is made from. Returns its EntryID; <paramref name="quoteToken"/> is a word
    /// that occurs only in its body.
    /// </summary>
    internal static string SeedSelfSentSource(Q149Run run, string label, out string quoteToken)
    {
        quoteToken = "Q149QUOTE" + run.Marker;
        string subject = run.TaggedSubject(label);
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(run.Hub, subject, "Origineel bericht " + quoteToken + "\r\nTweede regel.", attachmentPath: null);
        using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
        ComMailBrief arrived = LiveInboxArrival.WaitFor(session, run.Hub, subject, sentUtc);
        run.Allow(arrived.EntryId);
        return arrived.EntryId;
    }

    /// <summary>The hub's StoreID, read through <paramref name="client"/>.</summary>
    internal static string HubStoreId(Q149Run run, Q149OutlookClient client)
    {
        return client.Run((app, ns) =>
        {
            dynamic? store = FindStore(ns, run.Hub);
            try
            {
                return (string)store!.StoreID;
            }
            finally
            {
                Q149Com.Release(store);
            }
        });
    }

    /// <summary>
    /// The HTML-only draft (c): a new draft in the hub, its pictures attached and marked inline
    /// (<c>PR_ATTACH_CONTENT_ID</c>, hidden, MIME type), its <c>HTMLBody</c> written whole, saved - no
    /// inspector, no Word. <paramref name="buildHtml"/> gets the cid of every picture, in order, and the
    /// stub HTML Outlook gave the new item.
    /// </summary>
    internal static (string EntryId, Q149HtmlFacts Stub) CreateHtmlOnlyDraft(
        Q149Run run,
        Q149OutlookClient client,
        string label,
        IReadOnlyList<(string Path, string Cid, string Mime)> pictures,
        Func<string?, string> buildHtml)
    {
        string subject = run.TaggedSubject(label);
        Refuse(Q149ProbeGuards.RefuseSubject(subject, run.Marker));
        LiveStoreWriteGuard.Assert(run.Hub, StoreWriteKind.Draft, "q149 html-only draft");
        (string entryId, Q149HtmlFacts stub) = client.Run((app, ns) =>
        {
            dynamic? store = null;
            dynamic? drafts = null;
            dynamic? items = null;
            dynamic? mail = null;
            try
            {
                store = FindStore(ns, run.Hub) ?? throw new InvalidOperationException("Q149: the hub store is not in the profile.");
                drafts = store.GetDefaultFolder(LiveComposeProbe.Drafts);
                items = drafts.Items;
                mail = items.Add(0);
                mail.Subject = subject;
                SetSendUsingHub(ns, mail, run.HubAddress);
                AddRecipient(mail, run.HubAddress);
                string? stubHtml = (string?)mail.HTMLBody;
                AttachInline(mail, pictures);
                mail.HTMLBody = buildHtml(stubHtml);
                mail.Save();
                return ((string)mail.EntryID, Q149Html.Analyze(stubHtml));
            }
            finally
            {
                Q149Com.Release(mail);
                Q149Com.Release(items);
                Q149Com.Release(drafts);
                Q149Com.Release(store);
            }
        });
        run.Allow(entryId);
        return (entryId, stub);
    }

    /// <summary>
    /// A reply or forward made by the object model with no inspector: <c>Reply()</c> or <c>Forward()</c> on a
    /// mail this run sent itself, the hub as the sending account (and, for a forward, as the recipient),
    /// saved and released. What Outlook put in its HTML before any inspector existed is the C4 measurement.
    /// </summary>
    internal static (string EntryId, string Html) CreateDerivedWithoutInspector(
        Q149Run run, Q149OutlookClient client, string sourceEntryId, string storeId, bool forward, string label)
    {
        Refuse(Q149ProbeGuards.RefuseEntryId(sourceEntryId, run.Created));
        LiveStoreWriteGuard.Assert(run.Hub, StoreWriteKind.Draft, "q149 derived draft");
        (string entryId, string html) = client.Run((app, ns) =>
        {
            dynamic? source = null;
            dynamic? derived = null;
            try
            {
                source = ns.GetItemFromID(sourceEntryId, storeId);
                Refuse(Q149ProbeGuards.RefuseSubject((string?)source.Subject, run.Marker));
                derived = forward ? source.Forward() : source.Reply();
                if (Q149ProbeGuards.RefuseSubject((string?)derived.Subject, run.Marker) != null)
                {
                    derived.Subject = run.TaggedSubject(label);
                }

                SetSendUsingHub(ns, derived, run.HubAddress);
                if (forward)
                {
                    AddRecipient(derived, run.HubAddress);
                }

                derived.Save();
                return ((string)derived.EntryID, (string?)derived.HTMLBody ?? string.Empty);
            }
            finally
            {
                Q149Com.Release(derived);
                Q149Com.Release(source);
            }
        });
        run.Allow(entryId);
        return (entryId, html);
    }

    /// <summary>
    /// Rewrites one of this run's drafts' <c>HTMLBody</c> through <paramref name="transform"/>, attaches
    /// <paramref name="pictures"/> first when given, and saves - no inspector. Refused unless the EntryID
    /// is one this run created, the subject carries the tag and the marker, and the draft is unsent.
    /// </summary>
    internal static Q149HtmlFacts RewriteHtml(
        Q149Run run,
        Q149OutlookClient client,
        string entryId,
        string storeId,
        Func<string, string?> transform,
        IReadOnlyList<(string Path, string Cid, string Mime)>? pictures = null)
    {
        Refuse(Q149ProbeGuards.RefuseEntryId(entryId, run.Created));
        LiveStoreWriteGuard.Assert(run.Hub, StoreWriteKind.Draft, "q149 rewrite draft html");
        return client.Run((app, ns) =>
        {
            dynamic? item = null;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                RequireOwnUnsentDraft(run, item);
                string html = (string?)item.HTMLBody ?? string.Empty;
                string? rewritten = transform(html) ?? throw new InvalidOperationException("Q149: the HTML rewrite found nothing to rewrite.");
                if (pictures != null && pictures.Count > 0)
                {
                    AttachInline(item, pictures);
                }

                item.HTMLBody = rewritten;
                item.Save();
                return Q149Html.Analyze((string?)item.HTMLBody);
            }
            finally
            {
                Q149Com.Release(item);
            }
        });
    }

    /// <summary>Reads an item's HTMLBody, recipients, message class and form properties through its own fresh client - the separate-call re-read.</summary>
    internal static (string Html, IReadOnlyList<string> Recipients, string MessageClass, IReadOnlyList<string> FormPropertiesPresent, int Attachments, long AttachmentBytes, int HiddenAttachments)
        Read(Q149Run run, string entryId, string storeId)
    {
        using Q149OutlookClient reader = Q149OutlookClient.Connect(new Stopwatch(), allowStart: false);
        return reader.Run((app, ns) =>
        {
            dynamic? item = null;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                Refuse(Q149ProbeGuards.RefuseSubject((string?)item.Subject, run.Marker));
                string html = (string?)item.HTMLBody ?? string.Empty;
                List<string> recipients = ReadRecipients(item);
                string messageClass = (string?)item.MessageClass ?? string.Empty;
                IReadOnlyList<string> present = LiveComposeProbe.FormPropertiesPresent((object)item);
                (int count, long bytes, int hidden) = ReadAttachments((object)item);
                return (html, (IReadOnlyList<string>)recipients, messageClass, present, count, bytes, hidden);
            }
            finally
            {
                Q149Com.Release(item);
            }
        });
    }

    /// <summary>Deletes every item this run created, each by EntryID AND the ordinal tag and marker (rule 2), through the tested helper.</summary>
    internal static int DeleteCreated(Q149Run run)
    {
        int deleted = 0;
        foreach (string entryId in run.Created)
        {
            try
            {
                if (LiveOutlookTestMailer.DeleteItemByEntryId(run.Hub, entryId, run.Marker))
                {
                    deleted++;
                }
            }
            catch (InvalidOperationException ex)
            {
                // Already gone, or sent (a sent draft has a new EntryID in Sent Items): the marker sweep covers those.
                run.Log("delete by EntryID did not apply (" + ex.GetType().Name + ") - left to the marker sweep");
            }
        }

        return deleted;
    }

    internal static void RequireOwnUnsentDraft(Q149Run run, dynamic item)
    {
        Refuse(Q149ProbeGuards.RefuseSubject((string?)item.Subject, run.Marker));
        if ((bool)item.Sent)
        {
            throw new InvalidOperationException("Q149: refusing to change an item that has been sent.");
        }
    }

    internal static void Refuse(string? why)
    {
        if (why != null)
        {
            throw new InvalidOperationException("Q149 refused a mailbox write: " + why + ".");
        }
    }

    internal static dynamic? FindStore(dynamic ns, string displayName)
    {
        dynamic stores = ns.Stores;
        try
        {
            int count = stores.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic store = stores[i];
                string? name = null;
                try
                {
                    name = (string?)store.DisplayName;
                }
                catch (COMException)
                {
                }

                if (string.Equals(name, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    return store;
                }

                Q149Com.Release(store);
            }

            return null;
        }
        finally
        {
            Q149Com.Release(stores);
        }
    }

    internal static List<string> ReadRecipients(dynamic item)
    {
        List<string> addresses = new();
        dynamic recipients = item.Recipients;
        try
        {
            int count = recipients.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic recipient = recipients[i];
                try
                {
                    addresses.Add((string?)recipient.Address ?? (string?)recipient.Name ?? string.Empty);
                }
                finally
                {
                    Q149Com.Release(recipient);
                }
            }
        }
        finally
        {
            Q149Com.Release(recipients);
        }

        return addresses;
    }

    private static (int Count, long Bytes, int Hidden) ReadAttachments(object itemObject)
    {
        int count = 0;
        long bytes = 0;
        int hidden = 0;
        dynamic attachments = ((dynamic)itemObject).Attachments;
        try
        {
            count = attachments.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic attachment = attachments[i];
                try
                {
                    try
                    {
                        bytes += (int)attachment.Size;
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }

                    dynamic accessor = attachment.PropertyAccessor;
                    try
                    {
                        if (accessor.GetProperty(LiveComposeProbe.PrAttachmentHidden) is bool isHidden && isHidden)
                        {
                            hidden++;
                        }
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Q149Com.Release(accessor);
                    }
                }
                finally
                {
                    Q149Com.Release(attachment);
                }
            }
        }
        finally
        {
            Q149Com.Release(attachments);
        }

        return (count, bytes, hidden);
    }

    private static void AttachInline(dynamic mail, IReadOnlyList<(string Path, string Cid, string Mime)> pictures)
    {
        dynamic attachments = mail.Attachments;
        try
        {
            foreach ((string path, string cid, string mime) in pictures)
            {
                dynamic attachment = attachments.Add(path, 1, 0, Path.GetFileName(path));
                dynamic accessor = attachment.PropertyAccessor;
                try
                {
                    accessor.SetProperty(LiveComposeProbe.PrAttachContentId, cid);
                    accessor.SetProperty(LiveComposeProbe.PrAttachMimeTag, mime);
                    accessor.SetProperty(LiveComposeProbe.PrAttachmentHidden, true);
                }
                finally
                {
                    Q149Com.Release(accessor);
                    Q149Com.Release(attachment);
                }
            }
        }
        finally
        {
            Q149Com.Release(attachments);
        }
    }

    private static void AddRecipient(dynamic mail, string address)
    {
        dynamic recipients = mail.Recipients;
        try
        {
            dynamic recipient = recipients.Add(address);
            Q149Com.Release(recipient);
            _ = recipients.ResolveAll();
        }
        finally
        {
            Q149Com.Release(recipients);
        }
    }

    /// <summary>The sending account by putref, as the product sets it - the account whose address is the hub's.</summary>
    private static void SetSendUsingHub(dynamic ns, dynamic mail, string hubAddress)
    {
        dynamic accounts = ns.Accounts;
        try
        {
            int count = accounts.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic account = accounts[i];
                try
                {
                    if (string.Equals((string?)account.SmtpAddress, hubAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        object target = mail;
                        target.GetType().InvokeMember(
                            "SendUsingAccount",
                            System.Reflection.BindingFlags.PutRefDispProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                            null,
                            target,
                            new object[] { account },
                            CultureInfo.InvariantCulture);
                        return;
                    }
                }
                finally
                {
                    Q149Com.Release(account);
                }
            }
        }
        finally
        {
            Q149Com.Release(accounts);
        }
    }
}

/// <summary>What one read of <c>Inspector.WordEditor</c> gave: the document, or the failure with its HRESULT.</summary>
internal sealed record Q149EditorRead(object? Document, long Ms, int Polls, Dictionary<string, object>? Failure)
{
    public bool Ok => Document != null;
}

/// <summary>The inspector and Word-editor steps the cells are made of. Every one runs on a probe client's STA.</summary>
internal static class Q149Editor
{
    /// <summary>
    /// Reads <c>WordEditor</c> until it answers or <paramref name="budgetMs"/> has passed, a pumped
    /// <paramref name="intervalMs"/> apart; one read when the budget is 0. Never a second GetInspector -
    /// asking for one after releasing the first wedged Outlook (D46).
    /// </summary>
    internal static Q149EditorRead ReadEditor(object inspectorObject, int budgetMs, int intervalMs)
    {
        dynamic inspector = inspectorObject;
        Stopwatch elapsed = Stopwatch.StartNew();
        int polls = 0;
        Dictionary<string, object>? failure = null;
        while (true)
        {
            polls++;
            try
            {
                object? document = inspector.WordEditor;
                if (document != null)
                {
                    return new Q149EditorRead(document, elapsed.ElapsedMilliseconds, polls, null);
                }

                failure = new Dictionary<string, object>(StringComparer.Ordinal) { ["type"] = "null" };
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                failure = Q149Com.Failure(ex);
            }

            if (elapsed.ElapsedMilliseconds + intervalMs > budgetMs)
            {
                return new Q149EditorRead(null, elapsed.ElapsedMilliseconds, polls, failure);
            }

            PumpedStaRunner.PumpedWait(intervalMs);
        }
    }

    /// <summary>IsWordMail() and EditorType, each null when Outlook would not say.</summary>
    internal static (bool? IsWordMail, int? EditorType) DescribeInspector(object inspectorObject)
    {
        dynamic inspector = inspectorObject;
        bool? isWordMail = null;
        int? editorType = null;
        try
        {
            isWordMail = (bool)inspector.IsWordMail();
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
        }

        try
        {
            editorType = (int)inspector.EditorType;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
        }

        return (isWordMail, editorType);
    }

    /// <summary>
    /// B3's commit, first half: <c>ModifiedFormPages.Add("Message")</c> on an inspector never activated, before
    /// its editor is read; the second half is <c>Close(olSave)</c> after the edit. Records <c>formPageAdded</c>
    /// or <c>formPageFailure</c> on <paramref name="attempt"/>.
    /// </summary>
    internal static void AddMessageFormPage(object inspectorObject, Q149Attempt attempt)
    {
        dynamic inspector = inspectorObject;
        dynamic? pages = null;
        try
        {
            pages = inspector.ModifiedFormPages;
            object page = pages.Add("Message");
            Q149Com.Release(page);
            attempt.Set("formPageAdded", true);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            attempt.Set("formPageFailure", Q149Com.Failure(ex));
        }
        finally
        {
            Q149Com.Release(pages);
        }
    }

    /// <summary>Replaces the whole text of a mail document with <paramref name="text"/> - the edit every commit cell makes.</summary>
    internal static void ReplaceText(object document, string text)
    {
        dynamic doc = document;
        dynamic content = doc.Content;
        dynamic? range = null;
        try
        {
            int end = (int)content.End;
            range = doc.Range(0, Math.Max(0, end - 1));
            range.Text = text;
        }
        finally
        {
            Q149Com.Release(range);
            Q149Com.Release(content);
        }
    }

    /// <summary>
    /// How Word lays the document out, paragraph by paragraph (at most <paramref name="cap"/>) - what a person
    /// sees when Outlook opens the item. Text is normalized; only its length and hash ever leave the test.
    /// </summary>
    internal static List<WordParagraphLook> ReadLook(object document, int cap = 300)
    {
        dynamic doc = document;
        List<WordParagraphLook> look = new();
        dynamic paragraphs = doc.Paragraphs;
        try
        {
            int count = Math.Min((int)paragraphs.Count, cap);
            for (int i = 1; i <= count; i++)
            {
                dynamic paragraph = paragraphs[i];
                dynamic? range = null;
                dynamic? font = null;
                dynamic? shapes = null;
                try
                {
                    range = paragraph.Range;
                    font = range.Font;
                    shapes = range.InlineShapes;
                    int shapeCount = (int)shapes.Count;
                    double width = 0;
                    double height = 0;
                    if (shapeCount > 0)
                    {
                        dynamic first = shapes[1];
                        try
                        {
                            width = Convert.ToDouble(first.Width, CultureInfo.InvariantCulture);
                            height = Convert.ToDouble(first.Height, CultureInfo.InvariantCulture);
                        }
                        finally
                        {
                            Q149Com.Release(first);
                        }
                    }

                    look.Add(new WordParagraphLook(
                        Q149WordLook.NormalizeText((string?)range.Text),
                        TopBorder(paragraph),
                        StyleName(paragraph),
                        (string?)font.Name ?? string.Empty,
                        Convert.ToDouble(font.Size, CultureInfo.InvariantCulture),
                        Convert.ToInt32(font.Bold, CultureInfo.InvariantCulture),
                        Convert.ToInt32(font.Italic, CultureInfo.InvariantCulture),
                        Convert.ToInt32(font.Color, CultureInfo.InvariantCulture),
                        Convert.ToInt32(paragraph.Alignment, CultureInfo.InvariantCulture),
                        Convert.ToDouble(paragraph.SpaceBefore, CultureInfo.InvariantCulture),
                        Convert.ToDouble(paragraph.SpaceAfter, CultureInfo.InvariantCulture),
                        Convert.ToDouble(paragraph.LineSpacing, CultureInfo.InvariantCulture),
                        Convert.ToDouble(paragraph.LeftIndent, CultureInfo.InvariantCulture),
                        shapeCount,
                        width,
                        height));
                }
                finally
                {
                    Q149Com.Release(shapes);
                    Q149Com.Release(font);
                    Q149Com.Release(range);
                    Q149Com.Release(paragraph);
                }
            }
        }
        finally
        {
            Q149Com.Release(paragraphs);
        }

        return look;
    }

    /// <summary>
    /// Opens one of this run's items by EntryID in an inspector, has the product's own invisible promotion
    /// hand over the editor, reads how Word lays it out, and closes the inspector discarding - nothing is
    /// written. Null with the reason when the editor could not be had.
    /// </summary>
    internal static (List<WordParagraphLook>? Look, string? Failure) ReadLookOf(Q149Run run, Q149OutlookClient client, string entryId, string storeId, bool readLayout = true)
    {
        return client.Run<(List<WordParagraphLook>?, string?)>((app, ns) =>
        {
            dynamic? item = null;
            object? inspector = null;
            object? document = null;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                Q149Drafts.Refuse(Q149ProbeGuards.RefuseSubject((string?)item.Subject, run.Marker));
                inspector = item.GetInspector;
                document = ComposeSurface.PromoteForWordEditor(inspector!, out string? error);
                if (document == null)
                {
                    return (null, error ?? "NoWordEditor");
                }

                // The hunt's layout ablation (HuntAblation.Layout): the item opened and promoted as ever, its paragraphs never walked.
                return readLayout ? (ReadLook(document), null) : (null, "layout not read (ablation)");
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return (null, OutlookComSession.DescribeComFailure(ex));
            }
            finally
            {
                Q149Com.Release(document);
                if (inspector != null)
                {
                    try
                    {
                        ((dynamic)inspector).Close(LiveComposeProbe.OlDiscard);
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }
                }

                Q149Com.Release(inspector);
                Q149Com.Release(item);
            }
        });
    }

    /// <summary>The paragraph's top border line style (wdBorderTop), 0 for none - a reply's separator line is one.</summary>
    private static int TopBorder(dynamic paragraph)
    {
        dynamic? borders = null;
        dynamic? top = null;
        try
        {
            borders = paragraph.Borders;
            top = borders.Item(-1);
            return Convert.ToInt32(top.LineStyle, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return -1;
        }
        finally
        {
            Q149Com.Release(top);
            Q149Com.Release(borders);
        }
    }

    private static string StyleName(dynamic paragraph)
    {
        try
        {
            object style = paragraph.Style;
            if (style is string name)
            {
                return name;
            }

            try
            {
                return (string?)((dynamic)style).NameLocal ?? string.Empty;
            }
            finally
            {
                Q149Com.Release(style);
            }
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return "?";
        }
    }
}
