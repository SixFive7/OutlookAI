using System.Globalization;
using System.Text;

namespace OutlookAI.RemediationTools;

/// <summary>
/// Which curated fixture population a plan describes. Absent (null on
/// <see cref="CorpusPlanOptions.Population"/>) means the measurement corpus, whose shape is a
/// weighted mixture and is left exactly as it was.
/// </summary>
public enum CorpusPopulationKind
{
    /// <summary>
    /// The test hub: the POP3 account's delivery store, which the live tier writes to. Small
    /// on purpose - every hub test that pages, caps or walks it was written against a "tiny
    /// test-hub store" of a few dozen items - and carrying the shapes the index tier reads off
    /// the FIRST indexed store: senders, attachments of the kinds the old kind filter dropped,
    /// unread mail, multi-member conversations, and the subject-only probe population.
    /// </summary>
    Hub = 1,

    /// <summary>
    /// The count tripwire's bystander: a few hundred items, below the census identity budget
    /// in every folder, with a mail folder that has populated subfolders.
    /// </summary>
    Bystander = 2,

    /// <summary>
    /// The identity account's own delivery store: a handful of addressed items, so the index
    /// knows the store exists and <c>outlook_health</c> does not report it missing.
    /// </summary>
    Identity = 3,
}

/// <summary>
/// The person a population's mailbox belongs to. Received mail is addressed TO the owner and
/// sent mail is FROM the owner, which is what makes a small store discoverable in the search
/// index at all: the targeted discovery looks for mail whose recipient address or sender
/// contains the store's display name (<c>IndexSearchService.TryDiscoverStoreScopeByAddress</c>).
/// </summary>
/// <param name="Name">Display name the owner's messages carry.</param>
/// <param name="Address">SMTP address. Always under the RFC 2606 <c>.invalid</c> top-level domain.</param>
public sealed record CorpusMailboxOwner(string Name, string Address)
{
    /// <summary>The domain an owner gets when the store is not named as an address.</summary>
    public const string FallbackDomain = "population.invalid";

    /// <summary>
    /// The owner of the store called <paramref name="storeDisplayName"/>. A store named as an
    /// address - the hub and the identity store must be, and every other small indexed store
    /// should be - IS its owner's address. Any other name becomes the owner's display name, with
    /// an address derived from it under <see cref="FallbackDomain"/>, so the sender-name half of
    /// the targeted discovery can still find the store.
    /// </summary>
    public static CorpusMailboxOwner ForStore(string storeDisplayName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(storeDisplayName);
        string name = storeDisplayName.Trim();
        if (IsAddress(name))
        {
            return new CorpusMailboxOwner(name, name);
        }

        var slug = new StringBuilder();
        foreach (char c in name.ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(c))
            {
                slug.Append(c);
            }
            else if (slug.Length > 0 && slug[^1] != '.')
            {
                slug.Append('.');
            }
        }

        string local = slug.ToString().Trim('.');
        if (local.Length == 0)
        {
            throw new ArgumentException(
                "A store name with no ASCII letter or digit in it gives a population owner no address.",
                nameof(storeDisplayName));
        }

        return new CorpusMailboxOwner(name, local + "@" + FallbackDomain);
    }

    /// <summary>Whether <paramref name="text"/> is shaped like an SMTP address: one '@', a dot after it, no space.</summary>
    public static bool IsAddress(string? text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Contains(' ', StringComparison.Ordinal))
        {
            return false;
        }

        int at = text.IndexOf('@', StringComparison.Ordinal);
        return at > 0
            && at == text.LastIndexOf('@')
            && text.IndexOf('.', at) > at + 1
            && !text.EndsWith('.');
    }
}

/// <summary>A person in a population's synthetic directory. Every address is under <c>.invalid</c>.</summary>
/// <param name="Name">Display name.</param>
/// <param name="Address">SMTP address.</param>
public sealed record CorpusCorrespondent(string Name, string Address)
{
    /// <summary>The <c>Name &lt;address&gt;</c> form Outlook resolves to a one-off recipient.</summary>
    public string ToAddressSpec() => Name + " <" + Address + ">";
}

/// <summary>Which recipient row an address goes in.</summary>
public enum CorpusRecipientKind
{
    /// <summary>olTo.</summary>
    To = 1,

    /// <summary>olCC.</summary>
    Cc = 2,
}

/// <summary>One recipient of a population item.</summary>
/// <param name="Person">Who.</param>
/// <param name="Kind">To or Cc.</param>
public sealed record CorpusRecipient(CorpusCorrespondent Person, CorpusRecipientKind Kind);

/// <summary>
/// The attachment kinds a population carries, chosen from what the live tier measures: the old
/// <c>System.Kind IN ('email','document')</c> predicate dropped images, embedded messages and
/// <c>.ics</c> invites, and a plain text file is the document kind it always kept.
/// </summary>
public enum CorpusAttachmentKind
{
    /// <summary>A PNG image. <c>System.Kind</c> picture - a kind the old filter dropped.</summary>
    Png = 1,

    /// <summary>An iCalendar invite. <c>System.Kind</c> calendar - a kind the old filter dropped.</summary>
    Calendar = 2,

    /// <summary>An RFC 822 message saved as <c>.eml</c>, the way a forwarded message is attached.</summary>
    Message = 3,

    /// <summary>A plain-text file - the document kind, and the one whose text the index reads.</summary>
    Text = 4,
}

/// <summary>
/// One attachment, content included. A pure function of the plan: the same seed and ordinal
/// always produce the same bytes, which is what lets a rebuild reproduce a population exactly.
/// </summary>
public sealed class CorpusAttachment : IEquatable<CorpusAttachment>
{
    private readonly byte[] _content;

    /// <summary>Creates an attachment description. The content is copied.</summary>
    public CorpusAttachment(string fileName, CorpusAttachmentKind kind, byte[] content)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        ArgumentNullException.ThrowIfNull(content);
        FileName = fileName;
        Kind = kind;
        _content = (byte[])content.Clone();
    }

    /// <summary>The file name Outlook shows and the index records.</summary>
    public string FileName { get; }

    /// <summary>What kind of file it is.</summary>
    public CorpusAttachmentKind Kind { get; }

    /// <summary>The exact bytes. A copy - the description is immutable.</summary>
    public byte[] Content => (byte[])_content.Clone();

    /// <summary>How many bytes the file holds.</summary>
    public int Length => _content.Length;

    /// <summary>The content as text, for the kinds that are text. Null for the image.</summary>
    public string? Text => Kind == CorpusAttachmentKind.Png ? null : Encoding.ASCII.GetString(_content);

    /// <inheritdoc/>
    public bool Equals(CorpusAttachment? other)
        => other != null
            && string.Equals(FileName, other.FileName, StringComparison.Ordinal)
            && Kind == other.Kind
            && _content.AsSpan().SequenceEqual(other._content);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => Equals(obj as CorpusAttachment);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(FileName, Kind, _content.Length);
}

/// <summary>
/// Everything a population item carries beyond the measurement corpus's subject, body, dates and
/// read state. Null for a measurement-corpus item.
/// </summary>
public sealed class CorpusItemEnrichment : IEquatable<CorpusItemEnrichment>
{
    private readonly byte[]? _conversationIndex;

    /// <summary>Creates an enrichment. Lists and the conversation index are copied.</summary>
    public CorpusItemEnrichment(
        CorpusCorrespondent sender,
        IReadOnlyList<CorpusRecipient> recipients,
        IReadOnlyList<CorpusAttachment> attachments,
        int? threadKey,
        string? conversationTopic,
        byte[]? conversationIndex)
    {
        ArgumentNullException.ThrowIfNull(sender);
        ArgumentNullException.ThrowIfNull(recipients);
        ArgumentNullException.ThrowIfNull(attachments);
        Sender = sender;
        Recipients = recipients.ToList();
        Attachments = attachments.ToList();
        ThreadKey = threadKey;
        ConversationTopic = conversationTopic;
        _conversationIndex = conversationIndex == null ? null : (byte[])conversationIndex.Clone();
    }

    /// <summary>Who the item is from.</summary>
    public CorpusCorrespondent Sender { get; }

    /// <summary>To and Cc rows, in the order they are added.</summary>
    public IReadOnlyList<CorpusRecipient> Recipients { get; }

    /// <summary>Attachments, in the order they are added.</summary>
    public IReadOnlyList<CorpusAttachment> Attachments { get; }

    /// <summary>Which conversation the item belongs to, or null for a single message.</summary>
    public int? ThreadKey { get; }

    /// <summary><c>PR_CONVERSATION_TOPIC</c> for a conversation member, else null.</summary>
    public string? ConversationTopic { get; }

    /// <summary><c>PR_CONVERSATION_INDEX</c> for a conversation member, else null. A copy.</summary>
    public byte[]? ConversationIndex => _conversationIndex == null ? null : (byte[])_conversationIndex.Clone();

    /// <summary>The conversation index as upper-case hex, or null.</summary>
    public string? ConversationIndexHex => _conversationIndex == null ? null : Convert.ToHexString(_conversationIndex);

    /// <summary>
    /// A canonical text rendering, used for equality and for the population digest a T1 test
    /// pins. Attachment content is rendered as its SHA-256, so the digest covers every byte.
    /// </summary>
    public string Render()
    {
        var sb = new StringBuilder();
        sb.Append("from=").Append(Sender.Name).Append('<').Append(Sender.Address).Append('>');
        foreach (CorpusRecipient r in Recipients)
        {
            sb.Append("|").Append(r.Kind == CorpusRecipientKind.To ? "to=" : "cc=")
                .Append(r.Person.Name).Append('<').Append(r.Person.Address).Append('>');
        }

        foreach (CorpusAttachment a in Attachments)
        {
            sb.Append("|att=").Append(a.FileName).Append(':').Append(a.Kind).Append(':')
                .Append(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(a.Content)));
        }

        if (ThreadKey != null)
        {
            sb.Append("|thread=").Append(ThreadKey.Value.ToString(CultureInfo.InvariantCulture))
                .Append("|topic=").Append(ConversationTopic)
                .Append("|index=").Append(ConversationIndexHex);
        }

        return sb.ToString();
    }

    /// <inheritdoc/>
    public bool Equals(CorpusItemEnrichment? other)
        => other != null && string.Equals(Render(), other.Render(), StringComparison.Ordinal);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => Equals(obj as CorpusItemEnrichment);

    /// <inheritdoc/>
    public override int GetHashCode() => Render().GetHashCode(StringComparison.Ordinal);
}

/// <summary>
/// A folder a population creates under one of the store's default folders. Its id is SYNTHETIC -
/// at or above <see cref="CorpusPopulation.SubfolderIdBase"/>, far above every Outlook
/// default-folder id - so the manifest, the scan and the census can carry it exactly as they carry
/// a default folder.
/// </summary>
/// <param name="FolderId">The synthetic id.</param>
/// <param name="ParentFolderId">The Outlook default-folder id the folder is created under.</param>
/// <param name="Name">The folder name. Always starts with <see cref="CorpusManifest.CreatedFolderPrefix"/>, which is half of what teardown checks before removing it.</param>
/// <param name="ParentName">The parent's English default name, for reports and the store-relative path.</param>
public sealed record CorpusPopulationFolder(int FolderId, int ParentFolderId, string Name, string ParentName)
{
    /// <summary>The store-relative path, as the live-test settings and the search tool spell it.</summary>
    public string Path => ParentName + "/" + Name;
}

/// <summary>
/// Where the SF-6 subject-only probe population lives in the hub population - the four
/// coordinates the live-test settings' <c>subjectOnlyProbe</c> block names, three of them fixed by
/// the plan and the fourth being the hub store's own name.
/// </summary>
/// <param name="FolderId">The synthetic folder id holding the population.</param>
/// <param name="FolderPath">Store-relative folder path.</param>
/// <param name="SubjectTerm">A term in every member's subject and in no body.</param>
/// <param name="SenderFragment">A fragment of the one sender every member - and nothing else in that folder - comes from.</param>
public sealed record CorpusSubjectOnlyProbe(int FolderId, string FolderPath, string SubjectTerm, string SenderFragment);

/// <summary>
/// A CURATED fixture population: the small, tagged mini-corpora the generator builds into the
/// live tier's hub, bystander and identity stores so the tests that read those stores have
/// something to read. Decided 2026-09-24 (Q70): the hub PST the design used to leave empty, and
/// the bystander a person used to have to populate by hand, are both generated.
/// <para>
/// <b>Why a curated layout and not the corpus's weighted mixture.</b> The measurement corpus is a
/// DISTRIBUTION - sizes, ages and folders drawn by weight - because what it exists for is volume.
/// A fixture population exists for the opposite reason: the live tests that read it make exact
/// demands of it - at most 99 searchable hub items, a folder whose every member carries one term in
/// its subject and none in its body, a conversation the newest item belongs to, an attachment of
/// each kind the old kind filter dropped. A distribution satisfies those only on average. So the
/// STRUCTURE here is fixed per kind (which ordinal is threaded, which carries which attachment,
/// which folder holds what) and the SEED decides only the text, the correspondents and the ages
/// inside their bands. T1's <c>CorpusPopulationTests</c> pins every demand against the plan itself.
/// </para>
/// <para>
/// <b>Vocabularies are disjoint by construction.</b> Subject words never occur in a body, body
/// words never occur in a subject, and attachment and file-name words occur in neither - so a term
/// the completeness oracle derives from a subject or body can only ever match an item whose own
/// plain text carries it. The one deliberate overlap is <see cref="ProbeTerm"/>, which is a body
/// word AND the text of one attachment, and whose parent's body always carries it too.
/// </para>
/// <para>
/// Everything else is the measurement corpus's machinery unchanged: the subject tags, the ordinal
/// parse every delete and rewrite rests on, the manifest, the store and profile guards, the probes
/// and the census. A population is built, censused and torn down with the same commands, from the
/// account-less profile, and nothing here is reachable without <c>--population</c>.
/// </para>
/// </summary>
public sealed class CorpusPopulation
{
    /// <summary>
    /// The population format version, folded into the shape key. Bump it whenever a change here
    /// would make a rebuild produce a different population from the same seed, so an old manifest
    /// is refused rather than extended with items of a different shape.
    /// </summary>
    public const int Version = 1;

    /// <summary>The lowest synthetic folder id. Outlook's default-folder ids stop in the low forties.</summary>
    public const int SubfolderIdBase = 1000;

    /// <summary>
    /// The probe term the live tier's <c>probeTerm</c> setting names on a test guest: a word that is
    /// in the body vocabulary here, in the measurement corpus's vocabulary, and in the text of one hub
    /// attachment - so it hits every indexed store and produces an attachment hit. A DECISION recorded
    /// in <c>Testbed/testbed.json</c>, and T1 holds the two equal.
    /// </summary>
    public const string ProbeTerm = "invoice";

    /// <summary>The subject term of the hub's SF-6 probe population.</summary>
    public const string SubjectOnlyTerm = "bulletin";

    /// <summary>The sender fragment of the hub's SF-6 probe population: in the sender's name AND address.</summary>
    public const string SubjectOnlySenderFragment = "noticebot";

    /// <summary>
    /// The only sender of the SF-6 probe population, and of nothing else. The fragment is a whole
    /// word of the display name as well as the address's local part, because the index's sender
    /// predicate is a word CONTAINS over both columns and how the word breaker splits an address is
    /// not something to lean on.
    /// </summary>
    public static CorpusCorrespondent NoticeSender { get; } = new("Noticebot Relay", "noticebot@alerts.invalid");

    /// <summary>
    /// Words that only ever appear in SUBJECTS. None is a substring of any body word or filler, and
    /// none is a substring of anything an attachment or ICS/EML header carries - "summary" is not
    /// here for exactly that reason: every <c>.ics</c> says SUMMARY.
    /// </summary>
    private static readonly string[] SubjectWords =
    {
        "quarterly", "roadmap", "kickoff", "milestone", "workshop", "briefing", "agenda", "proposal",
        "timeline", "checklist", "handbook", "keynote", "offsite", "retrospective", "headcount",
        "newsletter", "townhall", "scorecard", "benchmark", "playbook", "blueprint", "showcase", "forecast",
    };

    /// <summary>Words that only ever appear in BODIES. Contains <see cref="ProbeTerm"/>.</summary>
    private static readonly string[] BodyWords =
    {
        "invoice", "delivery", "shipment", "warehouse", "pallet", "freight", "customs", "payment",
        "balance", "ledger", "receipt", "credit", "refund", "discount", "pricing", "vendor",
        "supplier", "contract", "clause", "deadline", "meeting", "printer", "network", "server",
        "backup", "license", "renewal", "support", "ticket", "priority", "escalation", "hardware",
        "laptop", "monitor", "keyboard", "parking", "catering", "travel", "hotel", "flight",
    };

    /// <summary>Short connecting words bodies use between body words. Held to the same substring rules.</summary>
    private static readonly string[] BodyFiller = { "please", "confirm", "regarding", "update", "thanks", "the", "and" };

    /// <summary>Words that only ever appear inside attachments.</summary>
    private static readonly string[] AttachmentWords =
    {
        "turbine", "gasket", "valve", "piston", "bearing", "coupling", "flange", "impeller",
        "sensor", "actuator", "bracket", "spindle", "gearbox", "hydraulic", "pneumatic", "sprocket",
    };

    /// <summary>The fixed first word of each attachment kind's file name.</summary>
    private static readonly string[] FileNameWords = { "diagram", "invite", "forwarded", "notes" };

    /// <summary>The words the SF-6 probe subjects add around <see cref="SubjectOnlyTerm"/>.</summary>
    private static readonly string[] NoticeSubjectWords = { "weekly" };

    /// <summary>The synthetic directory every correspondent is drawn from.</summary>
    private static readonly CorpusCorrespondent[] Directory =
    {
        new("Alder Finch", "alder.finch@northwind.invalid"),
        new("Briony Hale", "briony.hale@contoso.invalid"),
        new("Cedric Oakes", "cedric.oakes@fabrikam.invalid"),
        new("Dalia Moss", "dalia.moss@litware.invalid"),
        new("Emrys Vane", "emrys.vane@adatum.invalid"),
        new("Freya Holt", "freya.holt@tailspin.invalid"),
        new("Gideon Pike", "gideon.pike@wingtip.invalid"),
        new("Hazel Crane", "hazel.crane@proseware.invalid"),
        new("Ivor Lark", "ivor.lark@lucerne.invalid"),
        new("Juno Frost", "juno.frost@alpine.invalid"),
        new("Kester Wren", "kester.wren@margie.invalid"),
        new("Linnea Ash", "linnea.ash@fourthcoffee.invalid"),
    };

    // Independent hash streams, far away from CorpusPlan's own so neither can shift the other.
    private const int StreamCorrespondent = 200_001;
    private const int StreamCc = 200_002;
    private const int StreamCcWho = 200_003;
    private const int StreamAge = 200_004;
    private const int StreamTransport = 200_005;
    private const int StreamSubject = 210_000;
    private const int StreamBody = 220_000;
    private const int StreamAttachment = 230_000;
    private const int StreamThread = 240_000;

    private readonly CorpusPlanOptions _options;
    private readonly Slot[] _slots;
    private readonly Dictionary<int, CorpusPopulationFolder> _folders;

    private CorpusPopulation(CorpusPlanOptions options, CorpusPopulationKind kind, CorpusMailboxOwner owner, Slot[] slots, IReadOnlyList<CorpusPopulationFolder> folders)
    {
        _options = options;
        Kind = kind;
        Owner = owner;
        _slots = slots;
        _folders = folders.ToDictionary(f => f.FolderId);
        Folders = folders;
        SubjectOnlyProbe = kind == CorpusPopulationKind.Hub
            ? new CorpusSubjectOnlyProbe(
                HubNoticesFolderId,
                _folders[HubNoticesFolderId].Path,
                SubjectOnlyTerm,
                SubjectOnlySenderFragment)
            : null;
    }

    /// <summary>Which population this is.</summary>
    public CorpusPopulationKind Kind { get; }

    /// <summary>The mailbox owner items are addressed to and sent from.</summary>
    public CorpusMailboxOwner Owner { get; }

    /// <summary>How many items the population holds. Fixed by the kind - it is not a parameter.</summary>
    public int ItemCount => _slots.Length;

    /// <summary>The folders the population creates, in the order a build first needs them.</summary>
    public IReadOnlyList<CorpusPopulationFolder> Folders { get; }

    /// <summary>The SF-6 probe coordinates, for the hub; null for every other kind.</summary>
    public CorpusSubjectOnlyProbe? SubjectOnlyProbe { get; }

    /// <summary>The words subjects are built from. Exposed for the invariants T1 holds.</summary>
    public static IReadOnlyList<string> SubjectVocabulary => SubjectWords;

    /// <summary>The words bodies are built from.</summary>
    public static IReadOnlyList<string> BodyVocabulary => BodyWords;

    /// <summary>The connecting words bodies use.</summary>
    public static IReadOnlyList<string> BodyFillerWords => BodyFiller;

    /// <summary>The words attachments are built from.</summary>
    public static IReadOnlyList<string> AttachmentVocabulary => AttachmentWords;

    /// <summary>The first word of every attachment's file name.</summary>
    public static IReadOnlyList<string> AttachmentFileNameWords => FileNameWords;

    /// <summary>The words the SF-6 subjects carry besides the subject term.</summary>
    public static IReadOnlyList<string> NoticeWords => NoticeSubjectWords;

    /// <summary>Everybody a population can address.</summary>
    public static IReadOnlyList<CorpusCorrespondent> Correspondents => Directory;

    /// <summary>The shape-key fragment that makes a population's manifest unmistakable for any other.</summary>
    public string ShapeKeySuffix => ShapeKeySuffixFor(Kind, Owner);

    /// <summary>
    /// The shape-key fragment for <paramref name="kind"/> owned by <paramref name="owner"/>. The owner
    /// is part of the shape because it is part of the CONTENT - every item is addressed to or sent
    /// from it - so the same seed built into two differently named stores is two different populations.
    /// </summary>
    public static string ShapeKeySuffixFor(CorpusPopulationKind kind, CorpusMailboxOwner owner)
    {
        ArgumentNullException.ThrowIfNull(owner);
        return "|p:" + kind.ToString().ToLowerInvariant() + ":v" + Version.ToString(CultureInfo.InvariantCulture)
            + "|o:" + owner.Name + "<" + owner.Address + ">";
    }

    /// <summary>
    /// Builds the population <paramref name="options"/> asks for. Refuses a corpus id that contains
    /// a subject or body word: the id sits in every subject, and a body word inside it would put
    /// that word in every subject, which is the one thing the disjoint vocabularies promise not to do.
    /// </summary>
    public static CorpusPopulation Create(CorpusPlanOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        CorpusPopulationKind kind = options.Population
            ?? throw new ArgumentException("The options name no population.", nameof(options));
        CorpusMailboxOwner owner = options.Owner
            ?? throw new ArgumentException(
                "A population needs its mailbox owner - pass the store (--store) the population is for.",
                nameof(options));
        if (!CorpusMailboxOwner.IsAddress(owner.Address))
        {
            throw new ArgumentException("The population owner's address is not an address.", nameof(options));
        }

        string id = options.CorpusId.ToLowerInvariant();
        foreach (string word in SubjectWords.Concat(BodyWords).Concat(BodyFiller.Where(w => w.Length >= 4)))
        {
            if (id.Contains(word, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    $"Corpus id '{options.CorpusId}' contains the population word '{word}'. The id is in every subject, "
                    + "so that word would stop being confined to one side of the subject/body split the live tests rely on. "
                    + "Choose an id without it.",
                    nameof(options));
            }
        }

        (Slot[] slots, IReadOnlyList<CorpusPopulationFolder> folders) = kind switch
        {
            CorpusPopulationKind.Hub => HubLayout(),
            CorpusPopulationKind.Bystander => BystanderLayout(),
            CorpusPopulationKind.Identity => IdentityLayout(),
            _ => throw new ArgumentOutOfRangeException(nameof(options), "Unknown population kind."),
        };

        return new CorpusPopulation(options, kind, owner, slots, folders);
    }

    /// <summary>Parses a population name as the command line spells it. Case is ignored.</summary>
    public static bool TryParseKind(string? text, out CorpusPopulationKind kind)
    {
        kind = default;
        switch ((text ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "hub":
                kind = CorpusPopulationKind.Hub;
                return true;
            case "bystander":
                kind = CorpusPopulationKind.Bystander;
                return true;
            case "identity":
                kind = CorpusPopulationKind.Identity;
                return true;
            default:
                return false;
        }
    }

    /// <summary>The folder a synthetic id stands for, or null when the id is not one of this population's.</summary>
    public CorpusPopulationFolder? FolderOf(int folderId) => _folders.TryGetValue(folderId, out CorpusPopulationFolder? f) ? f : null;

    /// <summary>
    /// The synthetic id of the folder called <paramref name="name"/> under default folder
    /// <paramref name="parentFolderId"/>, or null. What a scan with no manifest uses to put an item it
    /// finds in a created folder back where the plan counts it.
    /// </summary>
    public int? FolderIdOf(int parentFolderId, string name)
    {
        foreach (CorpusPopulationFolder folder in Folders)
        {
            if (folder.ParentFolderId == parentFolderId && string.Equals(folder.Name, name, StringComparison.Ordinal))
            {
                return folder.FolderId;
            }
        }

        return null;
    }

    /// <summary>What item <paramref name="ordinal"/> is. Depends only on the seed, the anchor, the owner and the kind.</summary>
    public CorpusItemSpec Describe(int ordinal)
    {
        Slot slot = SlotOf(ordinal);
        string body = BuildBody(ordinal);
        long ageSeconds = AgeSecondsOf(ordinal, slot);
        DateTime receivedUtc = DateTime.SpecifyKind(_options.AnchorUtc, DateTimeKind.Utc).AddSeconds(-ageSeconds);
        bool outbound = slot.Role is SlotRole.Outbound;
        int transport = outbound ? 0 : 30 + (int)(Draw(ordinal, StreamTransport) % 570UL);
        DateTime sentUtc = receivedUtc.AddSeconds(-transport);
        return new CorpusItemSpec(
            ordinal,
            slot.FolderId,
            BuildSubject(ordinal),
            body.Length,
            receivedUtc,
            sentUtc,
            !slot.Unread,
            "population-" + Kind.ToString().ToLowerInvariant(),
            slot.Segment);
    }

    /// <summary>
    /// The subject: both corpus tags first, exactly as the measurement corpus writes them, so every
    /// delete, rewrite and scan predicate the corpus has applies unchanged. A reply carries "RE: "
    /// after the tags. The SF-6 members carry <see cref="SubjectOnlyTerm"/>.
    /// </summary>
    public string BuildSubject(int ordinal)
    {
        Slot slot = SlotOf(ordinal);
        var sb = new StringBuilder(96);
        sb.Append(CorpusPlan.SubjectTag).Append(CorpusPlan.CorpusTagOpen).Append(_options.CorpusId).Append('#')
            .Append(ordinal.ToString("D7", CultureInfo.InvariantCulture)).Append("] ");

        if (slot.Thread != null)
        {
            if (slot.Thread.Position > 0)
            {
                sb.Append("RE: ");
            }

            sb.Append(ThreadTopic(slot.Thread.Index));
            return sb.ToString();
        }

        if (slot.Role == SlotRole.Notice)
        {
            sb.Append(char.ToUpperInvariant(NoticeSubjectWords[0][0])).Append(NoticeSubjectWords[0], 1, NoticeSubjectWords[0].Length - 1)
                .Append(' ').Append(SubjectOnlyTerm).Append(' ')
                .Append((slot.Index + 1).ToString(CultureInfo.InvariantCulture)).Append(' ');
            AppendWords(sb, SubjectWords, ordinal, StreamSubject, 2);
            return sb.ToString();
        }

        int words = 3 + (int)(Draw(ordinal, StreamSubject) % 3UL);
        AppendWords(sb, SubjectWords, ordinal, StreamSubject + 1, words);
        return sb.ToString();
    }

    /// <summary>
    /// The plain-text body: sentences of body words joined by filler, never quoting a subject - a
    /// quoted subject would put subject words in a body. Items flagged to carry the probe term carry
    /// it in the first sentence.
    /// </summary>
    public string BuildBody(int ordinal)
    {
        Slot slot = SlotOf(ordinal);
        int target = slot.Role == SlotRole.Notice
            ? 160 + (int)(Draw(ordinal, StreamBody) % 240UL)
            : 220 + (int)(Draw(ordinal, StreamBody) % 1180UL);
        var sb = new StringBuilder(target + 96);
        int sentence = 0;
        if (slot.ProbeTerm)
        {
            sb.Append("Please confirm the ").Append(ProbeTerm).Append(" and the ")
                .Append(BodyWords[(int)(Draw(ordinal, StreamBody + 1) % (ulong)BodyWords.Length)]).Append(".\n");
        }

        while (sb.Length < target)
        {
            int words = 5 + (int)(Draw(ordinal, StreamBody + 10 + sentence) % 7UL);
            string opener = BodyFiller[(int)(Draw(ordinal, StreamBody + 200 + sentence) % 5UL)];
            sb.Append(char.ToUpperInvariant(opener[0])).Append(opener, 1, opener.Length - 1).Append(' ');
            for (int w = 0; w < words; w++)
            {
                if (w > 0)
                {
                    sb.Append(w % 3 == 0 ? " and " : " ");
                }

                sb.Append(BodyWords[(int)(Draw(ordinal, StreamBody + 1000 + (sentence * 32) + w) % (ulong)BodyWords.Length)]);
            }

            sb.Append(sentence % 3 == 2 ? ".\n\n" : ".\n");
            sentence++;
        }

        sb.Append("Thanks\n");
        return sb.ToString();
    }

    /// <summary>Everything item <paramref name="ordinal"/> carries beyond the corpus basics.</summary>
    public CorpusItemEnrichment Enrich(int ordinal)
    {
        Slot slot = SlotOf(ordinal);
        CorpusCorrespondent owner = new(Owner.Name, Owner.Address);
        CorpusCorrespondent counterpart = slot.Thread != null
            ? Directory[(int)(DrawThread(slot.Thread.Index, 1) % (ulong)Directory.Length)]
            : Directory[(int)(Draw(ordinal, StreamCorrespondent) % (ulong)Directory.Length)];

        CorpusCorrespondent sender;
        var recipients = new List<CorpusRecipient>();
        switch (slot.Role)
        {
            case SlotRole.Outbound:
                sender = owner;
                recipients.Add(new CorpusRecipient(counterpart, CorpusRecipientKind.To));
                break;
            case SlotRole.Notice:
                sender = NoticeSender;
                recipients.Add(new CorpusRecipient(owner, CorpusRecipientKind.To));
                break;
            default:
                sender = counterpart;
                recipients.Add(new CorpusRecipient(owner, CorpusRecipientKind.To));
                if (slot.Thread == null && Draw(ordinal, StreamCc) % 3UL == 0UL)
                {
                    int senderIndex = Array.IndexOf(Directory, counterpart);
                    int cc = (senderIndex + 1 + (int)(Draw(ordinal, StreamCcWho) % (ulong)(Directory.Length - 1))) % Directory.Length;
                    recipients.Add(new CorpusRecipient(Directory[cc], CorpusRecipientKind.Cc));
                }

                break;
        }

        var attachments = new List<CorpusAttachment>();
        for (int i = 0; i < slot.Attachments.Length; i++)
        {
            attachments.Add(BuildAttachment(ordinal, i, slot.Attachments[i], slot.ProbeTerm, sender, owner));
        }

        return new CorpusItemEnrichment(
            sender,
            recipients,
            attachments,
            slot.Thread?.Index,
            slot.Thread == null ? null : ThreadTopic(slot.Thread.Index),
            slot.Thread == null ? null : ConversationIndexOf(slot.Thread, ordinal));
    }

    /// <summary>A human label for a synthetic folder id, or null when it is not one of this population's.</summary>
    public string? FolderLabel(int folderId) => FolderOf(folderId)?.Path;

    // ------------------------------------------------------------------ layouts

    /// <summary>The synthetic id of the hub's SF-6 folder.</summary>
    private const int HubNoticesFolderId = SubfolderIdBase + 2;

    private const int Inbox = 6;
    private const int SentItems = 5;

    private static (Slot[] Slots, IReadOnlyList<CorpusPopulationFolder> Folders) HubLayout()
    {
        var projects = new CorpusPopulationFolder(SubfolderIdBase + 1, Inbox, CorpusManifest.CreatedFolderPrefix + "-Projects", "Inbox");
        var notices = new CorpusPopulationFolder(HubNoticesFolderId, Inbox, CorpusManifest.CreatedFolderPrefix + "-Notices", "Inbox");
        var slots = new List<Slot>();

        // 1-16: four conversations of four, alternating received and sent. Conversation 0 holds the
        // NEWEST item of the whole store, one minute before the anchor, so a "most recent hit with a
        // conversation id" lands in a conversation that has members to walk - and the index frontier
        // is one minute old on a population built against the moment the tier starts.
        for (int t = 0; t < 4; t++)
        {
            for (int p = 0; p < 4; p++)
            {
                bool received = p % 2 == 0;
                long age = 60L + (t * 216_000L) + ((3 - p) * 21_600L);
                slots.Add(new Slot(
                    received ? SlotRole.Inbound : SlotRole.Outbound,
                    received ? Inbox : SentItems,
                    "thread",
                    slots.Count,
                    new ThreadSlot(t, p),
                    t == 1 && p == 0 ? new[] { CorpusAttachmentKind.Message } : Array.Empty<CorpusAttachmentKind>(),
                    Unread: t == 2 && p == 2,
                    ProbeTerm: false,
                    FixedAgeSeconds: age,
                    AgeFromSeconds: 0,
                    AgeToSeconds: 0));
            }
        }

        // 17-32: sixteen received singles in the Inbox, four unread, eight attachments of every kind -
        // one of them a text file carrying the probe term, and one mail carrying three at once.
        for (int i = 0; i < 16; i++)
        {
            CorpusAttachmentKind[] attachments = i switch
            {
                1 => new[] { CorpusAttachmentKind.Png },
                3 => new[] { CorpusAttachmentKind.Calendar },
                5 => new[] { CorpusAttachmentKind.Message },
                7 => new[] { CorpusAttachmentKind.Text },
                9 => new[] { CorpusAttachmentKind.Png, CorpusAttachmentKind.Calendar, CorpusAttachmentKind.Text },
                11 => new[] { CorpusAttachmentKind.Text },
                _ => Array.Empty<CorpusAttachmentKind>(),
            };
            slots.Add(Single(SlotRole.Inbound, Inbox, "inbox", i, attachments, unread: i % 4 == 0, probeTerm: i == 7, 1, 45));
        }

        // 33-38: six sent singles.
        for (int i = 0; i < 6; i++)
        {
            slots.Add(Single(SlotRole.Outbound, SentItems, "sent", i, Array.Empty<CorpusAttachmentKind>(), unread: false, probeTerm: false, 2, 50));
        }

        // 39-44: six received items one folder down, so the hub has a subtree of its own.
        for (int i = 0; i < 6; i++)
        {
            CorpusAttachmentKind[] attachments = i switch
            {
                1 => new[] { CorpusAttachmentKind.Png },
                3 => new[] { CorpusAttachmentKind.Message },
                _ => Array.Empty<CorpusAttachmentKind>(),
            };
            slots.Add(Single(SlotRole.Inbound, projects.FolderId, "projects", i, attachments, unread: i == 4, probeTerm: false, 3, 60));
        }

        // 45-56: the SF-6 subject-only probe population - one sender, one term in every subject and
        // no body, no attachment, and nothing else in the folder.
        for (int i = 0; i < 12; i++)
        {
            slots.Add(Single(SlotRole.Notice, notices.FolderId, "notices", i, Array.Empty<CorpusAttachmentKind>(), unread: i % 3 == 2, probeTerm: false, 1, 40));
        }

        return (slots.ToArray(), new[] { projects, notices });
    }

    private static (Slot[] Slots, IReadOnlyList<CorpusPopulationFolder> Folders) BystanderLayout()
    {
        var projects = new CorpusPopulationFolder(SubfolderIdBase + 1, Inbox, CorpusManifest.CreatedFolderPrefix + "-Projects", "Inbox");
        var suppliers = new CorpusPopulationFolder(SubfolderIdBase + 2, Inbox, CorpusManifest.CreatedFolderPrefix + "-Suppliers", "Inbox");
        var slots = new List<Slot>();

        // 1-18: six conversations of three.
        for (int t = 0; t < 6; t++)
        {
            for (int p = 0; p < 3; p++)
            {
                bool received = p % 2 == 0;
                long age = 86_400L + (t * 1_728_000L) + ((2 - p) * 86_400L);
                slots.Add(new Slot(
                    received ? SlotRole.Inbound : SlotRole.Outbound,
                    received ? Inbox : SentItems,
                    "thread",
                    slots.Count,
                    new ThreadSlot(t, p),
                    Array.Empty<CorpusAttachmentKind>(),
                    Unread: false,
                    ProbeTerm: false,
                    FixedAgeSeconds: age,
                    AgeFromSeconds: 0,
                    AgeToSeconds: 0));
            }
        }

        // 19-178: 160 received singles. A quarter are recent, the rest spread over two years; one in
        // five unread; one in eight carries an attachment, cycling through the kinds.
        CorpusAttachmentKind[] cycle =
        {
            CorpusAttachmentKind.Png, CorpusAttachmentKind.Calendar, CorpusAttachmentKind.Message, CorpusAttachmentKind.Text,
        };
        for (int i = 0; i < 160; i++)
        {
            CorpusAttachmentKind[] attachments = i % 8 == 3
                ? new[] { cycle[(i / 8) % cycle.Length] }
                : Array.Empty<CorpusAttachmentKind>();
            bool recent = i % 4 == 0;
            slots.Add(Single(SlotRole.Inbound, Inbox, "inbox", i, attachments, unread: i % 5 == 0, probeTerm: false, recent ? 1 : 30, recent ? 30 : 730));
        }

        // 179-228: fifty sent singles.
        for (int i = 0; i < 50; i++)
        {
            slots.Add(Single(SlotRole.Outbound, SentItems, "sent", i, Array.Empty<CorpusAttachmentKind>(), unread: false, probeTerm: false, 1, 730));
        }

        // 229-268 and 269-300: two populated subfolders of the Inbox - the "mail folder with populated
        // children" the exclude-subfolders measurement looks for in the first non-hub indexed store.
        for (int i = 0; i < 40; i++)
        {
            CorpusAttachmentKind[] attachments = i % 10 == 5 ? new[] { CorpusAttachmentKind.Text } : Array.Empty<CorpusAttachmentKind>();
            slots.Add(Single(SlotRole.Inbound, projects.FolderId, "projects", i, attachments, unread: i % 7 == 0, probeTerm: false, 2, 400));
        }

        for (int i = 0; i < 32; i++)
        {
            slots.Add(Single(SlotRole.Inbound, suppliers.FolderId, "suppliers", i, Array.Empty<CorpusAttachmentKind>(), unread: false, probeTerm: false, 5, 600));
        }

        return (slots.ToArray(), new[] { projects, suppliers });
    }

    private static (Slot[] Slots, IReadOnlyList<CorpusPopulationFolder> Folders) IdentityLayout()
    {
        var slots = new List<Slot>();
        for (int i = 0; i < 5; i++)
        {
            slots.Add(Single(SlotRole.Inbound, Inbox, "inbox", i, Array.Empty<CorpusAttachmentKind>(), unread: i == 0, probeTerm: false, 1, 20));
        }

        for (int i = 0; i < 3; i++)
        {
            slots.Add(Single(SlotRole.Outbound, SentItems, "sent", i, Array.Empty<CorpusAttachmentKind>(), unread: false, probeTerm: false, 2, 25));
        }

        return (slots.ToArray(), Array.Empty<CorpusPopulationFolder>());
    }

    private static Slot Single(
        SlotRole role, int folderId, string segment, int index, CorpusAttachmentKind[] attachments,
        bool unread, bool probeTerm, int fromDays, int toDays)
        => new(role, folderId, segment, index, null, attachments, unread, probeTerm, 0, fromDays * 86_400L, toDays * 86_400L);

    // ------------------------------------------------------------------ content

    private Slot SlotOf(int ordinal)
    {
        if (ordinal < 1 || ordinal > _slots.Length)
        {
            throw new ArgumentOutOfRangeException(
                nameof(ordinal), $"The {Kind} population holds ordinals 1..{_slots.Length}; {ordinal} is not one of them.");
        }

        return _slots[ordinal - 1];
    }

    private long AgeSecondsOf(int ordinal, Slot slot)
    {
        if (slot.FixedAgeSeconds > 0)
        {
            return slot.FixedAgeSeconds;
        }

        long span = Math.Max(1L, slot.AgeToSeconds - slot.AgeFromSeconds);
        return slot.AgeFromSeconds + (long)(Draw(ordinal, StreamAge) % (ulong)span);
    }

    private string ThreadTopic(int thread)
    {
        var sb = new StringBuilder();
        int words = 3 + (int)(DrawThread(thread, 2) % 2UL);
        for (int i = 0; i < words; i++)
        {
            if (i > 0)
            {
                sb.Append(' ');
            }

            sb.Append(SubjectWords[(int)(DrawThread(thread, 10 + i) % (ulong)SubjectWords.Length)]);
        }

        return sb.ToString();
    }

    /// <summary>
    /// <c>PR_CONVERSATION_INDEX</c> for a conversation member (MS-OXOMSG 2.2.1.3): a 22-byte header -
    /// a reserved 0x01, the high 40 bits of the thread's first FILETIME, and a 16-byte GUID - then one
    /// 5-byte child block per reply. Every member of a thread shares the header, which is what makes
    /// them one conversation; the GUID comes from the seed, so a rebuild reproduces it.
    /// </summary>
    private byte[] ConversationIndexOf(ThreadSlot thread, int ordinal)
    {
        int first = ordinal - thread.Position;
        DateTime start = Describe(first).SentUtc;
        long startTicks = start.ToFileTimeUtc();
        var bytes = new List<byte>(22 + (5 * thread.Position)) { 0x01 };
        long high = startTicks >> 24;
        for (int shift = 32; shift >= 0; shift -= 8)
        {
            bytes.Add((byte)((high >> shift) & 0xFF));
        }

        ulong a = DrawThread(thread.Index, 100);
        ulong b = DrawThread(thread.Index, 101);
        for (int i = 0; i < 8; i++)
        {
            bytes.Add((byte)(a >> (i * 8)));
        }

        for (int i = 0; i < 8; i++)
        {
            bytes.Add((byte)(b >> (i * 8)));
        }

        for (int p = 1; p <= thread.Position; p++)
        {
            long delta = Math.Max(0L, Describe(first + p).SentUtc.ToFileTimeUtc() - startTicks);
            uint timeDelta = (uint)((delta >> 18) & 0x7FFF_FFFF);
            bytes.Add((byte)(timeDelta >> 24));
            bytes.Add((byte)(timeDelta >> 16));
            bytes.Add((byte)(timeDelta >> 8));
            bytes.Add((byte)timeDelta);
            bytes.Add((byte)(((DrawThread(thread.Index, 200 + p) & 0x0F) << 4) | (uint)(p & 0x0F)));
        }

        return bytes.ToArray();
    }

    private CorpusAttachment BuildAttachment(
        int ordinal, int index, CorpusAttachmentKind kind, bool probeTerm, CorpusCorrespondent sender, CorpusCorrespondent owner)
    {
        int stream = StreamAttachment + (index * 1000);
        string number = ordinal.ToString("D4", CultureInfo.InvariantCulture);
        DateTime when = Describe(ordinal).ReceivedUtc;
        switch (kind)
        {
            case CorpusAttachmentKind.Png:
                return new CorpusAttachment(
                    FileNameWords[0] + "-" + number + ".png",
                    kind,
                    CorpusAttachmentContent.Png(16, 16, (x, y, channel) => (byte)Draw(ordinal, stream + (((y * 16) + x) * 3) + channel)));

            case CorpusAttachmentKind.Calendar:
                return new CorpusAttachment(
                    FileNameWords[1] + "-" + number + ".ics",
                    kind,
                    CorpusAttachmentContent.Calendar(
                        _options.CorpusId + "-" + number + "-" + index.ToString(CultureInfo.InvariantCulture),
                        when.AddDays(2),
                        Words(AttachmentWords, ordinal, stream, 3),
                        Words(AttachmentWords, ordinal, stream + 100, 8)));

            case CorpusAttachmentKind.Message:
                return new CorpusAttachment(
                    FileNameWords[2] + "-" + number + ".eml",
                    kind,
                    CorpusAttachmentContent.Message(
                        Directory[(int)(Draw(ordinal, stream + 1) % (ulong)Directory.Length)],
                        sender.Address == owner.Address ? Directory[(int)(Draw(ordinal, stream + 2) % (ulong)Directory.Length)] : sender,
                        Words(AttachmentWords, ordinal, stream + 10, 3),
                        when.AddDays(-1),
                        Words(AttachmentWords, ordinal, stream + 200, 14)));

            default:
                string text = Words(AttachmentWords, ordinal, stream + 300, 12);
                if (probeTerm)
                {
                    text = ProbeTerm + " " + text;
                }

                return new CorpusAttachment(FileNameWords[3] + "-" + number + ".txt", kind, CorpusAttachmentContent.Text(text));
        }
    }

    private string Words(string[] vocabulary, int ordinal, int stream, int count)
    {
        var sb = new StringBuilder();
        AppendWords(sb, vocabulary, ordinal, stream, count);
        return sb.ToString();
    }

    private void AppendWords(StringBuilder sb, string[] vocabulary, int ordinal, int stream, int count)
    {
        for (int i = 0; i < count; i++)
        {
            if (i > 0)
            {
                sb.Append(' ');
            }

            sb.Append(vocabulary[(int)(Draw(ordinal, stream + i) % (ulong)vocabulary.Length)]);
        }
    }

    private ulong Draw(int ordinal, int stream) => CorpusPlan.Draw(_options.Seed, ordinal, stream);

    /// <summary>Per-THREAD draws: a thread's topic, correspondent and GUID do not depend on which member is asking.</summary>
    private ulong DrawThread(int thread, int stream) => CorpusPlan.Draw(_options.Seed, 1_000_000 + thread, StreamThread + stream);

    // ------------------------------------------------------------------ slot model

    private enum SlotRole
    {
        Inbound,
        Outbound,
        Notice,
    }

    private sealed record ThreadSlot(int Index, int Position);

    private sealed record Slot(
        SlotRole Role,
        int FolderId,
        string Segment,
        int Index,
        ThreadSlot? Thread,
        CorpusAttachmentKind[] Attachments,
        bool Unread,
        bool ProbeTerm,
        long FixedAgeSeconds,
        long AgeFromSeconds,
        long AgeToSeconds);
}

/// <summary>
/// Which folder id an item found in a CREATED folder counts under, when no manifest says - the
/// <c>corpus-reindex</c> case. Pure, so T1 pins the naming rules both halves of the builder use.
/// </summary>
public static class CorpusFolderIds
{
    /// <summary>The Outlook default-folder id of Junk Email, whose stand-in is named "-Junk" rather than by number.</summary>
    public const int JunkFolderId = 23;

    /// <summary>
    /// The name the builder gives the stand-in for default folder <paramref name="folderId"/> when the
    /// store has no such default folder: <c>OutlookAI-Corpus-Folder-Junk</c>, or the id itself.
    /// </summary>
    public static string StandInName(int folderId)
        => CorpusManifest.CreatedFolderPrefix + "-"
            + (folderId == JunkFolderId ? "Junk" : folderId.ToString(CultureInfo.InvariantCulture));

    /// <summary>
    /// The id a created folder called <paramref name="name"/> counts under. Under the store root it
    /// is a stand-in, and its name says which default folder it stands in for. Under a default
    /// folder it is a population subfolder, which only the population can name. Anything
    /// unrecognised is 0 - "a created folder" - which a census reports as misplaced rather than
    /// guessing.
    /// </summary>
    /// <param name="parentFolderId">The default folder it was found under, or null for the store root.</param>
    /// <param name="name">The folder's name.</param>
    /// <param name="population">The population being scanned for, when there is one.</param>
    public static int ForCreatedFolder(int? parentFolderId, string name, CorpusPopulation? population)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (parentFolderId != null)
        {
            return population?.FolderIdOf(parentFolderId.Value, name) ?? 0;
        }

        string prefix = CorpusManifest.CreatedFolderPrefix + "-";
        if (!name.StartsWith(prefix, StringComparison.Ordinal))
        {
            return 0;
        }

        string suffix = name.Substring(prefix.Length);
        if (string.Equals(suffix, "Junk", StringComparison.Ordinal))
        {
            return JunkFolderId;
        }

        return int.TryParse(suffix, NumberStyles.None, CultureInfo.InvariantCulture, out int id)
            && id > 0 && id < CorpusPopulation.SubfolderIdBase
            ? id
            : 0;
    }
}
