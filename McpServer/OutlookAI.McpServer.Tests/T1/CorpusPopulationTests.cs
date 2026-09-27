using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the curated fixture POPULATIONS the corpus generator builds into a test guest's hub,
/// bystander and identity stores (Q70, 2026-09-24) - and pins them against the live tests that
/// read those stores rather than against a description of them.
/// <para>
/// <b>Why the demands are restated here.</b> Every one of the 21 hub tests, 15 corpus-content
/// tests and 4 weak-on-a-VM tests the populations exist for makes an exact demand of its store:
/// at most 99 searchable hub rows, a term in every subject of one folder and in no body, a
/// conversation the newest item belongs to, an attachment of each kind the old kind filter
/// dropped. A population that misses one does not fail at build time - it fails, or worse passes
/// having proved nothing, in a live run on a guest, which is the most expensive place to find out.
/// So each demand is checked here against the plan itself, with the SAME selection rules the live
/// test uses - <see cref="HubCorpus.RankedCleanTerms"/> is called directly on a synthetic walk of
/// the plan - and a change to the generator that breaks one fails CI instead.
/// </para>
/// <para>
/// Pure: no Outlook, no mailbox, no guest. The populations' COM half is guest-only and is listed
/// in the report the work behind this class produced.
/// </para>
/// </summary>
public sealed class CorpusPopulationTests
{
    private static readonly DateTime Anchor = new(2026, 9, 24, 9, 0, 0, DateTimeKind.Utc);

    private const string HubStore = "tier@vm.invalid";
    private const string BystanderStore = "bystander@vm.invalid";
    private const string IdentityStore = "identity@vm.invalid";

    /// <summary>
    /// The live search tool's top-100 ceiling on hub hits (Phase7...Search_TopOne_OnHubStore wants
    /// 2..99), minus room for the tagged drafts other tests leave in the hub while a run is going.
    /// </summary>
    private const int HubSearchRowCeiling = 80;

    private static CorpusPlan Plan(CorpusPopulationKind kind, string store, long seed = 8181, string? id = null)
        => new(new CorpusPlanOptions(id ?? kind.ToString().ToLowerInvariant() + "-synthetic", seed, Anchor)
        {
            Population = kind,
            Owner = CorpusMailboxOwner.ForStore(store),
        });

    private static CorpusPlan Hub() => Plan(CorpusPopulationKind.Hub, HubStore);

    private static CorpusPlan Bystander() => Plan(CorpusPopulationKind.Bystander, BystanderStore, 8282);

    private static CorpusPlan Identity() => Plan(CorpusPopulationKind.Identity, IdentityStore, 8383);

    private static IEnumerable<int> Ordinals(CorpusPlan plan) => Enumerable.Range(1, plan.FixedItemCount!.Value);

    private static string Body(CorpusPlan plan, int ordinal) => plan.BuildBody(plan.Describe(ordinal));

    /// <summary>
    /// What <c>OutlookComSession.WalkStoreMailItems</c> would return for the built store: every MAIL
    /// item (OlObjectClass 43) - the dated mail; since population version 2 there are no undated
    /// drafts - and none of the appointments, contacts and tasks, which are other classes.
    /// </summary>
    private static List<ComWalkedItem> MailWalk(CorpusPlan plan)
        => Ordinals(plan)
            .Where(o => plan.Describe(o).Kind is CorpusItemKind.Mail)
            .Select(o =>
            {
                CorpusItemSpec spec = plan.Describe(o);
                return new ComWalkedItem(
                    "E" + o.ToString("D8", CultureInfo.InvariantCulture),
                    spec.Subject,
                    Body(plan, o),
                    spec.IsUndated ? null : spec.ReceivedUtc,
                    plan.FolderLabel(spec.FolderId),
                    43);
            }).ToList();

    /// <summary>
    /// The ground truth the completeness oracle and <c>LivePhase3Fixture.TestHubCorpus</c> derive
    /// terms from: the mail walk with ReceivedTime - every item of it, now that no draft is planned.
    /// </summary>
    private static List<ComWalkedItem> Walk(CorpusPlan plan)
        => MailWalk(plan).Where(i => i.ReceivedTime.HasValue).ToList();

    /// <summary>The ordinals of the dated items - everything <see cref="CorpusPlan.Enrich"/> answers for.</summary>
    private static IEnumerable<int> Dated(CorpusPlan plan) => Ordinals(plan).Where(o => !plan.Describe(o).IsUndated);

    /// <summary>The ordinals of the UNDATED items.</summary>
    private static IEnumerable<int> Undated(CorpusPlan plan) => Ordinals(plan).Where(o => plan.Describe(o).IsUndated);

    private static IEnumerable<string> Tokens(string text, int min = 4, int max = int.MaxValue)
        => Regex.Matches(text, "[A-Za-z]+", RegexOptions.CultureInvariant)
            .Select(m => m.Value.ToLowerInvariant())
            .Where(t => t.Length >= min && t.Length <= max);

    // ================================================================ the default is untouched

    [Fact]
    public void TheMeasurementCorpusShapeKey_IsExactlyTheOneTestbedJsonRecords()
    {
        // The corpus every published measurement rests on is reproduced by four parameters and the
        // default shape. Adding populations must not move a single byte of that shape's key.
        string path = Path.Combine(RepoRoot(), "Testbed", "testbed.json");
        JsonObject corpus = JsonNode.Parse(File.ReadAllText(path))!["corpus"]!.AsObject();
        var options = new CorpusPlanOptions(
            corpus["corpusId"]!.GetValue<string>(),
            corpus["seed"]!.GetValue<long>(),
            CorpusOptions.ParseAnchor(corpus["anchor"]!.GetValue<string>()));

        Assert.Null(options.Population);
        Assert.Equal(corpus["shapeKey"]!.GetValue<string>(), options.ShapeKey);
        Assert.Null(new CorpusPlan(options).Population);
        Assert.Null(new CorpusPlan(options).Enrich(1));
        Assert.Null(new CorpusPlan(options).FixedItemCount);
    }

    [Fact]
    public void APopulationShapeKey_IsTheCorpusKeyPlusItsKindVersionAndOwner()
    {
        CorpusPlan hub = Hub();
        var bare = new CorpusPlanOptions(hub.Options.CorpusId, hub.Options.Seed, Anchor);

        Assert.StartsWith(bare.ShapeKey, hub.Options.ShapeKey, StringComparison.Ordinal);
        Assert.EndsWith("|p:hub:v" + CorpusPopulation.Version + "|o:" + HubStore + "<" + HubStore + ">", hub.Options.ShapeKey, StringComparison.Ordinal);
        Assert.NotEqual(Plan(CorpusPopulationKind.Bystander, HubStore).Options.ShapeKey, hub.Options.ShapeKey);
        Assert.NotEqual(Plan(CorpusPopulationKind.Hub, "other@vm.invalid").Options.ShapeKey, hub.Options.ShapeKey);
    }

    // ================================================================ determinism

    [Fact]
    public void APopulation_IsAPureFunctionOfItsParameters()
    {
        CorpusPlan a = Hub();
        CorpusPlan b = Hub();
        foreach (int o in Ordinals(a))
        {
            Assert.Equal(a.Describe(o), b.Describe(o));
            Assert.Equal(Body(a, o), Body(b, o));
            Assert.Equal(a.Enrich(o), b.Enrich(o));
        }

        CorpusPlan reseeded = Plan(CorpusPopulationKind.Hub, HubStore, seed: 9999);
        Assert.NotEqual(Render(a), Render(reseeded));
    }

    [Fact]
    public void TheHubPopulation_ContentDigestIsPinned()
    {
        // The whole population - subjects, bodies, dates, folders, read state, senders, recipients,
        // every attachment byte and every conversation index - rendered canonically and hashed. A
        // change here is a change to what a rebuild produces from the same seed: bump
        // CorpusPopulation.Version with it, so an old manifest is refused rather than extended.
        Assert.Equal(PinnedHubDigest, Sha256(Render(Hub())));
    }

    /// <summary>
    /// Recorded 2026-09-24 from population format version 2 - version 1 plus twelve undated items: four
    /// appointments, four contacts and four tasks, ordinals 57-68. (A first cut of version 2 carried three
    /// unsent drafts among them, digest 522DFD26...E668A8; it was never built, and the drafts went the same
    /// day - see <see cref="CorpusItemKind"/>.) Version 1's was
    /// D9B4032E9EF1C84264834E29C52E23921E9F1EB33548F057655D5ED70A972EC0; the fifty-six dated items are
    /// unchanged, which <see cref="TheDatedHubItems_AreExactlyVersionOnes"/> holds.
    /// </summary>
    private const string PinnedHubDigest = "A791C745CBF20C5505001E086E20003AB518790354D693F21E486FE670874D24";

    [Fact]
    public void TheDatedHubItems_AreExactlyVersionOnes()
    {
        // Adding the undated items must not have moved a single dated one: version 1's digest, taken
        // over the same fields version 1 rendered, is reproduced by the first fifty-six ordinals.
        CorpusPlan hub = Hub();
        var sb = new StringBuilder();
        foreach (int o in Dated(hub))
        {
            CorpusItemSpec s = hub.Describe(o);
            sb.Append(o).Append('|').Append(s.FolderId).Append('|').Append(s.Subject).Append('|')
                .Append(CorpusManifest.FormatUtc(s.ReceivedUtc)).Append('|').Append(CorpusManifest.FormatUtc(s.SentUtc))
                .Append('|').Append(s.IsRead).Append('|').Append(Body(hub, o)).Append('|')
                .Append(hub.Enrich(o)!.Render()).Append('\n');
        }

        Assert.Equal(Enumerable.Range(1, 56), Dated(hub));
        Assert.Equal("D9B4032E9EF1C84264834E29C52E23921E9F1EB33548F057655D5ED70A972EC0", Sha256(sb.ToString()));
    }

    private static string Render(CorpusPlan plan)
    {
        var sb = new StringBuilder();
        foreach (int o in Ordinals(plan))
        {
            CorpusItemSpec s = plan.Describe(o);
            sb.Append(o).Append('|').Append(s.FolderId).Append('|').Append(s.Kind).Append('|').Append(s.Subject).Append('|')
                .Append(CorpusManifest.FormatUtc(s.ReceivedUtc)).Append('|').Append(CorpusManifest.FormatUtc(s.SentUtc))
                .Append('|').Append(s.IsRead).Append('|').Append(Body(plan, o)).Append('|');
            CorpusUndatedDetail? detail = plan.UndatedDetail(o);
            if (detail != null)
            {
                sb.Append("undated:").Append(detail.Kind)
                    .Append(':').Append(detail.AppointmentStartUtc == null ? "-" : CorpusManifest.FormatUtc(detail.AppointmentStartUtc.Value))
                    .Append(':').Append(detail.AppointmentMinutes?.ToString(CultureInfo.InvariantCulture) ?? "-")
                    .Append(':').Append(detail.ContactFullName ?? "-");
            }
            else
            {
                sb.Append(plan.Enrich(o)!.Render());
            }

            sb.Append('\n');
        }

        return sb.ToString();
    }

    private static string Sha256(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));

    // ================================================================ every item is a corpus item

    [Theory]
    [InlineData(CorpusPopulationKind.Hub)]
    [InlineData(CorpusPopulationKind.Bystander)]
    [InlineData(CorpusPopulationKind.Identity)]
    public void EverySubject_CarriesTheCorpusTagsAndParsesBackToItsOrdinal_AndNeverTheArtifactTag(CorpusPopulationKind kind)
    {
        CorpusPlan plan = Plan(kind, HubStore);
        foreach (int o in Ordinals(plan))
        {
            string subject = plan.Describe(o).Subject;
            Assert.StartsWith(CorpusPlan.SubjectTag, subject, StringComparison.Ordinal);
            Assert.True(CorpusPlan.TryParseOrdinal(subject, plan.Options.CorpusId, out int parsed));
            Assert.Equal(o, parsed);
            Assert.Equal(subject, plan.BuildSubject(o));
            Assert.DoesNotContain(RemediationRules.DaslCountFragment, subject, StringComparison.Ordinal);
            Assert.False(HubCorpus.IsTestArtifact(subject));
        }
    }

    [Fact]
    public void AnOrdinalOutsideThePopulation_IsRefused()
    {
        CorpusPlan hub = Hub();
        Assert.Throws<ArgumentOutOfRangeException>(() => hub.Describe(hub.FixedItemCount!.Value + 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => hub.Describe(0));
    }

    [Fact]
    public void ACorpusIdContainingAPopulationWord_IsRefused()
    {
        // The id is in every subject, so a body word in it would be in every subject too.
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => Plan(CorpusPopulationKind.Hub, HubStore, id: "invoice-hub"));
        Assert.Contains("invoice", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APopulationWithoutAnOwner_IsRefused()
    {
        Assert.Throws<ArgumentException>(() => new CorpusPlan(
            new CorpusPlanOptions("hub-x", 1, Anchor) { Population = CorpusPopulationKind.Hub }));
    }

    [Fact]
    public void BodiesAndSubjectsAreAscii_AndSubmitNeverFollowsDelivery_AndNothingIsAfterTheAnchor()
    {
        foreach (CorpusPlan plan in new[] { Hub(), Bystander(), Identity() })
        {
            foreach (int o in Ordinals(plan))
            {
                CorpusItemSpec s = plan.Describe(o);
                string body = Body(plan, o);
                Assert.All(s.Subject + body, c => Assert.True(c < 0x7F));
                Assert.Equal(body.Length, s.BodyBytes);
                Assert.True(s.SentUtc <= s.ReceivedUtc);
                Assert.True(s.ReceivedUtc < Anchor);
            }
        }
    }

    // ================================================================ the owner, and addressing

    [Fact]
    public void AStoreNamedAsAnAddress_IsItsOwnersAddress_AndAnyOtherNameGetsAnInvalidAddress()
    {
        Assert.Equal(new CorpusMailboxOwner(HubStore, HubStore), CorpusMailboxOwner.ForStore(HubStore));
        CorpusMailboxOwner named = CorpusMailboxOwner.ForStore("OutlookAI Bystander");
        Assert.Equal("OutlookAI Bystander", named.Name);
        Assert.Equal("outlookai.bystander@" + CorpusMailboxOwner.FallbackDomain, named.Address);
        Assert.Throws<ArgumentException>(() => CorpusMailboxOwner.ForStore("!!!"));
    }

    [Fact]
    public void AnOwnerWhoseNameIsAnAddress_IsAddedAsTheBareAddress_AndEveryoneElseByNameAndAddress()
    {
        // OAI-UNINDEXED, 2026-09-24: Outlook's resolver refused "tier@vm.invalid <tier@vm.invalid>", and
        // every RECEIVED item of three populations carried one unresolved, address-less To row. The
        // CONTROL is the sent items' "Kester Wren <kester.wren@margie.invalid>", which resolved.
        CorpusMailboxOwner owner = CorpusMailboxOwner.ForStore(HubStore);
        Assert.Equal(HubStore, new CorpusCorrespondent(owner.Name, owner.Address).ToRecipientSpec());
        Assert.Equal("Kester Wren <kester.wren@margie.invalid>",
            new CorpusCorrespondent("Kester Wren", "kester.wren@margie.invalid").ToRecipientSpec());

        // Case-blind, and any name the resolver could read as an address part is dropped too.
        Assert.Equal("tier@vm.invalid", new CorpusCorrespondent("TIER@VM.INVALID", "tier@vm.invalid").ToRecipientSpec());
        Assert.Equal("x@vm.invalid", new CorpusCorrespondent("someone@else.invalid", "x@vm.invalid").ToRecipientSpec());
        Assert.Equal("x@vm.invalid", new CorpusCorrespondent("A <b>", "x@vm.invalid").ToRecipientSpec());

        // The header form is untouched: attachments and the plan's own rendering still carry Name <address>.
        Assert.Equal(HubStore + " <" + HubStore + ">", new CorpusCorrespondent(owner.Name, owner.Address).ToAddressSpec());

        // Every correspondent a population writes keeps its display name on the row.
        foreach (CorpusCorrespondent c in CorpusPopulation.Correspondents.Append(CorpusPopulation.NoticeSender))
        {
            Assert.Equal(c.ToAddressSpec(), c.ToRecipientSpec());
        }

        // And a hub item addressed to its owner is exactly the case: the owner is a To row of every received item.
        CorpusPlan hub = Hub();
        Assert.Contains(Dated(hub), o => hub.Enrich(o)!.Recipients.Any(r =>
            r.Kind == CorpusRecipientKind.To && string.Equals(r.Person.Address, HubStore, StringComparison.Ordinal)
            && r.Person.ToRecipientSpec() == HubStore));
    }

    [Fact]
    public void EveryAddressAPopulationCanWrite_IsUnderDotInvalid()
    {
        // RFC 2606: .invalid cannot resolve, so a population item can never address anybody real.
        foreach (CorpusCorrespondent c in CorpusPopulation.Correspondents.Append(CorpusPopulation.NoticeSender))
        {
            Assert.EndsWith(".invalid", c.Address, StringComparison.Ordinal);
        }

        Assert.EndsWith(".invalid", CorpusMailboxOwner.FallbackDomain, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(CorpusPopulationKind.Hub, HubStore)]
    [InlineData(CorpusPopulationKind.Bystander, BystanderStore)]
    [InlineData(CorpusPopulationKind.Identity, IdentityStore)]
    public void EveryReceivedItemIsAddressedToTheOwner_AndEverySentItemIsFromThem(CorpusPopulationKind kind, string store)
    {
        // What makes a small store findable at all: IndexSearchService.TryDiscoverStoreScopeByAddress
        // looks for mail whose recipient address or sender contains the store's name, and
        // MailService does the same for any store named as an address. The 2000-row sample that
        // finds big stores misses small ones.
        CorpusPlan plan = Plan(kind, store);
        int received = 0;
        int sent = 0;
        foreach (int o in Dated(plan))
        {
            CorpusItemEnrichment e = plan.Enrich(o)!;
            if (string.Equals(e.Sender.Address, store, StringComparison.OrdinalIgnoreCase))
            {
                sent++;
                Assert.Equal(5, plan.Describe(o).FolderId);
                Assert.NotEmpty(e.Recipients);
            }
            else
            {
                received++;
                Assert.Contains(e.Recipients, r => r.Kind == CorpusRecipientKind.To
                    && string.Equals(r.Person.Address, store, StringComparison.OrdinalIgnoreCase));
            }
        }

        Assert.True(received > 0 && sent > 0, $"{kind}: {received} received, {sent} sent");
    }

    // ================================================================ the hub, test by test

    [Fact]
    public void TheHubStaysSmallEnoughForEveryTestThatCountsIt_AndBigEnoughToPage()
    {
        CorpusPlan hub = Hub();
        int items = hub.FixedItemCount!.Value;
        int attachments = Dated(hub).Sum(o => hub.Enrich(o)!.Attachments.Count);

        // Phase7...Search_TopOne_OnHubStore: 2..99 search rows (messages AND attachment rows), with
        // room left for the drafts other tests leave in the hub mid-run. EVERY item counts, the
        // undated ones included: since gap B3 a store-scoped search admits every item class, so an
        // appointment is a search hit exactly as a mail is.
        Assert.InRange(items + attachments, 2, HubSearchRowCeiling);

        // LiveResumableScanTests: the unpaged control is top 100 and must be "complete"; the paged
        // run is pages of 2 and the superseded-token test needs a THIRD page to exist.
        Assert.InRange(items, 5, 100 - 20);

        // LiveExhaustiveSearchTests: ranked[0] - the most frequent clean term - is searched with
        // top 100 and must not truncate.
        List<ComWalkedItem> walk = Walk(hub);
        string top = HubCorpus.RankedCleanTerms(walk)[0];
        Assert.InRange(walk.Count(i => HubCorpus.WordRegex(top).IsMatch(HubCorpus.TextOf(i))), 1, 100);
    }

    [Fact]
    public void TheHubOffersTheOracleAndTheExhaustiveTestsTheirTerms()
    {
        // LiveCompletenessOracleTests wants at least three clean terms; LiveExhaustiveSearchTests
        // at least two, and a folder holding matches of the most frequent one.
        List<ComWalkedItem> walk = Walk(Hub());
        IReadOnlyList<string> ranked = HubCorpus.RankedCleanTerms(walk);
        Assert.True(ranked.Count >= 3, $"only {ranked.Count} clean terms");
        Assert.Contains(walk, i => HubCorpus.WordRegex(ranked[^1]).IsMatch(HubCorpus.TextOf(i)));

        // No undated item is MAIL any more (population version 2 dropped the drafts), so the COM walk
        // and its ReceivedTime-filtered ground truth are the same set: the appointments, contacts and
        // tasks are other classes and never reach the walk at all.
        List<ComWalkedItem> mail = MailWalk(Hub());
        Assert.Equal(0, mail.Count(i => i.ReceivedTime == null));
        Assert.Equal(mail.Count, walk.Count);
        Assert.All(walk, i => Assert.NotNull(i.ReceivedTime));
    }

    [Fact]
    public void SubjectWordsNeverOccurInABody_AndBodyWordsNeverInASubject()
    {
        // LiveSearchInTests.AllTiers_SubjectOnlyAndBodyOnlyTerms and LiveSweepScopeTests' two-term
        // probe both need a word that is in one field and - as a SUBSTRING, which is stricter than
        // the word breaker - nowhere in the other. Held for the whole vocabulary, not just "some".
        foreach (CorpusPlan plan in new[] { Hub(), Bystander(), Identity() })
        {
            List<string> subjects = Ordinals(plan).Select(o => plan.Describe(o).Subject.ToLowerInvariant()).ToList();
            List<string> bodies = Ordinals(plan).Select(o => Body(plan, o).ToLowerInvariant()).ToList();
            foreach (string word in CorpusPopulation.SubjectVocabulary.Append(CorpusPopulation.SubjectOnlyTerm))
            {
                Assert.DoesNotContain(bodies, b => b.Contains(word, StringComparison.Ordinal));
            }

            foreach (string word in CorpusPopulation.BodyVocabulary.Concat(CorpusPopulation.BodyFillerWords.Where(w => w.Length >= 4)))
            {
                Assert.DoesNotContain(subjects, s => s.Contains(word, StringComparison.Ordinal));
            }

            // And the undated items keep the same split within their own vocabulary.
            foreach (string word in CorpusPopulation.UndatedSubjectVocabulary)
            {
                Assert.DoesNotContain(bodies, b => b.Contains(word, StringComparison.Ordinal));
            }

            foreach (string word in CorpusPopulation.UndatedBodyVocabulary)
            {
                Assert.DoesNotContain(subjects, s => s.Contains(word, StringComparison.Ordinal));
            }
        }
    }

    [Fact]
    public void TheUndatedVocabulary_AndEverythingADatedItemCarries_NeverMeet()
    {
        // The live tests that derive terms from the hub's MAIL - the oracle, the exhaustive known
        // answer, the subject/body separation - compare against ground truth drawn from received
        // mail only. An undated row carrying one of those terms would be a hit that ground truth
        // cannot contain; a dated item carrying an undated word would be the reverse. So the two
        // vocabularies are disjoint as SUBSTRINGS, in both directions, against every text a dated
        // item carries - subject, body, attachment text and names, and every address and name.
        foreach (CorpusPlan plan in new[] { Hub(), Bystander(), Identity() })
        {
            var datedText = new StringBuilder();
            foreach (int o in Dated(plan))
            {
                CorpusItemEnrichment e = plan.Enrich(o)!;
                datedText.Append(plan.Describe(o).Subject).Append('\n').Append(Body(plan, o)).Append('\n')
                    .Append(e.Sender.Name).Append(' ').Append(e.Sender.Address).Append('\n');
                foreach (CorpusRecipient r in e.Recipients)
                {
                    datedText.Append(r.Person.Name).Append(' ').Append(r.Person.Address).Append('\n');
                }

                foreach (CorpusAttachment a in e.Attachments)
                {
                    datedText.Append(a.FileName).Append(' ').Append(a.Text).Append('\n');
                }
            }

            string dated = datedText.ToString().ToLowerInvariant();
            foreach (string word in CorpusPopulation.UndatedSubjectVocabulary.Concat(CorpusPopulation.UndatedBodyVocabulary))
            {
                Assert.DoesNotContain(word, dated, StringComparison.Ordinal);
            }

            // ...and no dated term of four letters or more - what every term picker extracts - is in
            // an undated item's text, the corpus tags aside: those are in every subject by design,
            // and every picker that meets them in an undated row excludes that row by class or by date.
            var tagTokens = new HashSet<string>(
                Tokens(CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + plan.Options.CorpusId), StringComparer.Ordinal);
            var datedTokens = new HashSet<string>(Tokens(dated), StringComparer.Ordinal);
            foreach (int o in Undated(plan))
            {
                string text = (plan.Describe(o).Subject + "\n" + Body(plan, o) + "\n" + plan.UndatedDetail(o)!.ContactFullName).ToLowerInvariant();
                foreach (string token in Tokens(text).Where(t => !tagTokens.Contains(t)))
                {
                    Assert.False(datedTokens.Contains(token), $"{plan.Population!.Kind} ordinal {o}: '{token}' is also a dated item's word");
                }
            }
        }
    }

    [Fact]
    public void TheUndatedItems_AreFourOfEachKindInTheHub_FourteenOfEachInTheBystander_NoneForIdentity()
    {
        // LiveOrderKeyCollationTests measure where a row with NO received date sorts. The hub carries
        // every undated kind a population can build - appointments, contacts and tasks, classes that
        // have no received date at all - and the bystander carries the VOLUME: WidenedSearch's
        // order-key refetch can only be told apart from its absence when more undated rows sort ahead
        // of a TOP 25 cut than its over-fetch to TOP 60 leaves room for, and the hub's search must
        // stay under 100 hits.
        CorpusPlan hub = Hub();
        Assert.Equal(new[] { CorpusItemKind.Appointment, CorpusItemKind.Contact, CorpusItemKind.Task }, CorpusItemKinds.Undated);
        Assert.Equal(CorpusItemKinds.Undated, hub.Population!.UndatedKinds);
        foreach (CorpusItemKind kind in CorpusItemKinds.Undated)
        {
            Assert.Equal(4, Undated(hub).Count(o => hub.Describe(o).Kind == kind));
        }

        Assert.Equal(Enumerable.Range(57, 12), Undated(hub));

        CorpusPlan bystander = Bystander();
        Assert.Equal(CorpusItemKinds.Undated, bystander.Population!.UndatedKinds);
        Assert.Equal(42, Undated(bystander).Count());
        Assert.True(Undated(bystander).Count() > 60 - 25, "the bystander must out-number the widened search's over-fetch");

        // No undated item is MAIL: an unsent mail item's first save is filed in the profile's DEFAULT
        // store's Drafts (OAI-UNINDEXED, 2026-09-24), and a population is never built in the default store.
        foreach (CorpusPlan plan in new[] { hub, bystander })
        {
            Assert.DoesNotContain(Undated(plan), o => plan.Describe(o).Kind == CorpusItemKind.Mail);
        }

        Assert.Empty(Undated(Identity()));
        Assert.Empty(Identity().Population!.UndatedKinds);
    }

    [Theory]
    [InlineData(CorpusPopulationKind.Hub, HubStore)]
    [InlineData(CorpusPopulationKind.Bystander, BystanderStore)]
    public void EveryUndatedItem_LivesInItsKindsFolder_CarriesNoDate_AndHasNoSenderOrRecipient(CorpusPopulationKind kind, string store)
    {
        CorpusPlan plan = Plan(kind, store);
        Assert.NotEmpty(Undated(plan));
        foreach (int o in Undated(plan))
        {
            CorpusItemSpec spec = plan.Describe(o);
            Assert.True(spec.IsUndated);
            Assert.Equal(CorpusItemKinds.FolderIdOf(spec.Kind), spec.FolderId);
            Assert.Equal(CorpusItemSpec.UndatedInstant, spec.ReceivedUtc);
            Assert.Equal(CorpusItemSpec.UndatedInstant, spec.SentUtc);
            Assert.True(spec.IsRead);
            Assert.StartsWith("undated-", spec.DateBand, StringComparison.Ordinal);
            Assert.Null(plan.Enrich(o));

            CorpusUndatedDetail detail = plan.UndatedDetail(o)!;
            Assert.Equal(spec.Kind, detail.Kind);
            if (spec.Kind == CorpusItemKind.Appointment)
            {
                // In the past relative to the anchor, so a reminder could never be due - and there is none.
                Assert.True(detail.AppointmentStartUtc < Anchor, $"ordinal {o} starts at {detail.AppointmentStartUtc:O}");
                Assert.Contains(detail.AppointmentMinutes, new int?[] { 30, 60 });
                Assert.Equal(0, detail.AppointmentStartUtc!.Value.Minute);
            }
            else
            {
                Assert.Null(detail.AppointmentStartUtc);
                Assert.Null(detail.AppointmentMinutes);
            }

            // A contact's name carries the subject, tags and all, so whichever of its fields Outlook
            // derives PR_SUBJECT from, the two-key rule can still find and delete it.
            Assert.Equal(spec.Kind == CorpusItemKind.Contact ? spec.Subject : null, detail.ContactFullName);
        }

        foreach (int o in Dated(plan))
        {
            Assert.Null(plan.UndatedDetail(o));
            Assert.NotNull(plan.Enrich(o));
        }
    }

    [Fact]
    public void EveryHubItemOffersTheCrossColumnSplitTheSweepScopeTestLooksFor()
    {
        // LiveSweepScopeTests.SelectCrossColumnProbe: a hub item with a UNIQUE subject, a 5-20 letter
        // subject token absent from its own body, and a body token absent from its own subject.
        CorpusPlan hub = Hub();
        List<string> subjects = Ordinals(hub).Select(o => hub.Describe(o).Subject).ToList();
        Assert.Equal(subjects.Count, subjects.Distinct(StringComparer.Ordinal).Count());
        foreach (int o in Ordinals(hub))
        {
            string subject = hub.Describe(o).Subject;
            string body = Body(hub, o);
            Assert.Contains(Tokens(subject, 5, 20), t => body.IndexOf(t, StringComparison.OrdinalIgnoreCase) < 0);
            Assert.Contains(Tokens(body, 5, 20), t => subject.IndexOf(t, StringComparison.OrdinalIgnoreCase) < 0);
        }
    }

    [Fact]
    public void NothingOutsideAnItemsOwnPlainText_CarriesATermTheOracleCouldPick()
    {
        // The completeness oracle derives its terms from subjects and bodies, and TOLERATES only a
        // bounded number of index matches on an item whose plain text lacks the term (attachment
        // text, address fields, file names - System.Search.Contents reads all of them). So no
        // subject or body token of the hub may appear in an item's attachments, file names or
        // addresses unless that item's own subject or body carries it too.
        CorpusPlan hub = Hub();
        var plainTokens = new HashSet<string>(
            Dated(hub).SelectMany(o => Tokens(hub.Describe(o).Subject + "\n" + Body(hub, o))), StringComparer.Ordinal);
        foreach (int o in Dated(hub))
        {
            CorpusItemEnrichment e = hub.Enrich(o)!;
            var own = new HashSet<string>(Tokens(hub.Describe(o).Subject + "\n" + Body(hub, o)), StringComparer.Ordinal);
            var hidden = new StringBuilder();
            hidden.Append(e.Sender.Name).Append(' ').Append(e.Sender.Address).Append(' ');
            foreach (CorpusRecipient r in e.Recipients)
            {
                hidden.Append(r.Person.Name).Append(' ').Append(r.Person.Address).Append(' ');
            }

            foreach (CorpusAttachment a in e.Attachments)
            {
                hidden.Append(a.FileName).Append(' ').Append(a.Text).Append(' ');
            }

            foreach (string token in Tokens(hidden.ToString()))
            {
                Assert.False(
                    plainTokens.Contains(token) && !own.Contains(token),
                    $"ordinal {o}: '{token}' is in the hub's plain text and in this item's hidden fields, but not in its own plain text");
            }
        }
    }

    [Fact]
    public void TheSubjectOnlyProbePopulation_IsExactlyWhatTheSf6TestsAssume()
    {
        CorpusPlan hub = Hub();
        CorpusSubjectOnlyProbe probe = hub.Population!.SubjectOnlyProbe!;
        Assert.Equal("Inbox/" + CorpusManifest.CreatedFolderPrefix + "-Notices", probe.FolderPath);
        Assert.Equal(CorpusPopulation.SubjectOnlyTerm, probe.SubjectTerm);
        Assert.Equal(CorpusPopulation.SubjectOnlySenderFragment, probe.SenderFragment);

        // Sf6DiscoveryCase_IndexTier_PrefixStemsWorkInTheSubjectColumnToo stems it by two letters.
        Assert.True(probe.SubjectTerm.Length >= 5);
        string stem = probe.SubjectTerm[..^2];

        List<int> members = Ordinals(hub).Where(o => hub.Describe(o).FolderId == probe.FolderId).ToList();
        Assert.True(members.Count >= 5, $"only {members.Count} members");
        foreach (int o in members)
        {
            CorpusItemEnrichment e = hub.Enrich(o)!;

            // Every member is found by the term in its SUBJECT - as a clean word...
            Assert.Matches(HubCorpus.WordRegex(probe.SubjectTerm), hub.Describe(o).Subject);

            // ...and never in its body, not even as the stem the prefix test searches for.
            string body = Body(hub, o).ToLowerInvariant();
            Assert.DoesNotContain(probe.SubjectTerm, body, StringComparison.Ordinal);
            Assert.DoesNotContain(Tokens(body, 1), t => t.StartsWith(stem, StringComparison.Ordinal));

            // The expectation is derived by SENDER: every member, and nothing else in the folder,
            // comes from the one sender whose name and address carry the fragment as a word.
            Assert.Equal(CorpusPopulation.NoticeSender, e.Sender);
            Assert.Contains(probe.SenderFragment, Tokens(e.Sender.Name, 1));
            Assert.StartsWith(probe.SenderFragment + "@", e.Sender.Address, StringComparison.Ordinal);

            // Attachment rows would enter both counts through a different door; there are none.
            Assert.Empty(e.Attachments);
        }

        // The fragment and the term appear nowhere else in the hub, so a from:/term count is the folder's.
        foreach (int o in Ordinals(hub).Except(members))
        {
            Assert.DoesNotContain(probe.SubjectTerm, hub.Describe(o).Subject, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(probe.SubjectTerm, Body(hub, o), StringComparison.OrdinalIgnoreCase);
            Assert.NotEqual(CorpusPopulation.NoticeSender, hub.Enrich(o)?.Sender);
        }
    }

    [Fact]
    public void TheProbeTerm_HitsBodiesAndOneTextAttachmentWhoseParentCarriesItToo()
    {
        // probeTerm on a guest names this word: LiveMailServiceTests.AttachmentHit_ReadParent_SaveToScratch
        // needs an ATTACHMENT hit for it, and the oracle guard above needs that attachment's parent to
        // carry it in plain text as well.
        CorpusPlan hub = Hub();
        Assert.Contains(CorpusPopulation.ProbeTerm, CorpusPopulation.BodyVocabulary);
        int carriers = 0;
        foreach (int o in Dated(hub))
        {
            foreach (CorpusAttachment a in hub.Enrich(o)!.Attachments.Where(a => a.Kind == CorpusAttachmentKind.Text))
            {
                if (a.Text!.Contains(CorpusPopulation.ProbeTerm, StringComparison.Ordinal))
                {
                    carriers++;
                    Assert.Contains(CorpusPopulation.ProbeTerm, Body(hub, o), StringComparison.Ordinal);
                }
            }
        }

        Assert.Equal(1, carriers);
        Assert.Contains(Ordinals(hub), o => Body(hub, o).Contains(CorpusPopulation.ProbeTerm, StringComparison.Ordinal));

        // Never in an undated item: QuerySetLatency and the MCP-shaped probe search it across stores,
        // and the attachment-hit tests read the parent of the hit it produces.
        Assert.DoesNotContain(Undated(hub), o => (hub.Describe(o).Subject + Body(hub, o)).Contains(CorpusPopulation.ProbeTerm, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void TheHubCarriesEveryAttachmentKindTheOldKindFilterDropped_AndAMultiAttachmentMail()
    {
        // LiveAttachmentKindRecallTests (recovered kinds must exist; one previously-dropped extension
        // must come back from search) and LiveIndexSearchTests.FilterShapes (has-attachments in the
        // FIRST indexed store, which is the hub).
        CorpusPlan hub = Hub();
        string[] previouslyDropped = { ".png", ".ics", ".eml" };
        List<CorpusAttachment> all = Dated(hub).SelectMany(o => hub.Enrich(o)!.Attachments).ToList();
        foreach (CorpusAttachmentKind kind in Enum.GetValues<CorpusAttachmentKind>())
        {
            Assert.Contains(all, a => a.Kind == kind);
        }

        foreach (string extension in previouslyDropped)
        {
            Assert.Contains(all, a => a.FileName.EndsWith(extension, StringComparison.Ordinal));
        }

        Assert.Contains(Dated(hub), o => hub.Enrich(o)!.Attachments.Count >= 3);
        Assert.All(all, a => Assert.True(a.Length > 0));
        Assert.Equal(all.Count, all.Select(a => a.FileName).Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void TheHubHasUnreadAndReadMail_ForTheUnreadFilterAndForShowMe()
    {
        CorpusPlan hub = Hub();
        Assert.Contains(Ordinals(hub), o => !hub.Describe(o).IsRead);
        Assert.Contains(Ordinals(hub), o => hub.Describe(o).IsRead);
    }

    [Fact]
    public void TheNewestHubItemIsOneMinuteBeforeTheAnchor_AndBelongsToAConversation()
    {
        // LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier can only catch a local-time
        // misreading of the frontier when the frontier is younger than the machine's UTC offset;
        // a hub rebuilt against the moment a run starts puts it one minute old. And
        // LiveMailServiceTests.Thread picks the most recent hit with a conversation id.
        CorpusPlan hub = Hub();
        int newest = Ordinals(hub).OrderByDescending(o => hub.Describe(o).ReceivedUtc).First();
        Assert.Equal(Anchor.AddMinutes(-1), hub.Describe(newest).ReceivedUtc);
        Assert.NotNull(hub.Enrich(newest)!.ThreadKey);
        Assert.Contains(Ordinals(hub), o => hub.Describe(o).ReceivedUtc > Anchor.AddDays(-30));
    }

    [Fact]
    public void Conversations_ShareOneHeaderPerThread_AndGrowOneChildBlockPerReply()
    {
        foreach (CorpusPlan plan in new[] { Hub(), Bystander() })
        {
            var headers = new Dictionary<int, string>();
            var members = new Dictionary<int, List<int>>();
            foreach (int o in Dated(plan))
            {
                CorpusItemEnrichment e = plan.Enrich(o)!;
                if (e.ThreadKey == null)
                {
                    Assert.Null(e.ConversationIndex);
                    continue;
                }

                byte[] index = e.ConversationIndex!;
                Assert.Equal(0x01, index[0]);
                Assert.Equal(0, (index.Length - 22) % 5);
                string header = Convert.ToHexString(index, 0, 22);
                if (headers.TryGetValue(e.ThreadKey.Value, out string? known))
                {
                    Assert.Equal(known, header);
                }
                else
                {
                    headers[e.ThreadKey.Value] = header;
                }

                if (!members.TryGetValue(e.ThreadKey.Value, out List<int>? list))
                {
                    members[e.ThreadKey.Value] = list = new List<int>();
                }

                list.Add(o);
            }

            Assert.True(members.Count >= 3);
            Assert.Equal(headers.Count, headers.Values.Distinct().Count());
            foreach (List<int> thread in members.Values)
            {
                Assert.True(thread.Count >= 3);
                List<int> byDate = thread.OrderBy(o => plan.Describe(o).SentUtc).ToList();
                for (int i = 0; i < byDate.Count; i++)
                {
                    Assert.Equal(22 + (5 * i), plan.Enrich(byDate[i])!.ConversationIndex!.Length);
                    Assert.Equal(i == 0 ? false : true, plan.Describe(byDate[i]).Subject.Contains("] RE: ", StringComparison.Ordinal));
                }
            }
        }
    }

    // ================================================================ the bystander and the identity store

    [Fact]
    public void TheBystanderIsAFewHundredItems_EveryFolderInsideTheCensusIdentityBudget()
    {
        // The census identity budget is 500 items per folder and 3,000 per store: inside it a store
        // is walked item by item, outside it only counted. The bystander is the one store the guard
        // can DECIDE on, so it must be walked (Docs/live-tier-on-the-vm.md section 1.3).
        CorpusPlan bystander = Bystander();
        int items = bystander.FixedItemCount!.Value;
        Assert.InRange(items, 200, 500);
        foreach (IGrouping<int, int> folder in Ordinals(bystander).GroupBy(o => bystander.Describe(o).FolderId))
        {
            Assert.InRange(folder.Count(), 1, 500);
        }
    }

    [Fact]
    public void TheBystanderHasAMailFolderWithPopulatedChildren_ForTheExcludeSubfoldersMeasurement()
    {
        // LiveFolderScopeTests.PrimaryStore_ExcludeSubfolders_NarrowsExactly: in the first non-hub
        // indexed store, a folder of 5..20,000 items with at least one populated child.
        CorpusPlan bystander = Bystander();
        var counts = Ordinals(bystander).GroupBy(o => bystander.Describe(o).FolderId).ToDictionary(g => g.Key, g => g.Count());
        CorpusPopulation population = bystander.Population!;
        IGrouping<int, CorpusPopulationFolder> byParent = population.Folders.GroupBy(f => f.ParentFolderId).First();
        Assert.InRange(counts[byParent.Key], 5, 20_000);
        Assert.Contains(byParent, child => counts.TryGetValue(child.FolderId, out int n) && n >= 1);
    }

    [Fact]
    public void TheIdentityPopulation_IsSmallAndAddressed()
    {
        CorpusPlan identity = Identity();
        Assert.InRange(identity.FixedItemCount!.Value, 3, 20);
        Assert.Empty(identity.Population!.Folders);
    }

    [Fact]
    public void EveryPopulationFolder_IsASyntheticIdUnderADefaultFolder_NamedWithTheCreatedPrefix()
    {
        foreach (CorpusPlan plan in new[] { Hub(), Bystander() })
        {
            foreach (CorpusPopulationFolder folder in plan.Population!.Folders)
            {
                Assert.True(folder.FolderId >= CorpusPopulation.SubfolderIdBase);
                Assert.StartsWith(CorpusManifest.CreatedFolderPrefix + "-", folder.Name, StringComparison.Ordinal);
                Assert.Equal(folder.Path, plan.FolderLabel(folder.FolderId));
                Assert.Equal(folder.FolderId, plan.Population.FolderIdOf(folder.ParentFolderId, folder.Name));
                Assert.Equal(folder.FolderId, CorpusFolderIds.ForCreatedFolder(folder.ParentFolderId, folder.Name, plan.Population));
            }
        }
    }

    [Fact]
    public void TheCorpusFolderPrefix_AndTheLiveTiersTestFolderPrefix_NeverMatchEachOther()
    {
        // LiveOutlookTestMailer.DeleteTestFolders removes folders by ITS prefix in the hub - which now
        // holds population subfolders. The two prefixes must not be able to select each other's
        // folders, in either direction, exactly as the two subject tags must not.
        string corpus = CorpusManifest.CreatedFolderPrefix;
        string artifact = LiveOutlookTestMailer.TestFolderNamePrefix;
        Assert.DoesNotContain(artifact, corpus, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(corpus, artifact, StringComparison.OrdinalIgnoreCase);
    }

    // ================================================================ created folders without a manifest

    [Theory]
    [InlineData("OutlookAI-Corpus-Folder-Junk", 23)]
    [InlineData("OutlookAI-Corpus-Folder-5", 5)]
    [InlineData("OutlookAI-Corpus-Folder-Projects", 0)]
    [InlineData("OutlookAI-Corpus-Folder-1001", 0)]
    [InlineData("Something Else", 0)]
    public void AStandInUnderTheRoot_CountsUnderTheFolderItsNameSays(string name, int expected)
    {
        Assert.Equal(expected, CorpusFolderIds.ForCreatedFolder(null, name, null));
    }

    [Fact]
    public void AStandInName_RoundTrips()
    {
        foreach (int id in new[] { 3, 5, 6, 23 })
        {
            Assert.Equal(id, CorpusFolderIds.ForCreatedFolder(null, CorpusFolderIds.StandInName(id), null));
        }
    }

    [Fact]
    public void ASubfolderOnlyAPopulationCanName_IsZeroWithoutIt()
    {
        CorpusPopulationFolder folder = Hub().Population!.Folders[0];
        Assert.Equal(0, CorpusFolderIds.ForCreatedFolder(folder.ParentFolderId, folder.Name, null));
    }

    // ================================================================ the census, with subfolders

    [Fact]
    public void ACensusOfAPopulation_CountsItsSubfoldersByName_AndCatchesAnItemInTheWrongOne()
    {
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        List<CorpusSighting> perfect = Ordinals(hub).Select(o => new CorpusSighting(o, hub.Describe(o).FolderId)).ToList();
        CorpusCensusReport clean = CorpusCensus.Compare(hub, count, perfect);
        (bool ok, string message) = CorpusCensus.Decide(clean);
        Assert.True(ok, message);
        Assert.Contains("Inbox/" + CorpusManifest.CreatedFolderPrefix + "-Notices=12/12", message, StringComparison.Ordinal);

        int notice = hub.Population!.SubjectOnlyProbe!.FolderId;
        List<CorpusSighting> moved = perfect.Select(s => s.FolderId == notice ? s with { FolderId = 6 } : s).ToList();
        CorpusCensusReport misplaced = CorpusCensus.Compare(hub, count, moved);
        Assert.Equal(12, misplaced.Misplaced);
        Assert.False(CorpusCensus.Decide(misplaced).Clean);
    }

    [Fact]
    public void ACensus_CountsTheUndatedKindsInTheirFolders_AndStillCatchesMailThatStrayedIntoDrafts()
    {
        // The census has always read an item in Drafts as the first real build's failure - every
        // item filed as a draft - and since population version 2 nothing is planned there at all: the
        // undated items are appointments, contacts and tasks, each counted in its own folder.
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        List<CorpusSighting> perfect = Ordinals(hub).Select(o => new CorpusSighting(o, hub.Describe(o).FolderId)).ToList();
        CorpusCensusReport clean = CorpusCensus.Compare(hub, count, perfect);
        Assert.Equal(0, clean.StrayDrafts);
        (bool ok, string message) = CorpusCensus.Decide(clean);
        Assert.True(ok, message);
        Assert.DoesNotContain("Drafts=", message, StringComparison.Ordinal);
        Assert.Contains("Calendar=4/4", message, StringComparison.Ordinal);
        Assert.Contains("Contacts=4/4", message, StringComparison.Ordinal);
        Assert.Contains("Tasks=4/4", message, StringComparison.Ordinal);

        int mail = Dated(hub).First();
        CorpusCensusReport strayed = CorpusCensus.Compare(
            hub, count, perfect.Select(s => s.Ordinal == mail ? s with { FolderId = CorpusCensus.DraftsFolderId } : s));
        Assert.Equal(1, strayed.StrayDrafts);
        Assert.Contains("in DRAFTS", CorpusCensus.Decide(strayed).Message, StringComparison.Ordinal);

        // An undated item found anywhere but its own folder is misplaced, and not a stray draft.
        int appointment = Undated(hub).First(o => hub.Describe(o).Kind == CorpusItemKind.Appointment);
        CorpusCensusReport appointmentMoved = CorpusCensus.Compare(
            hub, count, perfect.Select(s => s.Ordinal == appointment ? s with { FolderId = 6 } : s));
        Assert.Equal(0, appointmentMoved.StrayDrafts);
        Assert.Equal(1, appointmentMoved.Misplaced);
        Assert.False(CorpusCensus.Decide(appointmentMoved).Clean);
    }

    [Fact]
    public void TheMeasurementCorpusCensus_StillCountsEveryDraftAsAStray()
    {
        // Unchanged for the corpus, which plans nothing into Drafts: every item there is a stray.
        var plan = new CorpusPlan(new CorpusPlanOptions("vm-x", 7777, Anchor));
        List<CorpusSighting> sightings = Enumerable.Range(1, 40)
            .Select(o => new CorpusSighting(o, o <= 5 ? CorpusCensus.DraftsFolderId : plan.Describe(o).FolderId)).ToList();
        CorpusCensusReport report = CorpusCensus.Compare(plan, 40, sightings);
        Assert.Equal(5, report.StrayDrafts);
    }

    [Fact]
    public void ThePlanReport_CountsUndatedItems_AndLeavesThemOutOfEveryDateLine()
    {
        CorpusPlan hub = Hub();
        CorpusPlanReport report = hub.Report(1, hub.FixedItemCount!.Value);
        Assert.Equal(12, report.UndatedItems);
        Assert.Equal(hub.FixedItemCount!.Value, report.ByFolderId.Values.Sum());
        Assert.Equal(Anchor.AddMinutes(-1), report.NewestReceivedUtc);
        Assert.True(report.OldestReceivedUtc > Anchor.AddDays(-400));
        Assert.Equal(Dated(hub).Count(), report.WithinDays[365]);
        Assert.Equal(Dated(hub).Count(o => !hub.Describe(o).IsRead), report.UnreadItems);

        using var output = new StringWriter(CultureInfo.InvariantCulture);
        CorpusCommands.WriteReport(hub, report, output);
        Assert.Contains("undated               : 12", output.ToString(), StringComparison.Ordinal);
        Assert.Contains("undated items         : appointment=4, contact=4, task=4", output.ToString(), StringComparison.Ordinal);

        // And the measurement corpus's sheet gains no line.
        var corpus = new CorpusPlan(new CorpusPlanOptions("vm-x", 7777, Anchor));
        using var corpusOutput = new StringWriter(CultureInfo.InvariantCulture);
        CorpusCommands.WriteReport(corpus, corpus.Report(1, 200), corpusOutput);
        Assert.DoesNotContain("undated", corpusOutput.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Freshness_NeverCountsAnUndatedItemInAWindow()
    {
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        CorpusFreshnessReport atAnchor = CorpusFreshness.Evaluate(hub, count, TimeSpan.Zero, Anchor);
        Assert.Equal(Dated(hub).Count(), atAnchor.Windows.Single(w => w.Days == 365).PlannedCount);

        // Long after the anchor every window has emptied - the undated items do not keep one alive.
        CorpusFreshnessReport later = CorpusFreshness.Evaluate(hub, count, TimeSpan.Zero, Anchor.AddYears(3));
        Assert.All(later.Windows, w => Assert.Equal(0, w.StillInWindow));
    }

    [Fact]
    public void AScanOfAPopulationWithUndatedItems_WalksTheirFolders_WithDeletedItemsStillLast()
    {
        // Unchanged for the measurement corpus: a read-only scan must not ask a PST for a default
        // folder it lacks, because asking can create it.
        Assert.Equal(new[] { 16, 6, 5, 23, 4, 3 }, ComCorpusMailbox.ScanFolderIdsFor(null));
        Assert.Equal(new[] { 16, 6, 5, 23, 4, 3 }, ComCorpusMailbox.ScanFolderIdsFor(Array.Empty<int>()));

        IReadOnlyList<int> hub = ComCorpusMailbox.ScanFolderIdsFor(Hub().Population!.UndatedFolderIds);
        Assert.Equal(new[] { 16, 6, 5, 23, 4, 9, 10, 13, 3 }, hub);
        Assert.Equal(hub.Count, hub.Distinct().Count());
        Assert.Equal(new[] { 9, 10, 13 }, Bystander().Population!.UndatedFolderIds);
        Assert.Empty(Identity().Population!.UndatedFolderIds);
    }

    [Fact]
    public void TheDraftKind_IsGone_AndItsValueIsNotReused()
    {
        // An unsent draft's first save lands in the DEFAULT store's Drafts (OAI-UNINDEXED, 2026-09-24),
        // so a population - never built into the default store - cannot carry one without writing into
        // another store. Value 1 stays unused, so a manifest line of the old kind can never read as a new one.
        Assert.Equal(new[] { "Mail", "Appointment", "Contact", "Task" }, Enum.GetNames<CorpusItemKind>());
        Assert.False(Enum.IsDefined(typeof(CorpusItemKind), 1));
        Assert.Null(CorpusItemKinds.UndatedKindOfFolder(CorpusCensus.DraftsFolderId));
    }

    [Theory]
    [InlineData(CorpusItemKind.Appointment, 9, 1, "IPM.Appointment")]
    [InlineData(CorpusItemKind.Contact, 10, 2, "IPM.Contact")]
    [InlineData(CorpusItemKind.Task, 13, 3, "IPM.Task")]
    public void EachUndatedKind_HasItsFolder_ItsItemType_AndItsMessageClass(CorpusItemKind kind, int folder, int olType, string messageClass)
    {
        Assert.Equal(folder, CorpusItemKinds.FolderIdOf(kind));
        Assert.Equal(olType, CorpusItemKinds.OlItemTypeOf(kind));
        Assert.Equal(messageClass, CorpusItemKinds.MessageClassOf(kind));
        Assert.Equal(kind, CorpusItemKinds.UndatedKindOfFolder(folder));
        Assert.True(CorpusItemKinds.IsUndated(kind));
    }

    [Fact]
    public void Mail_IsTheOneDatedKind_AndHasNoFolderOfItsOwn()
    {
        Assert.False(CorpusItemKinds.IsUndated(CorpusItemKind.Mail));
        Assert.Throws<ArgumentOutOfRangeException>(() => CorpusItemKinds.FolderIdOf(CorpusItemKind.Mail));
        Assert.Null(CorpusItemKinds.UndatedKindOfFolder(6));
        Assert.Equal(CorpusItemKind.Mail, new CorpusPlan(new CorpusPlanOptions("vm-x", 7777, Anchor)).Describe(1).Kind);
    }

    [Theory]
    [InlineData("IPM.Note", "IPM.Note", true)]
    [InlineData("ipm.note", "IPM.Note", true)]
    [InlineData("IPM.Note.Draft", "IPM.Note", true)]
    [InlineData("IPM.Notes", "IPM.Note", false)]
    [InlineData("IPM.Appointment", "IPM.Note", false)]
    [InlineData(null, "IPM.Task", false)]
    public void AMessageClass_MatchesItsOwnClassAndItsSubclasses_Only(string? actual, string expected, bool matches)
    {
        Assert.Equal(matches, CorpusMessageFlags.ClassMatches(actual, expected));
    }

    // ================================================================ the read-back and the probe

    private static List<CorpusEnrichmentObservation> PerfectReadBack(CorpusPlan plan, Func<int, string?>? conversationId = null)
        => Ordinals(plan).Select(o =>
        {
            CorpusItemEnrichment? e = plan.Enrich(o);
            if (e == null)
            {
                // An undated item reads back as its own kind, with no delivery time.
                CorpusItemKind kind = plan.Describe(o).Kind;
                return new CorpusEnrichmentObservation(
                    o, null, null, null, null, null, null, null,
                    CorpusItemKinds.MessageClassOf(kind),
                    false);
            }

            return new CorpusEnrichmentObservation(
                o,
                e.Sender.Address,
                e.Sender.Name,
                e.Recipients.Select(r => new CorpusObservedRecipient(r.Person.Address, (int)r.Kind)).ToList(),
                e.Attachments.Select(a => a.FileName).ToList(),
                e.ConversationIndexHex,
                conversationId == null
                    ? (e.ThreadKey == null ? null : "CONV" + e.ThreadKey.Value.ToString(CultureInfo.InvariantCulture))
                    : conversationId(o));
        }).ToList();

    [Fact]
    public void AReadBackOfEveryPlannedWrite_IsClean()
    {
        CorpusPlan hub = Hub();
        CorpusEnrichmentReport report = CorpusEnrichmentCheck.Compare(hub, hub.FixedItemCount!.Value, PerfectReadBack(hub));
        (bool clean, string message) = CorpusEnrichmentCheck.Decide(report);
        Assert.True(clean, message);
        Assert.Equal(4, report.Threads);
        Assert.Equal(4, report.ThreadsGroupedByStore);
    }

    [Fact]
    public void AReadBackCatchesAMissingSender_AWrongRecipient_AndALostAttachment()
    {
        CorpusPlan hub = Hub();
        List<CorpusEnrichmentObservation> seen = PerfectReadBack(hub);
        int withAttachment = Dated(hub).First(o => hub.Enrich(o)!.Attachments.Count > 0);
        seen[0] = seen[0] with { SenderAddress = null };
        seen[1] = seen[1] with { Recipients = new[] { new CorpusObservedRecipient("someone@else.invalid", 1) } };
        seen[withAttachment - 1] = seen[withAttachment - 1] with { AttachmentNames = Array.Empty<string>() };

        CorpusEnrichmentReport report = CorpusEnrichmentCheck.Compare(hub, hub.FixedItemCount!.Value, seen);
        Assert.Equal(1, report.SenderMismatches);
        Assert.Equal(1, report.RecipientMismatches);
        Assert.Equal(1, report.AttachmentMismatches);
        (bool clean, string message) = CorpusEnrichmentCheck.Decide(report);
        Assert.False(clean);
        Assert.Contains("SENDER", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AConversationTheStoreSplit_IsAFault_AndOneItCouldNotName_IsReportedNotFailed()
    {
        CorpusPlan hub = Hub();
        int firstThreadMember = Dated(hub).First(o => hub.Enrich(o)!.ThreadKey == 0);

        List<CorpusEnrichmentObservation> split = PerfectReadBack(hub, o =>
            hub.Enrich(o)?.ThreadKey is int t ? (o == firstThreadMember ? "OTHER" : "CONV" + t) : null);
        (bool splitClean, string splitMessage) =
            CorpusEnrichmentCheck.Decide(CorpusEnrichmentCheck.Compare(hub, hub.FixedItemCount!.Value, split));
        Assert.False(splitClean);
        Assert.Contains("DIFFERENT conversation ids", splitMessage, StringComparison.Ordinal);

        List<CorpusEnrichmentObservation> unnamed = PerfectReadBack(hub, _ => null);
        (bool unnamedClean, string unnamedMessage) =
            CorpusEnrichmentCheck.Decide(CorpusEnrichmentCheck.Compare(hub, hub.FixedItemCount!.Value, unnamed));
        Assert.True(unnamedClean);
        Assert.Contains("NOT ESTABLISHED", unnamedMessage, StringComparison.Ordinal);

        List<CorpusEnrichmentObservation> merged = PerfectReadBack(hub, o => hub.Enrich(o)?.ThreadKey == null ? null : "SAME");
        CorpusEnrichmentReport mergedReport = CorpusEnrichmentCheck.Compare(hub, hub.FixedItemCount!.Value, merged);
        Assert.Equal(3, mergedReport.ThreadsSharingAnId);
        Assert.False(CorpusEnrichmentCheck.Decide(mergedReport).Clean);
    }

    [Fact]
    public void TheEnrichmentProbe_RefusesAPopulationUnlessEveryWriteLanded()
    {
        Assert.True(CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, true, true, true, "C", null)).Proceed);

        (bool proceed, string message) =
            CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, true, true, true, null, null));
        Assert.True(proceed);
        Assert.Contains("not established", message, StringComparison.OrdinalIgnoreCase);

        Assert.False(CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(false, true, true, true, "C", null)).Proceed);
        Assert.False(CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, false, true, true, "C", null)).Proceed);
        Assert.False(CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, true, false, true, "C", null)).Proceed);
        Assert.False(CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, true, true, false, "C", null)).Proceed);
        (bool failed, string error) = CorpusEnrichmentFidelity.Decide(new CorpusEnrichmentProbe(true, true, true, true, "C", "boom"));
        Assert.False(failed);
        Assert.Contains("boom", error, StringComparison.Ordinal);
    }

    [Fact]
    public void AReadBack_CatchesAnUndatedItemOfTheWrongKind_ADatedOne_AndOneItCouldNotDate()
    {
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        List<CorpusEnrichmentObservation> seen = PerfectReadBack(hub);
        CorpusEnrichmentReport perfect = CorpusEnrichmentCheck.Compare(hub, count, seen);
        Assert.Equal(12, perfect.UndatedPlanned);
        (bool cleanPerfect, string perfectMessage) = CorpusEnrichmentCheck.Decide(perfect);
        Assert.True(cleanPerfect, perfectMessage);
        Assert.Contains("12 of them UNDATED", perfectMessage, StringComparison.Ordinal);

        int appointment = Undated(hub).First(o => hub.Describe(o).Kind == CorpusItemKind.Appointment);
        int task = Undated(hub).First(o => hub.Describe(o).Kind == CorpusItemKind.Task);
        int contact = Undated(hub).First(o => hub.Describe(o).Kind == CorpusItemKind.Contact);
        seen[appointment - 1] = seen[appointment - 1] with { MessageClass = "IPM.Note" };
        seen[task - 1] = seen[task - 1] with { HasDeliveryTime = true };
        seen[contact - 1] = seen[contact - 1] with { HasDeliveryTime = null };

        CorpusEnrichmentReport report = CorpusEnrichmentCheck.Compare(hub, count, seen);
        Assert.Equal(1, report.UndatedClassMismatches);
        Assert.Equal(1, report.UndatedCarryingADate);
        Assert.Equal(1, report.UndatedDateUnestablished);
        (bool clean, string message) = CorpusEnrichmentCheck.Decide(report);
        Assert.False(clean);
        Assert.Contains("CARRYING a delivery time", message, StringComparison.Ordinal);
        Assert.Contains("message class", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheUndatedProbe_RefusesUnlessEveryKindLandedInTheTargetStoreTaggedAndUndated()
    {
        IReadOnlyList<CorpusItemKind> kinds = CorpusItemKinds.Undated;
        List<CorpusUndatedProbe> good = kinds.Select(k => new CorpusUndatedProbe(k, true, true, true, true, true, true, null)).ToList();
        (bool proceed, string message) = CorpusUndatedFidelity.Decide(kinds, good);
        Assert.True(proceed, message);

        foreach (Func<CorpusUndatedProbe, CorpusUndatedProbe> spoil in new Func<CorpusUndatedProbe, CorpusUndatedProbe>[]
        {
            p => p with { FolderReachable = false },
            p => p with { InTheFolder = false },
            p => p with { SubjectTagParses = false },
            p => p with { HasNoDeliveryTime = false },
            p => p with { ClassMatches = false },
            p => p with { InTheTargetStore = false },
            p => p with { Error = "boom" },
        })
        {
            foreach (CorpusItemKind kind in kinds)
            {
                List<CorpusUndatedProbe> bad = good.Select(p => p.Kind == kind ? spoil(p) : p).ToList();
                (bool refused, string why) = CorpusUndatedFidelity.Decide(kinds, bad);
                Assert.False(refused);
                Assert.Contains(kind.ToString().ToUpperInvariant(), why, StringComparison.Ordinal);
            }
        }

        // A kind the probe never reached is a refusal, not a pass; a population with no undated kind probes nothing.
        Assert.False(CorpusUndatedFidelity.Decide(kinds, good.Skip(1).ToList()).Proceed);
        Assert.True(CorpusUndatedFidelity.Decide(Array.Empty<CorpusItemKind>(), Array.Empty<CorpusUndatedProbe>()).Proceed);
    }

    // ================================================================ is the population in the index yet

    private static List<CorpusIndexedRow> PerfectIndex(CorpusPlan plan)
        => Ordinals(plan).Select(o => new CorpusIndexedRow(o, plan.Describe(o).IsUndated ? null : plan.Describe(o).ReceivedUtc)).ToList();

    [Fact]
    public void IndexCoverage_IsCompleteOnlyWhenEveryOrdinalHasARow()
    {
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        CorpusIndexCoverageReport all = CorpusIndexCoverage.Compare(hub, count, PerfectIndex(hub), null);
        (bool complete, string message) = CorpusIndexCoverage.Decide(all);
        Assert.True(complete, message);
        Assert.Equal(count, all.Indexed);
        Assert.Equal(12, all.UndatedPlanned);
        Assert.Equal(0, all.DatedMismatched);
        Assert.Equal(Anchor.AddMinutes(-1), all.NewestIndexedUtc);
        Assert.Equal(all.NewestPlannedUtc, all.NewestIndexedUtc);

        CorpusIndexCoverageReport partial = CorpusIndexCoverage.Compare(hub, count, PerfectIndex(hub).Where(r => r.Ordinal % 10 != 0), null);
        (bool done, string notYet) = CorpusIndexCoverage.Decide(partial);
        Assert.False(done);
        Assert.Equal(Enumerable.Range(1, count).Where(o => o % 10 == 0), partial.Missing);
        Assert.Contains("NOT YET", notYet, StringComparison.Ordinal);

        // Rows of another corpus id never reach this: the command keeps only subjects that parse as this population's.
        Assert.False(CorpusIndexCoverage.Decide(CorpusIndexCoverage.Compare(hub, count, Array.Empty<CorpusIndexedRow>(), null)).Complete);
    }

    [Fact]
    public void IndexCoverage_ReportsAUtcOffsetShift_AndAnUndatedItemTheIndexDated_WithoutFailingOnThem()
    {
        // The dates are the staleness test's to judge; this only says when the index is ready.
        CorpusPlan hub = Hub();
        int count = hub.FixedItemCount!.Value;
        List<CorpusIndexedRow> shifted = PerfectIndex(hub)
            .Select(r => r.DateReceivedUtc == null ? r : r with { DateReceivedUtc = r.DateReceivedUtc.Value.AddHours(2) })
            .ToList();
        int appointment = Undated(hub).First(o => hub.Describe(o).Kind == CorpusItemKind.Appointment);
        shifted[appointment - 1] = shifted[appointment - 1] with { DateReceivedUtc = Anchor };

        CorpusIndexCoverageReport report = CorpusIndexCoverage.Compare(hub, count, shifted, null);
        Assert.Equal(Dated(hub).Count(), report.DatedMismatched);
        Assert.Equal(7200, report.ModalMismatchSeconds);
        Assert.Equal(1, report.UndatedIndexedWithADate);
        (bool complete, string message) = CorpusIndexCoverage.Decide(report);
        Assert.True(complete, message);
        Assert.Contains("7200 s", message, StringComparison.Ordinal);
        Assert.Contains("UNDATED", message, StringComparison.Ordinal);
    }

    [Fact]
    public void CorpusIndexed_RefusesTheMeasurementCorpus()
    {
        // A population's size is fixed; the corpus has corpus-verify, and 160,000 rows are not a wait.
        CorpusOptions corpus = CorpusOptions.Parse(new[] { "--corpus-id", "vm-x", "--seed", "1", "--anchor", "2026-09-24", "--count", "10" });
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        ArgumentException refused = Assert.Throws<ArgumentException>(() => CorpusCommands.RunIndexed(corpus, output));
        Assert.Contains("--population", refused.Message, StringComparison.Ordinal);

        CorpusOptions waiting = CorpusOptions.Parse(new[] { "--population", "hub", "--store", HubStore, "--corpus-id", "hub-x", "--seed", "1", "--anchor", "2026-09-24", "--wait-seconds", "900" });
        Assert.Equal(900, waiting.WaitSeconds);
    }

    // ================================================================ the attachment bytes

    [Fact]
    public void APopulationPng_IsAValidPngADecoderWouldAccept()
    {
        CorpusAttachment png = Dated(Hub()).SelectMany(o => Hub().Enrich(o)!.Attachments).First(a => a.Kind == CorpusAttachmentKind.Png);
        byte[] bytes = png.Content;
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, bytes[..8]);

        int offset = 8;
        var types = new List<string>();
        byte[]? idat = null;
        while (offset < bytes.Length)
        {
            int length = (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            byte[] typeAndData = bytes[(offset + 4)..(offset + 8 + length)];
            uint crc = (uint)((bytes[offset + 8 + length] << 24) | (bytes[offset + 9 + length] << 16)
                | (bytes[offset + 10 + length] << 8) | bytes[offset + 11 + length]);
            Assert.Equal(CorpusAttachmentContent.Crc32(typeAndData), crc);
            if (type == "IDAT")
            {
                idat = typeAndData[4..];
            }

            types.Add(type);
            offset += 12 + length;
        }

        Assert.Equal(new[] { "IHDR", "IDAT", "IEND" }, types);

        // The pixel data inflates - through the framework's own decoder - to 16 scanlines of 49 bytes.
        using var inflated = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(new MemoryStream(idat!), System.IO.Compression.CompressionMode.Decompress))
        {
            zlib.CopyTo(inflated);
        }

        Assert.Equal(16 * 49, inflated.Length);
    }

    [Fact]
    public void TheChecksums_MatchTheirPublishedTestVectors()
    {
        Assert.Equal(0xCBF43926u, CorpusAttachmentContent.Crc32(Encoding.ASCII.GetBytes("123456789")));
        Assert.Equal(0x11E60398u, CorpusAttachmentContent.Adler32(Encoding.ASCII.GetBytes("Wikipedia")));
    }

    [Fact]
    public void TextAttachments_AreAsciiWithCrlfLineEnds()
    {
        foreach (CorpusAttachment a in Dated(Hub()).SelectMany(o => Hub().Enrich(o)!.Attachments).Where(a => a.Kind != CorpusAttachmentKind.Png))
        {
            string text = a.Text!;
            Assert.EndsWith("\r\n", text, StringComparison.Ordinal);
            Assert.DoesNotContain("\n", text.Replace("\r\n", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);
            if (a.Kind == CorpusAttachmentKind.Calendar)
            {
                Assert.StartsWith("BEGIN:VCALENDAR\r\n", text, StringComparison.Ordinal);
                Assert.EndsWith("END:VCALENDAR\r\n", text, StringComparison.Ordinal);
            }
        }
    }

    // ================================================================ the command line

    [Fact]
    public void ThePopulationOption_ParsesAndNeedsTheStoreItIsFor()
    {
        CorpusOptions options = CorpusOptions.Parse(new[]
        {
            "--population", "Hub", "--corpus-id", "hub-x", "--seed", "1", "--anchor", "2026-09-24",
        });
        Assert.Equal(CorpusPopulationKind.Hub, options.Population);
        ArgumentException noStore = Assert.Throws<ArgumentException>(() => options.ToPlanOptions());
        Assert.Contains("--store", noStore.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() => CorpusOptions.Parse(new[] { "--population", "archive" }));
    }

    [Fact]
    public void APopulationFixesItsOwnCount_AndRefusesADifferentOne()
    {
        string[] args = { "--population", "hub", "--store", HubStore, "--corpus-id", "hub-x", "--seed", "1", "--anchor", "2026-09-24" };
        CorpusOptions options = CorpusOptions.Parse(args);
        var plan = new CorpusPlan(options.ToPlanOptions());
        Assert.Equal(plan.FixedItemCount, CorpusCommands.EffectiveCount(options, plan));

        CorpusOptions exact = CorpusOptions.Parse(args.Concat(new[] { "--count", plan.FixedItemCount!.Value.ToString(CultureInfo.InvariantCulture) }));
        Assert.Equal(plan.FixedItemCount, CorpusCommands.EffectiveCount(exact, plan));

        CorpusOptions wrong = CorpusOptions.Parse(args.Concat(new[] { "--count", "20000" }));
        Assert.Throws<ArgumentException>(() => CorpusCommands.EffectiveCount(wrong, plan));

        CorpusOptions corpus = CorpusOptions.Parse(new[] { "--corpus-id", "vm-x", "--seed", "1", "--anchor", "2026-09-24" });
        Assert.Throws<ArgumentException>(() => CorpusCommands.EffectiveCount(corpus, new CorpusPlan(corpus.ToPlanOptions())));
    }

    [Fact]
    public void ThePlanSheetOfAHubPopulation_NamesTheSettingsValuesTheGuestMustCarry()
    {
        CorpusPlan hub = Hub();
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        CorpusCommands.WriteReport(hub, hub.Report(1, hub.FixedItemCount!.Value), output);
        string sheet = output.ToString();
        Assert.Contains("population            : hub v" + CorpusPopulation.Version, sheet, StringComparison.Ordinal);
        Assert.Contains("probe term            : " + CorpusPopulation.ProbeTerm, sheet, StringComparison.Ordinal);
        Assert.Contains("subjectTerm=" + CorpusPopulation.SubjectOnlyTerm, sheet, StringComparison.Ordinal);
        Assert.Contains("senderFragment=" + CorpusPopulation.SubjectOnlySenderFragment, sheet, StringComparison.Ordinal);
    }

    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
