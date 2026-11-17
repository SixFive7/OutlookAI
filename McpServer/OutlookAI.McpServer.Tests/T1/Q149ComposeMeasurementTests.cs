using System.Reflection;
using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the pure half of the Q149 (d) measurement (<see cref="Q149Cells"/>): the guards every one of its
/// mailbox writes passes, the lifetime classification, the decision rules, the gate that keeps it out of a
/// whole-tier run, the attempt line, the HTML facts, the HTML-only composer it measures, and the layout
/// comparison. The live half runs only on a test guest; this is where CI can prove what it decides.
/// </summary>
public sealed class Q149ComposeMeasurementTests
{
    private const string Marker = "q9abcdef0123456789";
    private const string Hub = "tier@vm.invalid";

    // ------------------------------------------------------------------ the write guards (mailbox rules 1-3)

    [Fact]
    public void ASubjectMustCarryTheTagAndTheMarker_Ordinally()
    {
        Assert.Null(Q149ProbeGuards.RefuseSubject("[OutlookAI-McpTest] " + Marker + " q149 c2", Marker));
        Assert.Null(Q149ProbeGuards.RefuseSubject("RE: [OutlookAI-McpTest] " + Marker + " q149 c4-source", Marker));

        Assert.NotNull(Q149ProbeGuards.RefuseSubject(Marker + " q149 c2", Marker));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject("[OutlookAI-McpTest] q149 c2", Marker));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject("[outlookai-mcptest] " + Marker, Marker));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject("[OutlookAI-McpTest] " + Marker.ToUpperInvariant(), Marker));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject(null, Marker));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject(string.Empty, Marker));

        // A weak marker refuses everything, so a fixture with no marker can never match a real subject.
        Assert.NotNull(Q149ProbeGuards.RefuseSubject("[OutlookAI-McpTest] short", "short"));
        Assert.NotNull(Q149ProbeGuards.RefuseSubject("[OutlookAI-McpTest] ", null));
    }

    [Fact]
    public void AChangeOrDeleteNeedsAnEntryIdThisRunCreated()
    {
        string[] created = { "00000000ABCDEF", "00000000123456" };
        Assert.Null(Q149ProbeGuards.RefuseEntryId("00000000abcdef", created));
        Assert.NotNull(Q149ProbeGuards.RefuseEntryId("00000000FFFFFF", created));
        Assert.NotNull(Q149ProbeGuards.RefuseEntryId(null, created));
        Assert.NotNull(Q149ProbeGuards.RefuseEntryId(" ", created));
        Assert.NotNull(Q149ProbeGuards.RefuseEntryId("00000000ABCDEF", Array.Empty<string>()));
    }

    [Fact]
    public void NothingIsSentToAnyoneButTheHubItself()
    {
        Assert.Null(Q149ProbeGuards.RefuseRecipients(new[] { Hub }, Hub));
        Assert.Null(Q149ProbeGuards.RefuseRecipients(new[] { "TIER@VM.INVALID", " tier@vm.invalid " }, Hub));
        Assert.NotNull(Q149ProbeGuards.RefuseRecipients(new[] { Hub, "someone@example.com" }, Hub));
        Assert.NotNull(Q149ProbeGuards.RefuseRecipients(Array.Empty<string>(), Hub));
        Assert.NotNull(Q149ProbeGuards.RefuseRecipients(null, Hub));
        Assert.NotNull(Q149ProbeGuards.RefuseRecipients(new[] { Hub }, " "));
    }

    [Fact]
    public void TheRunRefusesAWeakMarker_AndTagsEverySubjectItMakes()
    {
        Assert.Throws<ArgumentException>(() => new Q149Run(new LiveTestSettings { TestHubStoreDisplayName = Hub }, "short", _ => { }));
        Q149Run run = new(new LiveTestSettings { TestHubStoreDisplayName = Hub }, Marker, _ => { });
        Assert.Null(Q149ProbeGuards.RefuseSubject(run.TaggedSubject("anything"), run.Marker));
        run.Allow("ABC");
        run.Allow(null);
        Assert.Equal(new[] { "ABC" }, run.Created);
    }

    // ------------------------------------------------------------------ the lifetime classification

    [Theory]
    [InlineData(true, null, false, Q149Lifetime.Survived)]
    [InlineData(true, null, true, Q149Lifetime.Parked)]
    [InlineData(false, 0, true, Q149Lifetime.OrderlyExit)]
    [InlineData(false, 0, false, Q149Lifetime.SilentExit)]
    [InlineData(false, unchecked((int)0xC0000005), true, Q149Lifetime.Crash)]
    [InlineData(false, unchecked((int)0xC0000409), false, Q149Lifetime.Crash)]
    [InlineData(false, null, false, Q149Lifetime.Unknown)]
    public void TheLifetimeIsClassifiedFromAliveExitCodeAndQuit(bool alive, int? exitCode, bool quit, Q149Lifetime expected)
    {
        Assert.Equal(expected, Q149LifetimeRules.Classify(alive, exitCode, quit));
    }

    [Fact]
    public void OnlyTheExitClassesCountAsExits()
    {
        Assert.True(Q149LifetimeRules.Exited(Q149Lifetime.OrderlyExit));
        Assert.True(Q149LifetimeRules.Exited(Q149Lifetime.SilentExit));
        Assert.True(Q149LifetimeRules.Exited(Q149Lifetime.Crash));
        Assert.False(Q149LifetimeRules.Exited(Q149Lifetime.Survived));
        Assert.False(Q149LifetimeRules.Exited(Q149Lifetime.Parked));
        Assert.False(Q149LifetimeRules.Exited(Q149Lifetime.Unknown));
        Assert.False(Q149LifetimeRules.IsExceptionCode(1));
        Assert.True(Q149LifetimeRules.IsExceptionCode(unchecked((int)0xC0000005)));
    }

    [Fact]
    public void ThePresenceSamplesAreThePlansOffsets_DerivedFromTheExitTime()
    {
        IReadOnlyList<int> offsets = Q149LifetimeRules.SampleOffsets(30_000);
        Assert.Equal(0, offsets[0]);
        Assert.Equal(50, offsets[1]);
        Assert.Contains(2000, offsets);
        Assert.Contains(5000, offsets);
        Assert.Contains(30_000, offsets);
        Assert.DoesNotContain(60_000, offsets);
        Assert.Equal(41 + 3, offsets.Count);

        IReadOnlyList<int[]> samples = Q149LifetimeRules.Samples(new[] { 0, 50, 100, 5000 }, exitedAtMs: 75);
        Assert.Equal(new[] { 1, 1, 0, 0 }, samples.Select(s => s[1]).ToArray());
        Assert.All(Q149LifetimeRules.Samples(offsets, exitedAtMs: null), s => Assert.Equal(1, s[1]));
    }

    // ------------------------------------------------------------------ the decision rules

    [Fact]
    public void D1_TheDecisiveCellDecidesThePin()
    {
        Assert.Contains("pin is required", Q149Decisions.Lifetime(20, 20), StringComparison.Ordinal);
        Assert.Contains("pin is required", Q149Decisions.Lifetime(18, 20), StringComparison.Ordinal);
        Assert.Contains("timing-dependent", Q149Decisions.Lifetime(17, 20), StringComparison.Ordinal);
        Assert.Contains("timing-dependent", Q149Decisions.Lifetime(3, 20), StringComparison.Ordinal);
        Assert.Contains("pin is unnecessary", Q149Decisions.Lifetime(2, 20), StringComparison.Ordinal);
        Assert.Contains("pin is unnecessary", Q149Decisions.Lifetime(0, 20), StringComparison.Ordinal);
        Assert.Equal("not measured", Q149Decisions.Lifetime(0, 0));
    }

    [Fact]
    public void D2_TheCheapestWayToTheEditorWins_InThePlansOrder()
    {
        Assert.StartsWith("D2: a bounded re-read", Q149Decisions.Availability((10, 10, 0), (0, 3, 0), (0, 10), (0, 10), (0, 9)), StringComparison.Ordinal);

        // A re-read that only worked because a window was shown is not window-free.
        Assert.DoesNotContain("bounded re-read", Q149Decisions.Availability((10, 10, 2), (0, 3, 0), (0, 10), (0, 10), (0, 9)), StringComparison.Ordinal);
        Assert.StartsWith("D2: a Windows-level show", Q149Decisions.Availability((0, 10, 0), (0, 3, 0), (9, 10), (0, 10), (9, 9)), StringComparison.Ordinal);
        Assert.StartsWith("D2: a Windows-level show", Q149Decisions.Availability((0, 10, 0), (0, 3, 0), (0, 10), (10, 10), (0, 9)), StringComparison.Ordinal);
        Assert.StartsWith("D2: only the product's promotion", Q149Decisions.Availability((0, 10, 0), (0, 3, 0), (0, 10), (0, 10), (27, 27)), StringComparison.Ordinal);
        Assert.StartsWith("D2: nothing short", Q149Decisions.Availability((0, 10, 0), (0, 3, 0), (8, 10), (0, 10), (0, 27)), StringComparison.Ordinal);
    }

    [Fact]
    public void D3_ACommitVariantMustCommitEveryDraftWithNoSideEffect()
    {
        Assert.Contains("B1", Q149Decisions.Commit(new[] { ("B1", 20, 20, 0), ("B2", 0, 20, 0) }), StringComparison.Ordinal);
        Assert.Contains("keep the Activate", Q149Decisions.Commit(new[] { ("B1", 19, 20, 0), ("B3", 20, 20, 1) }), StringComparison.Ordinal);
        Assert.Equal("not measured", Q149Decisions.Commit(Array.Empty<(string, int, int, int)>()));
    }

    [Fact]
    public void D4_TheHtmlOnlyDraftReplacesWordOnlyWhenEveryAspectIsTheSame()
    {
        Assert.Contains("(c) can replace Word", Q149Decisions.Parity(new[] { ("a", true), ("b", true) }), StringComparison.Ordinal);
        string verdict = Q149Decisions.Parity(new[] { ("stylesheet", true), ("looks", false) });
        Assert.Contains("looks", verdict, StringComparison.Ordinal);
        Assert.EndsWith("(b)", verdict, StringComparison.Ordinal);
        Assert.Equal("not measured", Q149Decisions.Parity(Array.Empty<(string, bool)>()));
    }

    // ------------------------------------------------------------------ the gate

    [Fact]
    public void TheMeasurementRunsOnlyInARunThatSelectedItAlone()
    {
        Assert.True(Q149RunGate.ShouldRun(new[] { LiveCollections.ComposeSurface }));
        Assert.False(Q149RunGate.ShouldRun(new[] { LiveCollections.ComposeSurface, LiveCollections.Phase4 }));
        Assert.False(Q149RunGate.ShouldRun(LiveCollections.All.ToList()));
        Assert.False(Q149RunGate.ShouldRun(Array.Empty<string>()));
        Assert.False(Q149RunGate.ShouldRun(null));
        Assert.Contains("NOT RUN", Q149RunGate.NotRunLine, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveQ149", Q149RunGate.NotRunLine, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryQ149LiveTest_AsksTheGateFirst_AndSitsInItsCollection()
    {
        List<Type> classes = typeof(Q149Cells).Assembly.GetTypes()
            .Where(t => t.Namespace == "OutlookAI.McpServer.Tests.T2" && t.Name.StartsWith("LiveQ149", StringComparison.Ordinal) && t.Name.EndsWith("Tests", StringComparison.Ordinal))
            .ToList();
        Assert.Equal(3, classes.Count);

        foreach (Type type in classes)
        {
            CollectionAttribute? collection = type.GetCustomAttribute<CollectionAttribute>();
            Assert.NotNull(collection);
            string source = File.ReadAllText(Path.Combine(TestProjectDir(), "T2", type.Name + ".cs"));
            foreach (MethodInfo method in type.GetMethods().Where(m => m.GetCustomAttribute<FactAttribute>() != null))
            {
                int at = source.IndexOf("public void " + method.Name + "(", StringComparison.Ordinal);
                Assert.True(at >= 0, type.Name + "." + method.Name + " not found in its source");
                int gate = source.IndexOf("if (!_fixture.Begin(_output))", at, StringComparison.Ordinal);
                int next = source.IndexOf("public void ", at + 1, StringComparison.Ordinal);
                Assert.True(gate > at && (next < 0 || gate < next), type.Name + "." + method.Name + " does not ask the gate first");
            }
        }
    }

    // ------------------------------------------------------------------ the attempt line

    [Fact]
    public void AnAttemptIsOnePrefixedJsonLine_WithSortedKeys_AndRoundTrips()
    {
        Q149Attempt attempt = new Q149Attempt("S1-L1-shipped-compose-no-pin", 3)
            .Set("zeta", 1)
            .Set("alpha", "x")
            .Set("gone", "y")
            .Set("gone", null)
            .Set("samples", new[] { new[] { 0, 1 }, new[] { 50, 0 } });
        string line = attempt.ToLine();
        Assert.StartsWith(Q149Cells.LinePrefix, line, StringComparison.Ordinal);
        Assert.DoesNotContain("\n", line, StringComparison.Ordinal);
        Assert.True(line.IndexOf("\"alpha\"", StringComparison.Ordinal) < line.IndexOf("\"zeta\"", StringComparison.Ordinal));
        Assert.DoesNotContain("gone", line, StringComparison.Ordinal);

        IReadOnlyDictionary<string, JsonElement>? parsed = Q149Attempt.TryParse("[xUnit.net 00:01:02]     " + line);
        Assert.NotNull(parsed);
        Assert.Equal("S1-L1-shipped-compose-no-pin", parsed!["cell"].GetString());
        Assert.Equal(3, parsed["iteration"].GetInt32());
        Assert.Equal(50, parsed["samples"][1][0].GetInt32());
        Assert.Null(Q149Attempt.TryParse("no line here"));
        Assert.Throws<ArgumentException>(() => new Q149Attempt(" ", 1));
    }

    // ------------------------------------------------------------------ the HTML facts

    [Fact]
    public void TheHtmlFactsFindTheMarkersTheStylesheetAndThePictures()
    {
        string html = "<html><head><style><!--\n p.MsoNormal {margin:0cm;}\n --></style></head><body><div class=WordSection1>"
            + "<p class=MsoNormal>BODYTOKEN</p><p class=MsoNormal><a name=\"_MailAutoSig\">sig</a></p>"
            + "<p><img src=\"cid:image001.png@x\"><img src=\"file:///C:/x.png\"></p>"
            + "<div style='border:none;border-top:solid #E1E1E1 1.0pt'>header</div><v:shape></v:shape></div></body></html>";
        Q149HtmlFacts facts = Q149Html.Analyze(html, new Dictionary<string, string> { ["body"] = "BODYTOKEN", ["none"] = "ABSENT" });
        Assert.True(facts.WordSectionAt > 0);
        Assert.Equal(2, facts.MsoNormal);
        Assert.Equal(1, facts.StyleBlocks);
        Assert.NotNull(facts.StylesheetHash);
        Assert.True(facts.MailAutoSigAt > facts.TokenAt["body"]);
        Assert.Equal(-1, facts.MailOriginalAt);
        Assert.True(facts.HeaderBlockAt > facts.MailAutoSigAt);
        Assert.Equal(2, facts.Images);
        Assert.Equal(1, facts.CidImages);
        Assert.Equal(1, facts.FileLinks);
        Assert.Equal(1, facts.VmlShapes);
        Assert.Equal(-1, facts.TokenAt["none"]);
        Assert.True(facts.ToFields().ContainsKey("at_body"));

        // The stylesheet hash ignores comment markers and whitespace, and nothing else.
        Assert.Equal(
            Q149Html.Analyze("<style>p.MsoNormal {margin:0cm;}</style>").StylesheetHash,
            Q149Html.Analyze("<style>\n<!--\np.MsoNormal   {margin:0cm;}\n-->\n</style>").StylesheetHash);
        Assert.NotEqual(
            Q149Html.Analyze("<style>p.MsoNormal {margin:0cm;}</style>").StylesheetHash,
            Q149Html.Analyze("<style>p.MsoNormal {margin:1cm;}</style>").StylesheetHash);
        Assert.Null(Q149Html.Analyze("<html><body>stub</body></html>").StylesheetHash);
        Assert.Equal(0, Q149Html.Analyze(null).Length);
    }

    // ------------------------------------------------------------------ the HTML-only composer

    private const string Reference =
        "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"><head><meta charset=utf-8><style>p.MsoNormal {margin:0cm;}</style></head>"
        + "<body lang=NL link=\"#0563C1\"><div class=WordSection1><p class=MsoNormal>old</p></div></body></html>";

    [Fact]
    public void TheFrameIsTheReferencesStartTagHeadAndBodyStart()
    {
        (string HtmlStart, string Head, string BodyStart)? frame = HtmlOnlyDraftComposer.ReadFrame(Reference);
        Assert.NotNull(frame);
        Assert.StartsWith("<html xmlns:o=", frame!.Value.HtmlStart, StringComparison.Ordinal);
        Assert.Contains("p.MsoNormal", frame.Value.Head, StringComparison.Ordinal);
        Assert.Equal("<body lang=NL link=\"#0563C1\">", frame.Value.BodyStart);
        Assert.Null(HtmlOnlyDraftComposer.ReadFrame("<html><body>no head</body></html>"));
        Assert.Null(HtmlOnlyDraftComposer.ReadFrame(null));
    }

    [Fact]
    public void TextBecomesWordsParagraphs_Escaped_WithLineBreaksAndBlankParagraphs()
    {
        string html = HtmlOnlyDraftComposer.ParagraphsFromText("a < b\nline two\n\nsecond");
        Assert.Contains("<p class=MsoNormal>a &lt; b<br>", html, StringComparison.Ordinal);
        Assert.Contains("line two<o:p></o:p></p>", html, StringComparison.Ordinal);
        Assert.Contains("<p class=MsoNormal>second<o:p></o:p></p>", html, StringComparison.Ordinal);
        Assert.Equal(HtmlOnlyDraftComposer.BlankParagraph, HtmlOnlyDraftComposer.ParagraphsFromText("one\n\n\n\ntwo").Split('\n')[1]);

        Assert.Equal("<p class=MsoNormal>x</p><p class=Other>y</p><h1>z</h1>", HtmlOnlyDraftComposer.ParagraphsFromHtml("<p>x</p><p class=Other>y</p><h1>z</h1>"));
    }

    [Fact]
    public void TheSignatureIsMarkedAndItsPicturesPointAtTheirAttachments()
    {
        string signature = "<html><head></head><body><p>Groet,</p><p><img src=\"Sig_files/sigimg.png\" alt=\"logo\"></p></body></html>";
        string block = HtmlOnlyDraftComposer.SignatureBlock(signature, src => src.EndsWith("sigimg.png", StringComparison.Ordinal) ? "image001.png@x" : null);
        Assert.StartsWith("<a name=\"_MailAutoSig\"></a>", block, StringComparison.Ordinal);
        Assert.Contains("src=\"cid:image001.png@x\"", block, StringComparison.Ordinal);
        Assert.DoesNotContain("Sig_files", block, StringComparison.Ordinal);
        Assert.Contains("<p class=MsoNormal>Groet,</p>", block, StringComparison.Ordinal);

        // A picture the caller does not map is left alone rather than pointed at nothing.
        Assert.Contains("src=\"other.png\"", HtmlOnlyDraftComposer.SignatureBlock("<body><img src=\"other.png\"></body>", _ => null), StringComparison.Ordinal);
    }

    [Fact]
    public void ANewMailIsTheFrame_ThenTheBody_ABlankLine_AndTheSignature_InsideWordSection1()
    {
        (string, string, string) frame = HtmlOnlyDraftComposer.ReadFrame(Reference)!.Value;
        string html = HtmlOnlyDraftComposer.ComposeNew(frame, "<p class=MsoNormal>BODY</p>", "<a name=\"_MailAutoSig\"></a><p>SIG</p>");
        Q149HtmlFacts facts = Q149Html.Analyze(html, new Dictionary<string, string> { ["body"] = "BODY", ["sig"] = "SIG" });
        Assert.True(facts.WordSectionAt >= 0);
        Assert.True(facts.TokenAt["body"] > facts.WordSectionAt);
        Assert.True(facts.MailAutoSigAt > facts.TokenAt["body"]);
        Assert.True(facts.TokenAt["sig"] > facts.MailAutoSigAt);
        Assert.Equal(Q149Html.Analyze(Reference).StylesheetHash, facts.StylesheetHash);
        Assert.EndsWith("</html>\n", html, StringComparison.Ordinal);
    }

    [Fact]
    public void AReplysBodyGoesAboveTheQuotedOriginal()
    {
        string reply = "<html><head></head><body><div class=WordSection1><p class=MsoNormal><o:p>&nbsp;</o:p></p>"
            + "<div style='border:none;border-top:solid #E1E1E1 1.0pt'><p><b>From:</b> x</p></div><p>QUOTE</p></div></body></html>";
        string? placed = HtmlOnlyDraftComposer.InsertAboveQuote(reply, "<p class=MsoNormal>BODY</p>", "<a name=\"_MailAutoSig\"></a><p>SIG</p>");
        Assert.NotNull(placed);
        Q149HtmlFacts facts = Q149Html.Analyze(placed, new Dictionary<string, string> { ["body"] = "BODY", ["sig"] = "SIG", ["quote"] = "QUOTE" });
        Assert.True(facts.TokenAt["body"] > facts.WordSectionAt);
        Assert.True(facts.TokenAt["sig"] > facts.TokenAt["body"]);
        Assert.True(facts.HeaderBlockAt > facts.TokenAt["sig"]);
        Assert.True(facts.TokenAt["quote"] > facts.HeaderBlockAt);
        Assert.Null(HtmlOnlyDraftComposer.InsertAboveQuote("<html><body>no section</body></html>", "x", null));
    }

    [Fact]
    public void AnUpdateReplacesEverythingAboveTheBoundary_AndKeepsTheBoundaryAndBelow()
    {
        string draft = "<html><head></head><body><div class=WordSection1><p class=MsoNormal>OLD1</p><p class=MsoNormal>OLD2</p>"
            + "<p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><a name=\"_MailAutoSig\">Groet</a></p><p>SIG</p></div></body></html>";
        string? updated = HtmlOnlyDraftComposer.ReplaceDraftRegion(draft, "<p class=MsoNormal>NEW</p>");
        Assert.NotNull(updated);
        Assert.DoesNotContain("OLD1", updated!, StringComparison.Ordinal);
        Assert.DoesNotContain("OLD2", updated, StringComparison.Ordinal);
        Assert.Contains("<p class=MsoNormal><a name=\"_MailAutoSig\">Groet</a></p><p>SIG</p>", updated, StringComparison.Ordinal);
        Assert.True(updated.IndexOf("NEW", StringComparison.Ordinal) < updated.IndexOf("_MailAutoSig", StringComparison.Ordinal));

        // With the quote as the only boundary, the header block and the quote survive.
        string reply = "<div class=WordSection1><p>OLD</p><div style='border:none;border-top:solid #E1E1E1 1.0pt'><p>From</p></div><p>QUOTE</p></div>";
        string? replyUpdated = HtmlOnlyDraftComposer.ReplaceDraftRegion(reply, "<p>NEW</p>");
        Assert.DoesNotContain("OLD", replyUpdated!, StringComparison.Ordinal);
        Assert.Contains("QUOTE", replyUpdated, StringComparison.Ordinal);
        Assert.Contains("border-top:solid #E1E1E1", replyUpdated, StringComparison.Ordinal);

        // No boundary: the whole section is the draft.
        string? whole = HtmlOnlyDraftComposer.ReplaceDraftRegion("<div class=WordSection1><p>OLD</p></div>", "<p>NEW</p>");
        Assert.Equal("<div class=WordSection1>\n<p>NEW</p></div>", whole);
        Assert.Null(HtmlOnlyDraftComposer.ReplaceDraftRegion("<p>no section</p>", "<p>x</p>"));
    }

    // ------------------------------------------------------------------ the layout comparison

    [Fact]
    public void TextIsComparedAsAReaderSeesIt()
    {
        Assert.Equal("a b c", Q149WordLook.NormalizeText(" a\u00A0 b\r\vc\a "));
        Assert.Equal(string.Empty, Q149WordLook.NormalizeText("\r"));
        Assert.Equal(string.Empty, Q149WordLook.NormalizeText(null));
    }

    [Fact]
    public void TwoLayoutsDifferInEveryPropertyAReaderCanSee_AndOnlyThere()
    {
        WordParagraphLook a = new("Hallo", "Normal", "Aptos", 11, 0, 0, -16777216, 0, 0, 0, 12, 0, 0, 0, 0);
        Assert.Equal(0, Q149WordLook.Compare(new[] { a }, new[] { a with { } }).Count);
        Assert.Equal(0, Q149WordLook.Compare(new[] { a }, new[] { a with { Size = 11.04 } }).Count);

        (int count, IReadOnlyList<string> listed) = Q149WordLook.Compare(
            new[] { a, a },
            new[] { a with { Style = "Normal (Web)", SpaceBefore = 14, SpaceAfter = 14 }, a with { Text = "Hallo!" } });
        Assert.Equal(4, count);
        Assert.Contains(listed, d => d.StartsWith("p1 style", StringComparison.Ordinal));
        Assert.Contains(listed, d => d.StartsWith("p2 text", StringComparison.Ordinal));
        Assert.DoesNotContain(listed, d => d.Contains("Hallo", StringComparison.Ordinal));

        Assert.Equal(1, Q149WordLook.Compare(new[] { a }, new[] { a, a }, max: 0).Count);
        Assert.Empty(Q149WordLook.Compare(new[] { a }, new[] { a, a }, max: 0).Listed);
    }

    private static string TestProjectDir()
    {
        return typeof(Q149ComposeMeasurementTests).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .First(a => a.Key == "TestProjectDir").Value!;
    }
}
