using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <c>Testbed/live-test-settings.template.json</c> - the committed, tokens-only template a
/// test guest's gitignored live-test settings are rendered from - against the live tier's own
/// loader, so the template cannot drift away from what that loader accepts.
/// <para>
/// <b>How the pieces fit.</b> <c>Testbed/host/New-LiveTestSettings.ps1</c> renders the template for
/// one guest from that guest's section of <c>Testbed/testbed.json</c> and refuses anything the tier
/// would refuse. <c>.github/scripts/check-testbed-references.ps1</c> check 8 holds the template's
/// FIELDS to the documented example's and runs on every pull request. Neither of those can say
/// whether the rendered shape still LOADS: the renderer re-states the loader's rules in PowerShell,
/// and a re-statement drifts. This class renders the template with the renderer's own rule - every
/// value is a double-brace token spelling its own JSON path, replaced by the value at that path, and
/// a block whose value is null is left out - and hands the result to
/// <see cref="LiveTestSettings.Parse"/>, then to the admission gate
/// <see cref="LiveStoreCountTripwire.EnsureBaseline"/> applies before its first COM call
/// (<see cref="TripwireWatchSoundness.Require"/>, with the same three arguments). It runs whenever the
/// loader changes, which is exactly when the template could stop loading.
/// </para>
/// <para>
/// <b>The controls are the point.</b> A test that only ever sees a good file passes whether or not
/// anything is checked. So each refusal the tier makes on a rendered file is driven here too, from
/// the same render: the unrendered template itself, a file declaring its bystanders only in the
/// census list, the hub declared a bystander, a half-written block, and a Production profile the
/// template has no fields for.
/// </para>
/// <para>
/// <b>One documented rule the loader does NOT enforce, recorded here because it surprised the work
/// that wrote this class.</b> <c>Docs/live-tier-on-the-vm.md</c> sections 1.3, 2.6 and 2.10 say a
/// store named in only one of <c>expectedStoreDisplayNames</c> and <c>bystanderStoreDisplayNames</c>
/// refuses the tier. Read against the code on 2026-09-24, that holds in one direction and only
/// conditionally: a store left OUT of the bystander list is inside the identity-draft grant, and is
/// refused only when that leaves the census nothing it could fail on - which is the control below.
/// A bystander left out of the census list is NOT refused at all:
/// <see cref="LiveStoreCountTripwire.WatchedStores"/> adds declared bystanders back in, deliberately
/// (<c>T1/TripwireBystanderStoreTests.TheCensusWatchesEveryDeclaredBystanderEvenOneNoOtherListNames</c>).
/// The renderer refuses that second shape itself. Which of the documents or the loader should change
/// is an open question and nothing here decides it.
/// </para>
/// <para>
/// Synthetic names only - RFC 2606 <c>.invalid</c> addresses and made-up store names. The one real
/// input is the committed <c>Testbed/testbed.json</c>, whose guest sections name a test guest's
/// synthetic stores and nothing else. No machine-local settings file is read, and nothing touches
/// Outlook or a mailbox.
/// </para>
/// </summary>
public sealed class LiveTestSettingsTemplateTests
{
    private const string Hub = "hub@render.invalid";
    private const string CorpusStore = "Synthetic Corpus";
    private const string Bystander = "Synthetic Bystander";
    private const string Identity = "identity@render.invalid";

    /// <summary>Every placeholder in testbed.json begins with this; the renderer refuses a section holding one.</summary>
    private const string PlaceholderMarker = "<FILL";

    private const string NoFailableStore = "NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE";

    private readonly ITestOutputHelper _output;

    public LiveTestSettingsTemplateTests(ITestOutputHelper output)
    {
        _output = output;
    }

    // ------------------------------------------------------------------------------ plumbing

    /// <summary>The repository root, located from the same assembly metadata the loader uses.</summary>
    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }

    private static string TemplateText()
    {
        string path = Path.Combine(RepoRoot(), "Testbed", "live-test-settings.template.json");
        Assert.True(File.Exists(path), "the committed settings template is missing: " + path);
        return File.ReadAllText(path);
    }

    /// <summary>
    /// What a guest's section of testbed.json holds, for a four-store guest: hub, corpus, plain
    /// bystander, and an identity account left out of the bystander list on purpose
    /// (Docs/live-tier-on-the-vm.md section 2.8b).
    /// </summary>
    private static JsonObject SyntheticValues()
    {
        return JsonNode.Parse(
            """
            {
              "_note": "comment keys are not settings",
              "machineProfile": "Portable",
              "testHubStoreDisplayName": "hub@render.invalid",
              "expectedStoreDisplayNames": [ "hub@render.invalid", "Synthetic Corpus", "Synthetic Bystander", "identity@render.invalid" ],
              "expectedDelegateStoreDisplayNames": [],
              "bystanderStoreDisplayNames": [ "Synthetic Bystander", "Synthetic Corpus" ],
              "corpus": {
                "storeDisplayName": "Synthetic Corpus",
                "manifestPath": "C:\\OutlookAI-Q5\\corpus-vm-synthetic.jsonl",
                "corpusId": "vm-synthetic",
                "seed": 4242,
                "anchorUtc": "2026-08-01T00:00:00Z",
                "itemCount": 1000,
                "windowDays": [ 7, 30, 60 ]
              },
              "mailSink": {
                "submitHost": "127.0.0.1",
                "submitPort": 2525,
                "retrieveHost": "127.0.0.1",
                "retrievePort": 1110,
                "connectTimeoutMs": 1500
              }
            }
            """)!.AsObject();
    }

    /// <summary>
    /// The renderer's substitution rule, restated: a token is replaced, quotes included, by the JSON
    /// value at the path it spells; a block whose value is null is left out; notes are not carried.
    /// It asserts the rule's one precondition as it goes - that every template value IS the token
    /// for its own path - because a token in the wrong slot would otherwise render a plausible file
    /// with a value in the wrong field.
    /// </summary>
    private static string Render(string templateJson, JsonObject values)
    {
        JsonObject template = JsonNode.Parse(templateJson)!.AsObject();
        return RenderObject(template, values, string.Empty).ToJsonString();
    }

    private static JsonObject RenderObject(JsonObject template, JsonObject values, string prefix)
    {
        JsonObject rendered = new();
        foreach (KeyValuePair<string, JsonNode?> entry in template)
        {
            if (entry.Key.StartsWith('_'))
            {
                continue;
            }

            string path = prefix.Length == 0 ? entry.Key : prefix + "." + entry.Key;
            Assert.True(values.TryGetPropertyValue(entry.Key, out JsonNode? value), "no value for " + path);

            if (entry.Value is JsonObject block)
            {
                if (value is null)
                {
                    continue;
                }

                rendered[entry.Key] = RenderObject(block, value.AsObject(), path);
                continue;
            }

            Assert.Equal("{{" + path + "}}", entry.Value!.GetValue<string>());
            rendered[entry.Key] = value?.DeepClone();
        }

        return rendered;
    }

    private static LiveTestSettings Load(JsonObject values)
    {
        return LiveTestSettings.Parse(Render(TemplateText(), values));
    }

    /// <summary>The gate EnsureBaseline applies before any COM call, with the arguments it uses.</summary>
    private static TripwireWatchReport Admit(LiveTestSettings settings)
    {
        return TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);
    }

    private static bool HoldsPlaceholder(JsonNode? node)
    {
        return node switch
        {
            JsonObject obj => obj.Any(p => !p.Key.StartsWith('_') && HoldsPlaceholder(p.Value)),
            JsonArray array => array.Any(HoldsPlaceholder),
            JsonValue value => value.TryGetValue(out string? text)
                && text.StartsWith(PlaceholderMarker, StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    // ------------------------------------------------------------------------ the good path

    [Fact]
    public void EveryTemplateValueIsTheTokenForItsOwnPath_AndNoNoteCarriesOne()
    {
        JsonObject template = JsonNode.Parse(TemplateText())!.AsObject();
        List<string> fields = new();
        foreach (KeyValuePair<string, JsonNode?> entry in template)
        {
            if (entry.Key.StartsWith('_'))
            {
                // A note that spelled a token would look like an unreplaced one to anything scanning
                // for them, so the notes describe tokens in words.
                Assert.DoesNotContain("{{", entry.Value!.GetValue<string>(), StringComparison.Ordinal);
                continue;
            }

            if (entry.Value is JsonObject block)
            {
                foreach (KeyValuePair<string, JsonNode?> inner in block)
                {
                    string path = entry.Key + "." + inner.Key;
                    Assert.Equal("{{" + path + "}}", inner.Value!.GetValue<string>());
                    fields.Add(path);
                }

                continue;
            }

            Assert.Equal("{{" + entry.Key + "}}", entry.Value!.GetValue<string>());
            fields.Add(entry.Key);
        }

        Assert.Equal(17, fields.Count);
    }

    [Fact]
    public void TheTemplateRenderedWithSyntheticValues_LoadsThroughTheLiveTiersOwnLoader()
    {
        LiveTestSettings settings = Load(SyntheticValues());

        Assert.Equal(LiveMachineProfile.Portable, settings.MachineProfile);
        Assert.Equal(Hub, settings.TestHubStoreDisplayName);
        Assert.Equal(new[] { Hub, CorpusStore, Bystander, Identity }, settings.ExpectedStoreDisplayNames);
        Assert.Empty(settings.ExpectedDelegateStoreDisplayNames);
        Assert.Equal(new[] { Bystander, CorpusStore }, settings.BystanderStoreDisplayNames);

        Assert.NotNull(settings.Corpus);
        Assert.True(settings.Corpus!.IsComplete);
        Assert.Equal(CorpusStore, settings.Corpus.StoreDisplayName);
        Assert.Equal("vm-synthetic", settings.Corpus.CorpusId);
        Assert.Equal(4242, settings.Corpus.Seed);
        Assert.Equal("2026-08-01T00:00:00Z", settings.Corpus.AnchorUtc);
        Assert.Equal(1000, settings.Corpus.ItemCount);
        Assert.Equal(new[] { 7, 30, 60 }, settings.Corpus.WindowDays);

        Assert.NotNull(settings.MailSink);
        Assert.True(settings.MailSink!.IsComplete);
        Assert.Equal(2525, settings.MailSink.SubmitPort);
        Assert.Equal(1110, settings.MailSink.RetrievePort);
        Assert.Equal(1500, settings.MailSink.ConnectTimeoutMs);

        Assert.Contains("machineProfile=Portable", settings.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void AndTheTierWouldAdmitIt_WithEveryDeclaredStorePolicedAndOnlyTheIdentityStoreWritable()
    {
        LiveTestSettings settings = Load(SyntheticValues());

        TripwireWatchReport report = Admit(settings);
        Assert.True(report.Usable);
        Assert.Equal(new[] { CorpusStore, Bystander }, report.Policed);
        Assert.Equal(new[] { Identity }, report.Writable);

        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);
        foreach (StoreWriteKind kind in Enum.GetValues<StoreWriteKind>())
        {
            Assert.True(allowlist.IsAllowed(Hub, kind));
            Assert.False(allowlist.IsAllowed(CorpusStore, kind));
            Assert.False(allowlist.IsAllowed(Bystander, kind));
        }

        Assert.Equal(new[] { Identity }, allowlist.IdentityAccountsAmong(settings.ExpectedStoreDisplayNames));
    }

    [Fact]
    public void WithBothOptionalBlocksDeclaredAbsent_ItLeavesThemOut_AndStillLoadsAndIsAdmitted()
    {
        // The shape both guests are expected to render to while no corpus is built and no sink is
        // installed (Docs/live-tier-on-the-vm.md section 1.4). Left OUT, not written as null: absent
        // is the documented shape, and the loader's default path.
        JsonObject values = SyntheticValues();
        values["corpus"] = null;
        values["mailSink"] = null;

        JsonObject rendered = JsonNode.Parse(Render(TemplateText(), values))!.AsObject();
        Assert.False(rendered.ContainsKey("corpus"));
        Assert.False(rendered.ContainsKey("mailSink"));

        LiveTestSettings settings = LiveTestSettings.Parse(rendered.ToJsonString());
        Assert.Null(settings.Corpus);
        Assert.Null(settings.MailSink);
        Assert.True(Admit(settings).Usable);
    }

    // -------------------------------------------------------------- the controls: it can fail

    [Fact]
    public void TheUnrenderedTemplateItself_IsRefusedByTheLoader()
    {
        // The template must never be mistaken for a settings file: copied into live-fixtures as it
        // stands, it has to stop the tier rather than load as something.
        Assert.ThrowsAny<JsonException>(() => LiveTestSettings.Parse(TemplateText()));
    }

    [Fact]
    public void StoresDeclaredOnlyInTheCensusList_LeaveNothingPoliced_AndTheTierRefuses()
    {
        // A bystander declared in only one of the two lists - the census list - is not a bystander:
        // it is inside the identity-draft grant. With every would-be bystander in that state the
        // census has nothing it could fail on, and the tier refuses. The file itself LOADS, which is
        // asserted first, so this is the admission gate refusing and not a malformed document.
        JsonObject values = SyntheticValues();
        values["bystanderStoreDisplayNames"] = new JsonArray();

        LiveTestSettings settings = Load(values);
        Assert.Empty(settings.BystanderStoreDisplayNames);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => Admit(settings));
        Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
        Assert.Contains(NoFailableStore, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubDeclaredABystander_IsRefusedByTheTier()
    {
        JsonObject values = SyntheticValues();
        values["bystanderStoreDisplayNames"]!.AsArray().Add(Hub);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => Admit(Load(values)));
        Assert.Contains("designated test hub", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AHalfWrittenCorpusBlock_IsRefusedByTheLoader()
    {
        JsonObject values = SyntheticValues();
        values["corpus"]!["manifestPath"] = string.Empty;

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => Load(values));
        Assert.Contains("partially filled 'corpus' block", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AHalfWrittenSinkBlock_IsRefusedByTheLoader()
    {
        JsonObject values = SyntheticValues();
        values["mailSink"]!["retrievePort"] = 0;

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => Load(values));
        Assert.Contains("partially filled 'mailSink' block", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AProductionProfile_IsRefused_BecauseTheTemplateHasNoFieldForItsProbes()
    {
        // Why the renderer refuses Production outright: the template is the Portable shape, like the
        // example, and a Production profile must also carry probeTerm and subjectOnlyProbe.
        JsonObject values = SyntheticValues();
        values["machineProfile"] = "Production";

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => Load(values));
        Assert.Contains("probeTerm", ex.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------- the committed guest values

    [Fact]
    public void EveryFilledGuestSectionOfTestbedJson_RendersLoadsAndIsAdmitted()
    {
        // The renderer checks a guest's values against its own restatement of the loader's rules.
        // This pushes them through the loader itself, so a restatement that has drifted cannot let a
        // committed guest through. A section still holding a placeholder is one the renderer refuses
        // outright; it is reported here as waiting rather than rendered.
        string path = Path.Combine(RepoRoot(), "Testbed", "testbed.json");
        JsonObject testbed = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        JsonObject sections = Assert.IsType<JsonObject>(testbed["liveTestSettings"]);

        string template = TemplateText();
        List<string> guests = new();
        List<string> waiting = new();
        foreach (KeyValuePair<string, JsonNode?> section in sections)
        {
            if (section.Key.StartsWith('_'))
            {
                continue;
            }

            guests.Add(section.Key);
            JsonObject values = Assert.IsType<JsonObject>(section.Value);
            if (HoldsPlaceholder(values))
            {
                waiting.Add(section.Key);
                continue;
            }

            LiveTestSettings settings = LiveTestSettings.Parse(Render(template, values));
            Assert.True(Admit(settings).Usable, section.Key + " renders to a file the tier would refuse.");
            _output.WriteLine(section.Key + ": rendered, loaded and admitted - " + settings.Describe() + ".");
        }

        Assert.NotEmpty(guests);
        if (waiting.Count > 0)
        {
            _output.WriteLine(
                "PROVED NOTHING about " + string.Join(", ", waiting) + ": still waiting on values only the "
                + "guest can supply, so the renderer refuses them and there is nothing to load yet.");
        }
    }
}
