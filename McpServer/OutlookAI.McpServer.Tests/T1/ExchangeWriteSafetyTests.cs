using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the pure parts of the Exchange VM's Phase 2 write-safety design as built on 2026-10-08 for its
/// test account: the recipient allowlist, the run marker, the per-item check, the test-side MCP client's
/// hub-only gate, the read-only stores' pre-run baseline and the hub's pre-run leftover reading.
/// </summary>
public sealed class RecipientAllowlistTests
{
    private const string Hub = "test-account@example.invalid";
    private static readonly RecipientAllowlist Allowlist = new(new[] { Hub });

    [Theory]
    [InlineData("test-account@example.invalid", true)]
    [InlineData("TEST-ACCOUNT@EXAMPLE.INVALID", true)]
    [InlineData("  test-account@example.invalid  ", true)]
    [InlineData("Test Account <test-account@example.invalid>", true)]
    [InlineData("SMTP:test-account@example.invalid", true)]
    [InlineData("someone@example.invalid", false)]
    [InlineData("Test Account", false)]
    [InlineData("test-account@example.invalid.test", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void OnlyTheHubsOwnAddressIsAllowed(string? entry, bool allowed)
    {
        Assert.Equal(allowed, Allowlist.IsAllowed(entry));
    }

    [Fact]
    public void AFieldIsSplitOnSemicolonsAndCommas_AndEveryEntryMustBeTheHub()
    {
        Assert.Equal(new[] { Hub, Hub }, RecipientAllowlist.Split(Hub + "; " + Hub + ",  "));
        Assert.Empty(RecipientAllowlist.Split(null));
        Assert.Empty(RecipientAllowlist.Split(" ; , "));

        Assert.Null(Allowlist.RefusalFor(new string?[] { Hub, null, "Hub <" + Hub + ">;" + Hub }, "new_draft"));

        string refusal = Allowlist.RefusalFor(new string?[] { Hub + "; stranger@example.invalid", "other@example.invalid" }, "forward_draft")!;
        Assert.StartsWith(RecipientAllowlist.Refused, refusal, StringComparison.Ordinal);
        Assert.Contains("'stranger@example.invalid'", refusal, StringComparison.Ordinal);
        Assert.Contains("'other@example.invalid'", refusal, StringComparison.Ordinal);
        Assert.Contains("forward_draft", refusal, StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => Allowlist.Assert(new string?[] { "stranger@example.invalid" }, "send"));
    }

    [Fact]
    public void AnItemsRecipientsAreCounted_NeverNamed_AndAnUnreadableOneIsRefused()
    {
        Assert.Null(Allowlist.RefusalForItem(new string?[] { Hub, Hub.ToUpperInvariant() }, "send"));

        string strangers = Allowlist.RefusalForItem(new string?[] { Hub, "real.person@example.invalid" }, "send")!;
        Assert.Contains("2 recipient(s), 1 of them not the hub", strangers, StringComparison.Ordinal);
        Assert.DoesNotContain("real.person", strangers, StringComparison.Ordinal);

        Assert.Contains("1 whose address could not be read", Allowlist.RefusalForItem(new string?[] { Hub, null }, "send")!, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubsPlusAddressesAreTheHub_AndAMarkedNameResolvesToNobody()
    {
        RecipientAllowlist marked = new(new[] { Hub }, "oair20261008153000");

        Assert.True(marked.IsAllowed("test-account+cc-x@example.invalid"));
        Assert.True(marked.IsAllowed("Cc <TEST-ACCOUNT+bcc@EXAMPLE.INVALID>"));
        Assert.False(marked.IsAllowed("test-account+x@other.invalid"));
        Assert.False(marked.IsAllowed("someone+test-account@example.invalid"));
        Assert.False(marked.IsAllowed("+x@example.invalid"));
        Assert.Equal("test-account+cc-1@example.invalid", marked.PlusAddress("cc-1"));

        Assert.True(marked.IsAllowed("not a valid address p4oair20261008153000"));
        Assert.False(marked.IsAllowed("not a valid address"));
        Assert.False(marked.IsAllowed("someone oair20261008153000@example.invalid"));
        Assert.False(Allowlist.IsAllowed("not a valid address p4oair20261008153000"));

        // Read off an item, a name with no address is refused: such a draft cannot be sent from here.
        Assert.NotNull(marked.RefusalForItem(new string?[] { Hub, "not a valid address p4oair20261008153000" }, "send"));
        Assert.Null(marked.RefusalForItem(new string?[] { Hub, "test-account+cc@example.invalid" }, "send"));
    }

    [Fact]
    public void TheAllowlistHoldsPlainAddresses_AndAtLeastOne()
    {
        Assert.Throws<ArgumentException>(() => new RecipientAllowlist(Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => new RecipientAllowlist(new[] { "Test Account" }));
        Assert.Throws<ArgumentException>(() => new RecipientAllowlist(new[] { Hub + ";other@example.invalid" }));
    }
}

/// <summary>Pins <see cref="LiveRunMarker"/>'s pure parts.</summary>
public sealed class LiveRunMarkerTests
{
    private static readonly DateTime When = new(2026, 10, 8, 15, 30, 0, DateTimeKind.Utc);

    [Fact]
    public void TheRunnersMarkerIsUsed_AndSaidToBeRecorded()
    {
        (string marker, bool recorded) = LiveRunMarker.Decide(" oair20261008153000 ", When, Guid.Empty);
        Assert.Equal("oair20261008153000", marker);
        Assert.True(recorded);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("short")]
    [InlineData("OAIR20261008153000")]
    [InlineData("oair-2026-10-08")]
    [InlineData("oair2026100815300'0")]
    public void WithoutAWellFormedOne_TheProcessMakesItsOwn_WhichNothingRecorded(string? fromRunner)
    {
        (string marker, bool recorded) = LiveRunMarker.Decide(fromRunner, When, new Guid("0123456789abcdef0123456789abcdef"));
        Assert.Equal("oaix2026100815300001234567", marker);
        Assert.False(recorded);
        Assert.True(LiveRunMarker.IsWellFormed(marker));
    }

    [Fact]
    public void PriorMarkers_AreTheWellFormedOnes_OnceEach_NeverThisRuns()
    {
        Assert.Equal(
            new[] { "oair20261007100000", "oair20261006100000" },
            LiveRunMarker.ParsePrior("oair20261007100000; bad , oair20261006100000,oair20261007100000,oair20261008153000", "oair20261008153000"));
        Assert.Empty(LiveRunMarker.ParsePrior(null, "oair20261008153000"));
        Assert.Empty(LiveRunMarker.ParsePrior("  ", "oair20261008153000"));
    }

    [Fact]
    public void OnlyTheExchangeTestAccount_RefusesAnUnrecordedMarker()
    {
        Assert.Contains(LiveRunMarker.EnvironmentVariable, LiveRunMarker.RefusalFor(LiveMachineProfile.ExchangeTestAccount, recorded: false)!, StringComparison.Ordinal);
        Assert.Null(LiveRunMarker.RefusalFor(LiveMachineProfile.ExchangeTestAccount, recorded: true));
        foreach (LiveMachineProfile other in new[] { LiveMachineProfile.Portable, LiveMachineProfile.ExchangeGuest, LiveMachineProfile.Production })
        {
            Assert.Null(LiveRunMarker.RefusalFor(other, recorded: false));
        }
    }

    [Fact]
    public void AFixturesMarker_EndsWithTheRunsAndIsLongEnoughForTheMailersDeleteFilter()
    {
        string fixture = LiveRunMarker.ForFixture("p4");
        Assert.StartsWith("p4", fixture, StringComparison.Ordinal);
        Assert.EndsWith(LiveRunMarker.Current, fixture, StringComparison.Ordinal);
        Assert.True(LiveRunMarker.Current.Length >= 12);
    }
}

/// <summary>Pins <see cref="LiveRunItemCheck"/>'s decision over what was read.</summary>
public sealed class LiveRunItemCheckTests
{
    private const string Hub = "test-account@example.invalid";
    private const string Marker = "oair20261008153000";
    private static readonly LiveRunItemCheck Check = new(Hub, Marker, new RecipientAllowlist(new[] { Hub }));
    private static readonly string Subject = LiveOutlookTestMailer.SubjectTag + " p4" + Marker + " seed";

    [Fact]
    public void ANewItemsSubjectMustCarryTheTagAndTheRunsMarker()
    {
        Assert.Null(Check.RefusalForNewSubject(Subject, "new_draft"));
        Assert.NotNull(Check.RefusalForNewSubject(LiveOutlookTestMailer.SubjectTag + " p4oair20261007100000 seed", "new_draft"));
        Assert.NotNull(Check.RefusalForNewSubject("p4" + Marker + " seed", "new_draft"));
        Assert.NotNull(Check.RefusalForNewSubject(null, "new_draft"));
    }

    [Fact]
    public void AnItemIsThisRunsOnly_InTheHub_WithTheTagAndTheMarker()
    {
        Assert.Null(Check.RefusalForReading(Hub, "RE: " + Subject, null, Array.Empty<string?>(), "update_draft", addressesMail: false));
        Assert.Contains(LiveRunItemCheck.Refused, Check.RefusalForReading("real-mailbox@example.invalid", Subject, null, Array.Empty<string?>(), "move_mail", false)!, StringComparison.Ordinal);
        Assert.Contains(LiveRunItemCheck.Refused, Check.RefusalForReading(Hub, "Real mail", null, Array.Empty<string?>(), "open_in_outlook", false)!, StringComparison.Ordinal);
        Assert.Contains(LiveRunItemCheck.Refused, Check.RefusalForReading(Hub, LiveOutlookTestMailer.SubjectTag + " an earlier run", null, Array.Empty<string?>(), "discard_draft", false)!, StringComparison.Ordinal);
    }

    [Fact]
    public void WhenTheOperationAddressesMail_SenderAndEveryRecipientMustBeTheHub()
    {
        Assert.Null(Check.RefusalForReading(Hub, Subject, Hub, new string?[] { Hub }, "replyall_draft", addressesMail: true));
        Assert.StartsWith(RecipientAllowlist.Refused, Check.RefusalForReading(Hub, Subject, Hub, new string?[] { Hub, "person@example.invalid" }, "replyall_draft", true)!, StringComparison.Ordinal);
        Assert.StartsWith(RecipientAllowlist.Refused, Check.RefusalForReading(Hub, Subject, "person@example.invalid", Array.Empty<string?>(), "reply_draft", true)!, StringComparison.Ordinal);
        Assert.StartsWith(RecipientAllowlist.Refused, Check.RefusalForReading(Hub, Subject, null, new string?[] { Hub }, "send", true)!, StringComparison.Ordinal);
    }
}

/// <summary>Pins <see cref="StdioHubOnlyGate"/>: what the test-side MCP client sends on the Exchange VM's test account.</summary>
public sealed class StdioHubOnlyGateTests
{
    private const string Hub = "test-account@example.invalid";
    private const string Marker = "oair20261008153000";
    private static readonly LiveRunItemCheck Check = new(Hub, Marker, new RecipientAllowlist(new[] { Hub }));
    private static readonly string Subject = LiveOutlookTestMailer.SubjectTag + " p4" + Marker + " seed";
    private const string EntryId = "00000000ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF01";

    private static JsonElement Args(object arguments) => JsonSerializer.SerializeToElement(arguments);

    [Fact]
    public void ANewDraft_ComesFromTheHub_GoesToTheHub_AndCarriesTheTagAndTheMarker()
    {
        StdioHubOnlyGate gate = new(Check, (_, _, _) => null);
        Assert.Null(gate.RefusalFor("new_draft", Args(new { account = Hub, to = Hub, cc = Hub, subject = Subject })));
        Assert.NotNull(gate.RefusalFor("new_draft", Args(new { account = "real-mailbox@example.invalid", to = Hub, subject = Subject })));
        Assert.NotNull(gate.RefusalFor("new_draft", Args(new { account = Hub, to = "person@example.invalid", subject = Subject })));
        Assert.NotNull(gate.RefusalFor("new_draft", Args(new { account = Hub, to = Hub, bcc = "person@example.invalid", subject = Subject })));
        Assert.NotNull(gate.RefusalFor("new_draft", Args(new { account = Hub, to = Hub, subject = "no tag" })));
    }

    [Fact]
    public void AReplyToAHit_NeedsAHitThisClientSaw_ThisRunsAndFromTheHub()
    {
        StdioHubOnlyGate gate = new(Check, (_, _, _) => null);
        Assert.Contains("no search or thread result", gate.RefusalFor("reply_draft", Args(new { id = "h12" }))!, StringComparison.Ordinal);

        gate.Observe(Args(new
        {
            hits = new object[]
            {
                new { id = "h12", subject = Subject, store = Hub, fromAddress = Hub },
                new { id = "h13", subject = Subject, store = Hub, fromAddress = "person@example.invalid" },
                new { id = "h14", subject = "Real mail", store = Hub, fromAddress = Hub },
                new { id = "h15", subject = Subject, store = "real-mailbox@example.invalid", fromAddress = Hub },
            },
        }));

        Assert.Null(gate.RefusalFor("reply_draft", Args(new { id = "h12" })));
        Assert.Null(gate.RefusalFor("replyall_draft", Args(new { id = "h12", cc = Hub })));
        Assert.NotNull(gate.RefusalFor("replyall_draft", Args(new { id = "h12", cc = "person@example.invalid" })));
        Assert.NotNull(gate.RefusalFor("reply_draft", Args(new { id = "h13" })));
        Assert.Null(gate.RefusalFor("forward_draft", Args(new { id = "h13", to = Hub })));
        Assert.NotNull(gate.RefusalFor("forward_draft", Args(new { id = "h12", to = "person@example.invalid" })));
        Assert.NotNull(gate.RefusalFor("open_in_outlook", Args(new { id = "h14" })));
        Assert.NotNull(gate.RefusalFor("move_mail", Args(new { ids = new[] { "h12", "h15" }, folder = "x" })));
        Assert.Null(gate.RefusalFor("archive_mail", Args(new { ids = new[] { "h12" } })));
        Assert.Contains("EntryID", gate.RefusalFor("send", Args(new { id = "h12" }))!, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEntryId_IsJudgedByReadingTheItem_ASendAsMail()
    {
        List<(string Tool, bool AddressesMail)> asked = new();
        StdioHubOnlyGate gate = new(Check, (id, tool, addressesMail) =>
        {
            asked.Add((tool, addressesMail));
            return tool == "discard_draft" ? "judged no" : null;
        });

        Assert.Null(gate.RefusalFor("send", Args(new { id = EntryId, confirm_token = "t" })));
        Assert.Equal("judged no", gate.RefusalFor("discard_draft", Args(new { id = EntryId })));
        Assert.Equal(new[] { ("send", true), ("discard_draft", false) }, asked);
    }

    [Fact]
    public void ReadsAndToolsThatAddressNoMailPass_AndAToolNobodyJudgedIsRefused()
    {
        StdioHubOnlyGate gate = new(Check, (_, _, _) => "never asked");
        foreach (string tool in new[] { "search", "read", "thread", "list_accounts", "list_folders", "list_signatures", "outlook_health", "audit_log", "goto_folder", "show_search_results", "save_attachment" })
        {
            Assert.Null(gate.RefusalFor(tool, Args(new { })));
        }

        Assert.StartsWith(StdioHubOnlyGate.Refused, gate.RefusalFor("brand_new_tool", Args(new { }))!, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("create")]
    [InlineData("update")]
    [InlineData("delete")]
    [InlineData("set_default")]
    [InlineData("")]
    public void ManageSignature_IsRefused_WhateverItsAction_UntilRoamingIsMeasured(string action)
    {
        // Q147 R1, decided by the maintainer 2026-11-17: on the Exchange VM a signature or a default set on the
        // test account could roam into the mailbox, where no sweep removes it. The gate refuses the tool whole,
        // so no action name it might grow can slip through.
        StdioHubOnlyGate gate = new(Check, (_, _, _) => "never asked");
        string? refusal = gate.RefusalFor("manage_signature", Args(new { action, name = "OutlookAI-McpTest-x" }));
        Assert.NotNull(refusal);
        Assert.StartsWith(StdioHubOnlyGate.Refused, refusal!, StringComparison.Ordinal);
        Assert.Contains(LiveRunFilters.LocalSignatures, refusal!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheClient_RefusesBeforeSending_UnderTheHubOnlyPosture()
    {
        StdioHubOnlyGate gate = new(Check, (_, _, _) => null);
        await using McpStdioClient client = McpStdioClient.Unstarted(StdioWritePosture.WritableHubOnly, outlookReachingToolsAllowed: true, gate);
        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync("new_draft", new { account = Hub, to = "person@example.invalid", subject = Subject, body = "x", display = false }));
        Assert.StartsWith(RecipientAllowlist.Refused, refused.Message, StringComparison.Ordinal);

        await using McpStdioClient gateless = McpStdioClient.Unstarted(StdioWritePosture.WritableHubOnly, outlookReachingToolsAllowed: true);
        InvalidOperationException noGate = await Assert.ThrowsAsync<InvalidOperationException>(
            () => gateless.CallToolAsync("send", new { id = EntryId }));
        Assert.StartsWith(StdioHubOnlyGate.Refused, noGate.Message, StringComparison.Ordinal);
    }
}

/// <summary>Pins the read-only stores' pre-run baseline in the sweep, and the hub's pre-run leftover reading.</summary>
public sealed class ArtifactSweepPreExistingTests
{
    private const string Hub = "test-account@example.invalid";
    private const string Bystander = "real-mailbox@example.invalid";

    private static ArtifactSweepPlan Plan() => ArtifactSweepPolicy.Assess(
        new[] { Hub, Bystander }, new StoreWriteAllowlist(Hub, null, null, new[] { Bystander }));

    private static readonly Dictionary<string, int> OneBefore = new(StringComparer.OrdinalIgnoreCase) { [Bystander] = 1 };

    [Fact]
    public void ATaggedItemABystanderHeldBeforeTheRun_IsReportedAndLeftAlone()
    {
        List<string> lines = new();
        ArtifactSweepPolicy.Run(Plan(), store => store == Bystander ? 1 : 0, _ => Assert.Fail("nothing to purge"), lines.Add, OneBefore);
        string line = lines.Single(l => l.StartsWith("sweep[" + Bystander + "]", StringComparison.Ordinal));
        Assert.Contains("taggedArtifacts=0", line, StringComparison.Ordinal);
        Assert.Contains("preExisting=1", line, StringComparison.Ordinal);
    }

    [Fact]
    public void OneMoreThanBefore_IsAnArrival_AndFails()
    {
        InvalidOperationException refused = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), store => store == Bystander ? 2 : 0, _ => { }, _ => { }, OneBefore));
        Assert.Contains(ArtifactSweepPolicy.Residue + ": store '" + Bystander + "' holds 1 ", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OneFewerThanBefore_LeftAStoreNothingMayDeleteFrom_AndFails()
    {
        InvalidOperationException refused = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), _ => 0, _ => { }, _ => { }, OneBefore));
        Assert.Contains("are gone", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubIgnoresABaseline_ItMustEndAtZero()
    {
        Dictionary<string, int> hubBefore = new(StringComparer.OrdinalIgnoreCase) { [Hub] = 3 };
        Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), store => store == Hub ? 3 : 0, _ => { }, _ => { }, hubBefore));
    }

    [Fact]
    public void TheHubsLeftovers_OfRecordedRunsAreLeftForTheSweep_AndAnUnrecordedOneRefuses()
    {
        const string Current = "oair20261008153000";
        Dictionary<string, int> byMarker = new(StringComparer.Ordinal) { [Current] = 0, ["oair20261007100000"] = 2 };

        HubLeftovers recorded = HubLeftovers.From(2, byMarker, testFolders: 1, Current);
        Assert.Equal(0, recorded.OfNoRecordedRun);
        Assert.Null(recorded.Refusal(Hub));
        Assert.Contains("2 of oair20261007100000", recorded.Describe(Hub), StringComparison.Ordinal);

        HubLeftovers unrecorded = HubLeftovers.From(3, byMarker, testFolders: 0, Current);
        Assert.Equal(1, unrecorded.OfNoRecordedRun);
        Assert.Contains(HubLeftovers.Unrecorded, unrecorded.Refusal(Hub)!, StringComparison.Ordinal);

        Assert.Contains("no tagged item, no test folder", HubLeftovers.From(0, byMarker, 0, Current).Describe(Hub), StringComparison.Ordinal);
    }
}
