using Microsoft.Win32;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The measurements G1-G3 of the signature research (2026-10-11, section 7) on a PST guest, run as
/// live tests so they go through the live tier and its guards (Q147): which profile's defaults a
/// draft gets when Outlook runs on a profile that is not the recorded default (G1); a deployed
/// default in <c>Common\MailSettings</c> against an account's own (G2, written while Outlook runs);
/// and a relocated signature folder (G3). Each states what OutlookAI does as a CONTRACT and what
/// Outlook's own signature engine reports as a MEASUREMENT (a "MEASURED" line - recorded, never
/// asserted, because it is what is being found out).
/// <para>
/// And the settings Outlook reads only when it starts - an administrator's policy, a deployed default
/// and the accounts' own defaults - set BEFORE it starts by the live runner's <c>-SignatureSettings</c>
/// step (Q200 (a), Q201): what Outlook ITSELF inserts in a new mail, a reply and a forward, measured,
/// and OutlookAI held to doing the same.
/// </para>
/// <para>
/// Every setting a test changes is user-writable, read first, and put back exactly in finally; the
/// Phase-4 fixture's registry snapshot fails the run if anything is not. Signatures carry the test
/// prefix; drafts carry the tag and the run marker and are deleted (S3).
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveSignatureSettingsTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSignatureSettingsTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "LocalSignatures")]
    public void G1_ADraftGetsTheRunningProfilesDefault_WhicheverProfileIsRecordedAsTheDefault()
    {
        string running = _fixture.VerifySession.GetProfileName();
        SignaturesOutcome before = Service.ListSignatures();
        Assert.Equal(running, before.RunningProfile, ignoreCase: true);
        SignatureDefaultsRow? hubRow = HubRow(running);
        if (hubRow == null)
        {
            _fixture.Settings.RequireProductionPopulation("the hub account in the running profile's account list");
            _output.WriteLine("PROVED NOTHING: the hub account is not in the running profile's account list.");
            return;
        }

        string? other = before.Profiles.Select(p => p.Name).FirstOrDefault(n => !ProfileAccountEntries.ProfileEquals(n, running));
        RawRegistryValues recordedDefault = RawRegistryValues.Capture(
            OutlookProfileRegistry.OutlookRootKeyPath, LiveProfileAccountRegistry.DefaultProfileValueName);
        RawRegistryValues hubDefaults = HubDefaults(hubRow);
        using TestSignature sig = TestSignature.Create(Marker, "G1");
        string? draftId = null;
        try
        {
            SetDefault(sig.Name, running, hubRow.Account, "new");
            if (other != null)
            {
                WriteValue(OutlookProfileRegistry.OutlookRootKeyPath, LiveProfileAccountRegistry.DefaultProfileValueName, other, RegistryValueKind.String);
            }
            else
            {
                _output.WriteLine("PROVED NOTHING (G1 half): this machine has one Outlook profile, so the running profile cannot differ "
                    + "from the recorded default here.");
            }

            SignaturesOutcome listing = Service.ListSignatures();
            Assert.Equal(running, listing.RunningProfile, ignoreCase: true);
            if (other != null)
            {
                Assert.Equal(other, listing.DefaultProfile, ignoreCase: true);
            }

            SignatureAccountView row = listing.Profiles.Single(p => p.Running == true).Accounts
                .Single(a => ProfileAccountEntries.AddressEquals(a.Account, hubRow.Account));
            Assert.Equal(sig.Name, row.NewMessage);
            Assert.Equal(sig.Name, row.EffectiveNewMessage);

            string agentText = "G1 default from the running profile " + Marker;
            DraftOutcome draft = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g1-running-profile"), agentText, display: false, signature: null);
            draftId = draft.EntryId;

            Assert.Equal(sig.Name, draft.Signature);
            Assert.Equal("accountDefault", draft.SignatureSource);
            Assert.Equal(running, draft.SignatureProfile, ignoreCase: true);
            Assert.True(draft.SignatureInjected, "the running profile's default must go in");
            ReadOutcome read = Service.Read(draft.EntryId);
            int agentAt = read.Body.IndexOf(agentText, StringComparison.Ordinal);
            int sigAt = read.Body.IndexOf(sig.BodyMarker, StringComparison.Ordinal);
            Assert.True(agentAt >= 0 && sigAt > agentAt, $"the default must follow the agent text (agent@{agentAt} sig@{sigAt})");

            OutlookSignatureEngineView engine = LiveOutlookTestMailer.ReadSignatureEngine(Hub, draft.EntryId, Marker);
            _output.WriteLine($"MEASURED G1: running='{running}' recordedDefault='{other ?? running}' draft={draft.Signature} "
                + $"source={draft.SignatureSource}; Outlook's engine new='{engine.NewMessageSignature ?? "-"}' "
                + $"reply='{engine.ReplyMessageSignature ?? "-"}' listsTestSignature={engine.Entries.Contains(sig.Name)} "
                + $"entries={engine.Entries.Count} error={engine.Error ?? "-"}");
        }
        finally
        {
            DeleteDraft(draftId);
            recordedDefault.Restore();
            hubDefaults.Restore();
        }

        AssertHubDefaultsAre(hubRow);

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "LocalSignatures")]
    public void G2_ADeployedDefault_GivesWayToTheAccountsOwn_AnotherChoiceStaysAllowed_AndOutlooksViewIsRecorded()
    {
        string running = _fixture.VerifySession.GetProfileName();
        SignatureDefaultsRow? hubRow = HubRow(running);
        if (hubRow == null)
        {
            _fixture.Settings.RequireProductionPopulation("the hub account in the running profile's account list");
            _output.WriteLine("PROVED NOTHING: the hub account is not in the running profile's account list.");
            return;
        }

        string preferenceKey = SignaturePolicy.KeyPath(SignaturePolicySource.Preference, OutlookProfileRegistry.OfficeVersion);
        RawRegistryValues original = RawRegistryValues.Capture(preferenceKey, SignaturePolicy.NewSignatureValueName);
        if (SignaturePolicy.Read(new LiveSignatureSettingsRegistry(), OutlookProfileRegistry.OfficeVersion).NewMessage != null)
        {
            _output.WriteLine("PROVED NOTHING: this machine already deploys a new-message default in Common\\MailSettings; "
                + "the test does not overwrite a real setting.");
            return;
        }

        RawRegistryValues hubDefaults = HubDefaults(hubRow);
        using TestSignature deployed = TestSignature.Create(Marker, "Pref");
        using TestSignature own = TestSignature.Create(Marker, "Own");
        List<string> drafts = new();
        try
        {
            // The account's own default, then a deployed one beside it, written while Outlook runs.
            SetDefault(own.Name, running, hubRow.Account, "new");
            WriteValue(preferenceKey, SignaturePolicy.NewSignatureValueName, deployed.Name, RegistryValueKind.String);

            SignaturesOutcome listing = Service.ListSignatures();
            Assert.Equal(deployed.Name, listing.Policy!.NewMessage);
            Assert.Equal("preference", listing.Policy.NewMessageFrom);
            SignatureAccountView row = listing.Profiles.Single(p => p.Running == true).Accounts
                .Single(a => ProfileAccountEntries.AddressEquals(a.Account, hubRow.Account));
            Assert.Equal(own.Name, row.NewMessage);

            // Q201, measured 2026-11-18 with the setting made before Outlook started (runbook 4.10): the account's own
            // default ranks above a deployed one. So it is the effective default, and it stays the account's to change.
            Assert.Equal(own.Name, row.EffectiveNewMessage);
            Assert.Equal("account", row.DefaultsFrom);
            Assert.True(row.DefaultsEditable);
            SetDefault(own.Name, running, hubRow.Account, "new");

            DraftOutcome byDefault = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g2-own-default"), "G2 own default " + Marker, display: false, signature: null);
            drafts.Add(byDefault.EntryId);
            Assert.Equal(own.Name, byDefault.Signature);
            Assert.Equal("accountDefault", byDefault.SignatureSource);
            Assert.Contains(own.BodyMarker, Service.Read(byDefault.EntryId).Body, StringComparison.Ordinal);

            // Another choice stays allowed - the deployed one included - and the account's own default is suggested.
            DraftOutcome chosen = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g2-chosen"), "G2 chosen " + Marker, display: false, signature: deployed.Name);
            drafts.Add(chosen.EntryId);
            Assert.Equal(deployed.Name, chosen.Signature);
            Assert.Equal(own.Name, chosen.SignatureSuggestion);

            OutlookSignatureEngineView engine = LiveOutlookTestMailer.ReadSignatureEngine(Hub, byDefault.EntryId, Marker);
            string verdict = engine.NewMessageSignature == null ? "unknown"
                : string.Equals(engine.NewMessageSignature, deployed.Name, StringComparison.OrdinalIgnoreCase) ? "deployed default"
                : string.Equals(engine.NewMessageSignature, own.Name, StringComparison.OrdinalIgnoreCase) ? "account's own default"
                : "something else";
            _output.WriteLine($"MEASURED G2 (preference, written while Outlook runs): OutlookAI uses the account's own default; "
                + $"Outlook's engine reports new='{engine.NewMessageSignature ?? "-"}' = {verdict}; error={engine.Error ?? "-"}");
        }
        finally
        {
            foreach (string draft in drafts)
            {
                DeleteDraft(draft);
            }

            original.Restore();
            hubDefaults.Restore();
        }

        AssertHubDefaultsAre(hubRow);

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    /// <summary>
    /// Q200 (a) and Q201, decided by the maintainer 2026-11-18. With a signature setting an administrator made BEFORE
    /// Outlook started - <see cref="LiveSignatureSettingsStep"/> - what Outlook ITSELF inserts in a new mail, a reply
    /// and a forward of the hub, whether its compose window leaves one message's signature to the user (its Signature
    /// button), and that OutlookAI inserts the same and allows the same.
    /// <list type="bullet">
    /// <item><b>GroupPolicy</b> (Q200): Group Policy names a default for new mail and one for replies and forwards.</item>
    /// <item><b>Preference</b> (Q201): Outlook's general mail settings name the same two, not as a policy.</item>
    /// </list>
    /// The hub account has one default of its own - for new mail under the deployed setting, for replies under the
    /// policy - so each run ranks the setting against the account's own for one scope and shows for the other what the
    /// setting does for an account that has none.
    /// Outlook's side is read from compose windows opened on the objects Outlook's own commands make - the kind that
    /// crashed Outlook on 2026-10-04 (runbook 4.6), which is why the runner arms full crash dumps. MEASURED lines first,
    /// then the contract. A run without the step proves nothing.
    /// </summary>
    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "Transport")]
    [Trait("Requires", "LocalSignatures")]
    public void SetBeforeOutlookStarted_WhatOutlookInsertsInANewMailAReplyAndAForward_IsWhatOutlookAIInserts()
    {
        string? scenario = LiveSignatureSettingsStep.Current;
        if (scenario == null)
        {
            _fixture.Settings.RequireProductionPopulation(
                "a signature setting the live runner's -SignatureSettings step made before Outlook started");
            _output.WriteLine("PROVED NOTHING: what Outlook itself inserts under an administrator's signature setting (Q200, "
                + "Q201) needs the live runner's -SignatureSettings step, which makes the setting before Outlook starts; this run "
                + "had none.");
            return;
        }

        string running = _fixture.VerifySession.GetProfileName();
        SignatureDefaultsRow hubRow = HubRow(running)
            ?? throw new InvalidOperationException("The hub account is not in the running profile's account list, where the "
                + "step set its own defaults.");

        // What the step promised, or there is nothing to read: both defaults of the setting from its own source, and
        // nothing but test signatures named.
        SignaturePolicySource source = scenario == LiveSignatureSettingsStep.GroupPolicy
            ? SignaturePolicySource.GroupPolicy
            : SignaturePolicySource.Preference;
        SignaturePolicy policy = SignaturePolicy.Read(new LiveSignatureSettingsRegistry(), OutlookProfileRegistry.OfficeVersion);
        Assert.False(policy.SignaturesDisabled, "the step never disables signatures");
        Assert.Equal(source, policy.NewMessageSource);
        Assert.Equal(source, policy.ReplyForwardSource);
        string settingNew = RequireTestSignature(policy.NewMessage, "the setting's new-mail default");
        string settingReply = RequireTestSignature(policy.ReplyForward, "the setting's reply and forward default");
        string? ownNew = hubRow.NewMessage == null ? null : RequireTestSignature(hubRow.NewMessage, "the hub account's own new-mail default");
        string? ownReply = hubRow.ReplyForward == null ? null : RequireTestSignature(hubRow.ReplyForward, "the hub account's own reply default");
        Assert.True(source == SignaturePolicySource.GroupPolicy ? ownNew == null && ownReply != null : ownNew != null && ownReply == null,
            "the step gives the hub one default of its own: for replies under the policy, for new mail under the deployed setting");

        Dictionary<string, string> roles = new(StringComparer.Ordinal)
        {
            [settingNew] = source == SignaturePolicySource.GroupPolicy ? "the policy's new-mail default" : "the deployed new-mail default",
            [settingReply] = source == SignaturePolicySource.GroupPolicy ? "the policy's reply default" : "the deployed reply default",
        };
        if (ownNew != null)
        {
            roles[ownNew] = "the account's own new-mail default";
        }

        if (ownReply != null)
        {
            roles[ownReply] = "the account's own reply default";
        }

        List<string> candidates = roles.Keys.ToList();
        foreach (string name in candidates)
        {
            Assert.True(File.Exists(Path.Combine(SignatureCatalog.DefaultSignatureDirectory, name + ".htm")), name + " is not in the signature folder");
        }

        string Role(string inserted) => roles.TryGetValue(inserted, out string? role) ? "'" + inserted + "' (" + role + ")" : inserted;

        // How the step wrote the two values - Outlook treats an expandable string and a plain one differently (one of
        // the things this measures).
        string settingKey = SignaturePolicy.KeyPath(source, OutlookProfileRegistry.OfficeVersion);
        string valueKinds = RegistryValueKinds(settingKey, SignaturePolicy.NewSignatureValueName, SignaturePolicy.ReplySignatureValueName);
        _output.WriteLine($"MEASURED {scenario}: HKCU\\{settingKey} holds {valueKinds}; the hub's own defaults are new '{ownNew ?? "-"}', "
            + $"reply '{ownReply ?? "-"}'");

        // A mail of this run, for the reply and the forward.
        string seedSubject = _fixture.TaggedSubject("signature-step-seed");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, "Seed for the signature setting measurement " + Marker, attachmentPath: null);
        ComMailBrief seed = LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, seedSubject, sentUtc);

        List<string> drafts = new();
        try
        {
            // OUTLOOK'S SIDE: what it inserts itself, and whether its compose window leaves the choice to the user.
            Dictionary<OutlookComposeKind, string> outlookInserted = new();
            Dictionary<OutlookComposeKind, bool?> signatureMenu = new();
            List<string> unread = new();
            foreach (OutlookComposeKind kind in new[] { OutlookComposeKind.New, OutlookComposeKind.Reply, OutlookComposeKind.Forward })
            {
                OutlookOwnSignatureReading reading = LiveOutlookTestMailer.ReadOutlookOwnSignature(
                    Hub, Hub, kind, kind == OutlookComposeKind.New ? null : seed.EntryId, Marker);
                string inserted = LiveSignatureSettingsStep.Classify(reading.SignatureRegionText, reading.DocumentText, candidates);
                outlookInserted[kind] = inserted;
                signatureMenu[kind] = reading.SignatureMenuEnabled;
                IReadOnlyList<string> listed = reading.Engine.Entries;
                _output.WriteLine($"MEASURED {scenario} {kind}: Outlook inserted {Role(inserted)}; its signature region "
                    + $"{(reading.SignatureRegionFound ? "is there" : "is not there")}; editor after {reading.EditorAttempts} attempt(s); "
                    + $"sends from {(ProfileAccountEntries.AddressEquals(reading.SendingAccount, Hub) ? "the hub" : reading.SendingAccount == null ? "(unread)" : "another account")}; "
                    + $"Signature button enabled={Show(reading.SignatureMenuEnabled)} visible={Show(reading.SignatureMenuVisible)}"
                    + (reading.SignatureMenuError == null ? string.Empty : " (" + reading.SignatureMenuError + ")")
                    + $"; Word's engine new='{ShowName(reading.Engine.NewMessageSignature)}' reply='{ShowName(reading.Engine.ReplyMessageSignature)}', "
                    + $"{listed.Count} listed, {listed.Count(n => candidates.Contains(n, StringComparer.Ordinal))} of the step's"
                    + (reading.Engine.Error == null ? string.Empty : " (" + reading.Engine.Error + ")")
                    + $"; error={reading.Error ?? "-"}");
                if (reading.Error != null || inserted == LiveSignatureSettingsStep.NotRead)
                {
                    unread.Add(kind + ": " + (reading.Error ?? "nothing to read"));
                }
            }

            // OUTLOOKAI'S SIDE: the default a draft gets, and whether one message may have another signature or none.
            Dictionary<OutlookComposeKind, string> outlookAiInserted = new();
            DraftOutcome newDraft = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("signature-step-new"), "Signature step new mail " + Marker, display: false, signature: null);
            drafts.Add(newDraft.EntryId);
            outlookAiInserted[OutlookComposeKind.New] = InsertedByOutlookAI(newDraft, candidates);
            DraftOutcome replyDraft = Service.ReplyDraft(seed.EntryId, "Signature step reply " + Marker, replyAll: false, display: false, signature: null);
            drafts.Add(replyDraft.EntryId);
            outlookAiInserted[OutlookComposeKind.Reply] = InsertedByOutlookAI(replyDraft, candidates);
            DraftOutcome forwardDraft = Service.ForwardDraft(seed.EntryId, "Signature step forward " + Marker, to: Hub, display: false, signature: null);
            drafts.Add(forwardDraft.EntryId);
            outlookAiInserted[OutlookComposeKind.Forward] = InsertedByOutlookAI(forwardDraft, candidates);

            // Another signature than the one the new mail got, by Outlook's reckoning and by OutlookAI's.
            string anotherNew = candidates.First(c => !string.Equals(c, outlookInserted[OutlookComposeKind.New], StringComparison.Ordinal)
                && !string.Equals(c, outlookAiInserted[OutlookComposeKind.New], StringComparison.Ordinal));
            (bool namedAllowed, string namedWhy) = TryDraft(drafts, () => Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("signature-step-named"), "Signature step named " + Marker, display: false, signature: anotherNew));
            (bool noneAllowed, string noneWhy) = TryDraft(drafts, () => Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("signature-step-none"), "Signature step none " + Marker, display: false, signature: SignatureManager.NoSignatureKeyword));
            foreach (OutlookComposeKind kind in outlookAiInserted.Keys)
            {
                _output.WriteLine($"MEASURED {scenario} {kind}: OutlookAI inserts {Role(outlookAiInserted[kind])}");
            }

            _output.WriteLine($"MEASURED {scenario}: OutlookAI {(namedAllowed ? "allows" : "refuses")} '{anotherNew}' in one new mail ({namedWhy}) "
                + $"and {(noneAllowed ? "allows" : "refuses")} none ({noneWhy})");

            // THE CONTRACT: OutlookAI does what Outlook does (Q156, Q201).
            Assert.True(unread.Count == 0, "Outlook's side could not be read: " + string.Join("; ", unread));
            foreach (OutlookComposeKind kind in outlookInserted.Keys)
            {
                Assert.True(
                    string.Equals(outlookInserted[kind], outlookAiInserted[kind], StringComparison.Ordinal),
                    $"{kind}: Outlook inserts {Role(outlookInserted[kind])}, OutlookAI {Role(outlookAiInserted[kind])}");
            }

            bool? outlookLeavesTheChoice = signatureMenu[OutlookComposeKind.New];
            Assert.True(outlookLeavesTheChoice != null, "the Signature button of Outlook's new mail could not be read");
            Assert.True(namedAllowed == outlookLeavesTheChoice.Value,
                $"Outlook's Signature button is {(outlookLeavesTheChoice.Value ? "enabled" : "disabled")}, and OutlookAI "
                + $"{(namedAllowed ? "allows" : "refuses")} another signature in one message ({namedWhy})");
            Assert.True(noneAllowed == outlookLeavesTheChoice.Value,
                $"Outlook's Signature button is {(outlookLeavesTheChoice.Value ? "enabled" : "disabled")}, and OutlookAI "
                + $"{(noneAllowed ? "allows" : "refuses")} no signature in one message ({noneWhy})");
        }
        finally
        {
            foreach (string draft in drafts)
            {
                DeleteDraft(draft);
            }

            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "LocalSignatures")]
    public void G3_ARelocatedSignatureFolder_IsWhereSignaturesAreListedWrittenAndInserted()
    {
        string key = SignatureFolder.RegistryKeyPath;
        RawRegistryValues original = RawRegistryValues.Capture(key, SignatureFolder.RegistryValueName);
        bool wasRecorded = SignatureFolder.Current().FromRegistry;
        string defaultFolder = SignatureCatalog.DefaultSignatureDirectory;
        string relocatedName = SignatureCatalog.TestSignaturePrefix + "Folder" + Marker;
        string relocatedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", relocatedName);
        string name = SignatureCatalog.TestSignaturePrefix + "Reloc" + Marker;
        string bodyMarker = "OutlookAI MCP verplaatste handtekening " + Marker;
        string? draftId = null;
        try
        {
            WriteValue(key, SignatureFolder.RegistryValueName, relocatedName, RegistryValueKind.String);

            SignaturesOutcome listing = Service.ListSignatures();
            Assert.Equal("registry", listing.Folder.Source);
            Assert.Equal(relocatedPath, listing.Folder.Path, ignoreCase: true);

            ManageSignatureOutcome created = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "create",
                Name = name,
                BodyText = "Met vriendelijke groet,\r\n" + bodyMarker,
            });
            Assert.Equal(relocatedPath, created.Folder, ignoreCase: true);
            Assert.True(File.Exists(Path.Combine(relocatedPath, name + ".htm")), "the signature must land in the relocated folder");
            Assert.False(File.Exists(Path.Combine(defaultFolder, name + ".htm")), "and not in the default one");
            Assert.Contains(Service.ListSignatures().Signatures, s => s.Name == name);

            DraftOutcome draft = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g3-relocated"), "G3 relocated folder " + Marker, display: false, signature: name);
            draftId = draft.EntryId;
            Assert.True(draft.SignatureApplied, $"the relocated signature must go in (error: {draft.SignatureError ?? "-"})");
            Assert.Contains(bodyMarker, Service.Read(draft.EntryId).Body, StringComparison.Ordinal);

            OutlookSignatureEngineView engine = LiveOutlookTestMailer.ReadSignatureEngine(Hub, draft.EntryId, Marker);
            _output.WriteLine($"MEASURED G3: value='{relocatedName}' (set while Outlook runs); Outlook's engine lists the relocated "
                + $"signature: {engine.Entries.Contains(name)}; lists {engine.Entries.Count} signature(s); error={engine.Error ?? "-"}");
        }
        finally
        {
            DeleteDraft(draftId);
            original.Restore();
            if (Directory.Exists(relocatedPath) && Path.GetFileName(relocatedPath).StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.Ordinal))
            {
                Directory.Delete(relocatedPath, recursive: true);
            }
        }

        Assert.False(Directory.Exists(relocatedPath), "the relocated test folder must be gone");
        Assert.Equal(wasRecorded ? "registry" : "default", Service.ListSignatures().Folder.Source);
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    // ------------------------------------------------------------------ helpers

    /// <summary>A name the step set, which must be a test signature's - the only kind a test may name (rule 6).</summary>
    private static string RequireTestSignature(string? name, string what)
    {
        Assert.True(name != null, what + " is not set - the step sets it");
        Assert.True(name!.StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.Ordinal),
            what + " is not a test signature - the step names only " + SignatureCatalog.TestSignaturePrefix + "* ones");
        return name;
    }

    /// <summary>Which candidate a draft of OutlookAI's carries, read back from the saved draft.</summary>
    private string InsertedByOutlookAI(DraftOutcome draft, IReadOnlyList<string> candidates)
    {
        return LiveSignatureSettingsStep.Classify(null, Service.Read(draft.EntryId).Body, candidates);
    }

    /// <summary>Makes a draft; true when OutlookAI made it, false with the reason when it refused the signature.</summary>
    private static (bool Allowed, string Why) TryDraft(List<string> drafts, Func<DraftOutcome> make)
    {
        try
        {
            DraftOutcome draft = make();
            drafts.Add(draft.EntryId);
            return (true, draft.SignatureSource ?? "-");
        }
        catch (SignatureRefusedException refusal)
        {
            return (false, refusal.Reason);
        }
    }

    private static string Show(bool? value) => value == null ? "(unread)" : value.Value ? "true" : "false";

    /// <summary>"NewSignature REG_EXPAND_SZ, ReplySignature REG_EXPAND_SZ" - each value's kind as it stands, or absent.</summary>
    private static string RegistryValueKinds(string keyPath, params string[] valueNames)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false);
        return string.Join(", ", valueNames.Select(name =>
        {
            if (key == null || key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) == null)
            {
                return name + " absent";
            }

            return name + " " + (key.GetValueKind(name) switch
            {
                RegistryValueKind.String => "REG_SZ",
                RegistryValueKind.ExpandString => "REG_EXPAND_SZ",
                RegistryValueKind.Binary => "REG_BINARY",
                RegistryValueKind kind => kind.ToString(),
            });
        }));
    }

    /// <summary>A signature name as a log may show it: a test signature's whole, any other as what it is.</summary>
    private static string ShowName(string? name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return "-";
        }

        return name.StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.Ordinal) ? name : "(not a test signature)";
    }

    private SignatureDefaultsRow? HubRow(string profile)
    {
        return new ProfileSignatureDefaultsStore().ReadAccounts().FirstOrDefault(r =>
            ProfileAccountEntries.ProfileEquals(r.Profile, profile) && ProfileAccountEntries.AddressEquals(r.Account, Hub));
    }

    private void SetDefault(string name, string profile, string account, string scope)
    {
        Service.ManageSignature(new ManageSignatureRequest
        {
            Action = "set_default",
            Name = name,
            DefaultForProfile = profile,
            DefaultForAccount = account,
            DefaultForScope = scope,
        });
    }

    /// <summary>The hub entry's two default values exactly as they stand - kind and data, or absent.</summary>
    private static RawRegistryValues HubDefaults(SignatureDefaultsRow row)
    {
        return RawRegistryValues.Capture(row.AccountKey, SignatureManager.NewSignatureValueName, SignatureManager.ReplyForwardSignatureValueName);
    }

    /// <summary>After the restore, the product reads the hub's defaults as it did before the test.</summary>
    private static void AssertHubDefaultsAre(SignatureDefaultsRow before)
    {
        SignatureDefaultsRow? now = new ProfileSignatureDefaultsStore().ReadAccounts()
            .FirstOrDefault(r => ProfileAccountEntries.KeyEquals(r.AccountKey, before.AccountKey));
        Assert.Equal(before.NewMessage, now?.NewMessage);
        Assert.Equal(before.ReplyForward, now?.ReplyForward);
    }

    /// <summary>Sets one value; the test's RawRegistryValues capture puts it back.</summary>
    private static void WriteValue(string key, string valueName, object value, RegistryValueKind kind)
    {
        using RegistryKey created = Registry.CurrentUser.CreateSubKey(key, writable: true);
        created.SetValue(valueName, value, kind);
    }

    private void DeleteDraft(string? entryId)
    {
        if (entryId == null)
        {
            return;
        }

        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                break;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }

        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
    }
}
