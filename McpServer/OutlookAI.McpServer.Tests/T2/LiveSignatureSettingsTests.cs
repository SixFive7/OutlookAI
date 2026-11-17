using Microsoft.Win32;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The measurements G1-G3 of the signature research (2026-10-11, section 7) on a PST guest, run as
/// live tests so they go through the live tier and its guards (Q147): which profile's defaults a
/// draft gets when Outlook runs on a profile that is not the recorded default (G1); a deployed
/// default in <c>Common\MailSettings</c> against an account's own (G2, the preference half - the
/// policy half needs an elevated step the tier does not have); and a relocated signature folder
/// (G3). Each states what OutlookAI does as a CONTRACT and what Outlook's own signature engine
/// reports as a MEASUREMENT (a "MEASURED" line - recorded, never asserted, because it is what is
/// being found out).
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
        object? recordedDefault = ReadValue(OutlookProfileRegistry.OutlookRootKeyPath, LiveProfileAccountRegistry.DefaultProfileValueName, out RegistryValueKind defaultKind);
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
            RestoreValue(OutlookProfileRegistry.OutlookRootKeyPath, LiveProfileAccountRegistry.DefaultProfileValueName, recordedDefault, defaultKind);
            RestoreDefaults(hubRow);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void G2_ADeployedDefault_IsWhatADraftGets_AnotherChoiceStaysAllowed_AndOutlooksViewIsRecorded()
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
        object? original = ReadValue(preferenceKey, SignaturePolicy.NewSignatureValueName, out RegistryValueKind originalKind);
        if (original != null)
        {
            _output.WriteLine("PROVED NOTHING: this machine already deploys a new-message default in Common\\MailSettings; "
                + "the test does not overwrite a real setting.");
            return;
        }

        using TestSignature deployed = TestSignature.Create(Marker, "Pref");
        using TestSignature own = TestSignature.Create(Marker, "Own");
        List<string> drafts = new();
        try
        {
            // The account's own default first: manage_signature refuses a default change while a deployed
            // one is in force, which is the point of the next assertions.
            SetDefault(own.Name, running, hubRow.Account, "new");
            WriteValue(preferenceKey, SignaturePolicy.NewSignatureValueName, deployed.Name, RegistryValueKind.String);

            SignaturesOutcome listing = Service.ListSignatures();
            Assert.Equal(deployed.Name, listing.Policy!.NewMessage);
            Assert.Equal("preference", listing.Policy.NewMessageFrom);
            SignatureAccountView row = listing.Profiles.Single(p => p.Running == true).Accounts
                .Single(a => ProfileAccountEntries.AddressEquals(a.Account, hubRow.Account));
            Assert.Equal(own.Name, row.NewMessage);
            Assert.Equal(deployed.Name, row.EffectiveNewMessage);
            Assert.False(row.DefaultsEditable);
            SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(() => SetDefault(own.Name, running, hubRow.Account, "new"));
            Assert.Equal("default_set_by_policy", refusal.Reason);

            DraftOutcome byDefault = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g2-deployed-default"), "G2 deployed default " + Marker, display: false, signature: null);
            drafts.Add(byDefault.EntryId);
            Assert.Equal(deployed.Name, byDefault.Signature);
            Assert.Equal("policyDefault", byDefault.SignatureSource);
            Assert.Contains(deployed.BodyMarker, Service.Read(byDefault.EntryId).Body, StringComparison.Ordinal);

            // A deployed setting is not a policy: a per-message choice stays allowed (Q156 forbids it only under a forced default).
            DraftOutcome chosen = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null,
                _fixture.TaggedSubject("g2-chosen"), "G2 chosen " + Marker, display: false, signature: own.Name);
            drafts.Add(chosen.EntryId);
            Assert.Equal(own.Name, chosen.Signature);
            Assert.Equal(deployed.Name, chosen.SignatureSuggestion);

            OutlookSignatureEngineView engine = LiveOutlookTestMailer.ReadSignatureEngine(Hub, byDefault.EntryId, Marker);
            string verdict = engine.NewMessageSignature == null ? "unknown"
                : string.Equals(engine.NewMessageSignature, deployed.Name, StringComparison.OrdinalIgnoreCase) ? "deployed default"
                : string.Equals(engine.NewMessageSignature, own.Name, StringComparison.OrdinalIgnoreCase) ? "account's own default"
                : "something else";
            _output.WriteLine($"MEASURED G2 (preference): OutlookAI uses the deployed default; Outlook's engine reports new='"
                + $"{engine.NewMessageSignature ?? "-"}' = {verdict}; error={engine.Error ?? "-"}");
        }
        finally
        {
            foreach (string draft in drafts)
            {
                DeleteDraft(draft);
            }

            RestoreValue(preferenceKey, SignaturePolicy.NewSignatureValueName, original, originalKind);
            RestoreDefaults(hubRow);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void G3_ARelocatedSignatureFolder_IsWhereSignaturesAreListedWrittenAndInserted()
    {
        string key = SignatureFolder.RegistryKeyPath;
        object? original = ReadValue(key, SignatureFolder.RegistryValueName, out RegistryValueKind originalKind);
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
            RestoreValue(key, SignatureFolder.RegistryValueName, original, originalKind);
            if (Directory.Exists(relocatedPath) && Path.GetFileName(relocatedPath).StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.Ordinal))
            {
                Directory.Delete(relocatedPath, recursive: true);
            }
        }

        Assert.False(Directory.Exists(relocatedPath), "the relocated test folder must be gone");
        Assert.Equal(original == null ? "default" : "registry", Service.ListSignatures().Folder.Source);
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    // ------------------------------------------------------------------ helpers

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

    /// <summary>The hub's two defaults exactly as they were - written back raw, not through manage_signature.</summary>
    private static void RestoreDefaults(SignatureDefaultsRow row)
    {
        ProfileSignatureDefaultsStore store = new ProfileSignatureDefaultsStore();
        Restore(store, row.AccountKey, SignatureManager.NewSignatureValueName, row.NewMessage);
        Restore(store, row.AccountKey, SignatureManager.ReplyForwardSignatureValueName, row.ReplyForward);
        SignatureDefaultsRow? now = store.ReadAccounts().FirstOrDefault(r => ProfileAccountEntries.KeyEquals(r.AccountKey, row.AccountKey));
        Assert.Equal(row.NewMessage, now?.NewMessage);
        Assert.Equal(row.ReplyForward, now?.ReplyForward);
    }

    private static void Restore(ProfileSignatureDefaultsStore store, string accountKey, string valueName, string? original)
    {
        if (original != null)
        {
            store.WriteDefault(accountKey, valueName, original);
        }
        else
        {
            store.ClearDefault(accountKey, valueName);
        }
    }

    private static object? ReadValue(string key, string valueName, out RegistryValueKind kind)
    {
        using RegistryKey? opened = Registry.CurrentUser.OpenSubKey(key, writable: false);
        object? value = opened?.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
        kind = value == null ? RegistryValueKind.String : opened!.GetValueKind(valueName);
        return value;
    }

    private readonly HashSet<string> _keysCreated = new(StringComparer.OrdinalIgnoreCase);

    private void WriteValue(string key, string valueName, object value, RegistryValueKind kind)
    {
        using (RegistryKey? existing = Registry.CurrentUser.OpenSubKey(key, writable: false))
        {
            if (existing == null)
            {
                _keysCreated.Add(key);
            }
        }

        using RegistryKey created = Registry.CurrentUser.CreateSubKey(key, writable: true);
        created.SetValue(valueName, value, kind);
    }

    /// <summary>
    /// Puts a value back exactly: rewritten with its kind, or deleted when it was absent - and a key
    /// this test had to create is removed again once it is empty.
    /// </summary>
    private void RestoreValue(string key, string valueName, object? original, RegistryValueKind kind)
    {
        if (original != null)
        {
            WriteValue(key, valueName, original, kind);
            return;
        }

        using (RegistryKey? opened = Registry.CurrentUser.OpenSubKey(key, writable: true))
        {
            opened?.DeleteValue(valueName, throwOnMissingValue: false);
        }

        if (_keysCreated.Remove(key))
        {
            using RegistryKey? left = Registry.CurrentUser.OpenSubKey(key, writable: false);
            if (left != null && left.ValueCount == 0 && left.SubKeyCount == 0)
            {
                left.Dispose();
                Registry.CurrentUser.DeleteSubKey(key, throwOnMissingSubKey: false);
            }
        }
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
