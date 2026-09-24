using System.Text;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Which profile-registry entry a signature default goes to - the defect measured on OAI-UNINDEXED
/// on 2026-09-24, reproduced and fixed against a fake of the registry.
/// <para>
/// <b>The measurement.</b> <c>Testbed/guest/Set-AccountSignature.ps1 -Account identity@vm.invalid</c>
/// drove the shipped <c>manage_signature</c> and printed "Verified" - but <c>New Signature</c> had
/// been written onto entry 00000005 of the tier profile's account-manager key, the identity PST's
/// DATA-FILE entry (clsid {ED475414-...}, <c>Account Name</c> <c>identity@vm.invalid</c>), and not
/// onto the POP3 account, entry 00000004 (<c>Account Name</c> <c>OutlookAI identity sink</c>,
/// <c>Email</c> <c>identity@vm.invalid</c>). The selection took any entry with an '@' in its
/// <c>Account Name</c>, and the read-back - <c>list_signatures</c> - used the same rule, so it
/// confirmed the wrong write.
/// </para>
/// <para>
/// <b>What is pinned.</b> <see cref="PreFixRule"/> replays the old selection and read-back verbatim:
/// on the measured layout it picks the data file and its read-back reports success - the control.
/// The fixed code, on the same layout, writes only the account and reads the account back; refuses
/// zero or several matching accounts before writing anything; will not write a non-account entry
/// even when handed one; and will not report a default as set when the entry written turns out not
/// to be the account - even with the broken selection plugged back in.
/// </para>
/// <para>
/// No test here touches the real registry or the real Signatures folder: every store runs over
/// <see cref="FakeProfileRegistry"/> and the signature files go to a per-test temp directory.
/// </para>
/// </summary>
public sealed class SignatureDefaultTargetTests : IDisposable
{
    private const string Identity = "identity@vm.invalid";
    private const string Tier = "tier@vm.invalid";
    private const string SignatureName = "Identity";

    private const string Pop3Clsid = "{ED475411-B0D6-11D2-8C3B-00104B2A6676}";
    private const string Imap4Clsid = "{ED475412-B0D6-11D2-8C3B-00104B2A6676}";
    private const string MapiServiceClsid = "{ED475414-B0D6-11D2-8C3B-00104B2A6676}";
    private const string LdapClsid = "{4DB5CBF2-3B77-4852-BC8E-BB81908861F3}";
    private const string HotmailClsid = "{4DB5CBF0-3B77-4852-BC8E-BB81908861F3}";

    private const string NewSig = SignatureManager.NewSignatureValueName;
    private const string ReplySig = SignatureManager.ReplyForwardSignatureValueName;

    /// <summary>The tier profile's account-manager key, as the live reader spells an entry's handle.</summary>
    private static readonly string AccountList =
        OutlookProfileRegistry.BuildOutlookRootKeyPath("16.0") + "\\" + OutlookProfileRegistry.ProfilesSubKeyName
        + "\\OutlookAI-Tier\\" + OutlookProfileRegistry.AccountsSubKeyName;

    private static readonly string TierAccountKey = AccountList + "\\00000001";
    private static readonly string TierDataFileKey = AccountList + "\\00000002";
    private static readonly string AddressBookKey = AccountList + "\\00000003";
    private static readonly string IdentityAccountKey = AccountList + "\\00000004";
    private static readonly string IdentityDataFileKey = AccountList + "\\00000005";

    private readonly string _dir;
    private readonly string _backupRoot;

    public SignatureDefaultTargetTests()
    {
        string root = Path.Combine(Path.GetTempPath(), "OutlookAI-SigDefaultTargetTests-" + Guid.NewGuid().ToString("N"));
        _dir = Path.Combine(root, "Signatures");
        _backupRoot = Path.Combine(root, "backups");
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path.GetDirectoryName(_dir)!, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    // ------------------------------------------------------------------ the rule

    [Theory]
    [InlineData(Pop3Clsid, null, ProfileEntryKind.MailAccount)]                                  // MEASURED, guest
    [InlineData("ed475411-b0d6-11d2-8c3b-00104b2a6676", null, ProfileEntryKind.MailAccount)]      // case and braces do not matter
    [InlineData(Imap4Clsid, null, ProfileEntryKind.MailAccount)]                                 // DOCUMENTED
    [InlineData(MapiServiceClsid, "MSEMS", ProfileEntryKind.MailAccount)]                        // MEASURED, host
    [InlineData(MapiServiceClsid, "MSUPST MS", ProfileEntryKind.DataFile)]                       // MEASURED, guest - the defect's entry
    [InlineData(MapiServiceClsid, "MSPST MS", ProfileEntryKind.DataFile)]
    [InlineData(MapiServiceClsid, "CONTAB", ProfileEntryKind.AddressBook)]                       // MEASURED, guest and host
    [InlineData(MapiServiceClsid, "EMABLT", ProfileEntryKind.AddressBook)]
    [InlineData(MapiServiceClsid, "MSPST AB", ProfileEntryKind.AddressBook)]
    [InlineData(LdapClsid, null, ProfileEntryKind.AddressBook)]                                  // DOCUMENTED
    [InlineData(MapiServiceClsid, "SOMETHING NEW", ProfileEntryKind.Unrecognized)]               // fail closed
    [InlineData(MapiServiceClsid, null, ProfileEntryKind.Unrecognized)]
    [InlineData(HotmailClsid, null, ProfileEntryKind.Unrecognized)]
    [InlineData("not a guid", null, ProfileEntryKind.Unrecognized)]
    [InlineData(null, null, ProfileEntryKind.Unrecognized)]
    public void ClassifyKind_TellsAccountsFromDataFilesAndAddressBooks_ByClassAndService(
        string? clsid, string? serviceName, ProfileEntryKind expected)
    {
        Assert.Equal(expected, ProfileAccountEntries.ClassifyKind(clsid, serviceName));
    }

    [Fact]
    public void Classify_APop3AccountIsKnownByItsEmail_NeverByItsAccountName()
    {
        // The measured identity account: its Account Name is a display name.
        ProfileAccountEntry measured = ProfileAccountEntries.Classify(
            IdentityAccountKey, Pop3Account("OutlookAI identity sink", Identity, "identity"));
        Assert.Equal(ProfileEntryKind.MailAccount, measured.Kind);
        Assert.Equal(Identity, measured.Address);
        Assert.Equal("OutlookAI identity sink", measured.AccountName);

        // An Account Name that is itself an address - of ANOTHER mailbox - still does not count.
        Dictionary<string, object?> renamed = Pop3Account("someone.else@vm.invalid", Identity, "identity");
        Assert.Equal(Identity, ProfileAccountEntries.Classify("k", renamed).Address);

        // And with no Email there is no address at all - unmatchable, never the display name.
        Dictionary<string, object?> noEmail = Pop3Account(Identity, Identity, "identity");
        noEmail.Remove("Email");
        Assert.Null(ProfileAccountEntries.Classify("k", noEmail).Address);
    }

    [Fact]
    public void Classify_AnExchangeAccountWithNoEmail_IsKnownByItsAccountName()
    {
        // The development host's shape (read 2026-09-24, value names only): MSEMS entries with no Email.
        ProfileAccountEntry exchange = ProfileAccountEntries.Classify("k", Wrapper("MSEMS", "hub@example.com"));
        Assert.Equal(ProfileEntryKind.MailAccount, exchange.Kind);
        Assert.Equal("hub@example.com", exchange.Address);

        // Renamed to something that is not an address: unmatchable, so refused - never guessed.
        Assert.Null(ProfileAccountEntries.Classify("k", Wrapper("MSEMS", "Work")).Address);
    }

    [Fact]
    public void Classify_ADataFileNamedLikeAnAddress_IsADataFile_WithNoAddress()
    {
        ProfileAccountEntry dataFile = ProfileAccountEntries.Classify(IdentityDataFileKey, DataFile(Identity));

        Assert.Equal(ProfileEntryKind.DataFile, dataFile.Kind);
        Assert.Null(dataFile.Address);
        Assert.Equal(Identity, dataFile.AccountName);
    }

    [Fact]
    public void Classify_DecodesRegBinaryValues_AsWellAsRegSz()
    {
        ProfileAccountEntry entry = ProfileAccountEntries.Classify("k", new Dictionary<string, object?>
        {
            ["clsid"] = Pop3Clsid,
            ["Email"] = Encoding.Unicode.GetBytes(Identity + "\0"),
            ["New Signature"] = Encoding.Unicode.GetBytes(SignatureName + "\0"),
            ["Reply-Forward Signature"] = "   ",
        });

        Assert.Equal(Identity, entry.Address);
        Assert.Equal(SignatureName, entry.NewSignature);
        Assert.Null(entry.ReplyForwardSignature);
    }

    // ------------------------------------------------------------------ the control

    [Fact]
    public void Control_ThePreFixRule_OnTheMeasuredProfile_WritesTheDataFile_AndItsReadBackReportsSuccess()
    {
        FakeProfileRegistry registry = MeasuredTierProfile();

        // The pre-fix selection cannot see the identity ACCOUNT at all - its Account Name has no
        // '@' - and finds the identity DATA FILE under the address instead.
        Assert.DoesNotContain(PreFixRule.Accounts(registry), r => r.Key == IdentityAccountKey);
        string? target = PreFixRule.SelectTarget(registry, Identity);
        Assert.Equal(IdentityDataFileKey, target);

        // What the shipped tool then did with that choice.
        registry.SetString(target!, NewSig, SignatureName);

        // The pre-fix read-back - list_signatures, as Set-AccountSignature.ps1 read it - reports
        // the default as set. That is the "Verified" the guest printed...
        (string Key, string Account, string? NewSignature) reported =
            Assert.Single(PreFixRule.Accounts(registry), r => r.Account == Identity);
        Assert.Equal(SignatureName, reported.NewSignature);

        // ...while the account Outlook sends from carries nothing.
        Assert.Null(registry.Value(IdentityAccountKey, NewSig));

        // The fixed read-back, on this SAME state, no longer confirms it.
        SignatureAssignment identity = Assert.Single(
            SignatureCatalog.ReadAccountAssignments(registry.ValueSets), a => a.Account == Identity);
        Assert.Null(identity.NewMessageSignature);
    }

    // ------------------------------------------------------------------ the fix

    [Theory]
    [InlineData("new", new[] { NewSig })]
    [InlineData("reply", new[] { ReplySig })]
    [InlineData("both", new[] { NewSig, ReplySig })]
    public void Fixed_OnTheMeasuredProfile_TheWriteAndTheReadBackBothLandOnTheAccount(string scope, string[] valueNames)
    {
        FakeProfileRegistry registry = MeasuredTierProfile();

        ManageSignatureOutcome outcome = Manage(
            CreateWithDefault(Identity, scope), new ProfileSignatureDefaultsStore(registry));

        Assert.Equal(Identity, outcome.DefaultSetForAccount);
        Assert.Equal(scope, outcome.DefaultSetScope);

        // Every write went to the POP3 account, 00000004, and nowhere else.
        Assert.Equal(valueNames.Select(v => (IdentityAccountKey, v, SignatureName)), registry.Sets);
        Assert.All(valueNames, v => Assert.Equal(SignatureName, registry.Value(IdentityAccountKey, v)));
        Assert.Null(registry.Value(IdentityDataFileKey, NewSig));
        Assert.Null(registry.Value(IdentityDataFileKey, ReplySig));

        // list_signatures reads the same entry back: both accounts by their Email, no data file,
        // no address book.
        IReadOnlyList<SignatureAssignment> assignments = SignatureCatalog.ReadAccountAssignments(registry.ValueSets);
        Assert.Equal(new[] { Tier, Identity }, assignments.Select(a => a.Account));
        Assert.Equal(scope == "reply" ? null : SignatureName, assignments[1].NewMessageSignature);
        Assert.Equal(scope == "new" ? null : SignatureName, assignments[1].ReplyForwardSignature);
        Assert.Null(assignments[0].NewMessageSignature);
    }

    [Fact]
    public void Fixed_AnExchangeAccountWithNoEmail_IsWrittenOnItsOwnEntry()
    {
        // The development host's shape: the Outlook Address Book, then MSEMS entries, no Email values.
        string hub = AccountList + "\\00000002";
        FakeProfileRegistry registry = new FakeProfileRegistry()
            .Add(AccountList + "\\00000001", Wrapper("CONTAB", "Outlook Address Book"))
            .Add(hub, Wrapper("MSEMS", "hub@example.com"))
            .Add(AccountList + "\\00000003", Wrapper("MSEMS", "other@example.com"));

        ManageSignatureOutcome outcome = Manage(
            CreateWithDefault("HUB@example.com", "both"), new ProfileSignatureDefaultsStore(registry));

        Assert.Equal("hub@example.com", outcome.DefaultSetForAccount);
        Assert.Equal(new[] { (hub, NewSig, SignatureName), (hub, ReplySig, SignatureName) }, registry.Sets);
    }

    // ------------------------------------------------------------------ refusals

    [Fact]
    public void Refuses_WhenOnlyADataFileCarriesTheAddress_AndTouchesNothing()
    {
        // The identity PST is attached but its account does not exist (yet): the address is in the
        // profile only as a data file's name. The old rule wrote there and reported success.
        FakeProfileRegistry registry = new FakeProfileRegistry()
            .Add(TierAccountKey, Pop3Account("OutlookAI tier sink", Tier, "tier"))
            .Add(IdentityDataFileKey, DataFile(Identity));

        ArgumentException refusal = Assert.Throws<ArgumentException>(() => Manage(
            CreateWithDefault(Identity, "new"), new ProfileSignatureDefaultsStore(registry)));

        Assert.Contains("was not found", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("data file", refusal.Message, StringComparison.Ordinal);
        Assert.Contains(Tier, refusal.Message, StringComparison.Ordinal);
        Assert.Empty(registry.Sets);
        Assert.False(File.Exists(Path.Combine(_dir, SignatureName + ".htm")), "the refusal must come before any file work");
    }

    [Fact]
    public void Refuses_WhenTwoMailAccountsCarryTheAddress_AndPicksNeither()
    {
        FakeProfileRegistry registry = MeasuredTierProfile()
            .Add(AccountList + "\\00000006", ImapAccount("identity over IMAP", Identity));

        ArgumentException refusal = Assert.Throws<ArgumentException>(() => Manage(
            CreateWithDefault(Identity, "new"), new ProfileSignatureDefaultsStore(registry)));

        Assert.Contains("2 mail accounts", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("00000004", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("00000006", refusal.Message, StringComparison.Ordinal);
        Assert.Empty(registry.Sets);
        Assert.False(File.Exists(Path.Combine(_dir, SignatureName + ".htm")), "the refusal must come before any file work");
    }

    [Theory]
    [InlineData("identity")]                    // a substring of the address
    [InlineData("vm.invalid")]
    [InlineData("OutlookAI identity sink")]     // the account's Account Name - a display name
    [InlineData("identity@vm.invalid.test")]    // a longer address that contains it
    public void Refuses_AnythingButTheWholeAddress(string account)
    {
        FakeProfileRegistry registry = MeasuredTierProfile();

        Assert.Throws<ArgumentException>(() => Manage(
            CreateWithDefault(account, "new"), new ProfileSignatureDefaultsStore(registry)));

        Assert.Empty(registry.Sets);
    }

    [Fact]
    public void TheStoreItself_WritesAndClearsOnlyTheTwoDefaultValues_OfAMailAccount()
    {
        FakeProfileRegistry registry = MeasuredTierProfile();
        ProfileSignatureDefaultsStore store = new ProfileSignatureDefaultsStore(registry);

        InvalidOperationException dataFile = Assert.Throws<InvalidOperationException>(
            () => store.WriteDefault(IdentityDataFileKey, NewSig, SignatureName));
        Assert.Contains("data file", dataFile.Message, StringComparison.Ordinal);

        InvalidOperationException addressBook = Assert.Throws<InvalidOperationException>(
            () => store.ClearDefault(AddressBookKey, ReplySig));
        Assert.Contains("address book", addressBook.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() => store.WriteDefault(IdentityAccountKey, "Account Name", SignatureName));
        Assert.Throws<InvalidOperationException>(() => store.WriteDefault(AccountList + "\\00000099", NewSig, SignatureName));

        Assert.Empty(registry.Sets);
        Assert.Empty(registry.Deletes);

        // What it does list: the two accounts, by Email - never the data files or the address book.
        Assert.Equal(
            new[] { (TierAccountKey, Tier), (IdentityAccountKey, Identity) },
            store.ReadAccounts().Select(r => (r.AccountKey, r.Account)));
    }

    // ------------------------------------------------------------------ the read-back

    [Fact]
    public void ReadBack_WillNotConfirmAWriteThatLandedOnADataFile_EvenWithTheBrokenSelectionPluggedBackIn()
    {
        FakeProfileRegistry registry = MeasuredTierProfile();

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => Manage(
            CreateWithDefault(Identity, "new"), new PreFixSelectionStore(registry)));

        // The broken selection still wrote the data file - that part is the replay...
        Assert.Equal(new[] { (IdentityDataFileKey, NewSig, SignatureName) }, registry.Sets);

        // ...and the read-back refused to call it set, and said what the entry really is.
        Assert.Equal(MutationOutcome.Applied, failure.Outcome);
        Assert.Contains("NOT reported as set", failure.Message, StringComparison.Ordinal);
        Assert.Contains("data file", failure.Message, StringComparison.Ordinal);
        Assert.Contains("00000005", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadBack_WillNotConfirmAWriteThatNeverLanded()
    {
        FakeProfileRegistry registry = MeasuredTierProfile();
        registry.DropWrites = true;

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => Manage(
            CreateWithDefault(Identity, "new"), new ProfileSignatureDefaultsStore(registry)));

        Assert.Equal(MutationOutcome.Applied, failure.Outcome);
        Assert.Contains("NOT reported as set", failure.Message, StringComparison.Ordinal);
        Assert.Contains("reads nothing", failure.Message, StringComparison.Ordinal);
        Assert.Contains("00000004", failure.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ delete sweep

    [Fact]
    public void Delete_ClearsTheAccountsDanglingDefault_AndNeverWritesTheDataFile()
    {
        // The guest after a pre-fix run and then a fixed one: the stray value on the data file,
        // the real one on the account.
        FakeProfileRegistry registry = MeasuredTierProfile();
        registry.SetString(IdentityDataFileKey, NewSig, SignatureName);
        registry.SetString(IdentityAccountKey, NewSig, SignatureName);
        registry.Sets.Clear();
        File.WriteAllText(Path.Combine(_dir, SignatureName + ".htm"), "<html><body><p>old</p></body></html>");

        ManageSignatureOutcome outcome = Manage(
            new ManageSignatureRequest { Action = "delete", Name = SignatureName }, new ProfileSignatureDefaultsStore(registry));

        Assert.Equal(new[] { Identity }, outcome.DefaultsClearedForAccounts);
        Assert.Equal(new[] { (IdentityAccountKey, NewSig) }, registry.Deletes);
        Assert.Null(registry.Value(IdentityAccountKey, NewSig));

        // Not an account, so not written - not even to tidy up. Outlook does not read it.
        Assert.Equal(SignatureName, registry.Value(IdentityDataFileKey, NewSig));
    }

    // ------------------------------------------------------------------ helpers

    private ManageSignatureOutcome Manage(ManageSignatureRequest request, ISignatureDefaultsStore store)
    {
        return SignatureManager.Manage(request, _dir, _backupRoot, store);
    }

    private static ManageSignatureRequest CreateWithDefault(string account, string scope)
    {
        return new ManageSignatureRequest
        {
            Action = "create",
            Name = SignatureName,
            BodyText = "OutlookAI testbed identity account.",
            DefaultForAccount = account,
            DefaultForScope = scope,
        };
    }

    /// <summary>
    /// The tier profile of OAI-UNINDEXED as it stood on 2026-09-24. Entries 00000004 and 00000005
    /// carry exactly what the defect report names; the other three are that profile's documented
    /// shape - its first POP3 account, that account's data file (also named after its address)
    /// and the Outlook Address Book - with illustrative subkey numbers. REG_SZ throughout, as
    /// measured on 16.0.17932.
    /// </summary>
    private static FakeProfileRegistry MeasuredTierProfile()
    {
        return new FakeProfileRegistry()
            .Add(TierAccountKey, Pop3Account("OutlookAI tier sink", Tier, "tier"))
            .Add(TierDataFileKey, DataFile(Tier))
            .Add(AddressBookKey, Wrapper("CONTAB", "Outlook Address Book"))
            .Add(IdentityAccountKey, Pop3Account("OutlookAI identity sink", Identity, "identity"))
            .Add(IdentityDataFileKey, DataFile(Identity));
    }

    private static Dictionary<string, object?> Pop3Account(string accountName, string email, string user)
    {
        return new Dictionary<string, object?>
        {
            ["clsid"] = Pop3Clsid,
            ["Account Name"] = accountName,
            ["Email"] = email,
            ["POP3 Server"] = "127.0.0.1",
            ["POP3 User"] = user,
            ["SMTP Server"] = "127.0.0.1",
        };
    }

    private static Dictionary<string, object?> ImapAccount(string accountName, string email)
    {
        return new Dictionary<string, object?>
        {
            ["clsid"] = Imap4Clsid,
            ["Account Name"] = accountName,
            ["Email"] = email,
            ["IMAP Server"] = "127.0.0.1",
            ["SMTP Server"] = "127.0.0.1",
        };
    }

    private static Dictionary<string, object?> DataFile(string storeName)
    {
        return Wrapper("MSUPST MS", storeName);
    }

    private static Dictionary<string, object?> Wrapper(string serviceName, string accountName)
    {
        return new Dictionary<string, object?>
        {
            ["clsid"] = MapiServiceClsid,
            ["Service Name"] = serviceName,
            ["Account Name"] = accountName,
        };
    }

    /// <summary>
    /// The PRE-FIX rule, replayed: the account selection and the read-back exactly as they stood
    /// at 142d793, frozen here so the defect is reproduced on the same data the fix is proven on.
    /// Not product code - a copy of the bug, kept only as the control.
    /// <list type="bullet">
    /// <item><c>ProfileSignatureDefaultsStore.ReadAccounts</c> (SignatureManager.cs:127-131) and
    /// <c>SignatureCatalog.ReadAccountAssignments</c> (SignatureCatalog.cs:173-180): an entry is a
    /// mail account when its Account Name contains '@', and that Account Name is its address -
    /// whatever its clsid.</item>
    /// <item><c>SignatureManager.Manage</c> (SignatureManager.cs:264-265): the FIRST such entry whose
    /// Account Name equals the requested address, case-insensitively.</item>
    /// </list>
    /// </summary>
    private static class PreFixRule
    {
        public static IReadOnlyList<(string Key, string Account, string? NewSignature)> Accounts(IProfileAccountRegistry registry)
        {
            List<(string Key, string Account, string? NewSignature)> rows = new();
            foreach (ProfileRegistryEntry entry in registry.ReadEntries())
            {
                string? address = SignatureCatalog.DecodeRegistryString(
                    entry.Values.TryGetValue("Account Name", out object? a) ? a : null);
                if (address == null || address.IndexOf('@') < 0)
                {
                    continue;
                }

                string? signature = SignatureCatalog.DecodeRegistryString(
                    entry.Values.TryGetValue("New Signature", out object? n) ? n : null);
                rows.Add((entry.Key, address.Trim(), string.IsNullOrWhiteSpace(signature) ? null : signature.Trim()));
            }

            return rows;
        }

        public static string? SelectTarget(IProfileAccountRegistry registry, string address)
        {
            return Accounts(registry)
                .Where(r => string.Equals(r.Account, address, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Key)
                .FirstOrDefault();
        }
    }

    /// <summary>
    /// A store that still SELECTS and WRITES by the pre-fix rule - rows from <see cref="PreFixRule"/>,
    /// writes with no guard, as the store did at 142d793 - but whose <c>ReadEntry</c> is the fixed
    /// store's independent re-read. Manage over it shows the read-back refusing to confirm a write
    /// that landed on a data file, even when the selection that chose it is the broken one.
    /// </summary>
    private sealed class PreFixSelectionStore : ISignatureDefaultsStore
    {
        private readonly FakeProfileRegistry _registry;
        private readonly ProfileSignatureDefaultsStore _fixedReader;

        public PreFixSelectionStore(FakeProfileRegistry registry)
        {
            _registry = registry;
            _fixedReader = new ProfileSignatureDefaultsStore(registry);
        }

        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            return PreFixRule.Accounts(_registry)
                .Select(r => new SignatureDefaultsRow(r.Key, r.Account, r.NewSignature, null))
                .ToList();
        }

        public ProfileAccountEntry? ReadEntry(string accountKey)
        {
            return _fixedReader.ReadEntry(accountKey);
        }

        public void WriteDefault(string accountKey, string valueName, string signatureName)
        {
            _registry.SetString(accountKey, valueName, signatureName);
        }

        public void ClearDefault(string accountKey, string valueName)
        {
            _registry.DeleteValue(accountKey, valueName);
        }
    }

    /// <summary>
    /// An in-memory account-manager key: raw value sets by entry, in the names and shapes Outlook
    /// stores them. Records every write so a test can say exactly where each one landed.
    /// </summary>
    private sealed class FakeProfileRegistry : IProfileAccountRegistry
    {
        private readonly List<(string Key, Dictionary<string, object?> Values)> _entries = new();

        public List<(string Key, string ValueName, string Value)> Sets { get; } = new();

        public List<(string Key, string ValueName)> Deletes { get; } = new();

        /// <summary>When set, a write is recorded and changes nothing - a write that silently did not happen.</summary>
        public bool DropWrites { get; set; }

        /// <summary>list_signatures' seam, fed from this same registry.</summary>
        public Func<IReadOnlyList<IReadOnlyDictionary<string, object?>>> ValueSets =>
            () => ReadEntries().Select(e => e.Values).ToList();

        public FakeProfileRegistry Add(string key, Dictionary<string, object?> values)
        {
            _entries.Add((key, new Dictionary<string, object?>(values, StringComparer.OrdinalIgnoreCase)));
            return this;
        }

        public object? Value(string key, string valueName)
        {
            Dictionary<string, object?>? entry = Find(key);
            return entry != null && entry.TryGetValue(valueName, out object? value) ? value : null;
        }

        public IReadOnlyList<ProfileRegistryEntry> ReadEntries()
        {
            // Copies: a reader never holds the live dictionaries, as a real registry read would not.
            return _entries
                .Select(e => new ProfileRegistryEntry(e.Key, new Dictionary<string, object?>(e.Values, StringComparer.OrdinalIgnoreCase)))
                .ToList();
        }

        public void SetString(string key, string valueName, string value)
        {
            Sets.Add((key, valueName, value));
            Dictionary<string, object?> entry = Find(key) ?? throw new InvalidOperationException("no such entry: " + key);
            if (!DropWrites)
            {
                entry[valueName] = value;
            }
        }

        public void DeleteValue(string key, string valueName)
        {
            Deletes.Add((key, valueName));
            Find(key)?.Remove(valueName);
        }

        private Dictionary<string, object?>? Find(string key)
        {
            foreach ((string Key, Dictionary<string, object?> Values) entry in _entries)
            {
                if (string.Equals(entry.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Values;
                }
            }

            return null;
        }
    }
}
