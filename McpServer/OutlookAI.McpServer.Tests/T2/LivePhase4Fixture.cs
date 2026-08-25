using System.Globalization;
using System.Reflection;
using Microsoft.Win32;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Phase-4 T2 live tier (drafts + signatures): the MailService
/// under test, an INDEPENDENT verify session (persisted-state re-opens, inspector
/// checks, window cleanup), one unique run marker for the S3 double-match delete rule,
/// and the resolved store ids / Drafts-folder identities of all three accounts.
/// Fixture disposal runs a final tag+marker cleanup on the test hub (belt - every test
/// also cleans up in finally).
/// </summary>
public sealed class LivePhase4Fixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _verifySession;

    public LivePhase4Fixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        RunMarker = "p4" + Guid.NewGuid().ToString("N").Substring(0, 14);
        _verifySession = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);

        // D38 guard: this collection contains signature-touching tests
        // (LiveSignatureTests), so the real signatures get the same snapshot
        // protection as the manage_signature suite - no snapshot, no suite.
        SignatureSnapshot = SignatureDirectorySnapshot.Capture();
    }

    /// <summary>Pre-suite Signatures-directory snapshot (D38: real signatures must stay bit-identical).</summary>
    public SignatureDirectorySnapshot SignatureSnapshot { get; }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Per-run unique marker; every artifact subject carries tag + marker (S3).</summary>
    public string RunMarker { get; }

    /// <summary>Independent COM session for persisted-state verification and window cleanup.</summary>
    public OutlookComSession VerifySession => _verifySession.Value;

    /// <summary>
    /// The business accounts for the identity-only checks (Q-it2-3a): the configured primaries
    /// the WRITE ALLOWLIST grants a draft in, minus the hub - announced on the way out, and
    /// refused when there are none.
    /// <para>
    /// Asked of the guard rather than filtered out of the settings, so the stores these tests
    /// write to and the stores they are permitted to write to are one answer. Derived
    /// separately - which is how it used to be - a store denied by the allowlist still ends up
    /// in this list, and the test throws at the guard halfway through instead of skipping a
    /// mailbox it was never entitled to touch. That is now the case for every declared
    /// BYSTANDER, which is exactly a configured primary that nothing may write to.
    /// </para>
    /// <para>
    /// <b>Why this takes arguments, and used to be a property.</b> The bystander declaration that
    /// keeps the suite out of the measurement corpus also empties this list on the documented
    /// three-store VM layout - hub plus two declared bystanders - and the two callers are
    /// <c>foreach</c> loops whose whole body is the test. An empty list therefore iterated
    /// nothing, asserted nothing and reported green, which is the failure this project keeps
    /// finding: a check that cannot fail reads as coverage in every report it appears in. The
    /// list is now unobtainable without a sink to say so through and a name for what will not
    /// run, so the announcement is not something a third caller can forget to make - see
    /// <see cref="IdentityDraftCoverage"/> for the decision itself, which is pure and pinned in
    /// CI because the tier that consumes it cannot be run there.
    /// </para>
    /// </summary>
    /// <param name="report">Where the coverage line goes - normally <c>ITestOutputHelper.WriteLine</c>.</param>
    /// <param name="whatWouldNotRun">What the caller was about to do, named as its reader would name it.</param>
    public IReadOnlyList<string> IdentityAccounts(Action<string> report, string whatWouldNotRun)
    {
        return IdentityDraftCoverage.AccountsToDraftIn(Settings, report, whatWouldNotRun);
    }

    /// <summary>Builds a tagged subject: [OutlookAI-McpTest] + run marker + label.</summary>
    public string TaggedSubject(string label)
    {
        return LiveOutlookTestMailer.SubjectTag + " " + RunMarker + " " + label;
    }

    /// <summary>StoreID of a store by display name (via the verify session).</summary>
    public string GetStoreId(string storeDisplayName)
    {
        return VerifySession.GetStores()
                .FirstOrDefault(s => string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId
            ?? throw new InvalidOperationException("Store not found by display name.");
    }

    /// <summary>Gitignored screenshots directory (v3.MD S6): McpServer/**/screenshots/.</summary>
    public string ScreenshotsDirectory
    {
        get
        {
            string testProjectDir =
                typeof(LivePhase4Fixture).Assembly
                    .GetCustomAttributes<AssemblyMetadataAttribute>()
                    .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
                ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
            return Path.Combine(testProjectDir, "screenshots");
        }
    }

    public void Dispose()
    {
        try
        {
            DisposeCore();
        }
        finally
        {
            // Outside the teardown on purpose, exactly like LiveLifecycleFixture's copy: a
            // tripwire that can be swallowed is not one. Whether this is where the run gets
            // verified depends on the FILTER, which LiveTierRunPlan knows and this fixture
            // does not.
            LiveStoreCountTripwire.CollectionFinished(LiveCollections.Phase4);
        }
    }

    /// <summary>This fixture's own teardown, separated so the tripwire signal cannot be skipped.</summary>
    private void DisposeCore()
    {
        try
        {
            // Final belt: purge anything of THIS run still tagged in the hub store.
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Settings.TestHubStoreDisplayName, RunMarker);
        }
        catch (Exception)
        {
            // Best-effort - each test already cleaned up and asserted in finally.
        }

        if (_verifySession.IsValueCreated)
        {
            // Releases COM references only - Outlook keeps running (S7: never kill/close).
            _verifySession.Value.Dispose();
        }

        Service.Dispose();

        // D38 guard: after everything (including the belt cleanup) the user's real
        // signatures must be bit-identical - only OutlookAI-McpTest-* entries may
        // have come and gone. Throws (failing the run) otherwise.
        SignatureSnapshot.VerifyRealSignaturesUntouched();
    }
}

[CollectionDefinition(LiveCollections.Phase4)]
public sealed class LivePhase4Collection : ICollectionFixture<LivePhase4Fixture>
{
}

/// <summary>
/// What the identity-only checks (Q-it2-3a) have to check on THIS machine: which configured
/// primaries the write allowlist would actually let a draft into, which it withholds, and
/// whether that leaves anything at all.
/// <para>
/// <b>The defect this closes.</b> Two live tests are a single <c>foreach</c> over that list -
/// <c>LiveDraftTests.IdentityDrafts_BusinessAccounts_RightStore_NeverDisplayed_DeletedImmediately</c>
/// and <c>LiveDraftOptionsTests.NewDraft_BusinessAccounts_BodyAboveTheirOwnIntactHtmlSignature</c>.
/// Declaring the measurement corpus and the tripwire's bystander store as BYSTANDERS - which is
/// what stops the suite drafting into a corpus and stops the artifact sweep deleting one - empties
/// the list on the documented three-store VM layout. Both tests then iterated nothing, asserted
/// nothing, and reported green: indistinguishable, in every report they appear in, from a run that
/// exercised the identity path.
/// </para>
/// <para>
/// <b>The answer is the idiom this repository already had</b>, not a new one:
/// <see cref="LiveTestSettings.RequireProductionPopulation"/> plus a <c>PROVED NOTHING:</c> line,
/// as used by <c>LiveManageSignatureTests</c> and <c>LiveStaleIndexRowTests</c>. On a Production
/// profile an empty identity set means the settings have drifted and the run refuses; on a
/// Portable one it is simply true of the machine, and the run says so in a line no reader can
/// mistake for a pass. Making it a hard failure everywhere was considered and rejected: it fails
/// machines these tests were never meant to run on, for a property of the machine.
/// </para>
/// <para>
/// <b>Partial coverage is announced, never refused.</b> A machine that grants one of three is
/// exercising the identity path, so calling that "proved nothing" would be false - and on a
/// Production profile it would refuse a run over a deliberate declaration. It gets a PARTIAL note
/// on the coverage line instead, so the number of accounts a run actually visited is in the log
/// rather than inferred from the test name.
/// </para>
/// <para>
/// Pure: names in, strings out, no COM and no settings file. That is deliberate - the live tier
/// cannot run in CI, so the decision lives here where CI can pin every branch of it, and the live
/// side is a one-line call.
/// </para>
/// </summary>
public static class IdentityDraftCoverage
{
    /// <summary>
    /// What a Production profile is told it is missing. One phrase, because the refusal text
    /// wraps it ("...where &lt;this&gt; is expected to exist") and the T1 pin greps for it.
    /// </summary>
    public const string Population = "a business account the identity tests may create a draft in";

    /// <summary>
    /// The identity accounts, with the coverage line reported first and an empty set refused
    /// (Production) or declared (Portable). The ONLY way to obtain the list.
    /// </summary>
    /// <param name="settings">The machine's live-test settings.</param>
    /// <param name="report">Where the coverage line goes - normally <c>ITestOutputHelper.WriteLine</c>.</param>
    /// <param name="whatWouldNotRun">What the caller was about to do, named as its reader would name it.</param>
    /// <returns>The stores to draft in; empty only after the run has been told so in writing.</returns>
    public static IReadOnlyList<string> AccountsToDraftIn(
        LiveTestSettings settings, Action<string> report, string whatWouldNotRun)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(report);

        IdentityDraftCoverageReport coverage = Assess(settings);
        report(coverage.Describe());
        if (coverage.ProvesNothing)
        {
            // Throws on Production - an empty set there means the settings have drifted, and a
            // green test would hide it. No-ops on Portable, where it is the truth about the
            // machine and the line below is the whole point.
            settings.RequireProductionPopulation(Population);
            report(coverage.ProvedNothing(whatWouldNotRun));
        }

        return coverage.Accounts;
    }

    /// <summary>Assesses the identity coverage of <paramref name="settings"/> as they stand.</summary>
    public static IdentityDraftCoverageReport Assess(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return Assess(settings.ExpectedStoreDisplayNames, LiveStoreWriteGuard.Build(settings));
    }

    /// <summary>
    /// Assesses the identity coverage of <paramref name="candidateStores"/> under
    /// <paramref name="allowlist"/>. The granted half is <see cref="StoreWriteAllowlist.IdentityAccountsAmong"/>
    /// itself rather than a second opinion about it: two derivations of "which stores may we draft
    /// in" is exactly how a test came to write somewhere it was not entitled to.
    /// </summary>
    public static IdentityDraftCoverageReport Assess(
        IEnumerable<string>? candidateStores, StoreWriteAllowlist allowlist)
    {
        ArgumentNullException.ThrowIfNull(allowlist);

        IReadOnlyList<string> accounts = allowlist.IdentityAccountsAmong(candidateStores);
        HashSet<string> granted = new(accounts, StringComparer.OrdinalIgnoreCase);
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        List<string> withheld = new();
        foreach (string store in candidateStores ?? [])
        {
            if (string.IsNullOrWhiteSpace(store) || allowlist.IsHub(store)
                || granted.Contains(store) || !seen.Add(store))
            {
                continue;
            }

            withheld.Add(store);
        }

        return new IdentityDraftCoverageReport(allowlist, accounts, withheld);
    }
}

/// <summary>
/// One machine's answer to "how much of the identity path can this run possibly exercise" -
/// see <see cref="IdentityDraftCoverage"/> for why it is asked at all.
/// </summary>
public sealed class IdentityDraftCoverageReport
{
    private readonly StoreWriteAllowlist _allowlist;

    internal IdentityDraftCoverageReport(
        StoreWriteAllowlist allowlist, IReadOnlyList<string> accounts, IReadOnlyList<string> withheld)
    {
        _allowlist = allowlist;
        Accounts = accounts;
        Withheld = withheld;
    }

    /// <summary>The stores the identity tests may draft in, in the order the settings name them.</summary>
    public IReadOnlyList<string> Accounts { get; }

    /// <summary>
    /// The configured non-hub primaries the allowlist withholds - declared bystanders, and
    /// anything else it does not grant a draft. Reported rather than silently dropped: the
    /// difference between "this machine has one business account" and "this machine has three and
    /// two are off limits" is invisible in a count of drafts created.
    /// </summary>
    public IReadOnlyList<string> Withheld { get; }

    /// <summary>Every configured primary but the hub, granted or not.</summary>
    public int NonHubStoreCount => Accounts.Count + Withheld.Count;

    /// <summary>
    /// True when the identity tests would iterate nothing - the state that used to report green.
    /// </summary>
    public bool ProvesNothing => Accounts.Count == 0;

    /// <summary>True when some, but not all, of the configured primaries are exercised.</summary>
    public bool Partial => Accounts.Count > 0 && Withheld.Count > 0;

    /// <summary>
    /// The coverage line, printed on EVERY run and not only the empty ones: a reader should be
    /// able to tell from the log how many accounts a passing identity test actually visited.
    /// </summary>
    public string Describe()
    {
        string line = "identity coverage: " + Accounts.Count + " of " + NonHubStoreCount
            + " non-hub store(s) the write allowlist grants an identity draft in";
        if (Withheld.Count == 0)
        {
            return line;
        }

        return line + "; " + Withheld.Count + " withheld (" + ExplainWithheld() + ")"
            + (Partial ? " - PARTIAL: this test exercises the granted ones only" : string.Empty);
    }

    /// <summary>
    /// The line a run prints instead of quietly passing. Written for whoever reads the log
    /// afterwards: what did not run, why there was nothing to run it against, and what a green
    /// result here does and does not mean.
    /// </summary>
    /// <param name="whatWouldNotRun">What the caller was about to do.</param>
    public string ProvedNothing(string whatWouldNotRun)
    {
        return "PROVED NOTHING: " + whatWouldNotRun + " iterated nothing - the write allowlist "
            + "grants an identity draft in none of this machine's " + NonHubStoreCount
            + " non-hub store(s)"
            + (Withheld.Count == 0
                ? ", because it declares none besides the hub"
                : " (" + ExplainWithheld() + ")")
            + ". Nothing about the identity path was verified here, so a green result on this "
            + "machine says only that there was no account to verify it in. To exercise it, give "
            + "this machine a second mail account and leave it out of 'bystanderStoreDisplayNames'.";
    }

    /// <summary>Each withheld store with the reason the allowlist withheld it.</summary>
    private string ExplainWithheld()
    {
        return string.Join(
            ", ",
            Withheld.Select(s => "'" + s + "' "
                + (_allowlist.IsBystander(s) ? "declared BYSTANDER" : "not granted a draft")));
    }
}

/// <summary>
/// Read-only probe of the machine's Outlook signature CONFIGURATION (names only - no
/// signature content is read or logged, S4): which .htm signature files exist and which
/// signature names the profile assigns to an account for New and Reply/Forward mail.
/// Best-effort - modern/roaming signature storage may not expose the registry values,
/// in which case the empirical signatureInjected flag from the draft result is the
/// authoritative finding.
/// </summary>
internal static class SignatureConfigProbe
{
    /// <summary>Signature .htm file base names present under %APPDATA%\Microsoft\Signatures.</summary>
    public static IReadOnlyList<string> SignatureFileBaseNames()
    {
        string dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Signatures");
        if (!Directory.Exists(dir))
        {
            return Array.Empty<string>();
        }

        return Directory.EnumerateFiles(dir, "*.htm")
            .Select(Path.GetFileNameWithoutExtension)
            .Where(n => n != null)
            .Select(n => n!)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// The (newSignatureName, replyForwardSignatureName) assigned to the account in the
    /// profile registry, or nulls when not determinable.
    /// </summary>
    public static (string? NewSignature, string? ReplySignature) AssignedSignatures(string accountSmtp)
    {
        try
        {
            // Addresses from OutlookProfileRegistry, not a fourth hand-typed copy of the
            // accounts GUID. The profile NAME stays a literal - "Outlook" is this machine's
            // default profile, and a fixture that guessed it would be worse, not better.
            using RegistryKey? profiles = Registry.CurrentUser.OpenSubKey(
                OutlookProfileRegistry.OutlookRootKeyPath
                + "\\" + OutlookProfileRegistry.ProfilesSubKeyName
                + "\\Outlook\\" + OutlookProfileRegistry.AccountsSubKeyName);
            if (profiles == null)
            {
                return (null, null);
            }

            foreach (string subKeyName in profiles.GetSubKeyNames())
            {
                using RegistryKey? sub = profiles.OpenSubKey(subKeyName);
                if (sub == null)
                {
                    continue;
                }

                string? accountName = ReadRegistryString(sub, "Account Name");
                if (accountName == null
                    || accountName.IndexOf(accountSmtp, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                return (ReadRegistryString(sub, "New Signature"), ReadRegistryString(sub, "Reply-Forward Signature"));
            }
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is IOException || ex is UnauthorizedAccessException)
        {
        }

        return (null, null);
    }

    private static string? ReadRegistryString(RegistryKey key, string valueName)
    {
        object? value = key.GetValue(valueName);
        if (value is string s)
        {
            return s.Trim('\0');
        }

        if (value is byte[] bytes && bytes.Length >= 2)
        {
            return System.Text.Encoding.Unicode.GetString(bytes).Trim('\0');
        }

        return null;
    }
}
