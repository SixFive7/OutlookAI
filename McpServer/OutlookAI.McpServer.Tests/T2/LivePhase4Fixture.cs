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

    /// <summary>The two business accounts for the identity-only checks (Q-it2-3a).</summary>
    public IReadOnlyList<string> IdentityAccounts =>
        Settings.ExpectedStoreDisplayNames
            .Where(s => !string.Equals(s, Settings.TestHubStoreDisplayName, StringComparison.OrdinalIgnoreCase))
            .ToList();

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
