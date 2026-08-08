using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state of the D38 manage_signature live suite: the MailService under test,
/// a unique run marker for the "OutlookAI-McpTest-" signature names, and - the
/// user-ordered ABSOLUTE guard - a full Signatures-directory snapshot (file list +
/// hashes) taken BEFORE any test runs. Construction REFUSES the suite (throws,
/// failing every collection test) when the snapshot cannot be taken; disposal
/// re-captures and throws on ANY non-test-prefixed difference, proving the user's
/// real signatures bit-identical. Disposal also belt-cleans leftover test-prefixed
/// signature entries (S3 extended to signature files).
/// </summary>
public sealed class LiveSignatureManageFixture : IDisposable
{
    public LiveSignatureManageFixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        RunMarker = "d38" + Guid.NewGuid().ToString("N").Substring(0, 12);

        // The guard: no snapshot, no suite.
        Snapshot = SignatureDirectorySnapshot.Capture();
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Per-run unique marker used in the test signature names.</summary>
    public string RunMarker { get; }

    /// <summary>The pre-suite Signatures-directory snapshot (real signatures baseline).</summary>
    public SignatureDirectorySnapshot Snapshot { get; }

    /// <summary>A test signature name of this run: OutlookAI-McpTest-Mgr&lt;marker&gt;&lt;label&gt;.</summary>
    public string TestSignatureName(string label)
    {
        return SignatureCatalog.TestSignaturePrefix + "Mgr" + RunMarker + label;
    }

    public void Dispose()
    {
        try
        {
            // Belt: no test-prefixed signature entry may survive the suite (each test
            // already cleans up; exact-prefix enumeration only - 7d incident
            // discipline keeps this away from any mailbox and any real signature).
            string directory = SignatureCatalog.DefaultSignatureDirectory;
            if (Directory.Exists(directory))
            {
                foreach (string entry in Directory.GetFileSystemEntries(directory))
                {
                    string name = Path.GetFileName(entry);
                    if (!name.StartsWith(SignatureCatalog.TestSignaturePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (Directory.Exists(entry))
                    {
                        Directory.Delete(entry, recursive: true);
                    }
                    else
                    {
                        File.Delete(entry);
                    }
                }
            }
        }
        catch (Exception)
        {
            // Best-effort - the verification below still decides pass/fail for the
            // real signatures; test-prefixed leftovers are reported by the tests.
        }
        finally
        {
            Service.Dispose();
        }

        // The user's real signatures MUST be bit-identical after the suite - throw
        // (visibly failing the run) when they are not.
        Snapshot.VerifyRealSignaturesUntouched();
    }
}

[CollectionDefinition("LiveSignatureManage")]
public sealed class LiveSignatureManageCollection : ICollectionFixture<LiveSignatureManageFixture>
{
}
