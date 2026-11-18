using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Q149 (d) measurement (<see cref="LiveCollections.ComposeSurface"/>): the settings, the
/// count tripwire's baseline (every guarded collection passes the funnel), the run marker every subject
/// carries, the EntryID allowlist, the Signatures-directory snapshot (the measurement makes one test
/// signature, D38), the product's MailService for the cells that compare against it, and the gate that
/// keeps the whole measurement out of any run that did not select it alone.
/// <para>
/// Its teardown deletes what the run created by EntryID, runs the tested artifact sweep over every watched
/// store (so a narrowed run still PROVES zero tagged artifacts - the runner reads its lines), checks the
/// user's signatures are bit-identical, and hands the collection to the count tripwire.
/// </para>
/// </summary>
public sealed class LiveQ149Fixture : IDisposable
{
    private ITestOutputHelper? _output;
    private MailService? _service;
    private Q180StartObserver? _starts;

    public LiveQ149Fixture()
    {
        Settings = LiveTestSettings.Load();
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        ShouldRun = Q149RunGate.ShouldRun(LiveTierRunPlan.Current);
        Run = new Q149Run(Settings, "q9" + Guid.NewGuid().ToString("N").Substring(0, 14), line => _output?.WriteLine(line));
        SignatureSnapshot = SignatureDirectorySnapshot.Capture();

        // Q180 (b): every start of Outlook while the measurement runs is logged, whoever causes it.
        _starts = ShouldRun ? new Q180StartObserver() : null;
    }

    public LiveTestSettings Settings { get; }

    public Q149Run Run { get; }

    /// <summary>Whether this run selected the measurement alone (<see cref="Q149RunGate"/>).</summary>
    public bool ShouldRun { get; }

    public SignatureDirectorySnapshot SignatureSnapshot { get; }

    /// <summary>The product's own service - created on first use, against whatever Outlook runs then.</summary>
    public MailService Service => _service ??= MailService.CreateDefault();

    /// <summary>
    /// Binds this test's output and says whether to run. A test the gate holds back writes one NOT RUN line
    /// and returns: a whole-tier run must not pay an hour of Outlook restarts for a measurement nobody asked for.
    /// </summary>
    public bool Begin(ITestOutputHelper output)
    {
        _output = output ?? throw new ArgumentNullException(nameof(output));
        if (!ShouldRun)
        {
            output.WriteLine(Q149RunGate.NotRunLine);
            return false;
        }

        output.WriteLine("q149: run marker " + Run.Marker + ", hub '" + Run.Hub + "'");
        return true;
    }

    /// <summary>Lets go of the product's session, so a window-less stage starts with no product client attached.</summary>
    public void ReleaseService()
    {
        _service?.Dispose();
        _service = null;
    }

    public void Dispose()
    {
        try
        {
            if (ShouldRun)
            {
                Teardown();
            }
        }
        finally
        {
            _starts?.Dispose();
            if (_starts != null)
            {
                Console.WriteLine("q149: Q180 observer logged " + _starts.Starts + " start(s) of Outlook");
            }

            ReleaseService();
            try
            {
                SignatureSnapshot.VerifyRealSignaturesUntouched();
            }
            finally
            {
                LiveStoreCountTripwire.CollectionFinished(LiveCollections.ComposeSurface);
            }
        }
    }

    private void Teardown()
    {
        _output = null;

        // Outlook in whatever state the last cell left it - the helpers below attach to it, or start it.
        int deleted = Q149Drafts.DeleteCreated(Run);
        Console.WriteLine("q149: deleted " + deleted + " of the " + Run.Created.Count + " item(s) this run recorded, by EntryID");

        // The tested sweep, over the count tripwire's own watched set: purge only where the allowlist permits a
        // delete, count everywhere, and fail on anything left. Its lines are what the guest runner reads to call
        // zero artifacts PROVEN for this narrowed run.
        ArtifactSweepPolicy.Run(
            Settings,
            store => LiveOutlookTestMailer.CountTaggedArtifacts(store, "OutlookAI-McpTest"),
            store => LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(store, "OutlookAI-McpTest"),
            Console.WriteLine);
    }
}

[CollectionDefinition(LiveCollections.ComposeSurface)]
public sealed class LiveQ149Collection : ICollectionFixture<LiveQ149Fixture>
{
}
