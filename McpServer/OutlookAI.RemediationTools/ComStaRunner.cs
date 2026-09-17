using System.Diagnostics;
using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>
/// Which half of an STA run a bound is measuring. The two are separated because they fail
/// for opposite reasons and a single flat bound cannot tell them apart - which is the
/// defect this type exists to close.
/// </summary>
public enum ComStaPhase
{
    /// <summary>
    /// From <c>Thread.Start</c> until the run says Outlook answered
    /// (<see cref="ComStaCheckpoint.OutlookBound"/>). Creating an
    /// <c>Outlook.Application</c> can COLD-START OUTLOOK.EXE, which is slow rather than
    /// stuck, and nothing inside that single COM call can report progress.
    /// </summary>
    StartingOutlook = 0,

    /// <summary>
    /// Outlook has answered and the run is doing what it was asked to do. Here silence IS a
    /// symptom: the work loops report a step at every safe point, so a run that has stopped
    /// reporting has stopped making progress.
    /// </summary>
    Working = 1,
}

/// <summary>
/// The two bounds an STA run is held to, plus how long it is given to acknowledge a stop.
/// <para>
/// <b>Why two.</b> The old runner had ONE flat bound covering everything from process
/// creation to the last delete, and on 2026-09-16 it expired twice on the test guest during
/// COLD OUTLOOK STARTS - reported as <c>TimeoutException: Corpus STA operation timed out</c>,
/// which reads as "the work is stuck" when the work had not begun. A cold start is not a
/// hang, so it gets its own allowance and the work bound does not start until Outlook has
/// answered. That is the same rule the product side already keeps in
/// <c>BudgetedSessionProxy</c>: the clock starts AFTER the session is connected.
/// </para>
/// <para>
/// <b><see cref="Work"/> is a SILENCE bound, not a duration.</b> It is reset by every
/// <see cref="ComStaCheckpoint.Step"/>, so a corpus build that legitimately runs for hours
/// never trips it while it is creating items, and a run wedged inside one COM call trips it
/// even if the run as a whole is young. That is strictly more detection than the old total
/// bound gave, and strictly fewer false expiries.
/// </para>
/// </summary>
/// <param name="Startup">
/// How long the run may take to reach <see cref="ComStaCheckpoint.OutlookBound"/>. Null is
/// unbounded.
/// </param>
/// <param name="Work">
/// How long the run may go without reporting a step once Outlook has answered. Null is
/// unbounded, which is what the corpus BUILD, RE-ANCHOR, TEARDOWN and SCAN use: a build of
/// tens of thousands of items runs for hours on purpose, and abandoning one mid-write is the
/// thing that leaves a store no manifest describes.
/// </param>
/// <param name="Grace">
/// How long the runner waits, after signalling, for the thread to acknowledge by ending.
/// </param>
public sealed record ComStaBudget(TimeSpan? Startup, TimeSpan? Work, TimeSpan Grace)
{
    /// <summary>Widens (or, at <c>0</c>, removes) the cold-start allowance, in milliseconds.</summary>
    public const string StartupVariable = "OUTLOOKAI_CORPUS_STA_STARTUP_MS";

    /// <summary>Widens (or, at <c>0</c>, removes) the work silence bound, in milliseconds.</summary>
    public const string WorkVariable = "OUTLOOKAI_CORPUS_STA_WORK_MS";

    /// <summary>
    /// How long a cold Outlook start may take before the run is abandoned: TEN MINUTES.
    /// <para>
    /// <b>Where the number comes from.</b> The only cold start anyone has timed on the
    /// measurement guest is the 2026-09-15 licence probe, with Outlook provably not running:
    /// <c>CreateObject returned in 3.7 s; full bind in 4.4 s</c>. The product's own budget for
    /// a connect that may cold-start Outlook is <c>ComOperationBudgets.ConnectDeadlineMs</c> =
    /// 180 s, which <c>Docs/magic-numbers.md</c> records as deliberately far above its
    /// measurement because "a large OST on a slow disk is far slower". This is 3.3x that, and
    /// ~135x the one measured start. It is therefore a CEILING TO BE NARROWED FROM
    /// MEASUREMENT, not a measured value - the same status that document gives the freshness
    /// sweep's budget. It is set high on purpose: this is an operator console with nobody
    /// waiting on latency, so paying ten minutes once beats reporting a start failure that was
    /// really a slow disk, and the incident that prompted this work was exactly that report.
    /// </para>
    /// </summary>
    public static readonly TimeSpan DefaultStartup = TimeSpan.FromMinutes(10);

    /// <summary>
    /// How long the thread is given to acknowledge a stop: SIXTY SECONDS.
    /// <para>
    /// Derived rather than measured: the checkpoints sit at loop boundaries, so acknowledging
    /// costs at most one item's worth of COM - a create, a save, a delete, a table walk. A
    /// minute is far more than any of those takes on a store that is responding at all, which
    /// makes "did not acknowledge within a minute" a statement about Outlook rather than about
    /// the bound being mean.
    /// </para>
    /// </summary>
    public static readonly TimeSpan DefaultGrace = TimeSpan.FromSeconds(60);

    /// <summary>
    /// The budget for a run whose work silence bound is <paramref name="work"/>, with the
    /// cold-start allowance and the grace taken from the environment or the defaults.
    /// </summary>
    public static ComStaBudget For(TimeSpan? work)
        => Resolve(
            work,
            Environment.GetEnvironmentVariable(StartupVariable),
            Environment.GetEnvironmentVariable(WorkVariable));

    /// <summary>
    /// PURE, so the T1 tier pins it: the budget that results from a caller's work bound and
    /// two raw environment values. A value is milliseconds; <c>0</c> means UNBOUNDED, which is
    /// the escape hatch for a machine where a cold start really does take longer than anyone
    /// has budgeted for; anything unparseable or negative is ignored and the default stands,
    /// which is the convention <c>ComHostPolicy</c> already uses for its own override.
    /// </summary>
    public static ComStaBudget Resolve(TimeSpan? work, string? startupRaw, string? workRaw)
    {
        TimeSpan? startup = TryReadBound(startupRaw, out TimeSpan? configuredStartup)
            ? configuredStartup
            : DefaultStartup;
        TimeSpan? effectiveWork = TryReadBound(workRaw, out TimeSpan? configuredWork) ? configuredWork : work;
        return new ComStaBudget(startup, effectiveWork, DefaultGrace);
    }

    /// <summary>The bounds in words, for the message an expiry prints.</summary>
    public string Describe()
        => "cold start " + Format(Startup) + ", work silence " + Format(Work) + ", grace " + Format(Grace);

    /// <summary>A bound as an operator reads it: <c>none</c>, <c>45s</c>, <c>10m</c>, <c>1m30s</c>.</summary>
    public static string Format(TimeSpan? bound)
    {
        if (bound == null)
        {
            return "none";
        }

        TimeSpan span = bound.Value;
        if (span.TotalMinutes < 1)
        {
            return span.TotalSeconds.ToString("0.#", CultureInfo.InvariantCulture) + "s";
        }

        int minutes = (int)span.TotalMinutes;
        return span.Seconds == 0
            ? minutes.ToString(CultureInfo.InvariantCulture) + "m"
            : minutes.ToString(CultureInfo.InvariantCulture) + "m"
                + span.Seconds.ToString(CultureInfo.InvariantCulture) + "s";
    }

    private static bool TryReadBound(string? raw, out TimeSpan? bound)
    {
        bound = null;
        if (!long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out long milliseconds)
            || milliseconds < 0)
        {
            return false;
        }

        bound = milliseconds == 0 ? null : TimeSpan.FromMilliseconds(milliseconds);
        return true;
    }
}

/// <summary>
/// The cancellation signal an STA run carries, and the only sanctioned way its loops step.
/// <para>
/// <b>It is one object doing two jobs on purpose.</b> Every safe point is both the place a
/// run may stop and the place it proves it is alive, so asking "may I continue?" and saying
/// "I am still here" must be the same call - otherwise a loop can be instrumented for one and
/// not the other, and the half that was forgotten is silent until the day it matters.
/// </para>
/// <para>
/// <b>Prefer <see cref="Steps{T}"/> to <see cref="Step"/>.</b> Iterating through
/// <see cref="Steps{T}"/> makes the check part of the loop's own enumerator: the loop cannot
/// be written without it, which is what stops this defect coming back the next time somebody
/// adds a pass over a collection. <see cref="Step"/> is for the <c>while</c> loops that have
/// no collection to iterate - the COM table walks.
/// </para>
/// <para>
/// <b>A stop is a BREAK, never an abort.</b> A .NET thread cannot be safely aborted and
/// <c>Thread.Interrupt</c> only unblocks waits, so cooperative is the only honest route; and
/// for a WRITE loop it is also the correct one, because breaking at a loop boundary leaves the
/// store in a state the manifest already describes. Every corpus write loop records its item
/// before the next iteration begins, so the boundary is exactly the safe point.
/// </para>
/// </summary>
public sealed class ComStaCheckpoint
{
    private readonly object _gate = new();
    private readonly CancellationToken _token;
    private long _lastSignOfLife = Stopwatch.GetTimestamp();
    private ComStaPhase _phase = ComStaPhase.StartingOutlook;
    private string? _label;
    private long _labelSteps;
    private long _totalSteps;
    private string? _profileName;

    internal ComStaCheckpoint(CancellationToken token) => _token = token;

    /// <summary>
    /// The run's cancellation token, for anything that already speaks
    /// <see cref="CancellationToken"/>. It is the SAME signal <see cref="Step"/> reports, not
    /// a second mechanism.
    /// </summary>
    public CancellationToken Token => _token;

    /// <summary>Whether the run has been asked to stop. <see cref="Step"/> returning false says the same thing.</summary>
    public bool Stopping => _token.IsCancellationRequested;

    /// <summary>
    /// The MAPI profile this run bound, as <c>NameSpace.CurrentProfileName</c> reported it, or
    /// null when it could not be read. Null also means the run never got that far.
    /// </summary>
    public string? ProfileName
    {
        get
        {
            lock (_gate)
            {
                return _profileName;
            }
        }
    }

    /// <summary>
    /// Records one safe point of <paramref name="what"/> and answers whether the run may
    /// continue. FALSE means stop now, at this boundary, and unwind through the
    /// <c>finally</c> blocks that release COM references and delete throwaway items.
    /// <para>
    /// The step is counted only when the answer is yes, so the count is what the run actually
    /// DID rather than what it was about to attempt - which is the number an operator needs
    /// when they are asking how far a cancelled write got.
    /// </para>
    /// </summary>
    public bool Step(string what)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(what);
        lock (_gate)
        {
            // Asking the question is itself a sign of life: a loop that is being told to stop
            // must not also look stalled while it unwinds.
            _lastSignOfLife = Stopwatch.GetTimestamp();
            if (_token.IsCancellationRequested)
            {
                return false;
            }

            if (!string.Equals(_label, what, StringComparison.Ordinal))
            {
                _label = what;
                _labelSteps = 0;
            }

            _labelSteps++;
            _totalSteps++;
            return true;
        }
    }

    /// <summary>
    /// <paramref name="source"/>, one <see cref="Step"/> per element, ending the iteration as
    /// soon as the run is asked to stop. This is the shape every corpus loop should use.
    /// </summary>
    public IEnumerable<T> Steps<T>(IEnumerable<T> source, string what)
    {
        ArgumentNullException.ThrowIfNull(source);
        foreach (T item in source)
        {
            if (!Step(what))
            {
                yield break;
            }

            yield return item;
        }
    }

    /// <summary>
    /// Says that Outlook has answered and names the profile it bound. This is the moment the
    /// cold-start allowance ends and the work silence bound begins, and it is deliberately
    /// called from ONE place (<c>ComCorpusMailbox.BindNamespace</c>) so it cannot be half
    /// applied.
    /// <para>
    /// If it were never called the run would stay under the cold-start allowance, which fails
    /// towards waiting longer rather than towards abandoning a live COM session - the right
    /// direction for a mistake in a file that writes to mailboxes.
    /// </para>
    /// </summary>
    public void OutlookBound(string? profileName)
    {
        lock (_gate)
        {
            _profileName = profileName;
            _phase = ComStaPhase.Working;
            _lastSignOfLife = Stopwatch.GetTimestamp();
        }
    }

    internal ComStaPhase Phase
    {
        get
        {
            lock (_gate)
            {
                return _phase;
            }
        }
    }

    internal TimeSpan Silence
    {
        get
        {
            lock (_gate)
            {
                return Stopwatch.GetElapsedTime(_lastSignOfLife);
            }
        }
    }

    /// <summary>How far the run got, in a sentence, for the message an expiry prints.</summary>
    internal string DescribeProgress()
    {
        lock (_gate)
        {
            if (_label == null)
            {
                return _phase == ComStaPhase.StartingOutlook
                    ? "It never got past starting Outlook, so it had attempted nothing in the store"
                    : "Outlook answered but the run reported no work step at all";
            }

            return string.Create(
                CultureInfo.InvariantCulture,
                $"It got as far as {_labelSteps:N0} '{_label}' step(s) ({_totalSteps:N0} step(s) in all)");
        }
    }
}

/// <summary>
/// An STA run that was abandoned because a bound expired - and, crucially, WHETHER THE THREAD
/// STOPPED. <see cref="Acknowledged"/> false is the dangerous state: a COM session is still
/// live against the store after the caller has been told the operation failed.
/// </summary>
public sealed class ComStaTimeoutException : TimeoutException
{
    /// <summary>Creates the exception. See the properties for what each part means.</summary>
    public ComStaTimeoutException(
        string message, bool acknowledged, ComStaPhase phase, string progress, Exception? innerException)
        : base(message, innerException)
    {
        Acknowledged = acknowledged;
        Phase = phase;
        Progress = progress;
    }

    /// <summary>Creates the exception with no detail. Present for the framework's exception shape only.</summary>
    public ComStaTimeoutException()
        : this("The STA operation timed out.", false, ComStaPhase.Working, string.Empty, null)
    {
    }

    /// <summary>Creates the exception with a message only. Present for the framework's exception shape only.</summary>
    public ComStaTimeoutException(string message)
        : this(message, false, ComStaPhase.Working, string.Empty, null)
    {
    }

    /// <summary>Creates the exception with a message and a cause. Present for the framework's exception shape only.</summary>
    public ComStaTimeoutException(string message, Exception? innerException)
        : this(message, false, ComStaPhase.Working, string.Empty, innerException)
    {
    }

    /// <summary>
    /// TRUE when the thread ended within the grace period: it unwound through its own
    /// <c>finally</c> blocks, released its Outlook references, and nothing from this run is
    /// still touching the store. FALSE when it did not - see the class summary.
    /// </summary>
    public bool Acknowledged { get; }

    /// <summary>Which bound expired.</summary>
    public ComStaPhase Phase { get; }

    /// <summary>How far the run had got, as <see cref="ComStaCheckpoint.DescribeProgress"/> describes it.</summary>
    public string Progress { get; }
}

/// <summary>
/// Runs a delegate on a dedicated STA thread under <see cref="ComStaBudget"/>, and - when a
/// bound expires - SIGNALS THE THREAD AND WAITS FOR IT, instead of walking away.
/// <para>
/// <b>The defect this replaces.</b> The previous runner threw <c>TimeoutException</c> the
/// moment <c>Thread.Join(timeout)</c> returned false and returned to the caller while the
/// thread kept running. The thread was <c>IsBackground = true</c>, so nothing joined it and
/// nothing cancelled it: it held its <c>Application</c>, <c>NameSpace</c>, <c>Store</c> and
/// <c>Folder</c> references and went on driving Outlook until the process exited. For a read
/// that is wasted work. For a WRITE - the placement and date probes both create and delete
/// mail - it is an abandoned COM session mutating a store the caller has been told it stopped
/// touching, and a caller that then retries puts TWO writers against one store.
/// </para>
/// <para>
/// <b>What it does instead.</b> It cancels, waits <see cref="ComStaBudget.Grace"/>, and
/// reports which of the two things happened: stopped cleanly after N steps, or did not
/// acknowledge - the second loudly, because that is the state where the abandoned session
/// really does exist and the operator must not start another one.
/// </para>
/// </summary>
public static class ComStaRunner
{
    /// <summary>
    /// How often the runner looks at the clock. Short enough that an expiry is reported
    /// promptly, long enough that watching costs nothing next to a COM call.
    /// </summary>
    internal static readonly TimeSpan PollInterval = TimeSpan.FromMilliseconds(250);

    /// <summary>
    /// Runs <paramref name="work"/> on a dedicated STA thread under <paramref name="budget"/>.
    /// </summary>
    /// <typeparam name="T">What the work returns.</typeparam>
    /// <param name="operation">
    /// What this run is, in words, for the message an expiry prints - "corpus build", "corpus
    /// placement probe". An operator who sees a timeout should not have to guess which of the
    /// eight corpus commands' STA runs produced it.
    /// </param>
    /// <param name="budget">The two bounds and the grace period.</param>
    /// <param name="work">
    /// The COM work. It is handed the checkpoint and must step through its loops with it; a
    /// body that never steps can only ever be stopped by not being waited for, which is the
    /// defect this class exists to remove.
    /// </param>
    /// <returns>What the work returned, when it finished within its bounds.</returns>
    /// <exception cref="ComStaTimeoutException">A bound expired. Read <see cref="ComStaTimeoutException.Acknowledged"/>.</exception>
    /// <exception cref="InvalidOperationException">The work threw.</exception>
    public static T Run<T>(string operation, ComStaBudget budget, Func<ComStaCheckpoint, T> work)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operation);
        ArgumentNullException.ThrowIfNull(budget);
        ArgumentNullException.ThrowIfNull(work);

        // Deliberately NOT a `using`. In the one case that matters - the thread did not
        // acknowledge - it is still reading this token, and disposing the source out from
        // under a live reader to save one object would be the same class of mistake as
        // walking away from the thread in the first place. It is disposed on every path where
        // the thread has provably finished.
        var stopping = new CancellationTokenSource();
        var checkpoint = new ComStaCheckpoint(stopping.Token);
        T result = default!;
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = work(checkpoint);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        })
        {
            IsBackground = true,
            Name = "OutlookAI.Corpus.Sta",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        if (WaitWithinBounds(thread, checkpoint, budget, out ComStaPhase expiredIn, out TimeSpan silence))
        {
            stopping.Dispose();
            if (failure != null)
            {
                throw new InvalidOperationException(operation + " failed.", failure);
            }

            return result;
        }

        // Signal, then WAIT. The whole point: the caller is not told this failed until the
        // question "is something still writing?" has an answer.
        stopping.Cancel();
        var acknowledging = Stopwatch.StartNew();
        bool acknowledged = thread.Join(budget.Grace);
        acknowledging.Stop();

        // ONE snapshot, used for both the message and the property: taken after the join so
        // an acknowledged run reports where it actually stopped, and read once so the two do
        // not disagree when the thread is still moving.
        string progress = checkpoint.DescribeProgress();
        if (acknowledged)
        {
            stopping.Dispose();
        }

        throw new ComStaTimeoutException(
            ExpiryMessage(operation, budget, progress, expiredIn, silence, acknowledged, acknowledging.Elapsed),
            acknowledged,
            expiredIn,
            progress,
            acknowledged ? failure : null);
    }

    /// <summary>
    /// Waits for the thread, watching whichever bound the run is currently under. True when
    /// the thread ended on its own; false when a bound expired, and then
    /// <paramref name="expiredIn"/> says which and <paramref name="silence"/> says how long
    /// the run had been quiet.
    /// </summary>
    private static bool WaitWithinBounds(
        Thread thread,
        ComStaCheckpoint checkpoint,
        ComStaBudget budget,
        out ComStaPhase expiredIn,
        out TimeSpan silence)
    {
        while (true)
        {
            if (thread.Join(PollInterval))
            {
                expiredIn = checkpoint.Phase;
                silence = TimeSpan.Zero;
                return true;
            }

            // Re-read the phase every tick: a cold start that finishes hands the run over to
            // the work bound, and a run that is still starting must never be judged by it.
            ComStaPhase phase = checkpoint.Phase;
            TimeSpan? bound = phase == ComStaPhase.StartingOutlook ? budget.Startup : budget.Work;
            TimeSpan quiet = checkpoint.Silence;
            if (bound != null && quiet >= bound.Value)
            {
                expiredIn = phase;
                silence = quiet;
                return false;
            }
        }
    }

    private static string ExpiryMessage(
        string operation,
        ComStaBudget budget,
        string progress,
        ComStaPhase expiredIn,
        TimeSpan silence,
        bool acknowledged,
        TimeSpan acknowledgingTook)
    {
        string what = expiredIn == ComStaPhase.StartingOutlook
            ? "was still starting Outlook after " + ComStaBudget.Format(silence)
                + " (the cold-start allowance; a cold start is slow, not stuck, so this bound is separate "
                + "from the work one)"
            : "reported no progress for " + ComStaBudget.Format(silence)
                + " (the work silence bound, which every safe point in the run resets)";

        string outcome = acknowledged
            ? "It ACKNOWLEDGED the stop after " + ComStaBudget.Format(acknowledgingTook)
                + ", unwound through its own finally blocks and released Outlook: nothing from this run is still "
                + "writing to the store."
            : "IT DID NOT ACKNOWLEDGE THE STOP within " + ComStaBudget.Format(budget.Grace)
                + ". THE STA THREAD IS STILL RUNNING AND STILL HOLDS OUTLOOK REFERENCES - an abandoned COM session "
                + "that may go on creating, moving or deleting items in this store until this process exits. Do NOT "
                + "re-run this command against the same store while this process lives: two writers against one "
                + "store is the outcome this bound exists to prevent. Let the process exit first.";

        return operation + " timed out: it " + what + ". " + progress + ". " + outcome
            + " Bounds: " + budget.Describe() + " - set " + ComStaBudget.StartupVariable + " or "
            + ComStaBudget.WorkVariable + " (milliseconds; 0 removes the bound) to widen them.";
    }
}
