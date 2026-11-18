using System.Globalization;
using OutlookAI.RemediationTools;

/// <summary>
/// Operator console for the SYNTHETIC MEASUREMENT CORPUS and the test guests' fixture populations.
/// Its commands build, check, age and remove them in a local .pst, which is how the freshness-sweep
/// and exhaustive-scan budgets get measured against known volume instead of modelled (see
/// Docs/corpus-measurement-plan.md), and how a test guest's hub, bystander and identity stores get
/// their curated contents (Docs/live-tier-on-the-vm.md section 3b). It is not part of the product:
/// the shipped MCP server keeps zero delete tools.
///
/// Until 2026-11-18 it also held the four commands of the July 2026 incident closure - audit, refile,
/// purge and dedupe - and the decision rules and COM helpers only they used. The maintainer closed that
/// incident and they were deleted (Q170 (a)); git history keeps them.
///
///   corpus-plan     --corpus-id ... --seed N --anchor yyyy-MM-dd --count N
///       Pure. Prints what the corpus would contain - per folder, per size class, per age
///       band, and how many items each measurement window selects. No Outlook.
///
///   corpus-probe    --store ... --allow-store ... --corpus-id ... --seed N --anchor ...
///                   [--population ... --undated-index-wait N]
///       Settles two things empirically, by writing throwaway items and reading them back,
///       and deletes every probe it creates. PLACEMENT first: whether an item can be made to
///       live in the folder the plan names and appear in that folder's table, which is what
///       the freshness sweep reads. Then DATES, built with the placement that verified - the
///       order matters, because a date probed against an item filed in another folder cannot
///       distinguish a date that does not select from an item that is not there.
///
///   corpus-build    ... --count N --manifest &lt;path&gt; [--allow-undated]
///                       [--allow-drafts-placement] [--execute]
///       Creates the corpus. Resumable and idempotent - it builds the ordinals the manifest
///       does not already record.
///
///   --population hub|bystander|identity, on any corpus verb
///       Builds a curated FIXTURE population for a test guest's hub, bystander or identity
///       store instead of the measurement corpus (CorpusPopulation; Docs/live-tier-on-the-vm.md
///       section 3b). Same generator, same guards, same tag, same manifest and teardown - plus
///       senders, recipients, attachments, conversations and subfolders, each proved by a probe
///       before the build and read back after it. Its item count is fixed by the kind, and
///       --store is required even by corpus-plan, because the population is addressed to the
///       store's owner. --undated-contacts gives the hub or bystander its undated CONTACTS (the
///       indexed guest's populations, Q98 (f)); it is part of the shape key, so every verb that
///       reads such a manifest needs it as well.
///
///   corpus-census   --store ... --allow-store ... --corpus-id ... --count N [--manifest ...]
///       READ-ONLY. Says whether the corpus in the store is the corpus the plan describes:
///       right count, right folders, one copy each, and nothing stranded in Drafts or the
///       Outbox. A build runs this on itself and fails if it is not clean.
///
///   corpus-verify   --corpus-id ... --seed N --anchor ... --count N --manifest ... [--window N]
///       PURE - no Outlook, nothing written. Says whether the corpus can STILL answer the
///       questions it exists for. A corpus anchored on a fixed date stops filling the narrow
///       measurement windows within weeks, and every test asking about them keeps passing,
///       because selecting nothing is a valid answer about an empty window.
///
///   corpus-reanchor RETIRED (2026-08-25) - prints why and refuses.
///       REBUILDING is the supported way to deal with a stale corpus: it is deterministic, and
///       the recorded build was 20 000 items in 13m25s. Re-anchoring is not to be used until
///       its write path is diagnosed - the date-write method is chosen by a probe that CREATES
///       THROWAWAY items and is then reused, unverified, on already-delivered ones, and a run
///       over 20 000 items once reported total success while dating every one of them inside
///       the six minutes it had been running. The command is kept, not deleted, because its
///       per-item write-landed guard is the thing that stops that repeating.
///
///   corpus-indexed  --population ... --store ... --corpus-id ... --seed N --anchor ...
///                   [--manifest ...] [--wait-seconds N]
///       READ-ONLY, no Outlook: asks the Windows Search index whether every item of a population
///       has a row, waiting up to --wait-seconds for the indexer. Exit 0 only when all do. What the
///       per-run hub rebuild waits on before a live run may start.
///
///   corpus-teardown --store ... --allow-store ... --corpus-id ... --manifest ... [--execute]
///       Removes exactly what the manifest records, by EntryID allowlist AND subject tag.
///
///   corpus-reindex  ... --manifest &lt;path&gt; [--execute]
///       Read-only recovery: rebuilds a candidate manifest by scanning the store.
///
/// Every corpus command refuses any store the caller did not name on --allow-store, any
/// store four independent COM facts do not agree is a local .pst, and any profile that has
/// mail accounts at all. That last one has no override: a build creates unsent items in
/// bulk, and the first real run put 5 532 of them into the target store's Outbox - inert
/// only because that profile could not send.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Exit code for the one failure a caller must not retry: an STA run was abandoned and
    /// the thread did not acknowledge the stop, so a COM session may still be writing to the
    /// store until this process exits. Everything else that fails still exits 1.
    /// </summary>
    internal const int AbandonedStaSessionExitCode = 3;

    private static int Main(string[] args)
    {
        // Every number this console prints is meant to be saved beside a measurement and
        // compared against another machine's run. Under the machine's own culture a
        // Dutch-locale VM writes 426.407.429 where an English one writes 426,407,429 - the
        // same figure and not the same string - so a reader comparing two runs has to work
        // out which convention each was written under. Several call sites already pass
        // CultureInfo.InvariantCulture explicitly and several do not; setting it once here
        // is what makes that impossible to get wrong in a line added later. It also makes
        // every bare int.Parse in the option parser invariant, which is stricter than the
        // machine default rather than looser.
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

        try
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return 1;
            }

            // The corpus commands take their own option parser: --allow-store is repeatable,
            // and that list is the guard deciding whether tens of thousands of items may be
            // written into a mailbox.
            string command = args[0].ToLowerInvariant();
            if (!command.StartsWith("corpus-", StringComparison.Ordinal))
            {
                return Fail($"Unknown command '{args[0]}'.");
            }

            CorpusOptions corpus = CorpusOptions.Parse(args.Skip(1));
            return command switch
            {
                "corpus-plan" => CorpusCommands.RunPlan(corpus, Console.Out),
                "corpus-probe" => CorpusCommands.RunProbe(corpus, Console.Out),
                "corpus-build" => CorpusCommands.RunBuild(corpus, Console.Out),
                "corpus-census" => CorpusCommands.RunCensus(corpus, Console.Out),
                "corpus-verify" => CorpusCommands.RunVerify(corpus, Console.Out),
                "corpus-reanchor" => CorpusCommands.RunReanchor(corpus, Console.Out),
                "corpus-teardown" => CorpusCommands.RunTeardown(corpus, Console.Out),
                "corpus-reindex" => CorpusCommands.RunReindex(corpus, Console.Out),
                "corpus-indexed" => CorpusCommands.RunIndexed(corpus, Console.Out),
                "corpus-folders" => CorpusCommands.RunFolders(corpus, Console.Out),
                _ => Fail($"Unknown command '{args[0]}'."),
            };
        }
        catch (ComStaTimeoutException ex)
        {
            // Its own exit code when the STA thread did NOT stop, because "this failed" and
            // "this failed AND something may still be writing to the store" call for opposite
            // next moves: the first invites a re-run, and a re-run during the second is two
            // writers against one store. Any wrapper testing for non-zero still sees a
            // failure; one that can tell the two apart now can.
            Console.Error.WriteLine($"FATAL: {ToolFailure.Describe(ex)}");
            return ex.Acknowledged ? 1 : AbandonedStaSessionExitCode;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"FATAL: {ToolFailure.Describe(ex)}");
            return 1;
        }
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine($"ERROR: {message}");
        return 1;
    }


    private static void PrintUsage()
    {
        Console.WriteLine("OutlookAI.RemediationTools - the measurement corpus and the test guests' fixture populations (see Program.cs).");
        Console.WriteLine("Measurement corpus: corpus-plan | corpus-probe | corpus-build | corpus-census");
        Console.WriteLine("                    corpus-verify | corpus-teardown | corpus-reindex | corpus-indexed");
        Console.WriteLine("                    corpus-reanchor is RETIRED - rebuild instead; run it to see why");
        Console.WriteLine("Common:   --corpus-id <id> --seed <n> --anchor <yyyy-MM-dd>   [--execute]");
        Console.WriteLine("Target:   --store <display name> --allow-store <display name> (repeatable; a local .pst only)");
        Console.WriteLine("Target:   the profile must have NO mail accounts (no override - see Program.cs)");
        Console.WriteLine("Build:    --count <n> --manifest <path> [--progress-every <n>]");
        Console.WriteLine("Fixture:  --population hub|bystander|identity   a curated test-guest population, not the corpus:");
        Console.WriteLine("          its count is fixed (--count optional), --store is required even by corpus-plan,");
        Console.WriteLine("          and every verb that reads its manifest needs the same --population");
        Console.WriteLine("          [--undated-contacts]  the hub or bystander WITH its undated contacts - the indexed guest's");
        Console.WriteLine("          populations (Q98 f); part of the shape, so every verb on such a manifest needs it too");
        Console.WriteLine("          [--all-kinds]  the hub or bystander WITH appointments, contacts and tasks - the indexed guest's");
        Console.WriteLine("          populations since D62 (b): contacts undated in the index, appointments and tasks dated at");
        Console.WriteLine("          planned instants older than every mail item; part of the shape, exclusive with --undated-contacts");
        Console.WriteLine("Indexed:  corpus-indexed --population ... [--manifest <path>] [--wait-seconds <n>]   (read-only, no Outlook)");
        Console.WriteLine("          corpus-probe --population ... --undated-index-wait <n>   holds each undated probe item up to");
        Console.WriteLine("          n s and prints the index's columns for it before deleting it (a NOT elevated Outlook must run)");
        Console.WriteLine("Folders:  corpus-folders --store <name> --allow-store <name>   (read-only: default folders + visible tree)");
        Console.WriteLine("Verify:   --count <n> --manifest <path> [--window <days> (repeatable)]   (pure - no Outlook)");
        Console.WriteLine("Stale:    rebuild - corpus-teardown --execute (or delete the .pst), then corpus-build");
        Console.WriteLine("Override: [--allow-undated] [--allow-drafts-placement]  (each says what it costs)");
        Console.WriteLine($"Tags:     corpus items carry {CorpusPlan.SubjectTag}, NOT the live tier's artifact tag");
        Console.WriteLine($"Bounds:   cold start {ComStaBudget.Format(ComStaBudget.DefaultStartup)} "
            + $"({ComStaBudget.StartupVariable}), work silence per command ({ComStaBudget.WorkVariable});"
            + " milliseconds, 0 removes the bound");
        Console.WriteLine($"Exit:     0 ok, 1 failed, {AbandonedStaSessionExitCode} an abandoned COM session may still "
            + "be writing - do NOT re-run until this process has exited");
    }
}
