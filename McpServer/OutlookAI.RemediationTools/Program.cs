using System.Globalization;
using System.Text.Json;
using OutlookAI.RemediationTools;

/// <summary>
/// Operator console for the 2026-07-25 incident closure (v3.MD SOAK FIX LOG entry 7;
/// user-approved remediation plan 2026-07-26). Four commands, each dry-run first:
///
///   audit  --settings &lt;live-test-settings.json&gt;
///       Read-only per-store/per-folder counts (total / tagged / untagged).
///
///   refile --settings ... --log &lt;incident-deletion-log.txt&gt; --server &lt;registered exe&gt; [--execute]
///       Step 1: re-derives the origin of every UNTAGGED item in the test hub's
///       Deleted Items (sender-is-hub vs received-by, signals must agree),
///       cross-checks the set against the deletion log's prefix multisets, and only
///       then moves them via the REGISTERED server's move_mail tool over raw stdio
///       (every move audit-logged by the product; no COM writes from this command).
///
///   purge  --settings ... [--execute]
///       Step 2: deletes every item whose subject ordinal-Contains the S3 tag across
///       the three primary stores' Drafts/Inbox/Sent/Deleted Items (+ hub Archive),
///       looping until stable zero. Delegate stores are structurally excluded.
///
///   dedupe --settings ... --store &lt;store&gt; [--execute]
///       Step 3: for every untagged item in the store's Deleted Items, deletes it
///       ONLY when its PR_INTERNET_MESSAGE_ID is non-empty AND an item with the same
///       Message-ID currently exists in that store's Inbox (re-verified per item at
///       delete time); everything else is skipped and reported.
///
/// Logging is S4-disciplined: counts/EntryIDs/booleans only for business stores;
/// subject prefixes appear only for the designated test hub.
///
/// Eight further commands build, check, age and remove a SYNTHETIC MEASUREMENT CORPUS in a
/// local .pst, which is how the freshness-sweep and exhaustive-scan budgets get measured
/// against known volume instead of modelled (see Docs/corpus-measurement-plan.md):
///
///   corpus-plan     --corpus-id ... --seed N --anchor yyyy-MM-dd --count N
///       Pure. Prints what the corpus would contain - per folder, per size class, per age
///       band, and how many items each measurement window selects. No Outlook.
///
///   corpus-probe    --store ... --allow-store ... --corpus-id ... --seed N --anchor ...
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

            string command = args[0].ToLowerInvariant();
            if (command.StartsWith("corpus-", StringComparison.Ordinal))
            {
                // The corpus commands take their own option parser: --allow-store is
                // repeatable, and that list is the guard deciding whether tens of thousands
                // of items may be written into a mailbox.
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
                    _ => Fail($"Unknown command '{args[0]}'."),
                };
            }

            Dictionary<string, string> options = ParseOptions(args.Skip(1));
            bool execute = options.ContainsKey("execute");
            return command switch
            {
                "audit" => RunAudit(LoadSettings(options)),
                "refile" => RunRefile(LoadSettings(options), Require(options, "log"), Require(options, "server"), execute),
                "purge" => RunPurge(LoadSettings(options), execute),
                "dedupe" => RunDedupe(LoadSettings(options), Require(options, "store"), execute),
                _ => Fail($"Unknown command '{args[0]}'."),
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"FATAL: {ex.GetType().Name}: {ex.Message}");
            return 1;
        }
    }

    private sealed record Settings(string HubStore, IReadOnlyList<string> Stores);

    private static Settings LoadSettings(Dictionary<string, string> options)
    {
        string path = Require(options, "settings");
        using JsonDocument doc = JsonDocument.Parse(File.ReadAllText(path));
        string hub = doc.RootElement.GetProperty("testHubStoreDisplayName").GetString()
            ?? throw new InvalidOperationException("testHubStoreDisplayName missing in settings.");
        List<string> stores = doc.RootElement.GetProperty("expectedStoreDisplayNames").EnumerateArray()
            .Select(e => e.GetString() ?? throw new InvalidOperationException("Null store name in settings."))
            .ToList();
        if (!stores.Contains(hub, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Settings hub store is not among expectedStoreDisplayNames.");
        }

        return new Settings(hub, stores);
    }

    private static int RunAudit(Settings settings)
    {
        Console.WriteLine("== Remediation audit (read-only) ==");
        Console.WriteLine($"{"store",-28} {"folder",-16} {"total",7} {"tagged",7} {"untagged",9}");
        foreach (string store in settings.Stores)
        {
            int[] folders = IsHub(settings, store) ? ComMailbox.HubSweepFolderIdsWithArchive : ComMailbox.SweepFolderIds;
            foreach (ComMailbox.FolderCounts c in ComMailbox.CountStoreFolders(store, folders))
            {
                Console.WriteLine($"{c.Store,-28} {c.FolderName,-16} {c.Total,7} {c.Tagged,7} {c.Total - c.Tagged,9}");
            }
        }

        return 0;
    }

    private static int RunRefile(Settings settings, string logPath, string serverExe, bool execute)
    {
        string hub = settings.HubStore;
        Console.WriteLine($"== Telefonie refile ({(execute ? "EXECUTE" : "dry-run")}) ==");

        IReadOnlyList<RemediationRules.DeletionLogEntry> log =
            RemediationRules.ParseDeletionLog(File.ReadAllLines(logPath));
        List<string> expectedSent = RemediationRules.ExpectedPrefixes(log, hub, 5);
        List<string> expectedInbox = RemediationRules.ExpectedPrefixes(log, hub, 6);
        Console.WriteLine($"Deletion log: {log.Count} lines; expected hub Sent-origin {expectedSent.Count}, Inbox-origin {expectedInbox.Count}.");

        List<ComMailbox.DeletedItemInfo> all = ComMailbox.ListDeletedItems(hub, withMessageIds: false, withSender: true);
        List<ComMailbox.DeletedItemInfo> untagged = all.Where(i => !i.Tagged).ToList();
        Console.WriteLine($"Hub Deleted Items: total {all.Count}, tagged {all.Count - untagged.Count}, untagged {untagged.Count}.");

        var sentIds = new List<string>();
        var inboxIds = new List<string>();
        var remainingSent = new List<string>(expectedSent);
        var remainingInbox = new List<string>(expectedInbox);
        bool ok = true;
        foreach (ComMailbox.DeletedItemInfo item in untagged)
        {
            RemediationRules.TelefonieOrigin? origin =
                RemediationRules.ClassifyOrigin(item.SenderSmtp, item.ReceivedByPresent, hub);
            if (origin == null)
            {
                Console.WriteLine($"  AMBIGUOUS {Short(item.EntryId)} senderIsHub/receivedBy signals disagree - ABORT.");
                ok = false;
                continue;
            }

            List<string> pool = origin == RemediationRules.TelefonieOrigin.SentOrigin ? remainingSent : remainingInbox;
            string? consumed = RemediationRules.TryConsumePrefixMatch(pool, item.Subject);
            if (consumed == null)
            {
                Console.WriteLine($"  UNMATCHED {Short(item.EntryId)} class={origin} subject8='{Prefix8(item.Subject)}' has no remaining log prefix - ABORT.");
                ok = false;
                continue;
            }

            Console.WriteLine($"  {origin,-11} prefix='{consumed}' senderIsHub={origin == RemediationRules.TelefonieOrigin.SentOrigin} receivedBy={item.ReceivedByPresent} id={Short(item.EntryId)}");
            (origin == RemediationRules.TelefonieOrigin.SentOrigin ? sentIds : inboxIds).Add(item.EntryId);
        }

        if (remainingSent.Count > 0 || remainingInbox.Count > 0)
        {
            Console.WriteLine($"  LOG PREFIXES UNCONSUMED: Sent-origin [{string.Join(", ", remainingSent)}], Inbox-origin [{string.Join(", ", remainingInbox)}] - ABORT.");
            ok = false;
        }

        Console.WriteLine($"Classified: {sentIds.Count} Sent-origin, {inboxIds.Count} Inbox-origin; cross-check {(ok ? "PASSED" : "FAILED")}.");
        if (!ok)
        {
            return 1;
        }

        if (!execute)
        {
            Console.WriteLine("Dry-run complete; nothing moved.");
            return 0;
        }

        string sentFolderName = ComMailbox.GetDefaultFolderName(hub, 5);
        string inboxFolderName = ComMailbox.GetDefaultFolderName(hub, 6);
        Console.WriteLine($"Move targets: Sent-origin -> '{sentFolderName}', Inbox-origin -> '{inboxFolderName}' (store '{hub}').");

        using var client = McpMoveClient.StartAndInitialize(serverExe);
        IReadOnlyList<string> tools = client.ListToolNames();
        Console.WriteLine($"Registered server: {tools.Count} tools; move_mail advertised: {tools.Contains("move_mail")}.");
        if (!tools.Contains("move_mail"))
        {
            return Fail("move_mail is not advertised by the server - wrong exe?");
        }

        bool moveOk = CallMove(client, sentIds, sentFolderName, hub)
            & CallMove(client, inboxIds, inboxFolderName, hub);
        bool cleanExit = client.CloseAndAwaitExit(TimeSpan.FromSeconds(30));
        Console.WriteLine($"Server exit on stdin EOF: {cleanExit}.");

        List<ComMailbox.DeletedItemInfo> after = ComMailbox.ListDeletedItems(hub, withMessageIds: false, withSender: false);
        int untaggedAfter = after.Count(i => !i.Tagged);
        Console.WriteLine($"Post-move hub Deleted Items: total {after.Count}, untagged {untaggedAfter} (expected 0).");
        return moveOk && untaggedAfter == 0 ? 0 : 1;
    }

    private static bool CallMove(McpMoveClient client, List<string> ids, string folder, string store)
    {
        if (ids.Count == 0)
        {
            return true;
        }

        JsonElement outcome = client.CallTool("move_mail", new { ids, folder, store });
        if (outcome.TryGetProperty("error", out JsonElement error))
        {
            Console.WriteLine($"  move_mail -> '{folder}': DOMAIN ERROR {error.GetRawText()}");
            return false;
        }

        int requested = outcome.GetProperty("requested").GetInt32();
        int moved = outcome.GetProperty("moved").GetInt32();
        int failed = outcome.GetProperty("failed").GetInt32();
        Console.WriteLine($"  move_mail -> '{folder}': requested {requested}, moved {moved}, failed {failed}.");
        foreach (JsonElement item in outcome.GetProperty("items").EnumerateArray())
        {
            bool itemOk = item.TryGetProperty("ok", out JsonElement okProp) && okProp.ValueKind == JsonValueKind.True;
            if (!itemOk)
            {
                string? err = item.TryGetProperty("error", out JsonElement e) ? e.GetString() : null;
                Console.WriteLine($"    FAILED {Short(item.GetProperty("id").GetString() ?? "?")}: {err}");
            }
        }

        return failed == 0 && moved == requested;
    }

    private static int RunPurge(Settings settings, bool execute)
    {
        Console.WriteLine($"== Tagged purge ({(execute ? "EXECUTE" : "dry-run")}) - ordinal '{RemediationRules.SubjectTag}' ==");
        bool allZero = true;
        foreach (string store in settings.Stores)
        {
            int[] folders = IsHub(settings, store) ? ComMailbox.HubSweepFolderIdsWithArchive : ComMailbox.SweepFolderIds;
            if (!execute)
            {
                foreach (ComMailbox.PurgeFolderResult r in ComMailbox.PurgeTaggedPass(store, folders, execute: false))
                {
                    Console.WriteLine($"  {r.Store,-28} {r.FolderName,-16} tagged {r.Matched,5}");
                }

                continue;
            }

            var totals = new Dictionary<string, int>(StringComparer.Ordinal);
            int pass = 0;

            // Stopwatch, not DateTime.UtcNow: both of these measure how long something has
            // been going on inside this run, so they must be read from a clock that only
            // moves forward. On the wall clock a backwards jump mid-purge extends the
            // 8-minute cap by the size of the jump and re-arms the stability window, and a
            // forwards jump ends the purge early with tagged items still in the mailbox -
            // which this tool then reports as a non-zero remaining count and a failure exit.
            System.Diagnostics.Stopwatch elapsed = System.Diagnostics.Stopwatch.StartNew();
            TimeSpan deadline = TimeSpan.FromMinutes(8);
            TimeSpan? zeroSince = null;
            while (elapsed.Elapsed < deadline)
            {
                pass++;
                List<ComMailbox.PurgeFolderResult> results = ComMailbox.PurgeTaggedPass(store, folders, execute: true);
                int deletedThisPass = results.Sum(r => r.Deleted);
                int failedThisPass = results.Sum(r => r.Failed);
                foreach (ComMailbox.PurgeFolderResult r in results.Where(r => r.Deleted > 0))
                {
                    totals[r.FolderName] = totals.TryGetValue(r.FolderName, out int prior) ? prior + r.Deleted : r.Deleted;
                    Console.WriteLine($"  pass {pass}: {r.Store,-28} {r.FolderName,-16} deleted {r.Deleted,5} (failed {r.Failed})");
                }

                if (failedThisPass > 0)
                {
                    Console.WriteLine($"  pass {pass}: {failedThisPass} delete failures (retried next pass).");
                }

                if (deletedThisPass == 0 && ComMailbox.CountTaggedInFolders(store, folders) == 0)
                {
                    zeroSince ??= elapsed.Elapsed;
                    if (elapsed.Elapsed - zeroSince.Value >= TimeSpan.FromSeconds(10))
                    {
                        break; // stable zero
                    }

                    Thread.Sleep(2000);
                    continue;
                }

                zeroSince = null;
            }

            int remaining = ComMailbox.CountTaggedInFolders(store, folders);
            allZero &= remaining == 0;
            Console.WriteLine($"  {store}: deleted {totals.Values.Sum()} deletions across {pass} passes "
                + $"[{string.Join(", ", totals.Select(t => $"{t.Key}={t.Value}"))}]; tagged remaining {remaining}.");
        }

        return !execute || allZero ? 0 : 1;
    }

    private static int RunDedupe(Settings settings, string store, bool execute)
    {
        if (!settings.Stores.Contains(store, StringComparer.OrdinalIgnoreCase))
        {
            return Fail("dedupe --store must be one of the three primary stores (delegates are permanently excluded).");
        }

        Console.WriteLine($"== Inbox-twin duplicate delete ({(execute ? "EXECUTE" : "dry-run")}) - store '{store}' ==");
        (HashSet<string> inboxIds, int inboxTotal, int inboxWithoutId) = ComMailbox.CollectInboxMessageIds(store);
        Console.WriteLine($"Inbox twin set: {inboxTotal} items, {inboxIds.Count} distinct Message-IDs, {inboxWithoutId} without one.");

        List<ComMailbox.DeletedItemInfo> deletedItems = ComMailbox.ListDeletedItems(store, withMessageIds: true, withSender: false);
        Console.WriteLine($"Deleted Items: total {deletedItems.Count}, tagged {deletedItems.Count(i => i.Tagged)}, untagged {deletedItems.Count(i => !i.Tagged)}.");

        var byDecision = new Dictionary<RemediationRules.DedupeDecision, int>();
        var toDelete = new List<string>();
        foreach (ComMailbox.DeletedItemInfo item in deletedItems)
        {
            RemediationRules.DedupeDecision decision =
                RemediationRules.DecideDuplicateDelete(item.Subject, item.InternetMessageId, inboxIds);
            byDecision[decision] = byDecision.TryGetValue(decision, out int prior) ? prior + 1 : 1;
            if (decision == RemediationRules.DedupeDecision.Delete)
            {
                toDelete.Add(item.EntryId);
            }
            else
            {
                Console.WriteLine($"  SKIP {decision} id={Short(item.EntryId)}");
            }
        }

        Console.WriteLine("Snapshot verification: "
            + string.Join(", ", byDecision.OrderBy(kv => kv.Key).Select(kv => $"{kv.Key}={kv.Value}")));
        if (!execute)
        {
            Console.WriteLine($"Dry-run complete; {toDelete.Count} items verified deletable, nothing deleted.");
            return 0;
        }

        int deleted = 0;
        int skipped = 0;
        int errors = 0;
        foreach (string entryId in toDelete)
        {
            ComMailbox.DedupeItemResult result = ComMailbox.DeleteVerifiedDuplicate(store, entryId, inboxIds, execute: true);
            if (result.Deleted)
            {
                deleted++;
            }
            else if (result.Error != null)
            {
                errors++;
                Console.WriteLine($"  NOT DELETED id={Short(entryId)}: {result.Error}");
            }
            else
            {
                skipped++;
                Console.WriteLine($"  RE-VERIFY SKIP {result.Decision} id={Short(entryId)}");
            }
        }

        Console.WriteLine($"Executed: deleted {deleted}, re-verify skips {skipped}, errors {errors} (of {toDelete.Count} verified candidates).");
        List<ComMailbox.DeletedItemInfo> after = ComMailbox.ListDeletedItems(store, withMessageIds: false, withSender: false);
        Console.WriteLine($"Post-state Deleted Items: total {after.Count}, tagged {after.Count(i => i.Tagged)}, untagged {after.Count(i => !i.Tagged)}.");
        return errors == 0 ? 0 : 1;
    }

    private static bool IsHub(Settings settings, string store)
        => string.Equals(store, settings.HubStore, StringComparison.OrdinalIgnoreCase);

    private static string Short(string entryId)
        => entryId.Length <= 16 ? entryId : entryId.Substring(entryId.Length - 16);

    private static string Prefix8(string? subject)
    {
        string trimmed = (subject ?? string.Empty).TrimStart();
        return trimmed.Length <= 8 ? trimmed : trimmed.Substring(0, 8);
    }

    private static Dictionary<string, string> ParseOptions(IEnumerable<string> args)
    {
        var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string? pending = null;
        foreach (string arg in args)
        {
            if (arg.StartsWith("--", StringComparison.Ordinal))
            {
                if (pending != null)
                {
                    options[pending] = "true";
                }

                pending = arg.Substring(2);
            }
            else if (pending != null)
            {
                options[pending] = arg;
                pending = null;
            }
            else
            {
                throw new ArgumentException($"Unexpected argument '{arg}'.");
            }
        }

        if (pending != null)
        {
            options[pending] = "true";
        }

        return options;
    }

    private static string Require(Dictionary<string, string> options, string name)
        => options.TryGetValue(name, out string? value) && value != "true"
            ? value
            : throw new ArgumentException($"--{name} <value> is required.");

    private static int Fail(string message)
    {
        Console.Error.WriteLine($"ERROR: {message}");
        return 1;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("OutlookAI.RemediationTools - operator console (see Program.cs doc comment).");
        Console.WriteLine("Incident-7 closure: audit | refile | purge | dedupe");
        Console.WriteLine("Common:   --settings <live-test-settings.json>   [--execute]");
        Console.WriteLine("refile:   --log <incident-deletion-log.txt> --server <registered OutlookAI.McpServer.exe>");
        Console.WriteLine("dedupe:   --store <primary store display name>");
        Console.WriteLine();
        Console.WriteLine("Measurement corpus: corpus-plan | corpus-probe | corpus-build | corpus-census");
        Console.WriteLine("                    corpus-verify | corpus-teardown | corpus-reindex");
        Console.WriteLine("                    corpus-reanchor is RETIRED - rebuild instead; run it to see why");
        Console.WriteLine("Common:   --corpus-id <id> --seed <n> --anchor <yyyy-MM-dd>   [--execute]");
        Console.WriteLine("Target:   --store <display name> --allow-store <display name> (repeatable; a local .pst only)");
        Console.WriteLine("Target:   the profile must have NO mail accounts (no override - see Program.cs)");
        Console.WriteLine("Build:    --count <n> --manifest <path> [--progress-every <n>]");
        Console.WriteLine("Verify:   --count <n> --manifest <path> [--window <days> (repeatable)]   (pure - no Outlook)");
        Console.WriteLine("Stale:    rebuild - corpus-teardown --execute (or delete the .pst), then corpus-build");
        Console.WriteLine("Override: [--allow-undated] [--allow-drafts-placement]  (each says what it costs)");
        Console.WriteLine($"Tags:     corpus items carry {CorpusPlan.SubjectTag}, NOT the live tier's artifact tag");
    }
}
