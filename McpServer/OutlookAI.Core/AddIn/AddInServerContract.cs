namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// THE ADD-IN / MAIL SERVER CONTRACT: the values the add-in writes and the mail server reads.
    ///
    /// Everything the two agree about is DATA: two HKCU keys the add-in writes and
    /// <c>outlook_health</c> reads, plus the Claude Code configuration file one of them repairs and
    /// the other reports on. Every value below used
    /// to be spelled out once per side, with a comment on the server saying it "mirrors" the
    /// add-in. A comment is not a mechanism - the audit in <c>Docs/magic-numbers.md</c> found one
    /// of those comments that had already become false - and this particular drift is silent:
    /// both halves go on working, they simply stop describing the same machine, so the health
    /// report says "tuning: not managed" about a machine the add-in has been tuning for months.
    ///
    /// One definition, in OutlookAI.Core's add-in namespace (Q163 (a), Q164 (a)): the add-in
    /// references Core's net48 build and takes these constants from here (Q167 (a)), and
    /// <c>HealthReporting</c> reads exactly what the add-in wrote. Until the code-sharing refactor's
    /// stage S3 the add-in compiled this file and Core linked it, internal in both.
    ///
    /// <para>
    /// Constants only: no probes, no registry access, no logic, nothing to test. The add-in compiles
    /// as C# 15 on .NET Framework 4.8 (C# 7.3 until Q172, 2026-11-18), so nothing here may need the
    /// modern runtime to use.
    /// </para>
    /// </summary>
    public static class AddInServerContract
    {
        /// <summary>
        /// The add-in's own key, under HKEY_CURRENT_USER: every key below is under it, and the
        /// add-in keeps its own settings in it (the installer's InstallDir, the chosen model). One
        /// spelling, where it used to be typed into four classes (Q154).
        /// </summary>
        public const string AppKeyPath = @"Software\OutlookAI";

        // ===== HKCU\Software\OutlookAI\Tuning =====
        // Written by OutlookTuningService (desired state, group toggles, reconcile bookkeeping);
        // read by HealthReporting.ReadTuningState for the tuning block of outlook_health.

        /// <summary>
        /// The tuning key, under HKEY_CURRENT_USER - the add-in's own state, which it can always
        /// write. (Since Q145 (a) every Outlook value it manages is the user's own setting too; one
        /// Windows refuses anyway is listed in <see cref="TuningNeedsAdministratorValueName"/>.)
        /// </summary>
        public const string TuningKeyPath = AppKeyPath + @"\Tuning";

        /// <summary>
        /// DWORD, written once when the add-in first seeds the toggles. Its ABSENCE is
        /// load-bearing on the server side: it is the difference between "tuning is off" and
        /// "the add-in has never run its tuning service on this machine", which health reports
        /// as unmanaged rather than as disabled.
        /// </summary>
        public const string TuningInitializedValueName = "Initialized";

        /// <summary>DWORD master switch for the whole tuning feature.</summary>
        public const string TuningEnabledValueName = "Enabled";

        /// <summary>DWORD per-group toggles, one per group id in OutlookTuningService's catalog.</summary>
        public const string TuningSearchEnabledValueName = "SearchEnabled";
        public const string TuningCachingEnabledValueName = "CachingEnabled";
        public const string TuningOstEnabledValueName = "OstEnabled";

        /// <summary>DWORD: a write has landed that only takes effect on the NEXT Outlook start.</summary>
        public const string TuningRestartNeededValueName = "RestartNeeded";

        /// <summary>
        /// String: ';'-joined entry ids the last reconcile left alone because a policy sets their
        /// value - Microsoft 365 Cloud Policy or Group Policy, an organization's or one an earlier
        /// OutlookAI wrote (Q145 (a)). Until then: the ids it backed off from because a policy
        /// reverted a value it had applied.
        /// </summary>
        public const string TuningPolicyConflictsValueName = "PolicyConflicts";

        /// <summary>
        /// String: ';'-joined entry ids whose write the last reconcile was REFUSED for lack of rights,
        /// and skipped (Q128, 2026-10-03). Then D25's five Cached Mode values under
        /// <c>HKCU\Software\Policies</c>, which a NOT elevated Outlook may read and not write; since
        /// Q145 (a) the add-in writes only the user's own settings, so it is normally empty. Empty
        /// when there are none. <c>outlook_health</c> reports it as <c>tuning.needsAdministrator</c>.
        /// </summary>
        public const string TuningNeedsAdministratorValueName = "NeedsAdministrator";

        /// <summary>
        /// String: ISO 8601 ("o") timestamp of the last reconcile. Also the add-in's start-up stamp for an
        /// add-in older than <see cref="AddInStartupCompleteUtcValueName"/>: the reconcile runs during every
        /// start of the add-in, so a value no older than the running Outlook says its add-in has started.
        /// </summary>
        public const string TuningLastReconcileUtcValueName = "LastReconcileUtc";

        // ===== HKCU\Software\OutlookAI\AddIn =====
        // Written by ThisAddIn at the end of every start of the add-in inside Outlook; read by the mail
        // server, which opens no mailbox in an Outlook that has just started until this says the add-in
        // has finished starting there (Q175 (a), decided by the maintainer 2026-11-18, after the hang of
        // 2026-11-15: an Outlook started for an agent deadlocked while its add-ins were still starting).

        /// <summary>The add-in's own start-up key, under HKEY_CURRENT_USER.</summary>
        public const string AddInKeyPath = @"Software\OutlookAI\AddIn";

        /// <summary>DWORD: the process id of the OUTLOOK.EXE whose add-in finished starting last.</summary>
        public const string AddInStartupProcessIdValueName = "StartupProcessId";

        /// <summary>
        /// String: ISO 8601 ("o") UTC time that start finished. The pair is this Outlook's stamp only when
        /// the id is this Outlook's and the time is not before it started, so a reader that comes between
        /// the add-in's two writes, or after a process id was reused, sees no stamp rather than a wrong one.
        /// </summary>
        public const string AddInStartupCompleteUtcValueName = "StartupCompleteUtc";

        // ===== HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI =====
        // Outlook's own registration of the add-in, written by the installer (Installer.iss). The server
        // reads it to know whether a start-up stamp is coming at all: with the add-in not set to load, it
        // does not wait for one.

        /// <summary>The add-in's registration with Outlook, under HKEY_CURRENT_USER.</summary>
        public const string OutlookAddInRegistrationKeyPath = @"Software\Microsoft\Office\Outlook\Addins\OutlookAI";

        /// <summary>DWORD: how Outlook loads the add-in.</summary>
        public const string OutlookAddInLoadBehaviorValueName = "LoadBehavior";

        /// <summary>The <see cref="OutlookAddInLoadBehaviorValueName"/> that loads it at every Outlook start.</summary>
        public const int OutlookAddInLoadAtStartup = 3;

        // ===== HKCU\Software\Microsoft\Office\<major>\Outlook\Search =====
        // OUTLOOK'S OWN key, not one of ours, and the only entry here that is not: the add-in's
        // tuning Search group (D22) WRITES the value below, and HealthReporting READS it back as
        // outlook_health's tuning.uiSearchBackend - which show_search_results then uses to decide
        // whether to tell an agent that the user's Outlook search box is looking somewhere other
        // than the corpus the agent searched. The KEY PATH is version-dependent and is therefore
        // built by OfficeVersions.OutlookSearchKeyPath, beside this file; only the VALUE NAME is a
        // constant, and it lives here with the rest of them.

        /// <summary>
        /// DWORD under the Outlook Search key: nonzero points Outlook's own search box at the
        /// local Windows Search index (the corpus the agent's <c>search</c> queries), absent or 0
        /// leaves it on Outlook's server-assisted default.
        /// <para>
        /// It is Outlook's name, not ours - we cannot rename it, only mistype it - and mistyping
        /// it is silent on both sides: the add-in writes a value Outlook ignores, or the server
        /// reads a value nobody wrote and reports <c>server-assisted</c> on a tuned machine for
        /// ever. It was spelled out six times before this constant existed (once in the add-in's
        /// tuning catalog, twice in <c>HealthReporting</c>, three times in the T2 live test).
        /// <c>Tools/Checks/check-pinned-constants.ps1</c> (#12) keeps it at one, and pins that
        /// one against the name <c>McpServer/README.md</c> publishes to users.
        /// </para>
        /// </summary>
        public const string DisableServerAssistedSearchValueName = "DisableServerAssistedSearch";

        // ===== HKCU\Software\OutlookAI\Mcp =====
        // Written by McpRegistrationService on every reconcile; read by
        // HealthReporting.ReadMcpRegistration so outlook_health can report what the add-in did
        // instead of guessing. This key IS the contract - see McpServer/README.md.

        /// <summary>The MCP registration key, under HKEY_CURRENT_USER.</summary>
        public const string McpKeyPath = AppKeyPath + @"\Mcp";

        /// <summary>String: one of McpRegistrationService's status codes.</summary>
        public const string McpStatusValueName = "Status";

        /// <summary>String: the human-readable detail behind that status (shown in Settings).</summary>
        public const string McpDetailValueName = "Detail";

        /// <summary>String: the command currently registered in the Claude Code configuration.</summary>
        public const string McpCommandValueName = "Command";

        /// <summary>String: the installed server path the reconcile resolved, when it found one.</summary>
        public const string McpResolvedServerPathValueName = "ResolvedServerPath";

        /// <summary>DWORD: whether that reconcile had to rewrite the registration.</summary>
        public const string McpHealedValueName = "Healed";

        /// <summary>String: ISO 8601 ("o") timestamp of the last reconcile.</summary>
        public const string McpLastReconcileUtcValueName = "LastReconcileUtc";

        /// <summary>
        /// Tri-state DWORD opt-in for user-scope registration: 1 on, 0 off, ABSENT never decided.
        /// The absent state is what separates "the user does not want this" from "the user has
        /// not been asked yet", and only the first may be acted on.
        /// </summary>
        public const string McpGlobalRegistrationValueName = "GlobalRegistrationEnabled";

        // ===== Claude Code's configuration =====
        // The add-in writes this file (McpConfigEditor's splice discipline); the server only ever
        // reads it, to answer "is this executable the one Claude Code is configured to spawn?".
        // Read-side and write-side therefore have to agree on the file and on three names in it.

        /// <summary>
        /// User-scope configuration file name. Both sides resolve it against
        /// <c>Environment.SpecialFolder.UserProfile</c>, so the whole path is the same on both
        /// sides once this name is.
        /// </summary>
        public const string ClaudeConfigFileName = ".claude.json";

        /// <summary>The top-level object holding the servers, in both the user- and project-scope shapes.</summary>
        public const string ServersProperty = "mcpServers";

        /// <summary>Our member under <see cref="ServersProperty"/>. Renaming it deregisters the server on every machine.</summary>
        public const string ServerName = "outlookai";

        /// <summary>The stdio server's executable, the one member of our entry both sides read.</summary>
        public const string CommandProperty = "command";
    }
}
