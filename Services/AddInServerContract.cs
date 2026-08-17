namespace OutlookAI.Services
{
    /// <summary>
    /// THE ADD-IN / MAIL SERVER CONTRACT: the values that must be identical on both sides of a
    /// boundary no compiler crosses.
    ///
    /// The net48 VSTO add-in cannot reference the .NET 10 mail server, and the server
    /// deliberately takes no add-in dependency, so everything they agree about is DATA: two HKCU
    /// keys the add-in writes and <c>outlook_health</c> reads, plus the Claude Code
    /// configuration file one of them repairs and the other reports on. Every value below used
    /// to be spelled out once per side, with a comment on the server saying it "mirrors" the
    /// add-in. A comment is not a mechanism - the audit in <c>Docs/magic-numbers.md</c> found one
    /// of those comments that had already become false - and this particular drift is silent:
    /// both halves go on working, they simply stop describing the same machine, so the health
    /// report says "tuning: not managed" about a machine the add-in has been tuning for months.
    ///
    /// One definition, compiled into both trees instead of compared across them:
    ///  - the add-in (<c>OutlookAI.csproj</c>) compiles this file directly;
    ///  - <c>OutlookAI.Core</c> LINKS it, which is how <c>HealthReporting</c> reads exactly what
    ///    the add-in wrote;
    ///  - the test project LINKS it as well, because it links <see cref="McpConfigEditor"/>,
    ///    which uses the Claude Code names below.
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL, and the intersection is narrow: this compiles as net48 (the add-in),
    /// as net48 AND net10 (Core's dual target) and as net10 (the test host), the last two
    /// nullable-enabled with warnings as errors. So no <c>string?</c>, no <c>return null</c> from
    /// a string member, no target-typed <c>new</c>, and an event would need an initialiser.
    /// Holding nothing but constants is the cheapest way to stay inside that intersection, and
    /// it is all this file is for: no probes, no registry access, no logic, nothing to test.
    /// </para>
    ///
    /// <para>
    /// INTERNAL, and it has to stay internal. A PUBLIC type in a linked file compiled into two
    /// assemblies that can see each other is CS0436, which is an error here
    /// (<c>TreatWarningsAsErrors</c>) - that is why the test project stopped linking
    /// <c>PromptStore.cs</c>. Core's copy is invisible to the test assembly precisely BECAUSE it
    /// is internal, so the two copies cannot collide. Whatever the server has to expose stays
    /// exposed by <c>HealthReporting</c>'s own public constants, which are initialised from here.
    /// For the same reason this file is NOT compiled into <c>OutlookAI.McpServer</c>: that
    /// project opens its internals to the test assembly, so a copy there WOULD collide with the
    /// linked one.
    /// </para>
    /// </summary>
    internal static class AddInServerContract
    {
        // ===== HKCU\Software\OutlookAI\Tuning =====
        // Written by OutlookTuningService (desired state, group toggles, reconcile bookkeeping);
        // read by HealthReporting.ReadTuningState for the tuning block of outlook_health.

        /// <summary>The tuning key, under HKEY_CURRENT_USER. No elevation anywhere in this feature.</summary>
        internal const string TuningKeyPath = @"Software\OutlookAI\Tuning";

        /// <summary>
        /// DWORD, written once when the add-in first seeds the toggles. Its ABSENCE is
        /// load-bearing on the server side: it is the difference between "tuning is off" and
        /// "the add-in has never run its tuning service on this machine", which health reports
        /// as unmanaged rather than as disabled.
        /// </summary>
        internal const string TuningInitializedValueName = "Initialized";

        /// <summary>DWORD master switch for the whole tuning feature.</summary>
        internal const string TuningEnabledValueName = "Enabled";

        /// <summary>DWORD per-group toggles, one per group id in OutlookTuningService's catalog.</summary>
        internal const string TuningSearchEnabledValueName = "SearchEnabled";
        internal const string TuningCachingEnabledValueName = "CachingEnabled";
        internal const string TuningOstEnabledValueName = "OstEnabled";

        /// <summary>DWORD: a write has landed that only takes effect on the NEXT Outlook start.</summary>
        internal const string TuningRestartNeededValueName = "RestartNeeded";

        /// <summary>String: ';'-joined entry ids the reconciler backed off from because a policy reverted them.</summary>
        internal const string TuningPolicyConflictsValueName = "PolicyConflicts";

        /// <summary>String: ISO 8601 ("o") timestamp of the last reconcile.</summary>
        internal const string TuningLastReconcileUtcValueName = "LastReconcileUtc";

        // ===== HKCU\Software\OutlookAI\Mcp =====
        // Written by McpRegistrationService on every reconcile; read by
        // HealthReporting.ReadMcpRegistration so outlook_health can report what the add-in did
        // instead of guessing. This key IS the contract - see McpServer/README.md.

        /// <summary>The MCP registration key, under HKEY_CURRENT_USER.</summary>
        internal const string McpKeyPath = @"Software\OutlookAI\Mcp";

        /// <summary>String: one of McpRegistrationService's status codes.</summary>
        internal const string McpStatusValueName = "Status";

        /// <summary>String: the human-readable detail behind that status (shown in Settings).</summary>
        internal const string McpDetailValueName = "Detail";

        /// <summary>String: the command currently registered in the Claude Code configuration.</summary>
        internal const string McpCommandValueName = "Command";

        /// <summary>String: the installed server path the reconcile resolved, when it found one.</summary>
        internal const string McpResolvedServerPathValueName = "ResolvedServerPath";

        /// <summary>DWORD: whether that reconcile had to rewrite the registration.</summary>
        internal const string McpHealedValueName = "Healed";

        /// <summary>String: ISO 8601 ("o") timestamp of the last reconcile.</summary>
        internal const string McpLastReconcileUtcValueName = "LastReconcileUtc";

        /// <summary>
        /// Tri-state DWORD opt-in for user-scope registration: 1 on, 0 off, ABSENT never decided.
        /// The absent state is what separates "the user does not want this" from "the user has
        /// not been asked yet", and only the first may be acted on.
        /// </summary>
        internal const string McpGlobalRegistrationValueName = "GlobalRegistrationEnabled";

        // ===== Claude Code's configuration =====
        // The add-in writes this file (McpConfigEditor's splice discipline); the server only ever
        // reads it, to answer "is this executable the one Claude Code is configured to spawn?".
        // Read-side and write-side therefore have to agree on the file and on three names in it.

        /// <summary>
        /// User-scope configuration file name. Both sides resolve it against
        /// <c>Environment.SpecialFolder.UserProfile</c>, so the whole path is the same on both
        /// sides once this name is.
        /// </summary>
        internal const string ClaudeConfigFileName = ".claude.json";

        /// <summary>The top-level object holding the servers, in both the user- and project-scope shapes.</summary>
        internal const string ServersProperty = "mcpServers";

        /// <summary>Our member under <see cref="ServersProperty"/>. Renaming it deregisters the server on every machine.</summary>
        internal const string ServerName = "outlookai";

        /// <summary>The stdio server's executable, the one member of our entry both sides read.</summary>
        internal const string CommandProperty = "command";
    }
}
