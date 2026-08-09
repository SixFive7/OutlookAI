using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// "OutlookAI Settings" — a small theme-aware dialog over OutlookTuningService: master and
    /// per-group toggles, the current effective registry values, a restart-needed indicator,
    /// and a flag line for values group policy has taken back. Opened modeless from the
    /// Explorer ribbon button (single instance) and from the COM automation hook; all calls
    /// arrive on Outlook's UI thread.
    /// </summary>
    public class SettingsDialog : Form
    {
        private static SettingsDialog _open;

        private readonly CheckBox chkMaster;
        private readonly CheckBox chkSearch;
        private readonly CheckBox chkCaching;
        private readonly CheckBox chkOst;
        private readonly CheckBox chkGlobalMcp;
        private readonly GroupBox grpSearch;
        private readonly GroupBox grpCaching;
        private readonly GroupBox grpOst;
        private readonly GroupBox grpClaude;
        private readonly Label lblHeader;
        private readonly Label lblSearchValues;
        private readonly Label lblSearchWarning;
        private readonly Label lblCachingValues;
        private readonly Label lblOstValues;
        private readonly Label lblRestart;
        private readonly Label lblGpo;
        private readonly Label lblGlobalMcpHelp;
        private readonly Label lblMcp;
        private readonly Button btnAddProject;
        private readonly Button btnCopyCommand;
        private readonly Button btnApply;
        private readonly Button btnClose;

        private bool _updating;
        private bool _disposedCustom;

        /// <summary>
        /// What a manual registration would name, refreshed by <see cref="RefreshFromState"/>.
        /// Cached in a field so the theme handler can redraw the status line without probing
        /// the disk again.
        /// </summary>
        private string _preferredCommand = "";

        /// <summary>
        /// The server's real path. The copy button uses this rather than the portable
        /// <c>${LOCALAPPDATA}</c> spelling on purpose: PowerShell expands <c>${NAME}</c>
        /// itself — quoted or not — so a copied command carrying that form would arrive at
        /// the CLI with the path blanked out. Claude Code expands it when it READS the
        /// config, which is why the file the button writes can use it and a shell command
        /// cannot.
        /// </summary>
        private string _resolvedServerPath = "";

        internal static bool IsOpen
        {
            get { return _open != null && !_open.IsDisposed; }
        }

        internal static void ShowSettings()
        {
            // Callers must be on the UI thread (ribbon callbacks are; the COM automation
            // surface marshals via the add-in's UI-thread control before calling this).
            SettingsDialog dlg = null;
            try
            {
                if (IsOpen)
                {
                    _open.Activate();
                    return;
                }
                dlg = new SettingsDialog();
                dlg.FormClosed += (s, e) => { if (ReferenceEquals(_open, dlg)) _open = null; };
                dlg.Show();
                dlg.Activate();
                _open = dlg;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowSettings: " + ex.Message);
                // Never leave a half-shown zombie registered as "open".
                if (ReferenceEquals(_open, dlg))
                    _open = null;
                try { dlg?.Dispose(); } catch { }
            }
        }

        internal static void CloseIfOpen()
        {
            try
            {
                if (IsOpen)
                    _open.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CloseIfOpen: " + ex.Message);
            }
        }

        /// <summary>
        /// Repaints an open dialog from stored state. Called when something OUTSIDE it changed
        /// the registration — the startup prompt being answered — so the tick box and the
        /// status line can never sit there contradicting what was just chosen. UI thread; a
        /// no-op when the dialog is closed. Never throws.
        /// </summary>
        internal static void RefreshIfOpen()
        {
            try
            {
                if (IsOpen)
                    _open.RefreshFromState();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("RefreshIfOpen: " + ex.Message);
            }
        }

        public SettingsDialog()
        {
            SuspendLayout();

            Name = "OutlookAISettingsForm";
            Text = "OutlookAI Settings";
            Font = new Font("Segoe UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(470, 803);

            int margin = 12;
            int innerWidth = ClientSize.Width - 2 * margin;

            lblHeader = new Label
            {
                Name = "lblHeader",
                Text = "OutlookAI keeps these Outlook settings applied: fast local search, a fully cached " +
                       "mailbox (sync slider = All), and enough OST size headroom for it.",
                Location = new Point(margin, 10),
                Size = new Size(innerWidth, 32),
            };

            chkMaster = new CheckBox
            {
                Name = "chkMaster",
                Text = "Manage Outlook tuning",
                Location = new Point(margin, 46),
                Size = new Size(innerWidth, 20),
            };

            // --- Search group ---
            grpSearch = new GroupBox
            {
                Name = "grpSearch",
                Text = "Search",
                Location = new Point(margin, 74),
                Size = new Size(innerWidth, 158),
            };
            chkSearch = new CheckBox
            {
                Name = "chkSearch",
                Text = "Keep local search tuning applied",
                Location = new Point(10, 20),
                Size = new Size(innerWidth - 20, 18),
            };
            lblSearchValues = new Label
            {
                Name = "lblSearchValues",
                Location = new Point(10, 42),
                Size = new Size(innerWidth - 20, 74),
            };
            lblSearchWarning = new Label
            {
                Name = "lblSearchWarning",
                Text = "Turning this off restores Outlook's online search: slower, capped results, and " +
                       "'show me' results may no longer match what the agent finds.",
                Location = new Point(10, 120),
                Size = new Size(innerWidth - 20, 32),
            };
            grpSearch.Controls.Add(chkSearch);
            grpSearch.Controls.Add(lblSearchValues);
            grpSearch.Controls.Add(lblSearchWarning);

            // --- Full caching group ---
            grpCaching = new GroupBox
            {
                Name = "grpCaching",
                Text = "Full caching (sync slider = All)",
                Location = new Point(margin, 238),
                Size = new Size(innerWidth, 175),
            };
            chkCaching = new CheckBox
            {
                Name = "chkCaching",
                Text = "Keep full Cached Mode sync applied",
                Location = new Point(10, 20),
                Size = new Size(innerWidth - 20, 18),
            };
            lblCachingValues = new Label
            {
                Name = "lblCachingValues",
                Location = new Point(10, 42),
                Size = new Size(innerWidth - 20, 125),
            };
            grpCaching.Controls.Add(chkCaching);
            grpCaching.Controls.Add(lblCachingValues);

            // --- OST headroom group ---
            grpOst = new GroupBox
            {
                Name = "grpOst",
                Text = "OST size headroom",
                Location = new Point(margin, 419),
                Size = new Size(innerWidth, 90),
            };
            chkOst = new CheckBox
            {
                Name = "chkOst",
                Text = "Keep raised OST size limits applied (100 GB max)",
                Location = new Point(10, 20),
                Size = new Size(innerWidth - 20, 18),
            };
            lblOstValues = new Label
            {
                Name = "lblOstValues",
                Location = new Point(10, 42),
                Size = new Size(innerWidth - 20, 40),
            };
            grpOst.Controls.Add(chkOst);
            grpOst.Controls.Add(lblOstValues);

            lblRestart = new Label
            {
                Name = "lblRestart",
                Text = "Restart Outlook to apply pending changes.",
                Location = new Point(margin, 517),
                Size = new Size(innerWidth, 18),
                Visible = false,
            };

            lblGpo = new Label
            {
                Name = "lblGpo",
                Location = new Point(margin, 537),
                Size = new Size(innerWidth, 30),
                Visible = false,
            };

            // --- Claude Code group: where the mail server is registered, and its state ---
            grpClaude = new GroupBox
            {
                Name = "grpClaude",
                Text = "Mail server in Claude Code",
                Location = new Point(margin, 571),
                Size = new Size(innerWidth, 186),
            };
            chkGlobalMcp = new CheckBox
            {
                Name = "chkGlobalMcp",
                Text = "Make available in all my Claude Code projects",
                Location = new Point(10, 20),
                Size = new Size(innerWidth - 20, 18),
            };
            lblGlobalMcpHelp = new Label
            {
                Name = "lblGlobalMcpHelp",
                Text = "Registers the mail server in your personal Claude Code configuration, so every " +
                       "project you open can use it. Turning this off removes that entry again.",
                Location = new Point(26, 40),
                Size = new Size(innerWidth - 36, 30),
            };
            // Always visible: "connected and pointing at the right place" is worth stating,
            // not just its absence.
            lblMcp = new Label
            {
                Name = "lblMcp",
                Location = new Point(10, 74),
                Size = new Size(innerWidth - 20, 64),
            };
            btnAddProject = new Button
            {
                Name = "btnAddProject",
                Text = "Add to a specific project…",
                Location = new Point(10, 146),
                Size = new Size(176, 26),
            };
            btnCopyCommand = new Button
            {
                Name = "btnCopyCommand",
                Text = "Copy CLI command",
                Location = new Point(194, 146),
                Size = new Size(176, 26),
            };
            grpClaude.Controls.Add(chkGlobalMcp);
            grpClaude.Controls.Add(lblGlobalMcpHelp);
            grpClaude.Controls.Add(lblMcp);
            grpClaude.Controls.Add(btnAddProject);
            grpClaude.Controls.Add(btnCopyCommand);

            btnApply = new Button
            {
                Name = "btnApply",
                Text = "Apply now",
                Location = new Point(ClientSize.Width - margin - 170, 767),
                Size = new Size(84, 26),
            };
            btnClose = new Button
            {
                Name = "btnClose",
                Text = "Close",
                Location = new Point(ClientSize.Width - margin - 80, 767),
                Size = new Size(80, 26),
            };

            Controls.Add(lblHeader);
            Controls.Add(chkMaster);
            Controls.Add(grpSearch);
            Controls.Add(grpCaching);
            Controls.Add(grpOst);
            Controls.Add(lblRestart);
            Controls.Add(lblGpo);
            Controls.Add(grpClaude);
            Controls.Add(btnApply);
            Controls.Add(btnClose);

            AcceptButton = btnClose;
            CancelButton = btnClose;

            chkMaster.CheckedChanged += OnToggleChanged;
            chkSearch.CheckedChanged += OnToggleChanged;
            chkCaching.CheckedChanged += OnToggleChanged;
            chkOst.CheckedChanged += OnToggleChanged;
            // Deliberately NOT OnToggleChanged: this one owns a different service, and
            // toggling it must not drag the Outlook tuning reconcile along with it.
            chkGlobalMcp.CheckedChanged += OnGlobalMcpChanged;
            btnAddProject.Click += OnAddProject;
            btnCopyCommand.Click += OnCopyCommand;
            btnApply.Click += (s, e) =>
            {
                OutlookTuningService.ReconcileFromUi();
                // Same button, same promise: re-check the mail-server registration too, so a
                // drift the user just fixed (installing the runtime, say) is picked up without
                // restarting Outlook. Never throws out of its public surface.
                McpRegistrationService.Reconcile();
                RefreshFromState();
            };
            btnClose.Click += (s, e) => Close();

            ResumeLayout(false);

            ApplyTheme();
            ThemeService.ThemeChanged += OnThemeChanged;

            RefreshFromState();
        }

        private void OnToggleChanged(object sender, EventArgs e)
        {
            if (_updating)
                return;
            try
            {
                OutlookTuningService.SetMasterEnabled(chkMaster.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupSearch, chkSearch.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupCaching, chkCaching.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupOst, chkOst.Checked);
                OutlookTuningService.ReconcileFromUi();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings toggle: " + ex.Message);
            }
            RefreshFromState();
        }

        // The "all my projects" toggle. Ticking or unticking it IS the user declaring their
        // intent, so it applies immediately rather than waiting for the next Outlook start —
        // a user who unticks it expects the entry gone now — and it never re-opens the
        // question the startup prompt asks: they just answered it.
        private void OnGlobalMcpChanged(object sender, EventArgs e)
        {
            if (_updating)
                return;
            try
            {
                Cursor = Cursors.WaitCursor;
                McpRegistrationService.ApplyUserChoice(chkGlobalMcp.Checked);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP toggle: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            RefreshFromState();
        }

        // Project scope: writes/merges .mcp.json in a folder the user picks. Never throws
        // out of the click handler - a failure is something to read, not a crash dialog.
        private void OnAddProject(object sender, EventArgs e)
        {
            try
            {
                string folder;
                using (var picker = new FolderBrowserDialog())
                {
                    picker.Description = "Choose the project folder that should get the OutlookAI mail server.";
                    picker.ShowNewFolderButton = false;
                    if (picker.ShowDialog(this) != DialogResult.OK)
                        return;
                    folder = picker.SelectedPath;
                }

                string configPath, error;
                bool ok;
                try
                {
                    Cursor = Cursors.WaitCursor;
                    ok = McpRegistrationService.TryRegisterInProject(folder, out configPath, out error);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

                if (!ok)
                {
                    MessageBox.Show(this, error, "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine("Written to:");
                sb.AppendLine(configPath);
                sb.AppendLine();
                sb.AppendLine("The first time you open Claude Code in that folder it will ask you to approve "
                              + "this server. That prompt is Claude Code's own security check and only you can "
                              + "answer it — the add-in deliberately does not.");
                sb.AppendLine();
                sb.Append(".mcp.json is normally committed to source control. ");
                if (McpConfigEditor.ContainsEnvironmentReference(_preferredCommand))
                {
                    sb.Append("Because the entry points at ${LOCALAPPDATA} rather than a fixed path, it is "
                              + "portable: teammates who have OutlookAI installed get a working mail server, "
                              + "and teammates who do not simply see a failed-connection warning for this one "
                              + "server — nothing else breaks.");
                }
                else
                {
                    sb.Append("This entry names a fixed path on this machine (the mail server is not in the "
                              + "default install location here), so it will not resolve on a teammate's "
                              + "machine — they would see a failed-connection warning for this one server.");
                }

                MessageBox.Show(this, sb.ToString(), "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Add to project: " + ex.Message);
                MessageBox.Show(this, ex.Message, "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // The same registration as a command, for people who would rather use the CLI.
        private void OnCopyCommand(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_resolvedServerPath))
                    return;

                string command = "claude mcp add --scope project outlookai -- \"" + _resolvedServerPath + "\"";
                Clipboard.SetText(command);
                MessageBox.Show(
                    this,
                    "Copied to the clipboard. Run it from the project folder:" + Environment.NewLine
                        + Environment.NewLine + command + Environment.NewLine + Environment.NewLine
                        + "Use --scope user instead of --scope project to cover every project, which is what "
                        + "the tick box above does." + Environment.NewLine + Environment.NewLine
                        + "This names the real path so it works in any shell. The \"Add to a specific project…\" "
                        + "button writes the portable ${LOCALAPPDATA} form instead, which is the one to keep if "
                        + "the file goes into source control.",
                    "OutlookAI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // The clipboard belongs to whatever grabbed it last; losing that race is not
                // worth a crash dialog.
                System.Diagnostics.Debug.WriteLine("Copy command: " + ex.Message);
            }
        }

        // Reads only what the last reconcile recorded, so opening the dialog never touches
        // Claude Code's config file. "Apply now" is what re-runs the reconcile.
        private void RefreshMcpLine()
        {
            try
            {
                var reg = McpRegistrationService.GetSnapshot();
                string text;
                switch (reg.Status)
                {
                    case McpRegistrationService.StatusOk:
                        text = "Registered for all your projects.";
                        break;
                    case McpRegistrationService.StatusHealed:
                        text = "Registration was missing or pointing elsewhere and has been repaired.";
                        break;
                    case McpRegistrationService.StatusDisabled:
                        text = "Not registered for all your projects. Individual projects with a .mcp.json are unaffected.";
                        break;
                    case McpRegistrationService.StatusRemoved:
                        text = "Removed from your personal Claude Code configuration. Individual projects with a .mcp.json are unaffected.";
                        break;
                    case McpRegistrationService.StatusNoClaude:
                        text = "Claude Code was not found on this machine, so there is nothing to register with.";
                        break;
                    case McpRegistrationService.StatusNoServer:
                        text = "The mail server is not installed alongside the add-in; any existing registration was left unchanged.";
                        break;
                    case McpRegistrationService.StatusNoRuntime:
                        text = "The mail server needs the .NET 10 runtime, which is not installed. Get it from "
                               + McpRegistrationService.DotnetRuntimeDownloadUrl + " and restart Outlook.";
                        break;
                    case McpRegistrationService.StatusParseFailed:
                        text = "Claude Code's configuration could not be read, so it was left untouched.";
                        break;
                    case McpRegistrationService.StatusAwaitingChoice:
                        // The detail carries what was actually found, and there is no useful
                        // shorter way to say it: nothing has been changed, and why.
                        text = string.IsNullOrEmpty(reg.Detail)
                            ? "Waiting for you to choose whether to register the mail server. Nothing has been changed."
                            : reg.Detail;
                        break;
                    default:
                        text = "Registration state unknown"
                               + (string.IsNullOrEmpty(reg.Detail) ? "." : (" — " + reg.Detail));
                        break;
                }

                if (!string.IsNullOrEmpty(reg.RegisteredCommand) &&
                    (reg.Status == McpRegistrationService.StatusOk || reg.Status == McpRegistrationService.StatusHealed))
                {
                    text += Environment.NewLine + reg.RegisteredCommand;
                }
                else if (!string.IsNullOrEmpty(_preferredCommand))
                {
                    // Labelled: an unadorned path under "not registered" reads like a claim
                    // that it IS registered.
                    text += Environment.NewLine + "Server: " + _preferredCommand;
                }

                lblMcp.Text = text;
                lblMcp.ForeColor = IsMcpProblem(reg.Status) ? ThemeService.StatusError : ThemeService.SecondaryText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP line refresh: " + ex.Message);
            }
        }

        private static bool IsMcpProblem(string status)
        {
            return status == McpRegistrationService.StatusNoRuntime
                || status == McpRegistrationService.StatusParseFailed
                || status == McpRegistrationService.StatusError;
        }

        private void RefreshFromState()
        {
            // The Claude Code group first and in its own guarded block: a tuning snapshot
            // that fails must not leave the registration controls unpainted.
            _updating = true;
            try
            {
                _preferredCommand = McpRegistrationService.ResolvePreferredCommand();
                _resolvedServerPath = McpRegistrationService.ResolveInstalledServerPath() ?? "";
                chkGlobalMcp.Checked = McpRegistrationService.GetSnapshot().GlobalRegistrationEnabled;
                bool haveServer = !string.IsNullOrEmpty(_preferredCommand);
                btnAddProject.Enabled = haveServer;
                btnCopyCommand.Enabled = haveServer && !string.IsNullOrEmpty(_resolvedServerPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP settings refresh: " + ex.Message);
            }
            finally
            {
                _updating = false;
            }

            RefreshMcpLine();

            OutlookTuningService.TuningSnapshot snap;
            try
            {
                snap = OutlookTuningService.GetSnapshot();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings refresh: " + ex.Message);
                return;
            }

            _updating = true;
            try
            {
                chkMaster.Checked = snap.MasterEnabled;
                chkSearch.Checked = snap.SearchEnabled;
                chkCaching.Checked = snap.CachingEnabled;
                chkOst.Checked = snap.OstEnabled;

                chkSearch.Enabled = snap.MasterEnabled;
                chkCaching.Enabled = snap.MasterEnabled;
                chkOst.Enabled = snap.MasterEnabled;

                lblSearchValues.Text = BuildGroupText(snap, OutlookTuningService.GroupSearch);
                lblCachingValues.Text = BuildGroupText(snap, OutlookTuningService.GroupCaching);
                lblOstValues.Text = BuildGroupText(snap, OutlookTuningService.GroupOst);

                lblRestart.Visible = snap.RestartNeeded;

                if (snap.PolicyConflicts.Count > 0)
                {
                    var names = new StringBuilder();
                    foreach (var v in snap.Values)
                    {
                        if (!v.BackedOff)
                            continue;
                        if (names.Length > 0)
                            names.Append(", ");
                        names.Append(v.Entry.ValueName);
                    }
                    lblGpo.Text = "Managed by your organization's policy (left unchanged): " + names;
                    lblGpo.Visible = true;
                }
                else
                {
                    lblGpo.Visible = false;
                }
            }
            finally
            {
                _updating = false;
            }
        }

        private static string BuildGroupText(OutlookTuningService.TuningSnapshot snap, string groupId)
        {
            var sb = new StringBuilder();
            foreach (var v in snap.Values)
            {
                if (v.Entry.GroupId != groupId)
                    continue;
                string prefix = "";
                if (groupId == OutlookTuningService.GroupCaching)
                    prefix = v.Entry.IsPolicyHive ? "policy: " : "user: ";
                string live = v.Live.HasValue ? v.Live.Value.ToString() : "(not set)";
                string status;
                if (v.BackedOff)
                    status = "(policy override, left unchanged)";
                else if (v.InSync)
                    status = "(OK)";
                else if (!v.GroupEnabled)
                    status = "(not managed)";
                else
                    status = "-> " + v.Desired + " on apply";
                sb.AppendLine(prefix + v.Entry.ValueName + " = " + live + "  " + status);
            }
            return sb.ToString().TrimEnd();
        }

        // ===== Theming (mirrors the AITaskPane pattern) =====

        private void OnThemeChanged(object sender, EventArgs e)
        {
            // ThemeService may raise this on a non-UI (SystemEvents / registry watcher) thread.
            if (_disposedCustom || IsDisposed || !IsHandleCreated)
                return;
            try { BeginInvoke((Action)ApplyTheme); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void ApplyTheme()
        {
            BackColor = ThemeService.Background;
            ForeColor = ThemeService.Text;

            lblHeader.ForeColor = ThemeService.SecondaryText;
            lblSearchWarning.ForeColor = ThemeService.SecondaryText;
            lblRestart.ForeColor = ThemeService.StatusError;
            lblGpo.ForeColor = ThemeService.SecondaryText;
            lblGlobalMcpHelp.ForeColor = ThemeService.SecondaryText;
            // Colour depends on the state, so let the refresh own it rather than pinning a
            // colour here that the next refresh would immediately overwrite.
            RefreshMcpLine();

            foreach (var grp in new[] { grpSearch, grpCaching, grpOst, grpClaude })
                grp.ForeColor = ThemeService.Text;

            foreach (var chk in new[] { chkMaster, chkSearch, chkCaching, chkOst, chkGlobalMcp })
                chk.ForeColor = ThemeService.Text;

            foreach (var lbl in new[] { lblSearchValues, lblCachingValues, lblOstValues })
                lbl.ForeColor = ThemeService.Text;

            foreach (var btn in new[] { btnApply, btnClose, btnAddProject, btnCopyCommand })
            {
                if (ThemeService.IsDarkMode)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = ThemeService.Border;
                    btn.BackColor = ThemeService.ButtonFace;
                    btn.ForeColor = ThemeService.ButtonText;
                }
                else
                {
                    btn.FlatStyle = FlatStyle.Standard;
                    btn.UseVisualStyleBackColor = true;
                    btn.ForeColor = ThemeService.ButtonText;
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposedCustom)
            {
                _disposedCustom = true;
                ThemeService.ThemeChanged -= OnThemeChanged;
            }
            base.Dispose(disposing);
        }
    }
}
