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
        private readonly GroupBox grpSearch;
        private readonly GroupBox grpCaching;
        private readonly GroupBox grpOst;
        private readonly Label lblHeader;
        private readonly Label lblSearchValues;
        private readonly Label lblCachingValues;
        private readonly Label lblOstValues;
        private readonly Label lblRestart;
        private readonly Label lblGpo;
        private readonly Button btnApply;
        private readonly Button btnClose;

        private bool _updating;
        private bool _disposedCustom;

        internal static bool IsOpen
        {
            get { return _open != null && !_open.IsDisposed; }
        }

        internal static void ShowSettings()
        {
            try
            {
                if (IsOpen)
                {
                    _open.Activate();
                    return;
                }
                var dlg = new SettingsDialog();
                _open = dlg;
                dlg.FormClosed += (s, e) => { if (ReferenceEquals(_open, dlg)) _open = null; };
                dlg.Show();
                dlg.Activate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowSettings: " + ex.Message);
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
            ClientSize = new Size(470, 532);

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
                Size = new Size(innerWidth, 110),
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
                Size = new Size(innerWidth - 20, 60),
            };
            grpSearch.Controls.Add(chkSearch);
            grpSearch.Controls.Add(lblSearchValues);

            // --- Full caching group ---
            grpCaching = new GroupBox
            {
                Name = "grpCaching",
                Text = "Full caching (sync slider = All)",
                Location = new Point(margin, 190),
                Size = new Size(innerWidth, 155),
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
                Size = new Size(innerWidth - 20, 105),
            };
            grpCaching.Controls.Add(chkCaching);
            grpCaching.Controls.Add(lblCachingValues);

            // --- OST headroom group ---
            grpOst = new GroupBox
            {
                Name = "grpOst",
                Text = "OST size headroom",
                Location = new Point(margin, 351),
                Size = new Size(innerWidth, 80),
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
                Size = new Size(innerWidth - 20, 30),
            };
            grpOst.Controls.Add(chkOst);
            grpOst.Controls.Add(lblOstValues);

            lblRestart = new Label
            {
                Name = "lblRestart",
                Text = "Restart Outlook to apply pending changes.",
                Location = new Point(margin, 437),
                Size = new Size(innerWidth, 18),
                Visible = false,
            };

            lblGpo = new Label
            {
                Name = "lblGpo",
                Location = new Point(margin, 457),
                Size = new Size(innerWidth, 30),
                Visible = false,
            };

            btnApply = new Button
            {
                Name = "btnApply",
                Text = "Apply now",
                Location = new Point(ClientSize.Width - margin - 170, 494),
                Size = new Size(84, 26),
            };
            btnClose = new Button
            {
                Name = "btnClose",
                Text = "Close",
                Location = new Point(ClientSize.Width - margin - 80, 494),
                Size = new Size(80, 26),
            };

            Controls.Add(lblHeader);
            Controls.Add(chkMaster);
            Controls.Add(grpSearch);
            Controls.Add(grpCaching);
            Controls.Add(grpOst);
            Controls.Add(lblRestart);
            Controls.Add(lblGpo);
            Controls.Add(btnApply);
            Controls.Add(btnClose);

            AcceptButton = btnClose;
            CancelButton = btnClose;

            chkMaster.CheckedChanged += OnToggleChanged;
            chkSearch.CheckedChanged += OnToggleChanged;
            chkCaching.CheckedChanged += OnToggleChanged;
            chkOst.CheckedChanged += OnToggleChanged;
            btnApply.Click += (s, e) => { OutlookTuningService.ReconcileFromUi(); RefreshFromState(); };
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

        private void RefreshFromState()
        {
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
            lblRestart.ForeColor = ThemeService.StatusError;
            lblGpo.ForeColor = ThemeService.SecondaryText;

            foreach (var grp in new[] { grpSearch, grpCaching, grpOst })
                grp.ForeColor = ThemeService.Text;

            foreach (var chk in new[] { chkMaster, chkSearch, chkCaching, chkOst })
                chk.ForeColor = ThemeService.Text;

            foreach (var lbl in new[] { lblSearchValues, lblCachingValues, lblOstValues })
                lbl.ForeColor = ThemeService.Text;

            foreach (var btn in new[] { btnApply, btnClose })
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
