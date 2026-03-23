using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using OutlookAI.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane
{
    public partial class AITaskPane : UserControl
    {
        private string _lastResult;
        private readonly bool _isInlineResponse;
        private readonly Outlook.Inspector _owningInspector;
        private readonly Timer _versionTimer;

        public AITaskPane(bool isInlineResponse = false, Outlook.Inspector inspector = null)
        {
            _isInlineResponse = isInlineResponse;
            _owningInspector = inspector;
            InitializeComponent();

            _versionTimer = new Timer();
            _versionTimer.Interval = 30000; // 30 seconds
            _versionTimer.Tick += (s, ev) => UpdateVersionLabel();
            _versionTimer.Start();
            UpdateVersionLabel();
        }

        private void UpdateVersionLabel()
        {
            var version = "v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var lastChecked = UpdateService.LastChecked;
            if (lastChecked == null)
            {
                lblVersion.Text = version;
                return;
            }

            var ago = DateTime.Now - lastChecked.Value;
            string agoText;
            if (ago.TotalSeconds < 60)
                agoText = "just now";
            else if (ago.TotalMinutes < 60)
                agoText = $"{(int)ago.TotalMinutes}m ago";
            else if (ago.TotalHours < 24)
                agoText = $"{(int)ago.TotalHours}h ago";
            else
                agoText = $"{(int)ago.TotalDays}d ago";

            lblVersion.Text = $"{version} \u2022 checked {agoText}";

            var error = UpdateService.LastError;
            lnkUpdateError.Visible = error != null;
        }

        private void lnkUpdateError_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var error = UpdateService.LastError;
            if (error != null)
                MessageBox.Show(error, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Call this when the task pane becomes visible for a new email
        /// </summary>
        public void ResetForNewEmail()
        {
            txtDraftPrompt.Text = "";
            txtCustomPrompt.Text = "";
            txtResult.Text = "";
            panelResult.Visible = false;
            lblStatus.Visible = false;
            _lastResult = null;
        }

        private async void btnProofread_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Proofread);
        }

        private async void btnRevise_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Revise);
        }

        private async void btnShorten_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Shorten);
        }

        private async void btnLengthen_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Lengthen);
        }

        private async void btnFormal_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Formal);
        }

        private async void btnFriendly_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Friendly);
        }

        private async void btnDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
            {
                ShowStatus("Please enter instructions for the email you want to draft.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Draft, txtDraftPrompt.Text);
        }

        private async void btnCustom_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCustomPrompt.Text))
            {
                ShowStatus("Please enter instructions for the custom action.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Custom, txtCustomPrompt.Text);
        }

        private async Task ProcessAction(ClaudeService.ActionType action, string prompt = "")
        {
            string emailContent = GetEmailBody();

            // For non-Draft/Custom actions, we need existing content to work with
            if (action != ClaudeService.ActionType.Draft && action != ClaudeService.ActionType.Custom && string.IsNullOrWhiteSpace(emailContent))
            {
                ShowStatus("No email content found. Please write something first.", true);
                return;
            }

            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
                string result = await ClaudeService.ProcessEmailAsync(action, emailContent, prompt);

                _lastResult = result;

                InvokeOnUI(() =>
                {
                    txtResult.Text = _lastResult;
                    panelResult.Visible = true;
                    ShowStatus("Done! Review the result below.", false);
                    SetUIEnabled(true);
                });
            }
            catch (Exception ex)
            {
                InvokeOnUI(() =>
                {
                    string msg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    ShowStatus(msg, true);
                    panelResult.Visible = false;
                    SetUIEnabled(true);
                });
            }
        }

        private void InvokeOnUI(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated)
                return;

            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastResult))
            {
                if (UpdateEmailBody(_lastResult, insert: true))
                {
                    panelResult.Visible = false;
                    txtDraftPrompt.Text = "";
                    ShowStatus("Draft inserted!", false);
                }
            }
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastResult))
            {
                if (UpdateEmailBody(_lastResult, insert: false))
                {
                    panelResult.Visible = false;
                    txtDraftPrompt.Text = "";
                    ShowStatus("Email replaced!", false);
                }
            }
        }

        private void btnDiscard_Click(object sender, EventArgs e)
        {
            _lastResult = null;
            panelResult.Visible = false;
            lblStatus.Visible = false;
        }

        private Outlook.MailItem GetCurrentMailItem()
        {
            if (_isInlineResponse)
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                return explorer?.ActiveInlineResponse as Outlook.MailItem;
            }

            // Use the owning inspector rather than ActiveInspector so that
            // clicking a button in a background window processes the right email.
            return _owningInspector?.CurrentItem as Outlook.MailItem;
        }

        private string GetEmailBody()
        {
            try
            {
                return GetCurrentMailItem()?.Body ?? "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetEmailBody error: " + ex.Message);
                return "";
            }
        }

        private bool UpdateEmailBody(string text, bool insert)
        {
            try
            {
                var mail = GetCurrentMailItem();
                if (mail == null)
                {
                    ShowStatus("Could not find active email window.", true);
                    return false;
                }
                mail.Body = insert ? text + "\n\n" + (mail.Body ?? "") : text;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UpdateEmailBody error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
        }

        private void ShowStatus(string message, bool isError)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = isError ? Color.DarkRed : Color.DarkGreen;
            lblStatus.Visible = true;
        }

        private void SetUIEnabled(bool enabled)
        {
            btnProofread.Enabled = enabled;
            btnRevise.Enabled = enabled;
            btnShorten.Enabled = enabled;
            btnLengthen.Enabled = enabled;
            btnFormal.Enabled = enabled;
            btnFriendly.Enabled = enabled;
            btnDraft.Enabled = enabled;
            txtDraftPrompt.Enabled = enabled;
            btnCustom.Enabled = enabled;
            txtCustomPrompt.Enabled = enabled;
        }

        partial void DisposeCustomResources()
        {
            _versionTimer?.Stop();
            _versionTimer?.Dispose();
            // Don't call ClaudeService.Shutdown() here -- it kills the shared
            // warm process that other panes still need. ThisAddIn_Shutdown
            // handles final cleanup.
        }
    }
}
