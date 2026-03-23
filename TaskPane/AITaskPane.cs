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
            _versionTimer.Interval = 1000; // 1 second
            _versionTimer.Tick += (s, ev) => UpdateVersionLabel();
            _versionTimer.Start();
            UpdateVersionLabel();
        }

        private void UpdateVersionLabel()
        {
            var version = "v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var lastChecked = UpdateService.LastChecked;
            var error = UpdateService.LastError;
            var status = UpdateService.Status;

            string suffix;
            if (status != null && status != "up to date")
                suffix = status;
            else if (lastChecked == null)
                suffix = error != null ? null : "checking\u2026";
            else
            {
                var ago = DateTime.Now - lastChecked.Value;
                if (ago.TotalSeconds < 60)
                    suffix = "checked just now";
                else if (ago.TotalMinutes < 60)
                    suffix = $"checked {(int)ago.TotalMinutes}m ago";
                else if (ago.TotalHours < 24)
                    suffix = $"checked {(int)ago.TotalHours}h ago";
                else
                    suffix = $"checked {(int)ago.TotalDays}d ago";
            }

            lblVersion.Text = suffix != null ? $"{version} - {suffix}" : version;
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

        private Outlook.MailItem GetCurrentMailItem(out Outlook.Explorer explorer)
        {
            explorer = null;
            if (_isInlineResponse)
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                object rawInline = explorer?.ActiveInlineResponse;
                var mail = rawInline as Outlook.MailItem;
                if (mail == null)
                    ThisAddIn.ReleaseCom(rawInline);
                return mail;
            }

            // Use the owning inspector rather than ActiveInspector so that
            // clicking a button in a background window processes the right email.
            object rawItem = _owningInspector?.CurrentItem;
            var mailItem = rawItem as Outlook.MailItem;
            if (mailItem == null)
                ThisAddIn.ReleaseCom(rawItem);
            return mailItem;
        }

        private string GetEmailBody()
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            try
            {
                mail = GetCurrentMailItem(out explorer);
                return mail?.Body ?? "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetEmailBody error: " + ex.Message);
                return "";
            }
            finally
            {
                ThisAddIn.ReleaseCom(mail);
                ThisAddIn.ReleaseCom(explorer);
            }
        }

        private bool UpdateEmailBody(string text, bool insert)
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            try
            {
                mail = GetCurrentMailItem(out explorer);
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
            finally
            {
                ThisAddIn.ReleaseCom(mail);
                ThisAddIn.ReleaseCom(explorer);
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
            ThisAddIn.ReleaseCom(_owningInspector);
            // Don't call ClaudeService.Shutdown() here -- it kills the shared
            // warm process that other panes still need. ThisAddIn_Shutdown
            // handles final cleanup.
        }
    }
}
