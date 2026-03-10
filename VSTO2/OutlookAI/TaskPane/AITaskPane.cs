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
        private readonly ClaudeService _claudeService;
        private string _lastResult;

        public AITaskPane()
        {
            InitializeComponent();
            _claudeService = new ClaudeService();
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

        // SetMailItem removed - we always use ActiveInspector now

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

            // For Draft, truncate chain to avoid token limits (keep ~4000 chars)
            if (action == ClaudeService.ActionType.Draft && emailContent.Length > 4000)
            {
                emailContent = emailContent.Substring(0, 4000) + "\n[... earlier messages truncated ...]";
            }

            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
                string result = await _claudeService.ProcessEmailAsync(action, emailContent, prompt);

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
                if (InsertEmailBody(_lastResult))
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
                if (SetEmailBody(_lastResult))
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

        private void btnSettings_Click(object sender, EventArgs e)
        {
            using (var settingsForm = new SettingsForm())
            {
                settingsForm.ShowDialog();
            }
        }

        private string GetEmailBody()
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        return mail.Body ?? "";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetEmailBody error: " + ex.Message);
            }

            return "";
        }

        private bool InsertEmailBody(string text)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        string existingBody = mail.Body ?? "";
                        // Insert at top with separator
                        mail.Body = text + "\n\n" + existingBody;
                        return true;
                    }
                }
                ShowStatus("Could not find active email window.", true);
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InsertEmailBody error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
        }

        private bool SetEmailBody(string text)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        mail.Body = text;
                        return true;
                    }
                }
                ShowStatus("Could not find active email window.", true);
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SetEmailBody error: " + ex.Message);
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
            ClaudeService.Shutdown();
        }
    }

    public class SettingsForm : Form
    {
        private TextBox txtPassword;
        private ComboBox cboModel;
        private NumericUpDown numMaxTokens;
        private TextBox txtNewPassword;
        private Button btnSave;
        private Panel panelSettings;
        private Label lblError;
        private bool _authenticated = false;

        public SettingsForm()
        {
            this.Text = "AI Assistant Settings";
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblPassword = new Label { Text = "Admin Password:", Location = new Point(20, 20), AutoSize = true };
            txtPassword = new TextBox { Location = new Point(20, 45), Width = 340, PasswordChar = '*' };
            var btnLogin = new Button { Text = "Login", Location = new Point(280, 75), Width = 80 };
            btnLogin.Click += BtnLogin_Click;

            lblError = new Label { Location = new Point(20, 80), AutoSize = true, ForeColor = Color.DarkRed, Visible = false };

            panelSettings = new Panel { Location = new Point(0, 110), Size = new Size(400, 160), Visible = false };

            var lblModel = new Label { Text = "Model:", Location = new Point(20, 10), AutoSize = true };
            cboModel = new ComboBox { Location = new Point(20, 30), Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };
            cboModel.Items.AddRange(Config.AvailableModels);
            cboModel.SelectedItem = Config.Model;

            var lblMaxTokens = new Label { Text = "Max Tokens:", Location = new Point(20, 60), AutoSize = true };
            numMaxTokens = new NumericUpDown { Location = new Point(20, 80), Width = 100, Minimum = 256, Maximum = 4096, Value = Config.MaxTokens };

            var lblNewPassword = new Label { Text = "New Password (leave blank to keep):", Location = new Point(20, 110), AutoSize = true };
            txtNewPassword = new TextBox { Location = new Point(20, 130), Width = 200, PasswordChar = '*' };

            btnSave = new Button { Text = "Save", Location = new Point(200, 160), Width = 80 };
            btnSave.Click += BtnSave_Click;

            var btnCancel = new Button { Text = "Cancel", Location = new Point(290, 160), Width = 80 };
            btnCancel.Click += (s, e) => this.Close();

            panelSettings.Controls.AddRange(new Control[] { lblModel, cboModel, lblMaxTokens, numMaxTokens, lblNewPassword, txtNewPassword, btnSave, btnCancel });
            this.Controls.AddRange(new Control[] { lblPassword, txtPassword, btnLogin, lblError, panelSettings });
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text == Config.AdminPassword)
            {
                _authenticated = true;
                panelSettings.Visible = true;
                lblError.Visible = false;
                txtPassword.Enabled = false;
            }
            else
            {
                lblError.Text = "Invalid password";
                lblError.Visible = true;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!_authenticated) return;
            if (cboModel.SelectedItem != null) Config.Model = cboModel.SelectedItem.ToString();
            Config.MaxTokens = (int)numMaxTokens.Value;
            if (!string.IsNullOrWhiteSpace(txtNewPassword.Text)) Config.AdminPassword = txtNewPassword.Text;
            Config.SaveConfig();
            MessageBox.Show("Settings saved!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}
