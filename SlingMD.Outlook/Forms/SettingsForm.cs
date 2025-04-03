using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Forms
{
    public partial class SettingsForm : Form
    {
        private readonly ObsidianSettings _settings;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.txtVaultName = new TextBox();
            this.txtVaultPath = new TextBox();
            this.txtInboxFolder = new TextBox();
            this.chkLaunchObsidian = new CheckBox();
            this.numDelay = new NumericUpDown();
            this.chkShowCountdown = new CheckBox();
            this.btnBrowse = new Button();
            this.btnSave = new Button();
            this.btnCancel = new Button();
            this.lblVaultName = new Label();
            this.lblVaultPath = new Label();
            this.lblInboxFolder = new Label();
            this.lblDelay = new Label();

            // Form settings
            this.Text = "Obsidian Settings";
            this.Size = new System.Drawing.Size(700, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Padding = new Padding(20);

            // Constants for layout
            const int labelX = 30;
            const int controlX = 200;
            const int startY = 40;
            const int lineHeight = 45;
            const int controlWidth = 350;
            const int labelWidth = 160;
            const int buttonHeight = 35;

            // Style all labels
            foreach (var label in new[] { lblVaultName, lblVaultPath, lblInboxFolder, lblDelay })
            {
                label.AutoSize = false;
                label.Size = new Size(labelWidth, 25);
                label.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
                label.TextAlign = ContentAlignment.MiddleLeft;
            }

            // Style all text boxes
            foreach (var textBox in new[] { txtVaultName, txtVaultPath, txtInboxFolder })
            {
                textBox.Font = new Font("Segoe UI", 10F);
                textBox.Size = new Size(controlWidth, 25);
            }

            // Labels
            this.lblVaultName.Text = "Vault Name:";
            this.lblVaultName.Location = new Point(labelX, startY);

            this.lblVaultPath.Text = "Vault Base Path:";
            this.lblVaultPath.Location = new Point(labelX, startY + lineHeight);

            this.lblInboxFolder.Text = "Inbox Folder:";
            this.lblInboxFolder.Location = new Point(labelX, startY + lineHeight * 2);

            this.lblDelay.Text = "Delay (seconds):";
            this.lblDelay.Location = new Point(labelX, startY + lineHeight * 4);

            // Controls
            this.txtVaultName.Location = new Point(controlX, startY);

            this.txtVaultPath.Location = new Point(controlX, startY + lineHeight);

            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.Location = new Point(controlX + controlWidth + 10, startY + lineHeight);
            this.btnBrowse.Size = new Size(100, 25);
            this.btnBrowse.Font = new Font("Segoe UI", 9F);
            this.btnBrowse.Click += new EventHandler(btnBrowse_Click);

            this.txtInboxFolder.Location = new Point(controlX, startY + lineHeight * 2);

            this.chkLaunchObsidian.Text = "Launch Obsidian after saving";
            this.chkLaunchObsidian.Location = new Point(controlX, startY + lineHeight * 3);
            this.chkLaunchObsidian.Font = new Font("Segoe UI", 10F);
            this.chkLaunchObsidian.AutoSize = true;

            this.numDelay.Location = new Point(controlX, startY + lineHeight * 4);
            this.numDelay.Size = new Size(80, 25);
            this.numDelay.Font = new Font("Segoe UI", 10F);
            this.numDelay.Minimum = 0;
            this.numDelay.Maximum = 10;

            this.chkShowCountdown.Text = "Show countdown";
            this.chkShowCountdown.Location = new Point(controlX, startY + lineHeight * 5);
            this.chkShowCountdown.Font = new Font("Segoe UI", 10F);
            this.chkShowCountdown.AutoSize = true;

            // Action Buttons at the bottom
            int bottomButtonY = this.ClientSize.Height - buttonHeight - 30;
            
            this.btnSave.Text = "Save";
            this.btnSave.DialogResult = DialogResult.OK;
            this.btnSave.Location = new Point(this.ClientSize.Width - 230, bottomButtonY);
            this.btnSave.Size = new Size(100, buttonHeight);
            this.btnSave.Font = new Font("Segoe UI", 9F);
            this.btnSave.Click += new EventHandler(btnSave_Click);

            this.btnCancel.Text = "Cancel";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new Point(this.ClientSize.Width - 120, bottomButtonY);
            this.btnCancel.Size = new Size(100, buttonHeight);
            this.btnCancel.Font = new Font("Segoe UI", 9F);

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                this.lblVaultName, this.txtVaultName,
                this.lblVaultPath, this.txtVaultPath, this.btnBrowse,
                this.lblInboxFolder, this.txtInboxFolder,
                this.chkLaunchObsidian,
                this.lblDelay, this.numDelay,
                this.chkShowCountdown,
                this.btnSave, this.btnCancel
            });

            this.AcceptButton = this.btnSave;
            this.CancelButton = this.btnCancel;
        }

        private void LoadSettings()
        {
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            chkLaunchObsidian.Checked = _settings.LaunchObsidian;
            numDelay.Value = _settings.ObsidianDelaySeconds;
            chkShowCountdown.Checked = _settings.ShowCountdown;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Obsidian Vault Base Directory";
                dialog.SelectedPath = txtVaultPath.Text;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtVaultPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _settings.VaultName = txtVaultName.Text;
            _settings.VaultBasePath = txtVaultPath.Text;
            _settings.InboxFolder = txtInboxFolder.Text;
            _settings.LaunchObsidian = chkLaunchObsidian.Checked;
            _settings.ObsidianDelaySeconds = (int)numDelay.Value;
            _settings.ShowCountdown = chkShowCountdown.Checked;
        }

        // Designer-generated variables
        private TextBox txtVaultName;
        private TextBox txtVaultPath;
        private TextBox txtInboxFolder;
        private CheckBox chkLaunchObsidian;
        private NumericUpDown numDelay;
        private CheckBox chkShowCountdown;
        private Button btnBrowse;
        private Button btnSave;
        private Button btnCancel;
        private Label lblVaultName;
        private Label lblVaultPath;
        private Label lblInboxFolder;
        private Label lblDelay;
    }
} 