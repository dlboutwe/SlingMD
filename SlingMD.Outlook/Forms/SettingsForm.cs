using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;
using System.Collections.Generic;

namespace SlingMD.Outlook.Forms
{
    public partial class SettingsForm : BaseForm
    {
        private readonly ObsidianSettings _settings;
        private ListBox lstPatterns;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnRemove;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void InitializeComponent()
        {
            // Create a panel to contain all controls
            var contentPanel = new Panel
            {
                AutoScroll = true,
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };

            this.txtVaultName = new TextBox();
            this.txtVaultPath = new TextBox();
            this.txtInboxFolder = new TextBox();
            this.chkLaunchObsidian = new CheckBox();
            this.numDelay = new NumericUpDown();
            this.chkShowCountdown = new CheckBox();
            this.chkCreateObsidianTask = new CheckBox();
            this.chkCreateOutlookTask = new CheckBox();
            this.numDefaultDueDays = new NumericUpDown();
            this.numDefaultReminderDays = new NumericUpDown();
            this.numDefaultReminderHour = new NumericUpDown();
            this.chkAskForDates = new CheckBox();
            this.btnBrowse = new Button();
            this.btnSave = new Button();
            this.btnCancel = new Button();
            this.lblVaultName = new Label();
            this.lblVaultPath = new Label();
            this.lblInboxFolder = new Label();
            this.lblDelay = new Label();
            this.lblFollowUpTasks = new Label();
            this.lblDefaultDueDays = new Label();
            this.lblDefaultReminderDays = new Label();
            this.lblDefaultReminderHour = new Label();
            this.lblDueDaysHelp = new Label();
            this.chkGroupEmailThreads = new CheckBox();
            this.lstPatterns = new ListBox();
            this.btnAdd = new Button();
            this.btnEdit = new Button();
            this.btnRemove = new Button();

            // Form settings
            this.Text = "Obsidian Settings";
            this.MinimumSize = new System.Drawing.Size(800, 600);
            this.Size = new System.Drawing.Size(900, 700); // Reduced default height
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Constants for layout
            const int labelX = 20;  // Reduced left margin
            const int controlX = 160;  // Reduced space between label and control
            const int startY = 20;  // Reduced top margin
            const int lineHeight = 30;  // Further reduced space between lines
            const int controlWidth = 350;
            const int labelWidth = 130;
            const int buttonHeight = 30;
            const int smallControlWidth = 80;
            const int helpTextX = controlX + smallControlWidth + 10;
            const int checkboxX = helpTextX + 200;

            // Style all labels
            foreach (var label in new[] { lblVaultName, lblVaultPath, lblInboxFolder, lblDelay, lblFollowUpTasks, 
                                        lblDefaultDueDays, lblDefaultReminderDays, lblDefaultReminderHour })
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

            this.lblFollowUpTasks.Text = "Follow-up Tasks:";
            this.lblFollowUpTasks.Location = new Point(labelX, startY + lineHeight * 6);

            this.lblDefaultDueDays.Text = "Due in Days:";
            this.lblDefaultDueDays.Location = new Point(labelX, startY + lineHeight * 8);

            this.lblDefaultReminderDays.Text = "Reminder Days:";
            this.lblDefaultReminderDays.Location = new Point(labelX, startY + lineHeight * 9);

            this.lblDefaultReminderHour.Text = "Reminder Hour:";
            this.lblDefaultReminderHour.Location = new Point(labelX, startY + lineHeight * 10);

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
            this.numDelay.Size = new Size(smallControlWidth, 25);
            this.numDelay.Font = new Font("Segoe UI", 10F);
            this.numDelay.Minimum = 0;
            this.numDelay.Maximum = 10;

            this.chkShowCountdown.Text = "Show countdown";
            this.chkShowCountdown.Location = new Point(controlX, startY + lineHeight * 5);
            this.chkShowCountdown.Font = new Font("Segoe UI", 10F);
            this.chkShowCountdown.AutoSize = true;

            this.chkCreateObsidianTask.Text = "Create task in Obsidian note";
            this.chkCreateObsidianTask.Location = new Point(controlX, startY + lineHeight * 6);
            this.chkCreateObsidianTask.Font = new Font("Segoe UI", 10F);
            this.chkCreateObsidianTask.AutoSize = true;

            this.chkCreateOutlookTask.Text = "Create task in Outlook";
            this.chkCreateOutlookTask.Location = new Point(controlX, startY + lineHeight * 7);
            this.chkCreateOutlookTask.Font = new Font("Segoe UI", 10F);
            this.chkCreateOutlookTask.AutoSize = true;

            // Due days settings
            this.numDefaultDueDays.Location = new Point(controlX, startY + lineHeight * 8);
            this.numDefaultDueDays.Size = new Size(smallControlWidth, 25);
            this.numDefaultDueDays.Font = new Font("Segoe UI", 10F);
            this.numDefaultDueDays.Minimum = 0;
            this.numDefaultDueDays.Maximum = 30;

            // Help text for due days
            this.lblDueDaysHelp.Text = "0 = Today, 1 = Tomorrow, etc.";
            this.lblDueDaysHelp.Location = new Point(helpTextX, startY + lineHeight * 8 + 3);
            this.lblDueDaysHelp.AutoSize = true;
            this.lblDueDaysHelp.Font = new Font("Segoe UI", 9F, FontStyle.Italic);

            // Reminder days settings
            this.numDefaultReminderDays.Location = new Point(controlX, startY + lineHeight * 9);
            this.numDefaultReminderDays.Size = new Size(smallControlWidth, 25);
            this.numDefaultReminderDays.Font = new Font("Segoe UI", 10F);
            this.numDefaultReminderDays.Minimum = 0;
            this.numDefaultReminderDays.Maximum = 30;

            // Help text for reminder days
            var lblReminderDaysHelp = new Label
            {
                Text = "Days before due date",
                Location = new Point(helpTextX, startY + lineHeight * 9 + 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Italic)
            };

            // Reminder hour settings
            this.numDefaultReminderHour.Location = new Point(controlX, startY + lineHeight * 10);
            this.numDefaultReminderHour.Size = new Size(smallControlWidth, 25);
            this.numDefaultReminderHour.Font = new Font("Segoe UI", 10F);
            this.numDefaultReminderHour.Minimum = 0;
            this.numDefaultReminderHour.Maximum = 23;

            // Help text for reminder hour
            var lblReminderHourHelp = new Label
            {
                Text = "(24-hour format)",
                Location = new Point(helpTextX, startY + lineHeight * 10 + 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Italic)
            };

            // Ask for dates checkbox
            this.chkAskForDates.Text = "Ask for dates and times each time";
            this.chkAskForDates.Location = new Point(checkboxX, startY + lineHeight * 9);
            this.chkAskForDates.Font = new Font("Segoe UI", 10F);
            this.chkAskForDates.AutoSize = true;

            // Group email threads checkbox
            this.chkGroupEmailThreads.Text = "Group email threads";
            this.chkGroupEmailThreads.Location = new Point(checkboxX, startY + lineHeight * 10);
            this.chkGroupEmailThreads.Font = new Font("Segoe UI", 10F);
            this.chkGroupEmailThreads.AutoSize = true;

            // Create main settings group
            var mainSettingsGroup = new GroupBox
            {
                Text = "General Settings",
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 0, 10)
            };

            // Add controls to the main settings group
            mainSettingsGroup.Controls.AddRange(new Control[] {
                this.lblVaultName, this.txtVaultName,
                this.lblVaultPath, this.txtVaultPath, this.btnBrowse,
                this.lblInboxFolder, this.txtInboxFolder,
                this.chkLaunchObsidian,
                this.lblDelay, this.numDelay,
                this.chkShowCountdown,
                this.lblFollowUpTasks,
                this.chkCreateObsidianTask,
                this.chkCreateOutlookTask,
                this.lblDefaultDueDays, this.numDefaultDueDays,
                this.lblDueDaysHelp,
                this.lblDefaultReminderDays, this.numDefaultReminderDays,
                lblReminderDaysHelp,
                this.lblDefaultReminderHour, this.numDefaultReminderHour,
                lblReminderHourHelp,
                this.chkAskForDates,
                this.chkGroupEmailThreads
            });

            // Add pattern management controls
            var patternGroup = new GroupBox
            {
                Text = "Subject Cleanup Patterns",
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                MinimumSize = new Size(0, 150)
            };

            // Create a container panel for the list and buttons
            var patternContentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(3, 30, 30, 3), // Added right padding of 10
                Margin = new Padding(3, 30, 30, 3) // Added right padding of 10
            };

            lstPatterns.Dock = DockStyle.Fill;
            lstPatterns.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
            lstPatterns.Margin = new Padding(0, 0, 20, 0); // Added right margin to the list

            var patternButtonPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 100,
                Padding = new Padding(20, 0, 0, 0), // Increased left padding from 10 to 20
                Margin = new Padding(20, 0, 0, 0) // Increased left padding from 10 to 20
            };

            btnAdd.Text = "Add";
            btnAdd.Dock = DockStyle.Top;
            btnAdd.Height = 30;
            btnAdd.Margin = new Padding(0, 0, 0, 5);
            btnAdd.Click += BtnAdd_Click;

            btnEdit.Text = "Edit";
            btnEdit.Dock = DockStyle.Top;
            btnEdit.Height = 30;
            btnEdit.Margin = new Padding(0, 0, 0, 5);
            btnEdit.Click += BtnEdit_Click;

            btnRemove.Text = "Remove";
            btnRemove.Dock = DockStyle.Top;
            btnRemove.Height = 30;
            btnRemove.Margin = new Padding(0, 0, 0, 5);
            btnRemove.Click += BtnRemove_Click;

            patternButtonPanel.Controls.Add(btnRemove);
            patternButtonPanel.Controls.Add(btnEdit);
            patternButtonPanel.Controls.Add(btnAdd);

            patternContentPanel.Controls.Add(lstPatterns);
            patternContentPanel.Controls.Add(patternButtonPanel);
            patternGroup.Controls.Add(patternContentPanel);

            // Create a panel for the bottom buttons that stays at the bottom
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Padding = new Padding(10)
            };

            // Configure Save/Cancel buttons
            this.btnSave.Text = "Save";
            this.btnSave.DialogResult = DialogResult.OK;
            this.btnSave.Size = new Size(100, 30);
            this.btnSave.Font = new Font("Segoe UI", 9F);
            this.btnSave.Click += new EventHandler(btnSave_Click);
            this.btnSave.Anchor = AnchorStyles.Right;

            this.btnCancel.Text = "Cancel";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Size = new Size(100, 30);
            this.btnCancel.Font = new Font("Segoe UI", 9F);
            this.btnCancel.Anchor = AnchorStyles.Right;

            // Position the buttons
            this.btnSave.Location = new Point(buttonPanel.ClientSize.Width - 230, 10);
            this.btnCancel.Location = new Point(buttonPanel.ClientSize.Width - 120, 10);
            buttonPanel.Controls.AddRange(new Control[] { this.btnSave, this.btnCancel });

            // Add panels to form in correct order
            contentPanel.Controls.Add(patternGroup);
            contentPanel.Controls.Add(mainSettingsGroup);
            this.Controls.Add(contentPanel);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = this.btnSave;
            this.CancelButton = this.btnCancel;

            // Handle form resize to adjust controls
            this.Resize += (s, e) => {
                buttonPanel.SuspendLayout();
                btnSave.Left = buttonPanel.ClientSize.Width - 230;
                btnCancel.Left = buttonPanel.ClientSize.Width - 120;
                buttonPanel.ResumeLayout();
            };
        }

        private void LoadSettings()
        {
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            chkLaunchObsidian.Checked = _settings.LaunchObsidian;
            numDelay.Value = _settings.ObsidianDelaySeconds;
            chkShowCountdown.Checked = _settings.ShowCountdown;
            chkCreateObsidianTask.Checked = _settings.CreateObsidianTask;
            chkCreateOutlookTask.Checked = _settings.CreateOutlookTask;
            chkAskForDates.Checked = _settings.AskForDates;
            chkGroupEmailThreads.Checked = _settings.GroupEmailThreads;
            numDefaultDueDays.Value = _settings.DefaultDueDays;
            numDefaultReminderDays.Value = _settings.DefaultReminderDays;
            numDefaultReminderHour.Value = _settings.DefaultReminderHour;

            // Load patterns
            lstPatterns.Items.Clear();
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                lstPatterns.Items.Add(pattern);
            }
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
            _settings.CreateObsidianTask = chkCreateObsidianTask.Checked;
            _settings.CreateOutlookTask = chkCreateOutlookTask.Checked;
            _settings.AskForDates = chkAskForDates.Checked;
            _settings.GroupEmailThreads = chkGroupEmailThreads.Checked;
            _settings.DefaultDueDays = (int)numDefaultDueDays.Value;
            _settings.DefaultReminderDays = (int)numDefaultReminderDays.Value;
            _settings.DefaultReminderHour = (int)numDefaultReminderHour.Value;

            // Save patterns
            _settings.SubjectCleanupPatterns.Clear();
            foreach (string pattern in lstPatterns.Items)
            {
                _settings.SubjectCleanupPatterns.Add(pattern);
            }

            _settings.Save();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Pattern", "Enter regex pattern:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    lstPatterns.Items.Add(form.InputText);
                }
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (lstPatterns.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Pattern", "Edit regex pattern:", lstPatterns.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstPatterns.SelectedIndex;
                        lstPatterns.Items[index] = form.InputText;
                    }
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lstPatterns.SelectedItem != null)
            {
                lstPatterns.Items.RemoveAt(lstPatterns.SelectedIndex);
            }
        }

        // Designer-generated variables
        private TextBox txtVaultName;
        private TextBox txtVaultPath;
        private TextBox txtInboxFolder;
        private CheckBox chkLaunchObsidian;
        private NumericUpDown numDelay;
        private CheckBox chkShowCountdown;
        private CheckBox chkCreateObsidianTask;
        private CheckBox chkCreateOutlookTask;
        private CheckBox chkAskForDates;
        private CheckBox chkGroupEmailThreads;
        private NumericUpDown numDefaultDueDays;
        private NumericUpDown numDefaultReminderDays;
        private NumericUpDown numDefaultReminderHour;
        private Button btnBrowse;
        private Button btnSave;
        private Button btnCancel;
        private Label lblVaultName;
        private Label lblVaultPath;
        private Label lblInboxFolder;
        private Label lblDelay;
        private Label lblFollowUpTasks;
        private Label lblDefaultDueDays;
        private Label lblDefaultReminderDays;
        private Label lblDefaultReminderHour;
        private Label lblDueDaysHelp;
    }
} 