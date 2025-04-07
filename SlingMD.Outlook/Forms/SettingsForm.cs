using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;
using System.Collections.Generic;
using System.Reflection;

namespace SlingMD.Outlook.Forms
{
    public partial class SettingsForm : Form
    {
        private readonly ObsidianSettings _settings;
        
        // Designer-generated fields
        private TextBox txtVaultName;
        private TextBox txtVaultPath;
        private TextBox txtInboxFolder;
        private TextBox txtContactsFolder;
        private CheckBox chkLaunchObsidian;
        private CheckBox chkEnableContactSaving;
        private CheckBox chkSearchEntireVaultForContacts;
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
        private Label lblContactsFolder;
        private Label lblDelay;
        private Label lblFollowUpTasks;
        private Label lblDefaultDueDays;
        private Label lblDefaultReminderDays;
        private Label lblDefaultReminderHour;
        private Label lblDueDaysHelp;
        private ListBox lstPatterns;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnRemove;
        private Label lblPatterns;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            
            // Create form elements
            this.txtVaultName = new TextBox();
            this.txtVaultPath = new TextBox(); 
            this.txtInboxFolder = new TextBox();
            this.txtContactsFolder = new TextBox();
            this.chkLaunchObsidian = new CheckBox();
            this.chkEnableContactSaving = new CheckBox();
            this.chkSearchEntireVaultForContacts = new CheckBox();
            this.numDelay = new NumericUpDown();
            this.chkShowCountdown = new CheckBox();
            this.chkCreateObsidianTask = new CheckBox();
            this.chkCreateOutlookTask = new CheckBox();
            this.chkAskForDates = new CheckBox();
            this.chkGroupEmailThreads = new CheckBox();
            this.numDefaultDueDays = new NumericUpDown();
            this.numDefaultReminderDays = new NumericUpDown();
            this.numDefaultReminderHour = new NumericUpDown();
            this.btnBrowse = new Button();
            this.btnSave = new Button();
            this.btnCancel = new Button();
            this.lblVaultName = new Label();
            this.lblVaultPath = new Label();
            this.lblInboxFolder = new Label();
            this.lblContactsFolder = new Label();
            this.lblDelay = new Label();
            this.lblFollowUpTasks = new Label();
            this.lblDefaultDueDays = new Label();
            this.lblDefaultReminderDays = new Label();
            this.lblDefaultReminderHour = new Label();
            this.lblDueDaysHelp = new Label();
            this.lstPatterns = new ListBox();
            this.btnAdd = new Button();
            this.btnEdit = new Button();
            this.btnRemove = new Button();
            this.lblPatterns = new Label();

            // Initialize numeric controls
            ((System.ComponentModel.ISupportInitialize)(this.numDelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultDueDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderHour)).BeginInit();

            this.SuspendLayout();

            // Configure Label: Vault Name
            this.lblVaultName.AutoSize = true;
            this.lblVaultName.Location = new System.Drawing.Point(12, 15);
            this.lblVaultName.Name = "lblVaultName";
            this.lblVaultName.Size = new System.Drawing.Size(72, 13);
            this.lblVaultName.Text = "Vault Name:";

            // Configure Textbox: Vault Name
            this.txtVaultName.Location = new System.Drawing.Point(184, 12);
            this.txtVaultName.Name = "txtVaultName";
            this.txtVaultName.Size = new System.Drawing.Size(350, 20);

            // Configure Label: Vault Path
            this.lblVaultPath.AutoSize = true;
            this.lblVaultPath.Location = new System.Drawing.Point(12, 45);
            this.lblVaultPath.Name = "lblVaultPath";
            this.lblVaultPath.Size = new System.Drawing.Size(82, 13);
            this.lblVaultPath.Text = "Vault Base Path:";

            // Configure Textbox: Vault Path
            this.txtVaultPath.Location = new System.Drawing.Point(184, 42);
            this.txtVaultPath.Name = "txtVaultPath";
            this.txtVaultPath.Size = new System.Drawing.Size(350, 20);

            // Configure Button: Browse
            this.btnBrowse.Location = new System.Drawing.Point(540, 42);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new EventHandler(this.btnBrowse_Click);

            // Configure Label: Inbox Folder
            this.lblInboxFolder.AutoSize = true;
            this.lblInboxFolder.Location = new System.Drawing.Point(12, 75);
            this.lblInboxFolder.Name = "lblInboxFolder";
            this.lblInboxFolder.Size = new System.Drawing.Size(71, 13);
            this.lblInboxFolder.Text = "Inbox Folder:";

            // Configure Textbox: Inbox Folder
            this.txtInboxFolder.Location = new System.Drawing.Point(184, 72);
            this.txtInboxFolder.Name = "txtInboxFolder";
            this.txtInboxFolder.Size = new System.Drawing.Size(350, 20);

            // Configure Label: Contacts Folder
            this.lblContactsFolder.AutoSize = true;
            this.lblContactsFolder.Location = new System.Drawing.Point(12, 105);
            this.lblContactsFolder.Name = "lblContactsFolder";
            this.lblContactsFolder.Size = new System.Drawing.Size(84, 13);
            this.lblContactsFolder.Text = "Contacts Folder:";

            // Configure Textbox: Contacts Folder
            this.txtContactsFolder.Location = new System.Drawing.Point(184, 102);
            this.txtContactsFolder.Name = "txtContactsFolder";
            this.txtContactsFolder.Size = new System.Drawing.Size(350, 20);

            // Configure Checkbox: Enable Contact Saving
            this.chkEnableContactSaving.AutoSize = true;
            this.chkEnableContactSaving.Location = new System.Drawing.Point(184, 128);
            this.chkEnableContactSaving.Name = "chkEnableContactSaving";
            this.chkEnableContactSaving.Size = new System.Drawing.Size(140, 17);
            this.chkEnableContactSaving.Text = "Enable Contact Saving";
            this.chkEnableContactSaving.UseVisualStyleBackColor = true;
            this.chkEnableContactSaving.CheckedChanged += new EventHandler(this.chkEnableContactSaving_CheckedChanged);

            // Configure Checkbox: Search Entire Vault For Contacts
            this.chkSearchEntireVaultForContacts = new CheckBox();
            this.chkSearchEntireVaultForContacts.AutoSize = true;
            this.chkSearchEntireVaultForContacts.Location = new System.Drawing.Point(340, 128);
            this.chkSearchEntireVaultForContacts.Name = "chkSearchEntireVaultForContacts";
            this.chkSearchEntireVaultForContacts.Size = new System.Drawing.Size(180, 17);
            this.chkSearchEntireVaultForContacts.Text = "Search entire vault for contacts";
            this.chkSearchEntireVaultForContacts.UseVisualStyleBackColor = true;

            // Configure Checkbox: Launch Obsidian
            this.chkLaunchObsidian.AutoSize = true;
            this.chkLaunchObsidian.Location = new System.Drawing.Point(184, 151);
            this.chkLaunchObsidian.Name = "chkLaunchObsidian";
            this.chkLaunchObsidian.Size = new System.Drawing.Size(140, 17);
            this.chkLaunchObsidian.Text = "Launch Obsidian after saving";
            this.chkLaunchObsidian.UseVisualStyleBackColor = true;

            // Configure Label: Delay
            this.lblDelay.AutoSize = true;
            this.lblDelay.Location = new System.Drawing.Point(12, 177);
            this.lblDelay.Name = "lblDelay";
            this.lblDelay.Size = new System.Drawing.Size(87, 13);
            this.lblDelay.Text = "Delay (seconds):";

            // Configure NumericUpDown: Delay
            this.numDelay.Location = new System.Drawing.Point(184, 175);
            this.numDelay.Name = "numDelay";
            this.numDelay.Size = new System.Drawing.Size(60, 20);
            this.numDelay.Minimum = 0;
            this.numDelay.Maximum = 10;
            
            // Configure Checkbox: Show Countdown
            this.chkShowCountdown.AutoSize = true;
            this.chkShowCountdown.Location = new System.Drawing.Point(267, 176);
            this.chkShowCountdown.Name = "chkShowCountdown";
            this.chkShowCountdown.Size = new System.Drawing.Size(109, 17);
            this.chkShowCountdown.Text = "Show countdown";
            this.chkShowCountdown.UseVisualStyleBackColor = true;

            // Configure Label: Follow-up Tasks
            this.lblFollowUpTasks.AutoSize = true;
            this.lblFollowUpTasks.Location = new System.Drawing.Point(12, 205);
            this.lblFollowUpTasks.Name = "lblFollowUpTasks";
            this.lblFollowUpTasks.Size = new System.Drawing.Size(87, 13);
            this.lblFollowUpTasks.Text = "Follow-up Tasks:";

            // Configure Checkbox: Create Obsidian Task
            this.chkCreateObsidianTask.AutoSize = true;
            this.chkCreateObsidianTask.Location = new System.Drawing.Point(184, 204);
            this.chkCreateObsidianTask.Name = "chkCreateObsidianTask";
            this.chkCreateObsidianTask.Size = new System.Drawing.Size(156, 17);
            this.chkCreateObsidianTask.Text = "Create task in Obsidian note";
            this.chkCreateObsidianTask.UseVisualStyleBackColor = true;

            // Configure Checkbox: Create Outlook Task
            this.chkCreateOutlookTask.AutoSize = true;
            this.chkCreateOutlookTask.Location = new System.Drawing.Point(184, 227);
            this.chkCreateOutlookTask.Name = "chkCreateOutlookTask";
            this.chkCreateOutlookTask.Size = new System.Drawing.Size(151, 17);
            this.chkCreateOutlookTask.Text = "Create task in Outlook";
            this.chkCreateOutlookTask.UseVisualStyleBackColor = true;

            // Configure Label: Default Due Days
            this.lblDefaultDueDays.AutoSize = true;
            this.lblDefaultDueDays.Location = new System.Drawing.Point(12, 254);
            this.lblDefaultDueDays.Name = "lblDefaultDueDays";
            this.lblDefaultDueDays.Size = new System.Drawing.Size(70, 13);
            this.lblDefaultDueDays.Text = "Due in Days:";

            // Configure NumericUpDown: Default Due Days
            this.numDefaultDueDays.Location = new System.Drawing.Point(184, 252);
            this.numDefaultDueDays.Name = "numDefaultDueDays";
            this.numDefaultDueDays.Size = new System.Drawing.Size(60, 20);
            this.numDefaultDueDays.Minimum = 0;
            this.numDefaultDueDays.Maximum = 30;

            // Configure Label: Due Days Help
            this.lblDueDaysHelp.AutoSize = true;
            this.lblDueDaysHelp.Location = new System.Drawing.Point(267, 254);
            this.lblDueDaysHelp.Name = "lblDueDaysHelp";
            this.lblDueDaysHelp.Size = new System.Drawing.Size(187, 13);
            this.lblDueDaysHelp.Text = "0 = Today, 1 = Tomorrow, etc.";

            // Configure Label: Default Reminder Days
            this.lblDefaultReminderDays.AutoSize = true;
            this.lblDefaultReminderDays.Location = new System.Drawing.Point(12, 280);
            this.lblDefaultReminderDays.Name = "lblDefaultReminderDays";
            this.lblDefaultReminderDays.Size = new System.Drawing.Size(81, 13);
            this.lblDefaultReminderDays.Text = "Reminder Days:";

            // Configure NumericUpDown: Default Reminder Days
            this.numDefaultReminderDays.Location = new System.Drawing.Point(184, 278);
            this.numDefaultReminderDays.Name = "numDefaultReminderDays";
            this.numDefaultReminderDays.Size = new System.Drawing.Size(60, 20);
            this.numDefaultReminderDays.Minimum = 0;
            this.numDefaultReminderDays.Maximum = 30;

            // Configure Label: Default Reminder Hour
            this.lblDefaultReminderHour.AutoSize = true;
            this.lblDefaultReminderHour.Location = new System.Drawing.Point(12, 306);
            this.lblDefaultReminderHour.Name = "lblDefaultReminderHour";
            this.lblDefaultReminderHour.Size = new System.Drawing.Size(81, 13);
            this.lblDefaultReminderHour.Text = "Reminder Hour:";

            // Configure NumericUpDown: Default Reminder Hour
            this.numDefaultReminderHour.Location = new System.Drawing.Point(184, 304);
            this.numDefaultReminderHour.Name = "numDefaultReminderHour";
            this.numDefaultReminderHour.Size = new System.Drawing.Size(60, 20);
            this.numDefaultReminderHour.Minimum = 0;
            this.numDefaultReminderHour.Maximum = 23;

            // Configure Checkbox: Ask for Dates
            this.chkAskForDates.AutoSize = true;
            this.chkAskForDates.Location = new System.Drawing.Point(267, 305);
            this.chkAskForDates.Name = "chkAskForDates";
            this.chkAskForDates.Size = new System.Drawing.Size(153, 17);
            this.chkAskForDates.Text = "Ask for dates and times each time";
            this.chkAskForDates.UseVisualStyleBackColor = true;

            // Configure Checkbox: Group Email Threads
            this.chkGroupEmailThreads.AutoSize = true;
            this.chkGroupEmailThreads.Location = new System.Drawing.Point(267, 328);
            this.chkGroupEmailThreads.Name = "chkGroupEmailThreads";
            this.chkGroupEmailThreads.Size = new System.Drawing.Size(124, 17);
            this.chkGroupEmailThreads.Text = "Group email threads";
            this.chkGroupEmailThreads.UseVisualStyleBackColor = true;

            // Configure Label: Subject Cleanup Patterns
            this.lblPatterns.AutoSize = true;
            this.lblPatterns.Location = new System.Drawing.Point(12, 350);
            this.lblPatterns.Name = "lblPatterns";
            this.lblPatterns.Size = new System.Drawing.Size(124, 13);
            this.lblPatterns.Text = "Subject Cleanup Patterns:";

            // Configure ListBox: Patterns
            this.lstPatterns.FormattingEnabled = true;
            this.lstPatterns.Location = new System.Drawing.Point(12, 370);
            this.lstPatterns.Name = "lstPatterns";
            this.lstPatterns.Size = new System.Drawing.Size(522, 160);
            this.lstPatterns.TabIndex = 0;

            // Configure Button: Add Pattern
            this.btnAdd.Location = new System.Drawing.Point(540, 370);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new EventHandler(this.BtnAdd_Click);

            // Configure Button: Edit Pattern
            this.btnEdit.Location = new System.Drawing.Point(540, 399);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 23);
            this.btnEdit.Text = "Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new EventHandler(this.BtnEdit_Click);

            // Configure Button: Remove Pattern
            this.btnRemove.Location = new System.Drawing.Point(540, 428);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new EventHandler(this.BtnRemove_Click);

            // Configure Button: Save
            this.btnSave.DialogResult = DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(447, 545);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new EventHandler(this.btnSave_Click);

            // Configure Button: Cancel
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(540, 545);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;

            // Form configuration
            this.AcceptButton = this.btnSave;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(635, 580);
            
            // Add all controls to the form
            this.Controls.Add(this.lblVaultName);
            this.Controls.Add(this.txtVaultName);
            this.Controls.Add(this.lblVaultPath);
            this.Controls.Add(this.txtVaultPath);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.lblInboxFolder);
            this.Controls.Add(this.txtInboxFolder);
            this.Controls.Add(this.lblContactsFolder);
            this.Controls.Add(this.txtContactsFolder);
            this.Controls.Add(this.chkEnableContactSaving);
            this.Controls.Add(this.chkSearchEntireVaultForContacts);
            this.Controls.Add(this.chkLaunchObsidian);
            this.Controls.Add(this.lblDelay);
            this.Controls.Add(this.numDelay);
            this.Controls.Add(this.chkShowCountdown);
            this.Controls.Add(this.lblFollowUpTasks);
            this.Controls.Add(this.chkCreateObsidianTask);
            this.Controls.Add(this.chkCreateOutlookTask);
            this.Controls.Add(this.lblDefaultDueDays);
            this.Controls.Add(this.numDefaultDueDays);
            this.Controls.Add(this.lblDueDaysHelp);
            this.Controls.Add(this.lblDefaultReminderDays);
            this.Controls.Add(this.numDefaultReminderDays);
            this.Controls.Add(this.lblDefaultReminderHour);
            this.Controls.Add(this.numDefaultReminderHour);
            this.Controls.Add(this.chkAskForDates);
            this.Controls.Add(this.chkGroupEmailThreads);
            this.Controls.Add(this.lblPatterns);
            this.Controls.Add(this.lstPatterns);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Text = "Obsidian Settings";
            this.Name = "SettingsForm";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // End initialization
            ((System.ComponentModel.ISupportInitialize)(this.numDelay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultDueDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderHour)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void LoadSettings()
        {
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            txtContactsFolder.Text = _settings.ContactsFolder;
            chkLaunchObsidian.Checked = _settings.LaunchObsidian;
            chkEnableContactSaving.Checked = _settings.EnableContactSaving;
            chkSearchEntireVaultForContacts.Checked = _settings.SearchEntireVaultForContacts;
            txtContactsFolder.Enabled = _settings.EnableContactSaving;
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
            _settings.ContactsFolder = txtContactsFolder.Text;
            _settings.LaunchObsidian = chkLaunchObsidian.Checked;
            _settings.EnableContactSaving = chkEnableContactSaving.Checked;
            _settings.SearchEntireVaultForContacts = chkSearchEntireVaultForContacts.Checked;
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

        private void chkEnableContactSaving_CheckedChanged(object sender, EventArgs e)
        {
            txtContactsFolder.Enabled = chkEnableContactSaving.Checked;
        }
    }
} 