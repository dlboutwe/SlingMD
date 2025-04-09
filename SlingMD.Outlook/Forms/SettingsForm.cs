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
            this.txtVaultName = new System.Windows.Forms.TextBox();
            this.txtVaultPath = new System.Windows.Forms.TextBox();
            this.txtInboxFolder = new System.Windows.Forms.TextBox();
            this.txtContactsFolder = new System.Windows.Forms.TextBox();
            this.chkLaunchObsidian = new System.Windows.Forms.CheckBox();
            this.chkEnableContactSaving = new System.Windows.Forms.CheckBox();
            this.chkSearchEntireVaultForContacts = new System.Windows.Forms.CheckBox();
            this.numDelay = new System.Windows.Forms.NumericUpDown();
            this.chkShowCountdown = new System.Windows.Forms.CheckBox();
            this.chkCreateObsidianTask = new System.Windows.Forms.CheckBox();
            this.chkCreateOutlookTask = new System.Windows.Forms.CheckBox();
            this.chkAskForDates = new System.Windows.Forms.CheckBox();
            this.chkGroupEmailThreads = new System.Windows.Forms.CheckBox();
            this.numDefaultDueDays = new System.Windows.Forms.NumericUpDown();
            this.numDefaultReminderDays = new System.Windows.Forms.NumericUpDown();
            this.numDefaultReminderHour = new System.Windows.Forms.NumericUpDown();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblVaultName = new System.Windows.Forms.Label();
            this.lblVaultPath = new System.Windows.Forms.Label();
            this.lblInboxFolder = new System.Windows.Forms.Label();
            this.lblContactsFolder = new System.Windows.Forms.Label();
            this.lblDelay = new System.Windows.Forms.Label();
            this.lblFollowUpTasks = new System.Windows.Forms.Label();
            this.lblDefaultDueDays = new System.Windows.Forms.Label();
            this.lblDefaultReminderDays = new System.Windows.Forms.Label();
            this.lblDefaultReminderHour = new System.Windows.Forms.Label();
            this.lblDueDaysHelp = new System.Windows.Forms.Label();
            this.lstPatterns = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.lblPatterns = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numDelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultDueDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderHour)).BeginInit();
            this.SuspendLayout();
            // 
            // txtVaultName
            // 
            this.txtVaultName.Location = new System.Drawing.Point(184, 12);
            this.txtVaultName.Name = "txtVaultName";
            this.txtVaultName.Size = new System.Drawing.Size(350, 26);
            this.txtVaultName.TabIndex = 1;
            // 
            // txtVaultPath
            // 
            this.txtVaultPath.Location = new System.Drawing.Point(184, 42);
            this.txtVaultPath.Name = "txtVaultPath";
            this.txtVaultPath.Size = new System.Drawing.Size(350, 26);
            this.txtVaultPath.TabIndex = 3;
            // 
            // txtInboxFolder
            // 
            this.txtInboxFolder.Location = new System.Drawing.Point(184, 72);
            this.txtInboxFolder.Name = "txtInboxFolder";
            this.txtInboxFolder.Size = new System.Drawing.Size(350, 26);
            this.txtInboxFolder.TabIndex = 6;
            // 
            // txtContactsFolder
            // 
            this.txtContactsFolder.Location = new System.Drawing.Point(184, 102);
            this.txtContactsFolder.Name = "txtContactsFolder";
            this.txtContactsFolder.Size = new System.Drawing.Size(350, 26);
            this.txtContactsFolder.TabIndex = 8;
            // 
            // chkLaunchObsidian
            // 
            this.chkLaunchObsidian.AutoSize = true;
            this.chkLaunchObsidian.Location = new System.Drawing.Point(184, 151);
            this.chkLaunchObsidian.Name = "chkLaunchObsidian";
            this.chkLaunchObsidian.Size = new System.Drawing.Size(240, 24);
            this.chkLaunchObsidian.TabIndex = 11;
            this.chkLaunchObsidian.Text = "Launch Obsidian after saving";
            this.chkLaunchObsidian.UseVisualStyleBackColor = true;
            // 
            // chkEnableContactSaving
            // 
            this.chkEnableContactSaving.AutoSize = true;
            this.chkEnableContactSaving.Location = new System.Drawing.Point(184, 128);
            this.chkEnableContactSaving.Name = "chkEnableContactSaving";
            this.chkEnableContactSaving.Size = new System.Drawing.Size(197, 24);
            this.chkEnableContactSaving.TabIndex = 9;
            this.chkEnableContactSaving.Text = "Enable Contact Saving";
            this.chkEnableContactSaving.UseVisualStyleBackColor = true;
            this.chkEnableContactSaving.CheckedChanged += new System.EventHandler(this.chkEnableContactSaving_CheckedChanged);
            // 
            // chkSearchEntireVaultForContacts
            // 
            this.chkSearchEntireVaultForContacts.AutoSize = true;
            this.chkSearchEntireVaultForContacts.Location = new System.Drawing.Point(387, 128);
            this.chkSearchEntireVaultForContacts.Name = "chkSearchEntireVaultForContacts";
            this.chkSearchEntireVaultForContacts.Size = new System.Drawing.Size(255, 24);
            this.chkSearchEntireVaultForContacts.TabIndex = 10;
            this.chkSearchEntireVaultForContacts.Text = "Search entire vault for contacts";
            this.chkSearchEntireVaultForContacts.UseVisualStyleBackColor = true;
            // 
            // numDelay
            // 
            this.numDelay.Location = new System.Drawing.Point(184, 175);
            this.numDelay.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numDelay.Name = "numDelay";
            this.numDelay.Size = new System.Drawing.Size(60, 26);
            this.numDelay.TabIndex = 13;
            // 
            // chkShowCountdown
            // 
            this.chkShowCountdown.AutoSize = true;
            this.chkShowCountdown.Location = new System.Drawing.Point(267, 176);
            this.chkShowCountdown.Name = "chkShowCountdown";
            this.chkShowCountdown.Size = new System.Drawing.Size(157, 24);
            this.chkShowCountdown.TabIndex = 14;
            this.chkShowCountdown.Text = "Show countdown";
            this.chkShowCountdown.UseVisualStyleBackColor = true;
            // 
            // chkCreateObsidianTask
            // 
            this.chkCreateObsidianTask.AutoSize = true;
            this.chkCreateObsidianTask.Location = new System.Drawing.Point(184, 204);
            this.chkCreateObsidianTask.Name = "chkCreateObsidianTask";
            this.chkCreateObsidianTask.Size = new System.Drawing.Size(235, 24);
            this.chkCreateObsidianTask.TabIndex = 16;
            this.chkCreateObsidianTask.Text = "Create task in Obsidian note";
            this.chkCreateObsidianTask.UseVisualStyleBackColor = true;
            // 
            // chkCreateOutlookTask
            // 
            this.chkCreateOutlookTask.AutoSize = true;
            this.chkCreateOutlookTask.Location = new System.Drawing.Point(184, 227);
            this.chkCreateOutlookTask.Name = "chkCreateOutlookTask";
            this.chkCreateOutlookTask.Size = new System.Drawing.Size(192, 24);
            this.chkCreateOutlookTask.TabIndex = 17;
            this.chkCreateOutlookTask.Text = "Create task in Outlook";
            this.chkCreateOutlookTask.UseVisualStyleBackColor = true;
            // 
            // chkAskForDates
            // 
            this.chkAskForDates.AutoSize = true;
            this.chkAskForDates.Location = new System.Drawing.Point(267, 305);
            this.chkAskForDates.Name = "chkAskForDates";
            this.chkAskForDates.Size = new System.Drawing.Size(275, 24);
            this.chkAskForDates.TabIndex = 25;
            this.chkAskForDates.Text = "Ask for dates and times each time";
            this.chkAskForDates.UseVisualStyleBackColor = true;
            // 
            // chkGroupEmailThreads
            // 
            this.chkGroupEmailThreads.AutoSize = true;
            this.chkGroupEmailThreads.Location = new System.Drawing.Point(267, 328);
            this.chkGroupEmailThreads.Name = "chkGroupEmailThreads";
            this.chkGroupEmailThreads.Size = new System.Drawing.Size(179, 24);
            this.chkGroupEmailThreads.TabIndex = 26;
            this.chkGroupEmailThreads.Text = "Group email threads";
            this.chkGroupEmailThreads.UseVisualStyleBackColor = true;
            // 
            // numDefaultDueDays
            // 
            this.numDefaultDueDays.Location = new System.Drawing.Point(184, 252);
            this.numDefaultDueDays.Maximum = new decimal(new int[] {
            30,
            0,
            0,
            0});
            this.numDefaultDueDays.Name = "numDefaultDueDays";
            this.numDefaultDueDays.Size = new System.Drawing.Size(60, 26);
            this.numDefaultDueDays.TabIndex = 19;
            // 
            // numDefaultReminderDays
            // 
            this.numDefaultReminderDays.Location = new System.Drawing.Point(184, 278);
            this.numDefaultReminderDays.Maximum = new decimal(new int[] {
            30,
            0,
            0,
            0});
            this.numDefaultReminderDays.Name = "numDefaultReminderDays";
            this.numDefaultReminderDays.Size = new System.Drawing.Size(60, 26);
            this.numDefaultReminderDays.TabIndex = 22;
            // 
            // numDefaultReminderHour
            // 
            this.numDefaultReminderHour.Location = new System.Drawing.Point(184, 304);
            this.numDefaultReminderHour.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.numDefaultReminderHour.Name = "numDefaultReminderHour";
            this.numDefaultReminderHour.Size = new System.Drawing.Size(60, 26);
            this.numDefaultReminderHour.TabIndex = 24;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(540, 42);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 33);
            this.btnBrowse.TabIndex = 4;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnSave
            // 
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(447, 545);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 40);
            this.btnSave.TabIndex = 31;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(540, 545);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 40);
            this.btnCancel.TabIndex = 32;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // lblVaultName
            // 
            this.lblVaultName.AutoSize = true;
            this.lblVaultName.Location = new System.Drawing.Point(12, 15);
            this.lblVaultName.Name = "lblVaultName";
            this.lblVaultName.Size = new System.Drawing.Size(96, 20);
            this.lblVaultName.TabIndex = 0;
            this.lblVaultName.Text = "Vault Name:";
            // 
            // lblVaultPath
            // 
            this.lblVaultPath.AutoSize = true;
            this.lblVaultPath.Location = new System.Drawing.Point(12, 45);
            this.lblVaultPath.Name = "lblVaultPath";
            this.lblVaultPath.Size = new System.Drawing.Size(128, 20);
            this.lblVaultPath.TabIndex = 2;
            this.lblVaultPath.Text = "Vault Base Path:";
            // 
            // lblInboxFolder
            // 
            this.lblInboxFolder.AutoSize = true;
            this.lblInboxFolder.Location = new System.Drawing.Point(12, 75);
            this.lblInboxFolder.Name = "lblInboxFolder";
            this.lblInboxFolder.Size = new System.Drawing.Size(101, 20);
            this.lblInboxFolder.TabIndex = 5;
            this.lblInboxFolder.Text = "Inbox Folder:";
            // 
            // lblContactsFolder
            // 
            this.lblContactsFolder.AutoSize = true;
            this.lblContactsFolder.Location = new System.Drawing.Point(12, 105);
            this.lblContactsFolder.Name = "lblContactsFolder";
            this.lblContactsFolder.Size = new System.Drawing.Size(126, 20);
            this.lblContactsFolder.TabIndex = 7;
            this.lblContactsFolder.Text = "Contacts Folder:";
            // 
            // lblDelay
            // 
            this.lblDelay.AutoSize = true;
            this.lblDelay.Location = new System.Drawing.Point(12, 177);
            this.lblDelay.Name = "lblDelay";
            this.lblDelay.Size = new System.Drawing.Size(127, 20);
            this.lblDelay.TabIndex = 12;
            this.lblDelay.Text = "Delay (seconds):";
            // 
            // lblFollowUpTasks
            // 
            this.lblFollowUpTasks.AutoSize = true;
            this.lblFollowUpTasks.Location = new System.Drawing.Point(12, 205);
            this.lblFollowUpTasks.Name = "lblFollowUpTasks";
            this.lblFollowUpTasks.Size = new System.Drawing.Size(127, 20);
            this.lblFollowUpTasks.TabIndex = 15;
            this.lblFollowUpTasks.Text = "Follow-up Tasks:";
            // 
            // lblDefaultDueDays
            // 
            this.lblDefaultDueDays.AutoSize = true;
            this.lblDefaultDueDays.Location = new System.Drawing.Point(12, 254);
            this.lblDefaultDueDays.Name = "lblDefaultDueDays";
            this.lblDefaultDueDays.Size = new System.Drawing.Size(99, 20);
            this.lblDefaultDueDays.TabIndex = 18;
            this.lblDefaultDueDays.Text = "Due in Days:";
            // 
            // lblDefaultReminderDays
            // 
            this.lblDefaultReminderDays.AutoSize = true;
            this.lblDefaultReminderDays.Location = new System.Drawing.Point(12, 280);
            this.lblDefaultReminderDays.Name = "lblDefaultReminderDays";
            this.lblDefaultReminderDays.Size = new System.Drawing.Size(122, 20);
            this.lblDefaultReminderDays.TabIndex = 21;
            this.lblDefaultReminderDays.Text = "Reminder Days:";
            // 
            // lblDefaultReminderHour
            // 
            this.lblDefaultReminderHour.AutoSize = true;
            this.lblDefaultReminderHour.Location = new System.Drawing.Point(12, 306);
            this.lblDefaultReminderHour.Name = "lblDefaultReminderHour";
            this.lblDefaultReminderHour.Size = new System.Drawing.Size(121, 20);
            this.lblDefaultReminderHour.TabIndex = 23;
            this.lblDefaultReminderHour.Text = "Reminder Hour:";
            // 
            // lblDueDaysHelp
            // 
            this.lblDueDaysHelp.AutoSize = true;
            this.lblDueDaysHelp.Location = new System.Drawing.Point(267, 254);
            this.lblDueDaysHelp.Name = "lblDueDaysHelp";
            this.lblDueDaysHelp.Size = new System.Drawing.Size(216, 20);
            this.lblDueDaysHelp.TabIndex = 20;
            this.lblDueDaysHelp.Text = "0 = Today, 1 = Tomorrow, etc.";
            // 
            // lstPatterns
            // 
            this.lstPatterns.FormattingEnabled = true;
            this.lstPatterns.ItemHeight = 20;
            this.lstPatterns.Location = new System.Drawing.Point(12, 370);
            this.lstPatterns.Name = "lstPatterns";
            this.lstPatterns.Size = new System.Drawing.Size(522, 144);
            this.lstPatterns.TabIndex = 0;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(540, 370);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(92, 39);
            this.btnAdd.TabIndex = 28;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(540, 415);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(92, 37);
            this.btnEdit.TabIndex = 29;
            this.btnEdit.Text = "Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.BtnEdit_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(540, 458);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(92, 37);
            this.btnRemove.TabIndex = 30;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.BtnRemove_Click);
            // 
            // lblPatterns
            // 
            this.lblPatterns.AutoSize = true;
            this.lblPatterns.Location = new System.Drawing.Point(12, 350);
            this.lblPatterns.Name = "lblPatterns";
            this.lblPatterns.Size = new System.Drawing.Size(194, 20);
            this.lblPatterns.TabIndex = 27;
            this.lblPatterns.Text = "Subject Cleanup Patterns:";
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnSave;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(648, 597);
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
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Obsidian Settings";
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