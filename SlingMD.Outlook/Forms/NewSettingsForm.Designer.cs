namespace SlingMD.Outlook.Forms
{
    partial class Settings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.grpVaultSettings = new System.Windows.Forms.GroupBox();
            this.btnBrowseVaultPath = new System.Windows.Forms.Button();
            this.txtAppointmentsFolder = new System.Windows.Forms.TextBox();
            this.txtContactsFolder = new System.Windows.Forms.TextBox();
            this.txtInboxFolder = new System.Windows.Forms.TextBox();
            this.txtVaultBasePath = new System.Windows.Forms.TextBox();
            this.txtVaultName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.nbrDefaultReminderHour = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.nbrObsidianDelaySeconds = new System.Windows.Forms.NumericUpDown();
            this.nbrDefaultReminderDays = new System.Windows.Forms.NumericUpDown();
            this.nbrDefaultDueDays = new System.Windows.Forms.NumericUpDown();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbShowCountdown = new System.Windows.Forms.CheckBox();
            this.cbSearchEntireVaultForContacts = new System.Windows.Forms.CheckBox();
            this.cbLaunchObsidian = new System.Windows.Forms.CheckBox();
            this.cbEnableContactSaving = new System.Windows.Forms.CheckBox();
            this.cbGroupEmailThreads = new System.Windows.Forms.CheckBox();
            this.cbCreateOutlookTask = new System.Windows.Forms.CheckBox();
            this.cbAskForDates = new System.Windows.Forms.CheckBox();
            this.cbCreateObsidianTask = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnRemoveSubjectCleanupPattern = new System.Windows.Forms.Button();
            this.btnEditSubjectCleanupPattern = new System.Windows.Forms.Button();
            this.btnAddSubjectCleanupPattern = new System.Windows.Forms.Button();
            this.lstSubjectCleanupPatterns = new System.Windows.Forms.ListBox();
            this.tabGrp = new System.Windows.Forms.TabControl();
            this.tabMailSettings = new System.Windows.Forms.TabPage();
            this.cbMoveDateToFrontInThread = new System.Windows.Forms.CheckBox();
            this.cbNoteTitleIncludeDate = new System.Windows.Forms.CheckBox();
            this.nbrNoteTitleMaxLength = new System.Windows.Forms.NumericUpDown();
            this.txtNoteTitleFormat = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.tabAppointmentSettings = new System.Windows.Forms.TabPage();
            this.cbAppointmentSaveAttachments = new System.Windows.Forms.CheckBox();
            this.nbrAppointmentNoteTitleMaxLength = new System.Windows.Forms.NumericUpDown();
            this.txtAppointmentNoteTitleFormat = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.tabDeveloperSettings = new System.Windows.Forms.TabPage();
            this.cbShowThreadDebug = new System.Windows.Forms.CheckBox();
            this.cbShowDevelopmentSettings = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.tabTaskSettings = new System.Windows.Forms.TabPage();
            this.lstDefaultNoteTags = new System.Windows.Forms.ListBox();
            this.btnRemoveNoteTag = new System.Windows.Forms.Button();
            this.btnEditNoteTag = new System.Windows.Forms.Button();
            this.btnAddNoteTag = new System.Windows.Forms.Button();
            this.btnRemoveAppointmentTag = new System.Windows.Forms.Button();
            this.btnEditAppointmentTag = new System.Windows.Forms.Button();
            this.lstDefaultAppointmentTags = new System.Windows.Forms.ListBox();
            this.btnAddAppointmentTag = new System.Windows.Forms.Button();
            this.btnRemoveTaskTag = new System.Windows.Forms.Button();
            this.btnEditTaskTag = new System.Windows.Forms.Button();
            this.lstDefaultTaskTags = new System.Windows.Forms.ListBox();
            this.btnAddTaskTag = new System.Windows.Forms.Button();
            this.grpVaultSettings.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultReminderHour)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrObsidianDelaySeconds)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultReminderDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultDueDays)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabGrp.SuspendLayout();
            this.tabMailSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrNoteTitleMaxLength)).BeginInit();
            this.tabAppointmentSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrAppointmentNoteTitleMaxLength)).BeginInit();
            this.tabDeveloperSettings.SuspendLayout();
            this.tabTaskSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpVaultSettings
            // 
            this.grpVaultSettings.Controls.Add(this.btnBrowseVaultPath);
            this.grpVaultSettings.Controls.Add(this.txtAppointmentsFolder);
            this.grpVaultSettings.Controls.Add(this.txtContactsFolder);
            this.grpVaultSettings.Controls.Add(this.txtInboxFolder);
            this.grpVaultSettings.Controls.Add(this.txtVaultBasePath);
            this.grpVaultSettings.Controls.Add(this.txtVaultName);
            this.grpVaultSettings.Controls.Add(this.label5);
            this.grpVaultSettings.Controls.Add(this.label4);
            this.grpVaultSettings.Controls.Add(this.label3);
            this.grpVaultSettings.Controls.Add(this.label2);
            this.grpVaultSettings.Controls.Add(this.label1);
            this.grpVaultSettings.Location = new System.Drawing.Point(8, 8);
            this.grpVaultSettings.Margin = new System.Windows.Forms.Padding(2);
            this.grpVaultSettings.Name = "grpVaultSettings";
            this.grpVaultSettings.Padding = new System.Windows.Forms.Padding(2);
            this.grpVaultSettings.Size = new System.Drawing.Size(438, 128);
            this.grpVaultSettings.TabIndex = 0;
            this.grpVaultSettings.TabStop = false;
            this.grpVaultSettings.Text = "Vault Settings";
            // 
            // btnBrowseVaultPath
            // 
            this.btnBrowseVaultPath.BackgroundImage = global::SlingMD.Outlook.Properties.Resources.folder_open_regular_full;
            this.btnBrowseVaultPath.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseVaultPath.Location = new System.Drawing.Point(92, 37);
            this.btnBrowseVaultPath.Margin = new System.Windows.Forms.Padding(2);
            this.btnBrowseVaultPath.Name = "btnBrowseVaultPath";
            this.btnBrowseVaultPath.Size = new System.Drawing.Size(21, 21);
            this.btnBrowseVaultPath.TabIndex = 1;
            this.btnBrowseVaultPath.UseVisualStyleBackColor = true;
            this.btnBrowseVaultPath.Click += new System.EventHandler(this.btnBrowseVaultPath_Click);
            // 
            // txtAppointmentsFolder
            // 
            this.txtAppointmentsFolder.Location = new System.Drawing.Point(140, 97);
            this.txtAppointmentsFolder.Margin = new System.Windows.Forms.Padding(2);
            this.txtAppointmentsFolder.Name = "txtAppointmentsFolder";
            this.txtAppointmentsFolder.Size = new System.Drawing.Size(288, 20);
            this.txtAppointmentsFolder.TabIndex = 9;
            // 
            // txtContactsFolder
            // 
            this.txtContactsFolder.Location = new System.Drawing.Point(140, 77);
            this.txtContactsFolder.Margin = new System.Windows.Forms.Padding(2);
            this.txtContactsFolder.Name = "txtContactsFolder";
            this.txtContactsFolder.Size = new System.Drawing.Size(288, 20);
            this.txtContactsFolder.TabIndex = 8;
            // 
            // txtInboxFolder
            // 
            this.txtInboxFolder.Location = new System.Drawing.Point(140, 57);
            this.txtInboxFolder.Margin = new System.Windows.Forms.Padding(2);
            this.txtInboxFolder.Name = "txtInboxFolder";
            this.txtInboxFolder.Size = new System.Drawing.Size(288, 20);
            this.txtInboxFolder.TabIndex = 7;
            // 
            // txtVaultBasePath
            // 
            this.txtVaultBasePath.Location = new System.Drawing.Point(114, 37);
            this.txtVaultBasePath.Margin = new System.Windows.Forms.Padding(2);
            this.txtVaultBasePath.Name = "txtVaultBasePath";
            this.txtVaultBasePath.Size = new System.Drawing.Size(314, 20);
            this.txtVaultBasePath.TabIndex = 6;
            // 
            // txtVaultName
            // 
            this.txtVaultName.Location = new System.Drawing.Point(92, 17);
            this.txtVaultName.Margin = new System.Windows.Forms.Padding(2);
            this.txtVaultName.Name = "txtVaultName";
            this.txtVaultName.Size = new System.Drawing.Size(336, 20);
            this.txtVaultName.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 101);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(131, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Appointments Folder Path:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 81);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(109, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Contacts Folder Path:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 61);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Inbox Folder Path:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 41);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Vault Base Path:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Vault Name:";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.UserProfile;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.nbrDefaultReminderHour);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.nbrObsidianDelaySeconds);
            this.groupBox1.Controls.Add(this.nbrDefaultReminderDays);
            this.groupBox1.Controls.Add(this.nbrDefaultDueDays);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(8, 258);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(226, 101);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Timing Settings";
            // 
            // nbrDefaultReminderHour
            // 
            this.nbrDefaultReminderHour.Location = new System.Drawing.Point(88, 77);
            this.nbrDefaultReminderHour.Margin = new System.Windows.Forms.Padding(2);
            this.nbrDefaultReminderHour.Name = "nbrDefaultReminderHour";
            this.nbrDefaultReminderHour.Size = new System.Drawing.Size(43, 20);
            this.nbrDefaultReminderHour.TabIndex = 7;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(4, 81);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(81, 13);
            this.label9.TabIndex = 6;
            this.label9.Text = "Reminder Hour:";
            // 
            // nbrObsidianDelaySeconds
            // 
            this.nbrObsidianDelaySeconds.Location = new System.Drawing.Point(88, 17);
            this.nbrObsidianDelaySeconds.Margin = new System.Windows.Forms.Padding(2);
            this.nbrObsidianDelaySeconds.Name = "nbrObsidianDelaySeconds";
            this.nbrObsidianDelaySeconds.Size = new System.Drawing.Size(43, 20);
            this.nbrObsidianDelaySeconds.TabIndex = 5;
            this.toolTip1.SetToolTip(this.nbrObsidianDelaySeconds, "Number of seconds to delay when launching obsidian");
            // 
            // nbrDefaultReminderDays
            // 
            this.nbrDefaultReminderDays.Location = new System.Drawing.Point(88, 57);
            this.nbrDefaultReminderDays.Margin = new System.Windows.Forms.Padding(2);
            this.nbrDefaultReminderDays.Name = "nbrDefaultReminderDays";
            this.nbrDefaultReminderDays.Size = new System.Drawing.Size(43, 20);
            this.nbrDefaultReminderDays.TabIndex = 4;
            // 
            // nbrDefaultDueDays
            // 
            this.nbrDefaultDueDays.Location = new System.Drawing.Point(88, 37);
            this.nbrDefaultDueDays.Margin = new System.Windows.Forms.Padding(2);
            this.nbrDefaultDueDays.Name = "nbrDefaultDueDays";
            this.nbrDefaultDueDays.Size = new System.Drawing.Size(43, 20);
            this.nbrDefaultDueDays.TabIndex = 3;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(4, 61);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(82, 13);
            this.label8.TabIndex = 2;
            this.label8.Text = "Reminder Days:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(4, 41);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Due in Days:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 21);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Delay:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbShowCountdown);
            this.groupBox2.Controls.Add(this.cbSearchEntireVaultForContacts);
            this.groupBox2.Controls.Add(this.cbLaunchObsidian);
            this.groupBox2.Controls.Add(this.cbEnableContactSaving);
            this.groupBox2.Location = new System.Drawing.Point(8, 140);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(226, 114);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "General Settings";
            // 
            // cbShowCountdown
            // 
            this.cbShowCountdown.AutoSize = true;
            this.cbShowCountdown.Location = new System.Drawing.Point(4, 66);
            this.cbShowCountdown.Name = "cbShowCountdown";
            this.cbShowCountdown.Size = new System.Drawing.Size(109, 17);
            this.cbShowCountdown.TabIndex = 5;
            this.cbShowCountdown.Text = "Show countdown";
            this.cbShowCountdown.UseVisualStyleBackColor = true;
            // 
            // cbSearchEntireVaultForContacts
            // 
            this.cbSearchEntireVaultForContacts.AutoSize = true;
            this.cbSearchEntireVaultForContacts.Location = new System.Drawing.Point(4, 42);
            this.cbSearchEntireVaultForContacts.Name = "cbSearchEntireVaultForContacts";
            this.cbSearchEntireVaultForContacts.Size = new System.Drawing.Size(174, 17);
            this.cbSearchEntireVaultForContacts.TabIndex = 4;
            this.cbSearchEntireVaultForContacts.Text = "Search entire vault for contacts";
            this.cbSearchEntireVaultForContacts.UseVisualStyleBackColor = true;
            // 
            // cbLaunchObsidian
            // 
            this.cbLaunchObsidian.AutoSize = true;
            this.cbLaunchObsidian.Location = new System.Drawing.Point(4, 89);
            this.cbLaunchObsidian.Name = "cbLaunchObsidian";
            this.cbLaunchObsidian.Size = new System.Drawing.Size(164, 17);
            this.cbLaunchObsidian.TabIndex = 1;
            this.cbLaunchObsidian.Text = "Launch Obsidian after saving";
            this.cbLaunchObsidian.UseVisualStyleBackColor = true;
            // 
            // cbEnableContactSaving
            // 
            this.cbEnableContactSaving.AutoSize = true;
            this.cbEnableContactSaving.Location = new System.Drawing.Point(4, 19);
            this.cbEnableContactSaving.Name = "cbEnableContactSaving";
            this.cbEnableContactSaving.Size = new System.Drawing.Size(132, 17);
            this.cbEnableContactSaving.TabIndex = 0;
            this.cbEnableContactSaving.Text = "Enable contact saving";
            this.cbEnableContactSaving.UseVisualStyleBackColor = true;
            this.cbEnableContactSaving.CheckedChanged += new System.EventHandler(this.cbEnableContactSaving_CheckedChanged);
            // 
            // cbGroupEmailThreads
            // 
            this.cbGroupEmailThreads.AutoSize = true;
            this.cbGroupEmailThreads.Location = new System.Drawing.Point(4, 57);
            this.cbGroupEmailThreads.Name = "cbGroupEmailThreads";
            this.cbGroupEmailThreads.Size = new System.Drawing.Size(120, 17);
            this.cbGroupEmailThreads.TabIndex = 7;
            this.cbGroupEmailThreads.Text = "Group email threads";
            this.cbGroupEmailThreads.UseVisualStyleBackColor = true;
            // 
            // cbCreateOutlookTask
            // 
            this.cbCreateOutlookTask.AutoSize = true;
            this.cbCreateOutlookTask.Location = new System.Drawing.Point(4, 9);
            this.cbCreateOutlookTask.Name = "cbCreateOutlookTask";
            this.cbCreateOutlookTask.Size = new System.Drawing.Size(131, 17);
            this.cbCreateOutlookTask.TabIndex = 6;
            this.cbCreateOutlookTask.Text = "Create task in Outlook";
            this.cbCreateOutlookTask.UseVisualStyleBackColor = true;
            // 
            // cbAskForDates
            // 
            this.cbAskForDates.AutoSize = true;
            this.cbAskForDates.Location = new System.Drawing.Point(4, 55);
            this.cbAskForDates.Name = "cbAskForDates";
            this.cbAskForDates.Size = new System.Drawing.Size(166, 30);
            this.cbAskForDates.TabIndex = 3;
            this.cbAskForDates.Text = "Ask for dates and times each \r\ntime";
            this.cbAskForDates.UseVisualStyleBackColor = true;
            // 
            // cbCreateObsidianTask
            // 
            this.cbCreateObsidianTask.AutoSize = true;
            this.cbCreateObsidianTask.Location = new System.Drawing.Point(4, 32);
            this.cbCreateObsidianTask.Name = "cbCreateObsidianTask";
            this.cbCreateObsidianTask.Size = new System.Drawing.Size(159, 17);
            this.cbCreateObsidianTask.TabIndex = 2;
            this.cbCreateObsidianTask.Text = "Create task in Obsidian note";
            this.cbCreateObsidianTask.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnRemoveSubjectCleanupPattern);
            this.groupBox3.Controls.Add(this.btnEditSubjectCleanupPattern);
            this.groupBox3.Controls.Add(this.btnAddSubjectCleanupPattern);
            this.groupBox3.Controls.Add(this.lstSubjectCleanupPatterns);
            this.groupBox3.Location = new System.Drawing.Point(240, 142);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(206, 217);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Subject Cleanup Patterns";
            // 
            // btnRemoveSubjectCleanupPattern
            // 
            this.btnRemoveSubjectCleanupPattern.Location = new System.Drawing.Point(136, 186);
            this.btnRemoveSubjectCleanupPattern.Name = "btnRemoveSubjectCleanupPattern";
            this.btnRemoveSubjectCleanupPattern.Size = new System.Drawing.Size(63, 23);
            this.btnRemoveSubjectCleanupPattern.TabIndex = 3;
            this.btnRemoveSubjectCleanupPattern.Text = "Remove";
            this.btnRemoveSubjectCleanupPattern.UseVisualStyleBackColor = true;
            this.btnRemoveSubjectCleanupPattern.Click += new System.EventHandler(this.btnRemoveSubjectCleanupPattern_Click);
            // 
            // btnEditSubjectCleanupPattern
            // 
            this.btnEditSubjectCleanupPattern.Location = new System.Drawing.Point(71, 186);
            this.btnEditSubjectCleanupPattern.Name = "btnEditSubjectCleanupPattern";
            this.btnEditSubjectCleanupPattern.Size = new System.Drawing.Size(63, 23);
            this.btnEditSubjectCleanupPattern.TabIndex = 2;
            this.btnEditSubjectCleanupPattern.Text = "Edit";
            this.btnEditSubjectCleanupPattern.UseVisualStyleBackColor = true;
            this.btnEditSubjectCleanupPattern.Click += new System.EventHandler(this.btnEditSubjectCleanupPattern_Click);
            // 
            // btnAddSubjectCleanupPattern
            // 
            this.btnAddSubjectCleanupPattern.Location = new System.Drawing.Point(6, 186);
            this.btnAddSubjectCleanupPattern.Name = "btnAddSubjectCleanupPattern";
            this.btnAddSubjectCleanupPattern.Size = new System.Drawing.Size(63, 23);
            this.btnAddSubjectCleanupPattern.TabIndex = 1;
            this.btnAddSubjectCleanupPattern.Text = "Add";
            this.btnAddSubjectCleanupPattern.UseVisualStyleBackColor = true;
            this.btnAddSubjectCleanupPattern.Click += new System.EventHandler(this.btnAddSubjectCleanupPattern_Click);
            // 
            // lstSubjectCleanupPatterns
            // 
            this.lstSubjectCleanupPatterns.FormattingEnabled = true;
            this.lstSubjectCleanupPatterns.Location = new System.Drawing.Point(6, 17);
            this.lstSubjectCleanupPatterns.Name = "lstSubjectCleanupPatterns";
            this.lstSubjectCleanupPatterns.Size = new System.Drawing.Size(193, 160);
            this.lstSubjectCleanupPatterns.TabIndex = 0;
            // 
            // tabGrp
            // 
            this.tabGrp.Controls.Add(this.tabMailSettings);
            this.tabGrp.Controls.Add(this.tabAppointmentSettings);
            this.tabGrp.Controls.Add(this.tabTaskSettings);
            this.tabGrp.Controls.Add(this.tabDeveloperSettings);
            this.tabGrp.Location = new System.Drawing.Point(8, 364);
            this.tabGrp.Name = "tabGrp";
            this.tabGrp.SelectedIndex = 0;
            this.tabGrp.Size = new System.Drawing.Size(438, 189);
            this.tabGrp.TabIndex = 4;
            // 
            // tabMailSettings
            // 
            this.tabMailSettings.Controls.Add(this.btnRemoveNoteTag);
            this.tabMailSettings.Controls.Add(this.btnEditNoteTag);
            this.tabMailSettings.Controls.Add(this.lstDefaultNoteTags);
            this.tabMailSettings.Controls.Add(this.btnAddNoteTag);
            this.tabMailSettings.Controls.Add(this.cbGroupEmailThreads);
            this.tabMailSettings.Controls.Add(this.cbMoveDateToFrontInThread);
            this.tabMailSettings.Controls.Add(this.cbNoteTitleIncludeDate);
            this.tabMailSettings.Controls.Add(this.nbrNoteTitleMaxLength);
            this.tabMailSettings.Controls.Add(this.txtNoteTitleFormat);
            this.tabMailSettings.Controls.Add(this.label13);
            this.tabMailSettings.Controls.Add(this.label12);
            this.tabMailSettings.Controls.Add(this.label11);
            this.tabMailSettings.Location = new System.Drawing.Point(4, 22);
            this.tabMailSettings.Name = "tabMailSettings";
            this.tabMailSettings.Padding = new System.Windows.Forms.Padding(3);
            this.tabMailSettings.Size = new System.Drawing.Size(430, 163);
            this.tabMailSettings.TabIndex = 0;
            this.tabMailSettings.Text = "Email Settings";
            this.tabMailSettings.UseVisualStyleBackColor = true;
            // 
            // cbMoveDateToFrontInThread
            // 
            this.cbMoveDateToFrontInThread.AutoSize = true;
            this.cbMoveDateToFrontInThread.Location = new System.Drawing.Point(4, 103);
            this.cbMoveDateToFrontInThread.Name = "cbMoveDateToFrontInThread";
            this.cbMoveDateToFrontInThread.Size = new System.Drawing.Size(170, 30);
            this.cbMoveDateToFrontInThread.TabIndex = 9;
            this.cbMoveDateToFrontInThread.Text = "Move date to front of filename \r\nwhen grouping threads";
            this.cbMoveDateToFrontInThread.UseVisualStyleBackColor = true;
            // 
            // cbNoteTitleIncludeDate
            // 
            this.cbNoteTitleIncludeDate.AutoSize = true;
            this.cbNoteTitleIncludeDate.Location = new System.Drawing.Point(4, 80);
            this.cbNoteTitleIncludeDate.Name = "cbNoteTitleIncludeDate";
            this.cbNoteTitleIncludeDate.Size = new System.Drawing.Size(115, 17);
            this.cbNoteTitleIncludeDate.TabIndex = 8;
            this.cbNoteTitleIncludeDate.Text = "Include date in title";
            this.cbNoteTitleIncludeDate.UseVisualStyleBackColor = true;
            this.cbNoteTitleIncludeDate.CheckedChanged += new System.EventHandler(this.cbNoteTitleIncludeDate_CheckedChanged);
            // 
            // nbrNoteTitleMaxLength
            // 
            this.nbrNoteTitleMaxLength.Location = new System.Drawing.Point(111, 31);
            this.nbrNoteTitleMaxLength.Name = "nbrNoteTitleMaxLength";
            this.nbrNoteTitleMaxLength.Size = new System.Drawing.Size(40, 20);
            this.nbrNoteTitleMaxLength.TabIndex = 7;
            // 
            // txtNoteTitleFormat
            // 
            this.txtNoteTitleFormat.Location = new System.Drawing.Point(110, 5);
            this.txtNoteTitleFormat.Name = "txtNoteTitleFormat";
            this.txtNoteTitleFormat.Size = new System.Drawing.Size(314, 20);
            this.txtNoteTitleFormat.TabIndex = 6;
            this.toolTip1.SetToolTip(this.txtNoteTitleFormat, "Use {Subject}, {Sender}, {Date}");
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(4, 35);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(89, 13);
            this.label13.TabIndex = 3;
            this.label13.Text = "Max Title Length:";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(4, 9);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(91, 13);
            this.label12.TabIndex = 2;
            this.label12.Text = "Note Title Format:\r\n";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(253, 41);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(94, 13);
            this.label11.TabIndex = 1;
            this.label11.Text = "Default Note Tags";
            this.toolTip1.SetToolTip(this.label11, "Comma Separated");
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(234, 45);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(130, 13);
            this.label10.TabIndex = 0;
            this.label10.Text = "Default Appointment Tags";
            this.toolTip1.SetToolTip(this.label10, "Comma Separated");
            // 
            // tabAppointmentSettings
            // 
            this.tabAppointmentSettings.Controls.Add(this.btnRemoveAppointmentTag);
            this.tabAppointmentSettings.Controls.Add(this.btnEditAppointmentTag);
            this.tabAppointmentSettings.Controls.Add(this.lstDefaultAppointmentTags);
            this.tabAppointmentSettings.Controls.Add(this.btnAddAppointmentTag);
            this.tabAppointmentSettings.Controls.Add(this.cbAppointmentSaveAttachments);
            this.tabAppointmentSettings.Controls.Add(this.nbrAppointmentNoteTitleMaxLength);
            this.tabAppointmentSettings.Controls.Add(this.txtAppointmentNoteTitleFormat);
            this.tabAppointmentSettings.Controls.Add(this.label14);
            this.tabAppointmentSettings.Controls.Add(this.label15);
            this.tabAppointmentSettings.Controls.Add(this.label10);
            this.tabAppointmentSettings.Location = new System.Drawing.Point(4, 22);
            this.tabAppointmentSettings.Name = "tabAppointmentSettings";
            this.tabAppointmentSettings.Padding = new System.Windows.Forms.Padding(3);
            this.tabAppointmentSettings.Size = new System.Drawing.Size(430, 163);
            this.tabAppointmentSettings.TabIndex = 1;
            this.tabAppointmentSettings.Text = "Appointment Settings";
            this.tabAppointmentSettings.UseVisualStyleBackColor = true;
            // 
            // cbAppointmentSaveAttachments
            // 
            this.cbAppointmentSaveAttachments.AutoSize = true;
            this.cbAppointmentSaveAttachments.Location = new System.Drawing.Point(4, 61);
            this.cbAppointmentSaveAttachments.Name = "cbAppointmentSaveAttachments";
            this.cbAppointmentSaveAttachments.Size = new System.Drawing.Size(113, 17);
            this.cbAppointmentSaveAttachments.TabIndex = 16;
            this.cbAppointmentSaveAttachments.Text = "Save Attachments";
            this.cbAppointmentSaveAttachments.UseVisualStyleBackColor = true;
            // 
            // nbrAppointmentNoteTitleMaxLength
            // 
            this.nbrAppointmentNoteTitleMaxLength.Location = new System.Drawing.Point(110, 35);
            this.nbrAppointmentNoteTitleMaxLength.Name = "nbrAppointmentNoteTitleMaxLength";
            this.nbrAppointmentNoteTitleMaxLength.Size = new System.Drawing.Size(40, 20);
            this.nbrAppointmentNoteTitleMaxLength.TabIndex = 15;
            // 
            // txtAppointmentNoteTitleFormat
            // 
            this.txtAppointmentNoteTitleFormat.Location = new System.Drawing.Point(109, 9);
            this.txtAppointmentNoteTitleFormat.Name = "txtAppointmentNoteTitleFormat";
            this.txtAppointmentNoteTitleFormat.Size = new System.Drawing.Size(314, 20);
            this.txtAppointmentNoteTitleFormat.TabIndex = 14;
            this.toolTip1.SetToolTip(this.txtAppointmentNoteTitleFormat, "Use {Subject}, {Sender}, {Date}");
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(4, 39);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(89, 13);
            this.label14.TabIndex = 11;
            this.label14.Text = "Max Title Length:";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(4, 13);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(91, 13);
            this.label15.TabIndex = 10;
            this.label15.Text = "Note Title Format:\r\n";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(258, 45);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(95, 13);
            this.label17.TabIndex = 8;
            this.label17.Text = "Default Task Tags";
            this.toolTip1.SetToolTip(this.label17, "Comma Separated");
            // 
            // tabDeveloperSettings
            // 
            this.tabDeveloperSettings.Controls.Add(this.cbShowThreadDebug);
            this.tabDeveloperSettings.Controls.Add(this.cbShowDevelopmentSettings);
            this.tabDeveloperSettings.Location = new System.Drawing.Point(4, 22);
            this.tabDeveloperSettings.Name = "tabDeveloperSettings";
            this.tabDeveloperSettings.Size = new System.Drawing.Size(430, 163);
            this.tabDeveloperSettings.TabIndex = 2;
            this.tabDeveloperSettings.Text = "Developer Settings";
            this.tabDeveloperSettings.UseVisualStyleBackColor = true;
            // 
            // cbShowThreadDebug
            // 
            this.cbShowThreadDebug.AutoSize = true;
            this.cbShowThreadDebug.Location = new System.Drawing.Point(5, 27);
            this.cbShowThreadDebug.Name = "cbShowThreadDebug";
            this.cbShowThreadDebug.Size = new System.Drawing.Size(119, 17);
            this.cbShowThreadDebug.TabIndex = 1;
            this.cbShowThreadDebug.Text = "Show thread debug";
            this.cbShowThreadDebug.UseVisualStyleBackColor = true;
            // 
            // cbShowDevelopmentSettings
            // 
            this.cbShowDevelopmentSettings.AutoSize = true;
            this.cbShowDevelopmentSettings.Location = new System.Drawing.Point(5, 4);
            this.cbShowDevelopmentSettings.Name = "cbShowDevelopmentSettings";
            this.cbShowDevelopmentSettings.Size = new System.Drawing.Size(156, 17);
            this.cbShowDevelopmentSettings.TabIndex = 0;
            this.cbShowDevelopmentSettings.Text = "Show development settings";
            this.cbShowDevelopmentSettings.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(371, 555);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(290, 555);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // tabTaskSettings
            // 
            this.tabTaskSettings.Controls.Add(this.cbCreateObsidianTask);
            this.tabTaskSettings.Controls.Add(this.cbCreateOutlookTask);
            this.tabTaskSettings.Controls.Add(this.btnRemoveTaskTag);
            this.tabTaskSettings.Controls.Add(this.btnEditTaskTag);
            this.tabTaskSettings.Controls.Add(this.cbAskForDates);
            this.tabTaskSettings.Controls.Add(this.lstDefaultTaskTags);
            this.tabTaskSettings.Controls.Add(this.btnAddTaskTag);
            this.tabTaskSettings.Controls.Add(this.label17);
            this.tabTaskSettings.Location = new System.Drawing.Point(4, 22);
            this.tabTaskSettings.Name = "tabTaskSettings";
            this.tabTaskSettings.Size = new System.Drawing.Size(430, 163);
            this.tabTaskSettings.TabIndex = 3;
            this.tabTaskSettings.Text = "Task Settings";
            this.tabTaskSettings.UseVisualStyleBackColor = true;
            // 
            // lstDefaultNoteTags
            // 
            this.lstDefaultNoteTags.FormattingEnabled = true;
            this.lstDefaultNoteTags.Location = new System.Drawing.Point(177, 57);
            this.lstDefaultNoteTags.Name = "lstDefaultNoteTags";
            this.lstDefaultNoteTags.Size = new System.Drawing.Size(247, 69);
            this.lstDefaultNoteTags.TabIndex = 10;
            // 
            // btnRemoveNoteTag
            // 
            this.btnRemoveNoteTag.Location = new System.Drawing.Point(338, 132);
            this.btnRemoveNoteTag.Name = "btnRemoveNoteTag";
            this.btnRemoveNoteTag.Size = new System.Drawing.Size(63, 23);
            this.btnRemoveNoteTag.TabIndex = 6;
            this.btnRemoveNoteTag.Text = "Remove";
            this.btnRemoveNoteTag.UseVisualStyleBackColor = true;
            this.btnRemoveNoteTag.Click += new System.EventHandler(this.btnRemoveNoteTag_Click);
            // 
            // btnEditNoteTag
            // 
            this.btnEditNoteTag.Location = new System.Drawing.Point(269, 132);
            this.btnEditNoteTag.Name = "btnEditNoteTag";
            this.btnEditNoteTag.Size = new System.Drawing.Size(63, 23);
            this.btnEditNoteTag.TabIndex = 5;
            this.btnEditNoteTag.Text = "Edit";
            this.btnEditNoteTag.UseVisualStyleBackColor = true;
            this.btnEditNoteTag.Click += new System.EventHandler(this.btnEditNoteTag_Click);
            // 
            // btnAddNoteTag
            // 
            this.btnAddNoteTag.Location = new System.Drawing.Point(200, 132);
            this.btnAddNoteTag.Name = "btnAddNoteTag";
            this.btnAddNoteTag.Size = new System.Drawing.Size(63, 23);
            this.btnAddNoteTag.TabIndex = 4;
            this.btnAddNoteTag.Text = "Add";
            this.btnAddNoteTag.UseVisualStyleBackColor = true;
            this.btnAddNoteTag.Click += new System.EventHandler(this.btnAddNoteTag_Click);
            // 
            // btnRemoveAppointmentTag
            // 
            this.btnRemoveAppointmentTag.Location = new System.Drawing.Point(337, 136);
            this.btnRemoveAppointmentTag.Name = "btnRemoveAppointmentTag";
            this.btnRemoveAppointmentTag.Size = new System.Drawing.Size(63, 23);
            this.btnRemoveAppointmentTag.TabIndex = 19;
            this.btnRemoveAppointmentTag.Text = "Remove";
            this.btnRemoveAppointmentTag.UseVisualStyleBackColor = true;
            this.btnRemoveAppointmentTag.Click += new System.EventHandler(this.btnRemoveAppointmentTag_Click);
            // 
            // btnEditAppointmentTag
            // 
            this.btnEditAppointmentTag.Location = new System.Drawing.Point(268, 136);
            this.btnEditAppointmentTag.Name = "btnEditAppointmentTag";
            this.btnEditAppointmentTag.Size = new System.Drawing.Size(63, 23);
            this.btnEditAppointmentTag.TabIndex = 18;
            this.btnEditAppointmentTag.Text = "Edit";
            this.btnEditAppointmentTag.UseVisualStyleBackColor = true;
            this.btnEditAppointmentTag.Click += new System.EventHandler(this.btnEditAppointmentTag_Click);
            // 
            // lstDefaultAppointmentTags
            // 
            this.lstDefaultAppointmentTags.FormattingEnabled = true;
            this.lstDefaultAppointmentTags.Location = new System.Drawing.Point(176, 61);
            this.lstDefaultAppointmentTags.Name = "lstDefaultAppointmentTags";
            this.lstDefaultAppointmentTags.Size = new System.Drawing.Size(247, 69);
            this.lstDefaultAppointmentTags.TabIndex = 20;
            // 
            // btnAddAppointmentTag
            // 
            this.btnAddAppointmentTag.Location = new System.Drawing.Point(199, 136);
            this.btnAddAppointmentTag.Name = "btnAddAppointmentTag";
            this.btnAddAppointmentTag.Size = new System.Drawing.Size(63, 23);
            this.btnAddAppointmentTag.TabIndex = 17;
            this.btnAddAppointmentTag.Text = "Add";
            this.btnAddAppointmentTag.UseVisualStyleBackColor = true;
            this.btnAddAppointmentTag.Click += new System.EventHandler(this.btnAddAppointmentTag_Click);
            // 
            // btnRemoveTaskTag
            // 
            this.btnRemoveTaskTag.Location = new System.Drawing.Point(343, 136);
            this.btnRemoveTaskTag.Name = "btnRemoveTaskTag";
            this.btnRemoveTaskTag.Size = new System.Drawing.Size(63, 23);
            this.btnRemoveTaskTag.TabIndex = 23;
            this.btnRemoveTaskTag.Text = "Remove";
            this.btnRemoveTaskTag.UseVisualStyleBackColor = true;
            this.btnRemoveTaskTag.Click += new System.EventHandler(this.btnRemoveTaskTag_Click);
            // 
            // btnEditTaskTag
            // 
            this.btnEditTaskTag.Location = new System.Drawing.Point(274, 136);
            this.btnEditTaskTag.Name = "btnEditTaskTag";
            this.btnEditTaskTag.Size = new System.Drawing.Size(63, 23);
            this.btnEditTaskTag.TabIndex = 22;
            this.btnEditTaskTag.Text = "Edit";
            this.btnEditTaskTag.UseVisualStyleBackColor = true;
            this.btnEditTaskTag.Click += new System.EventHandler(this.btnEditTaskTag_Click);
            // 
            // lstDefaultTaskTags
            // 
            this.lstDefaultTaskTags.FormattingEnabled = true;
            this.lstDefaultTaskTags.Location = new System.Drawing.Point(187, 61);
            this.lstDefaultTaskTags.Name = "lstDefaultTaskTags";
            this.lstDefaultTaskTags.Size = new System.Drawing.Size(237, 69);
            this.lstDefaultTaskTags.TabIndex = 24;
            // 
            // btnAddTaskTag
            // 
            this.btnAddTaskTag.Location = new System.Drawing.Point(205, 136);
            this.btnAddTaskTag.Name = "btnAddTaskTag";
            this.btnAddTaskTag.Size = new System.Drawing.Size(63, 23);
            this.btnAddTaskTag.TabIndex = 21;
            this.btnAddTaskTag.Text = "Add";
            this.btnAddTaskTag.UseVisualStyleBackColor = true;
            this.btnAddTaskTag.Click += new System.EventHandler(this.btnAddTaskTag_Click);
            // 
            // Settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 585);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabGrp);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpVaultSettings);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Settings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Sling Settings";
            this.grpVaultSettings.ResumeLayout(false);
            this.grpVaultSettings.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultReminderHour)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrObsidianDelaySeconds)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultReminderDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbrDefaultDueDays)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.tabGrp.ResumeLayout(false);
            this.tabMailSettings.ResumeLayout(false);
            this.tabMailSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrNoteTitleMaxLength)).EndInit();
            this.tabAppointmentSettings.ResumeLayout(false);
            this.tabAppointmentSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nbrAppointmentNoteTitleMaxLength)).EndInit();
            this.tabDeveloperSettings.ResumeLayout(false);
            this.tabDeveloperSettings.PerformLayout();
            this.tabTaskSettings.ResumeLayout(false);
            this.tabTaskSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpVaultSettings;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtAppointmentsFolder;
        private System.Windows.Forms.TextBox txtContactsFolder;
        private System.Windows.Forms.TextBox txtInboxFolder;
        private System.Windows.Forms.TextBox txtVaultBasePath;
        private System.Windows.Forms.TextBox txtVaultName;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnBrowseVaultPath;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown nbrObsidianDelaySeconds;
        private System.Windows.Forms.NumericUpDown nbrDefaultReminderDays;
        private System.Windows.Forms.NumericUpDown nbrDefaultDueDays;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.NumericUpDown nbrDefaultReminderHour;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox cbEnableContactSaving;
        private System.Windows.Forms.CheckBox cbLaunchObsidian;
        private System.Windows.Forms.CheckBox cbCreateObsidianTask;
        private System.Windows.Forms.CheckBox cbAskForDates;
        private System.Windows.Forms.CheckBox cbSearchEntireVaultForContacts;
        private System.Windows.Forms.CheckBox cbShowCountdown;
        private System.Windows.Forms.CheckBox cbCreateOutlookTask;
        private System.Windows.Forms.CheckBox cbGroupEmailThreads;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ListBox lstSubjectCleanupPatterns;
        private System.Windows.Forms.Button btnAddSubjectCleanupPattern;
        private System.Windows.Forms.Button btnRemoveSubjectCleanupPattern;
        private System.Windows.Forms.Button btnEditSubjectCleanupPattern;
        private System.Windows.Forms.TabControl tabGrp;
        private System.Windows.Forms.TabPage tabMailSettings;
        private System.Windows.Forms.TabPage tabAppointmentSettings;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtNoteTitleFormat;
        private System.Windows.Forms.NumericUpDown nbrNoteTitleMaxLength;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox cbNoteTitleIncludeDate;
        private System.Windows.Forms.CheckBox cbMoveDateToFrontInThread;
        private System.Windows.Forms.CheckBox cbAppointmentSaveAttachments;
        private System.Windows.Forms.NumericUpDown nbrAppointmentNoteTitleMaxLength;
        private System.Windows.Forms.TextBox txtAppointmentNoteTitleFormat;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TabPage tabDeveloperSettings;
        private System.Windows.Forms.CheckBox cbShowDevelopmentSettings;
        private System.Windows.Forms.CheckBox cbShowThreadDebug;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TabPage tabTaskSettings;
        private System.Windows.Forms.Button btnRemoveNoteTag;
        private System.Windows.Forms.Button btnEditNoteTag;
        private System.Windows.Forms.ListBox lstDefaultNoteTags;
        private System.Windows.Forms.Button btnAddNoteTag;
        private System.Windows.Forms.Button btnRemoveAppointmentTag;
        private System.Windows.Forms.Button btnEditAppointmentTag;
        private System.Windows.Forms.ListBox lstDefaultAppointmentTags;
        private System.Windows.Forms.Button btnAddAppointmentTag;
        private System.Windows.Forms.Button btnRemoveTaskTag;
        private System.Windows.Forms.Button btnEditTaskTag;
        private System.Windows.Forms.ListBox lstDefaultTaskTags;
        private System.Windows.Forms.Button btnAddTaskTag;
    }
}