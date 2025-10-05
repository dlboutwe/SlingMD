using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

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
        private GroupBox grpDevelopment;
        private CheckBox chkShowDevelopmentSettings;
        private CheckBox chkShowThreadDebug;
        // New controls for note/tag customization
        private TextBox txtDefaultNoteTags;
        private TextBox txtDefaultTaskTags;
        private TextBox txtNoteTitleFormat;
        private NumericUpDown numNoteTitleMaxLength;
        private CheckBox chkNoteTitleIncludeDate;
        private Label lblDefaultNoteTags;
        private Label lblDefaultTaskTags;
        private Label lblNoteTitleFormat;
        private Label lblNoteTitleMaxLength;
        private Label lblNoteTitleIncludeDate;
        private GroupBox grpNoteCustomization;
        private ToolTip toolTip;
        // Refactored layout containers
        private TableLayoutPanel mainLayout;
        private System.ComponentModel.IContainer components;
        private GroupBox grpVault;
        private TableLayoutPanel vaultLayout;
        private GroupBox grpGeneral;
        private TableLayoutPanel generalLayout;
        private GroupBox grpTiming;
        private TableLayoutPanel timingLayout;
        private Label lblDueDays;
        private Label lblReminderDays;
        private Label lblReminderHour;
        private GroupBox grpPatterns;
        private TableLayoutPanel patternsLayout;
        private FlowLayoutPanel btnPanel;
        private TableLayoutPanel noteTagLayout;
        private FlowLayoutPanel dateOptionsLayout;
        private FlowLayoutPanel devLayout;
        private FlowLayoutPanel btnLayout;
        private CheckBox chkMoveDateToFrontInThread;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.mainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.grpVault = new System.Windows.Forms.GroupBox();
            this.vaultLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblVaultName = new System.Windows.Forms.Label();
            this.txtVaultName = new System.Windows.Forms.TextBox();
            this.lblVaultPath = new System.Windows.Forms.Label();
            this.txtVaultPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblInboxFolder = new System.Windows.Forms.Label();
            this.txtInboxFolder = new System.Windows.Forms.TextBox();
            this.lblContactsFolder = new System.Windows.Forms.Label();
            this.txtContactsFolder = new System.Windows.Forms.TextBox();
            this.grpGeneral = new System.Windows.Forms.GroupBox();
            this.generalLayout = new System.Windows.Forms.TableLayoutPanel();
            this.chkEnableContactSaving = new System.Windows.Forms.CheckBox();
            this.chkSearchEntireVaultForContacts = new System.Windows.Forms.CheckBox();
            this.chkLaunchObsidian = new System.Windows.Forms.CheckBox();
            this.chkShowCountdown = new System.Windows.Forms.CheckBox();
            this.chkCreateObsidianTask = new System.Windows.Forms.CheckBox();
            this.chkCreateOutlookTask = new System.Windows.Forms.CheckBox();
            this.chkAskForDates = new System.Windows.Forms.CheckBox();
            this.chkGroupEmailThreads = new System.Windows.Forms.CheckBox();
            this.grpTiming = new System.Windows.Forms.GroupBox();
            this.timingLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblDelay = new System.Windows.Forms.Label();
            this.numDelay = new System.Windows.Forms.NumericUpDown();
            this.lblDueDays = new System.Windows.Forms.Label();
            this.numDefaultDueDays = new System.Windows.Forms.NumericUpDown();
            this.lblReminderDays = new System.Windows.Forms.Label();
            this.numDefaultReminderDays = new System.Windows.Forms.NumericUpDown();
            this.lblReminderHour = new System.Windows.Forms.Label();
            this.numDefaultReminderHour = new System.Windows.Forms.NumericUpDown();
            this.grpPatterns = new System.Windows.Forms.GroupBox();
            this.patternsLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lstPatterns = new System.Windows.Forms.ListBox();
            this.btnPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.grpNoteCustomization = new System.Windows.Forms.GroupBox();
            this.noteTagLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblDefaultNoteTags = new System.Windows.Forms.Label();
            this.txtDefaultNoteTags = new System.Windows.Forms.TextBox();
            this.lblDefaultTaskTags = new System.Windows.Forms.Label();
            this.txtDefaultTaskTags = new System.Windows.Forms.TextBox();
            this.lblNoteTitleFormat = new System.Windows.Forms.Label();
            this.txtNoteTitleFormat = new System.Windows.Forms.TextBox();
            this.lblNoteTitleMaxLength = new System.Windows.Forms.Label();
            this.numNoteTitleMaxLength = new System.Windows.Forms.NumericUpDown();
            this.dateOptionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.lblNoteTitleIncludeDate = new System.Windows.Forms.Label();
            this.chkNoteTitleIncludeDate = new System.Windows.Forms.CheckBox();
            this.chkMoveDateToFrontInThread = new System.Windows.Forms.CheckBox();
            this.grpDevelopment = new System.Windows.Forms.GroupBox();
            this.devLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.chkShowDevelopmentSettings = new System.Windows.Forms.CheckBox();
            this.chkShowThreadDebug = new System.Windows.Forms.CheckBox();
            this.btnLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.mainLayout.SuspendLayout();
            this.grpVault.SuspendLayout();
            this.vaultLayout.SuspendLayout();
            this.grpGeneral.SuspendLayout();
            this.generalLayout.SuspendLayout();
            this.grpTiming.SuspendLayout();
            this.timingLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultDueDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderHour)).BeginInit();
            this.grpPatterns.SuspendLayout();
            this.patternsLayout.SuspendLayout();
            this.btnPanel.SuspendLayout();
            this.grpNoteCustomization.SuspendLayout();
            this.noteTagLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numNoteTitleMaxLength)).BeginInit();
            this.dateOptionsLayout.SuspendLayout();
            this.grpDevelopment.SuspendLayout();
            this.devLayout.SuspendLayout();
            this.btnLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainLayout
            // 
            this.mainLayout.AutoScroll = true;
            this.mainLayout.AutoSize = true;
            this.mainLayout.ColumnCount = 1;
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.mainLayout.Controls.Add(this.grpVault);
            this.mainLayout.Controls.Add(this.grpGeneral);
            this.mainLayout.Controls.Add(this.grpTiming);
            this.mainLayout.Controls.Add(this.grpPatterns);
            this.mainLayout.Controls.Add(this.grpNoteCustomization);
            this.mainLayout.Controls.Add(this.grpDevelopment);
            this.mainLayout.Controls.Add(this.btnLayout);
            this.mainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainLayout.Location = new System.Drawing.Point(0, 0);
            this.mainLayout.Name = "mainLayout";
            this.mainLayout.RowCount = 1;
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //this.mainLayout.Size = new System.Drawing.Size(1175, 840);
            this.mainLayout.TabIndex = 0;
            // 
            // grpVault
            // 
            this.grpVault.AutoSize = true;
            this.grpVault.Controls.Add(this.vaultLayout);
            this.grpVault.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpVault.Location = new System.Drawing.Point(3, 3);
            this.grpVault.Name = "grpVault";
            //this.grpVault.Size = new System.Drawing.Size(1156, 173);
            this.grpVault.TabIndex = 0;
            this.grpVault.TabStop = false;
            this.grpVault.Text = "Vault Settings";
            // 
            // vaultLayout
            // 
            this.vaultLayout.AutoSize = true;
            this.vaultLayout.ColumnCount = 2;
            this.vaultLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.vaultLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.vaultLayout.Controls.Add(this.lblVaultName, 0, 0);
            this.vaultLayout.Controls.Add(this.txtVaultName, 1, 0);
            this.vaultLayout.Controls.Add(this.lblVaultPath, 0, 1);
            this.vaultLayout.Controls.Add(this.txtVaultPath, 1, 1);
            this.vaultLayout.Controls.Add(this.btnBrowse, 1, 2);
            this.vaultLayout.Controls.Add(this.lblInboxFolder, 0, 3);
            this.vaultLayout.Controls.Add(this.txtInboxFolder, 1, 3);
            this.vaultLayout.Controls.Add(this.lblContactsFolder, 0, 4);
            this.vaultLayout.Controls.Add(this.txtContactsFolder, 1, 4);
            this.vaultLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.vaultLayout.Location = new System.Drawing.Point(3, 22);
            this.vaultLayout.Name = "vaultLayout";
            this.vaultLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.vaultLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.vaultLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.vaultLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.vaultLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //this.vaultLayout.Size = new System.Drawing.Size(1150, 148);
            this.vaultLayout.TabIndex = 0;
            // 
            // lblVaultName
            // 
            this.lblVaultName.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblVaultName.Location = new System.Drawing.Point(3, 6);
            this.lblVaultName.Name = "lblVaultName";
            //this.lblVaultName.Size = new System.Drawing.Size(100, 20);
            this.lblVaultName.AutoSize = true;
            this.lblVaultName.TabIndex = 0;
            this.lblVaultName.Text = "Vault Name:";
            // 
            // txtVaultName
            // 
            this.txtVaultName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtVaultName.Location = new System.Drawing.Point(348, 3);
            this.txtVaultName.Name = "txtVaultName";
            //this.txtVaultName.Size = new System.Drawing.Size(799, 26);
            this.txtVaultName.TabIndex = 1;
            // 
            // lblVaultPath
            // 
            this.lblVaultPath.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblVaultPath.Location = new System.Drawing.Point(3, 38);
            this.lblVaultPath.Name = "lblVaultPath";
            //this.lblVaultPath.Size = new System.Drawing.Size(100, 20);
            this.lblVaultPath .AutoSize = true;
            this.lblVaultPath.TabIndex = 2;
            this.lblVaultPath.Text = "Vault Base Path:";
            // 
            // txtVaultPath
            // 
            this.txtVaultPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtVaultPath.Location = new System.Drawing.Point(348, 35);
            this.txtVaultPath.Name = "txtVaultPath";
            //this.txtVaultPath.Size = new System.Drawing.Size(799, 26);
            this.txtVaultName .AutoSize = true;
            this.txtVaultPath.TabIndex = 3;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.btnBrowse.Location = new System.Drawing.Point(348, 67);
            this.btnBrowse.Name = "btnBrowse";
            //this.btnBrowse.Size = new System.Drawing.Size(75, 14);
            this.btnBrowse.AutoSize = true;
            this.btnBrowse.TabIndex = 4;
            this.btnBrowse.Text = "Browse...";
            // 
            // lblInboxFolder
            // 
            this.lblInboxFolder.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblInboxFolder.Location = new System.Drawing.Point(3, 90);
            this.lblInboxFolder.Name = "lblInboxFolder";
            //this.lblInboxFolder.Size = new System.Drawing.Size(100, 20);
            this.lblInboxFolder.TabIndex = 5;
            this.lblInboxFolder.Text = "Inbox Folder:";
            // 
            // txtInboxFolder
            // 
            this.txtInboxFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtInboxFolder.Location = new System.Drawing.Point(348, 87);
            this.txtInboxFolder.Name = "txtInboxFolder";
            //this.txtInboxFolder.Size = new System.Drawing.Size(799, 26);
            this.txtInboxFolder.TabIndex = 6;
            // 
            // lblContactsFolder
            // 
            this.lblContactsFolder.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblContactsFolder.Location = new System.Drawing.Point(3, 122);
            this.lblContactsFolder.Name = "lblContactsFolder";
            //this.lblContactsFolder.Size = new System.Drawing.Size(100, 20);
            this.lblContactsFolder.TabIndex = 7;
            this.lblContactsFolder.Text = "Contacts Folder:";
            // 
            // txtContactsFolder
            // 
            this.txtContactsFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtContactsFolder.Location = new System.Drawing.Point(348, 119);
            this.txtContactsFolder.Name = "txtContactsFolder";
            //this.txtContactsFolder.Size = new System.Drawing.Size(799, 26);
            
            this.txtContactsFolder.TabIndex = 8;
            // 
            // grpGeneral
            // 
            this.grpGeneral.AutoSize = true;
            this.grpGeneral.Controls.Add(this.generalLayout);
            this.grpGeneral.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpGeneral.Location = new System.Drawing.Point(3, 182);
            this.grpGeneral.Name = "grpGeneral";
            //this.grpGeneral.Size = new System.Drawing.Size(1156, 145);
            this.grpGeneral.TabIndex = 1;
            this.grpGeneral.TabStop = false;
            this.grpGeneral.Text = "General Settings";
            // 
            // generalLayout
            // 
            this.generalLayout.AutoSize = true;
            this.generalLayout.ColumnCount = 2;
            this.generalLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.generalLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.generalLayout.Controls.Add(this.chkEnableContactSaving, 0, 0);
            this.generalLayout.Controls.Add(this.chkSearchEntireVaultForContacts, 1, 0);
            this.generalLayout.Controls.Add(this.chkLaunchObsidian, 0, 1);
            this.generalLayout.Controls.Add(this.chkShowCountdown, 1, 1);
            this.generalLayout.Controls.Add(this.chkCreateObsidianTask, 0, 2);
            this.generalLayout.Controls.Add(this.chkCreateOutlookTask, 1, 2);
            this.generalLayout.Controls.Add(this.chkAskForDates, 0, 3);
            this.generalLayout.Controls.Add(this.chkGroupEmailThreads, 1, 3);
            this.generalLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.generalLayout.Location = new System.Drawing.Point(3, 22);
            this.generalLayout.Name = "generalLayout";
            this.generalLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.generalLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.generalLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.generalLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //this.generalLayout.Size = new System.Drawing.Size(1150, 120);
            this.generalLayout.TabIndex = 0;
            // 
            // chkEnableContactSaving
            // 
            this.chkEnableContactSaving.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkEnableContactSaving.AutoSize = true;
            this.chkEnableContactSaving.Location = new System.Drawing.Point(3, 3);
            this.chkEnableContactSaving.Name = "chkEnableContactSaving";
            //this.chkEnableContactSaving.Size = new System.Drawing.Size(569, 24);
            this.chkEnableContactSaving.TabIndex = 0;
            this.chkEnableContactSaving.Text = "Enable Contact Saving";
            // 
            // chkSearchEntireVaultForContacts
            // 
            this.chkSearchEntireVaultForContacts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkSearchEntireVaultForContacts.AutoSize = true;
            this.chkSearchEntireVaultForContacts.Location = new System.Drawing.Point(578, 3);
            this.chkSearchEntireVaultForContacts.Name = "chkSearchEntireVaultForContacts";
            //this.chkSearchEntireVaultForContacts.Size = new System.Drawing.Size(569, 24);
            this.chkSearchEntireVaultForContacts.TabIndex = 1;
            this.chkSearchEntireVaultForContacts.Text = "Search entire vault for contacts";
            // 
            // chkLaunchObsidian
            // 
            this.chkLaunchObsidian.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkLaunchObsidian.AutoSize = true;
            this.chkLaunchObsidian.Location = new System.Drawing.Point(3, 33);
            this.chkLaunchObsidian.Name = "chkLaunchObsidian";
            //this.chkLaunchObsidian.Size = new System.Drawing.Size(569, 24);
            this.chkLaunchObsidian.TabIndex = 2;
            this.chkLaunchObsidian.Text = "Launch Obsidian after saving";
            // 
            // chkShowCountdown
            // 
            this.chkShowCountdown.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkShowCountdown.AutoSize = true;
            this.chkShowCountdown.Location = new System.Drawing.Point(578, 33);
            this.chkShowCountdown.Name = "chkShowCountdown";
            //this.chkShowCountdown.Size = new System.Drawing.Size(569, 24);
            this.chkShowCountdown.TabIndex = 3;
            this.chkShowCountdown.Text = "Show countdown";
            // 
            // chkCreateObsidianTask
            // 
            this.chkCreateObsidianTask.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkCreateObsidianTask.AutoSize = true;
            this.chkCreateObsidianTask.Location = new System.Drawing.Point(3, 63);
            this.chkCreateObsidianTask.Name = "chkCreateObsidianTask";
            //this.chkCreateObsidianTask.Size = new System.Drawing.Size(569, 24);
            this.chkCreateObsidianTask.TabIndex = 4;
            this.chkCreateObsidianTask.Text = "Create task in Obsidian note";
            // 
            // chkCreateOutlookTask
            // 
            this.chkCreateOutlookTask.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkCreateOutlookTask.AutoSize = true;
            this.chkCreateOutlookTask.Location = new System.Drawing.Point(578, 63);
            this.chkCreateOutlookTask.Name = "chkCreateOutlookTask";
            //this.chkCreateOutlookTask.Size = new System.Drawing.Size(569, 24);
            this.chkCreateOutlookTask.TabIndex = 5;
            this.chkCreateOutlookTask.Text = "Create task in Outlook";
            // 
            // chkAskForDates
            // 
            this.chkAskForDates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkAskForDates.AutoSize = true;
            this.chkAskForDates.Location = new System.Drawing.Point(3, 93);
            this.chkAskForDates.Name = "chkAskForDates";
            //this.chkAskForDates.Size = new System.Drawing.Size(569, 24);
            this.chkAskForDates.TabIndex = 6;
            this.chkAskForDates.Text = "Ask for dates and times each time";
            // 
            // chkGroupEmailThreads
            // 
            this.chkGroupEmailThreads.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkGroupEmailThreads.AutoSize = true;
            this.chkGroupEmailThreads.Location = new System.Drawing.Point(578, 93);
            this.chkGroupEmailThreads.Name = "chkGroupEmailThreads";
            //this.chkGroupEmailThreads.Size = new System.Drawing.Size(569, 24);
            this.chkGroupEmailThreads.TabIndex = 7;
            this.chkGroupEmailThreads.Text = "Group email threads";
            // 
            // grpTiming
            // 
            this.grpTiming.AutoSize = true;
            this.grpTiming.Controls.Add(this.timingLayout);
            this.grpTiming.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpTiming.Location = new System.Drawing.Point(3, 333);
            this.grpTiming.Name = "grpTiming";
            //this.grpTiming.Size = new System.Drawing.Size(1156, 153);
            this.grpTiming.TabIndex = 2;
            this.grpTiming.TabStop = false;
            this.grpTiming.Text = "Timing Settings";
            // 
            // timingLayout
            // 
            this.timingLayout.AutoSize = true;
            this.timingLayout.ColumnCount = 2;
            this.timingLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.timingLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.timingLayout.Controls.Add(this.lblDelay, 0, 0);
            this.timingLayout.Controls.Add(this.numDelay, 1, 0);
            this.timingLayout.Controls.Add(this.lblDueDays, 0, 1);
            this.timingLayout.Controls.Add(this.numDefaultDueDays, 1, 1);
            this.timingLayout.Controls.Add(this.lblReminderDays, 0, 2);
            this.timingLayout.Controls.Add(this.numDefaultReminderDays, 1, 2);
            this.timingLayout.Controls.Add(this.lblReminderHour, 0, 3);
            this.timingLayout.Controls.Add(this.numDefaultReminderHour, 1, 3);
            this.timingLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.timingLayout.Location = new System.Drawing.Point(3, 22);
            this.timingLayout.Name = "timingLayout";
            this.timingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.timingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.timingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.timingLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //this.timingLayout.Size = new System.Drawing.Size(1150, 128);
            this.timingLayout.TabIndex = 0;
            // 
            // lblDelay
            // 
            this.lblDelay.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblDelay.Location = new System.Drawing.Point(3, 6);
            this.lblDelay.Name = "lblDelay";
            //this.lblDelay.Size = new System.Drawing.Size(14, 20);
            this.lblDelay.TabIndex = 0;
            this.lblDelay.Text = "Delay (seconds):";
            // 
            // numDelay
            // 
            this.numDelay.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.numDelay.Location = new System.Drawing.Point(23, 3);
            this.numDelay.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numDelay.Name = "numDelay";
            //this.numDelay.Size = new System.Drawing.Size(120, 26);
            this.numDelay.TabIndex = 1;
            // 
            // lblDueDays
            // 
            this.lblDueDays.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblDueDays.Location = new System.Drawing.Point(3, 38);
            this.lblDueDays.Name = "lblDueDays";
            //this.lblDueDays.Size = new System.Drawing.Size(14, 20);
            this.lblDueDays.TabIndex = 2;
            this.lblDueDays.Text = "Due in Days:";
            // 
            // numDefaultDueDays
            // 
            this.numDefaultDueDays.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.numDefaultDueDays.Location = new System.Drawing.Point(23, 35);
            this.numDefaultDueDays.Maximum = new decimal(new int[] {
            30,
            0,
            0,
            0});
            this.numDefaultDueDays.Name = "numDefaultDueDays";
            //this.numDefaultDueDays.Size = new System.Drawing.Size(120, 26);
            this.numDefaultDueDays.TabIndex = 3;
            // 
            // lblReminderDays
            // 
            this.lblReminderDays.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblReminderDays.Location = new System.Drawing.Point(3, 70);
            this.lblReminderDays.Name = "lblReminderDays";
            //this.lblReminderDays.Size = new System.Drawing.Size(14, 20);
            this.lblReminderDays.TabIndex = 4;
            this.lblReminderDays.Text = "Reminder Days:";
            // 
            // numDefaultReminderDays
            // 
            this.numDefaultReminderDays.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.numDefaultReminderDays.Location = new System.Drawing.Point(23, 67);
            this.numDefaultReminderDays.Maximum = new decimal(new int[] {
            30,
            0,
            0,
            0});
            this.numDefaultReminderDays.Name = "numDefaultReminderDays";
            //this.numDefaultReminderDays.Size = new System.Drawing.Size(120, 26);
            this.numDefaultReminderDays.TabIndex = 5;
            // 
            // lblReminderHour
            // 
            this.lblReminderHour.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblReminderHour.Location = new System.Drawing.Point(3, 102);
            this.lblReminderHour.Name = "lblReminderHour";
            //this.lblReminderHour.Size = new System.Drawing.Size(14, 20);
            this.lblReminderHour.TabIndex = 6;
            this.lblReminderHour.Text = "Reminder Hour:";
            // 
            // numDefaultReminderHour
            // 
            this.numDefaultReminderHour.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.numDefaultReminderHour.Location = new System.Drawing.Point(23, 99);
            this.numDefaultReminderHour.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.numDefaultReminderHour.Name = "numDefaultReminderHour";
            //this.numDefaultReminderHour.Size = new System.Drawing.Size(120, 26);
            this.numDefaultReminderHour.TabIndex = 7;
            // 
            // grpPatterns
            // 
            this.grpPatterns.AutoSize = true;
            this.grpPatterns.Controls.Add(this.patternsLayout);
            this.grpPatterns.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpPatterns.Location = new System.Drawing.Point(3, 492);
            this.grpPatterns.Name = "grpPatterns";
            //this.grpPatterns.Size = new System.Drawing.Size(1156, 118);
            this.grpPatterns.TabIndex = 3;
            this.grpPatterns.TabStop = false;
            this.grpPatterns.Text = "Subject Cleanup Patterns";
            // 
            // patternsLayout
            // 
            this.patternsLayout.AutoSize = true;
            this.patternsLayout.ColumnCount = 2;
            this.patternsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.patternsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.patternsLayout.Controls.Add(this.lstPatterns, 0, 0);
            this.patternsLayout.Controls.Add(this.btnPanel, 1, 0);
            this.patternsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.patternsLayout.Location = new System.Drawing.Point(3, 22);
            this.patternsLayout.Name = "patternsLayout";
            this.patternsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //this.patternsLayout.Size = new System.Drawing.Size(1150, 93);
            this.patternsLayout.TabIndex = 0;
            // 
            // lstPatterns
            // 
            this.lstPatterns.ItemHeight = 20;
            this.lstPatterns.Location = new System.Drawing.Point(3, 3);
            this.lstPatterns.Name = "lstPatterns";
            //this.lstPatterns.Size = new System.Drawing.Size(14, 4);
            this.lstPatterns.TabIndex = 0;
            // 
            // btnPanel
            // 
            this.btnPanel.AutoSize = true;
            this.btnPanel.Controls.Add(this.btnAdd);
            this.btnPanel.Controls.Add(this.btnEdit);
            this.btnPanel.Controls.Add(this.btnRemove);
            this.btnPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.btnPanel.Location = new System.Drawing.Point(23, 3);
            this.btnPanel.Name = "btnPanel";
            //this.btnPanel.Size = new System.Drawing.Size(81, 87);
            this.btnPanel.TabIndex = 1;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(3, 3);
            this.btnAdd.Name = "btnAdd";
            //this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 0;
            this.btnAdd.Text = "Add";
            this.btnAdd.AutoSize = true;
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(3, 32);
            this.btnEdit.Name = "btnEdit";
            //this.btnEdit.Size = new System.Drawing.Size(75, 23);
            this.btnEdit.TabIndex = 1;
            this.btnEdit.Text = "Edit";
            this.btnEdit.AutoSize = true;
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(3, 61);
            this.btnRemove.Name = "btnRemove";
            //this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.AutoSize = true;
            this.btnRemove.TabIndex = 2;
            this.btnRemove.Text = "Remove";
            // 
            // grpNoteCustomization
            // 
            this.grpNoteCustomization.AutoSize = true;
            this.grpNoteCustomization.Controls.Add(this.noteTagLayout);
            this.grpNoteCustomization.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpNoteCustomization.Location = new System.Drawing.Point(3, 616);
            this.grpNoteCustomization.Name = "grpNoteCustomization";
            //this.grpNoteCustomization.Size = new System.Drawing.Size(1156, 189);
            this.grpNoteCustomization.TabIndex = 4;
            this.grpNoteCustomization.TabStop = false;
            this.grpNoteCustomization.Text = "Note & Tag Customization";
            // 
            // noteTagLayout
            // 
            this.noteTagLayout.AutoSize = true;
            this.noteTagLayout.ColumnCount = 2;
            this.noteTagLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.noteTagLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.noteTagLayout.Controls.Add(this.lblDefaultNoteTags, 0, 0);
            this.noteTagLayout.Controls.Add(this.txtDefaultNoteTags, 1, 0);
            this.noteTagLayout.Controls.Add(this.lblDefaultTaskTags, 0, 1);
            this.noteTagLayout.Controls.Add(this.txtDefaultTaskTags, 1, 1);
            this.noteTagLayout.Controls.Add(this.lblNoteTitleFormat, 0, 2);
            this.noteTagLayout.Controls.Add(this.txtNoteTitleFormat, 1, 2);
            this.noteTagLayout.Controls.Add(this.lblNoteTitleMaxLength, 0, 3);
            this.noteTagLayout.Controls.Add(this.numNoteTitleMaxLength, 1, 3);
            this.noteTagLayout.Controls.Add(this.dateOptionsLayout, 1, 4);
            this.noteTagLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.noteTagLayout.Location = new System.Drawing.Point(3, 22);
            this.noteTagLayout.Name = "noteTagLayout";
            this.noteTagLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.noteTagLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.noteTagLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.noteTagLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.noteTagLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            //.noteTagLayout.Size = new System.Drawing.Size(1150, 164);
            this.noteTagLayout.TabIndex = 0;
            // 
            // lblDefaultNoteTags
            // 
            this.lblDefaultNoteTags.AutoEllipsis = true;
            this.lblDefaultNoteTags.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblDefaultNoteTags.Location = new System.Drawing.Point(3, 0);
            this.lblDefaultNoteTags.Name = "lblDefaultNoteTags";
            //this.lblDefaultNoteTags.Size = new System.Drawing.Size(396, 32);
            this.lblDefaultNoteTags.TabIndex = 0;
            this.lblDefaultNoteTags.Text = "Default Note Tags (comma-separated):";
            this.lblDefaultNoteTags.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtDefaultNoteTags
            // 
            this.txtDefaultNoteTags.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtDefaultNoteTags.Location = new System.Drawing.Point(405, 3);
            this.txtDefaultNoteTags.Name = "txtDefaultNoteTags";
            //this.txtDefaultNoteTags.Size = new System.Drawing.Size(742, 26);
            this.txtDefaultNoteTags.TabIndex = 1;
            // 
            // lblDefaultTaskTags
            // 
            this.lblDefaultTaskTags.AutoEllipsis = true;
            this.lblDefaultTaskTags.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblDefaultTaskTags.Location = new System.Drawing.Point(3, 32);
            this.lblDefaultTaskTags.Name = "lblDefaultTaskTags";
            //this.lblDefaultTaskTags.Size = new System.Drawing.Size(396, 32);
            this.lblDefaultTaskTags.TabIndex = 2;
            this.lblDefaultTaskTags.Text = "Default Task Tags (comma-separated, will be rendered as #tags):";
            this.lblDefaultTaskTags.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtDefaultTaskTags
            // 
            this.txtDefaultTaskTags.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtDefaultTaskTags.Location = new System.Drawing.Point(405, 35);
            this.txtDefaultTaskTags.Name = "txtDefaultTaskTags";
            //this.txtDefaultTaskTags.Size = new System.Drawing.Size(742, 26);
            this.txtDefaultTaskTags.TabIndex = 3;
            // 
            // lblNoteTitleFormat
            // 
            this.lblNoteTitleFormat.AutoEllipsis = true;
            this.lblNoteTitleFormat.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblNoteTitleFormat.Location = new System.Drawing.Point(3, 64);
            this.lblNoteTitleFormat.Name = "lblNoteTitleFormat";
            //this.lblNoteTitleFormat.Size = new System.Drawing.Size(396, 32);
            this.lblNoteTitleFormat.TabIndex = 4;
            this.lblNoteTitleFormat.Text = "Note Title Format (use {Subject}, {Sender}, {Date}):";
            this.lblNoteTitleFormat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtNoteTitleFormat
            // 
            this.txtNoteTitleFormat.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtNoteTitleFormat.Location = new System.Drawing.Point(405, 67);
            this.txtNoteTitleFormat.Name = "txtNoteTitleFormat";
            //this.txtNoteTitleFormat.Size = new System.Drawing.Size(742, 26);
            this.txtNoteTitleFormat.TabIndex = 5;
            // 
            // lblNoteTitleMaxLength
            // 
            this.lblNoteTitleMaxLength.AutoEllipsis = true;
            this.lblNoteTitleMaxLength.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblNoteTitleMaxLength.Location = new System.Drawing.Point(3, 96);
            this.lblNoteTitleMaxLength.Name = "lblNoteTitleMaxLength";
           // this.lblNoteTitleMaxLength.Size = new System.Drawing.Size(396, 32);
            this.lblNoteTitleMaxLength.TabIndex = 6;
            this.lblNoteTitleMaxLength.Text = "Max Title Length:";
            this.lblNoteTitleMaxLength.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // numNoteTitleMaxLength
            // 
            this.numNoteTitleMaxLength.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numNoteTitleMaxLength.Location = new System.Drawing.Point(405, 99);
            this.numNoteTitleMaxLength.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.numNoteTitleMaxLength.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numNoteTitleMaxLength.Name = "numNoteTitleMaxLength";
            //this.numNoteTitleMaxLength.Size = new System.Drawing.Size(742, 26);
            this.numNoteTitleMaxLength.TabIndex = 7;
            this.numNoteTitleMaxLength.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // dateOptionsLayout
            // 
            this.dateOptionsLayout.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.dateOptionsLayout.AutoSize = true;
            this.dateOptionsLayout.Controls.Add(this.lblNoteTitleIncludeDate);
            this.dateOptionsLayout.Controls.Add(this.chkNoteTitleIncludeDate);
            this.dateOptionsLayout.Controls.Add(this.chkMoveDateToFrontInThread);
            this.dateOptionsLayout.Location = new System.Drawing.Point(405, 131);
            this.dateOptionsLayout.Name = "dateOptionsLayout";
            //this.dateOptionsLayout.Size = new System.Drawing.Size(687, 30);
            this.dateOptionsLayout.TabIndex = 8;
            // 
            // lblNoteTitleIncludeDate
            // 
            this.lblNoteTitleIncludeDate.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblNoteTitleIncludeDate.AutoSize = true;
            this.lblNoteTitleIncludeDate.Location = new System.Drawing.Point(3, 5);
            this.lblNoteTitleIncludeDate.Name = "lblNoteTitleIncludeDate";
            //this.lblNoteTitleIncludeDate.Size = new System.Drawing.Size(153, 20);
            this.lblNoteTitleIncludeDate.TabIndex = 0;
            this.lblNoteTitleIncludeDate.Text = "Include Date in Title:";
            this.lblNoteTitleIncludeDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkNoteTitleIncludeDate
            // 
            this.chkNoteTitleIncludeDate.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.chkNoteTitleIncludeDate.Location = new System.Drawing.Point(162, 3);
            this.chkNoteTitleIncludeDate.Name = "chkNoteTitleIncludeDate";
            //this.chkNoteTitleIncludeDate.Size = new System.Drawing.Size(104, 24);
            this.chkNoteTitleIncludeDate.TabIndex = 1;
            this.chkNoteTitleIncludeDate.CheckedChanged += new System.EventHandler(this.chkNoteTitleIncludeDate_CheckedChanged);
            // 
            // chkMoveDateToFrontInThread
            // 
            this.chkMoveDateToFrontInThread.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.chkMoveDateToFrontInThread.AutoSize = true;
            this.chkMoveDateToFrontInThread.Location = new System.Drawing.Point(272, 3);
            this.chkMoveDateToFrontInThread.Name = "chkMoveDateToFrontInThread";
            //this.chkMoveDateToFrontInThread.Size = new System.Drawing.Size(412, 24);
            this.chkMoveDateToFrontInThread.TabIndex = 2;
            this.chkMoveDateToFrontInThread.Text = "Move date to front of filename when grouping threads";
            // 
            // grpDevelopment
            // 
            this.grpDevelopment.AutoSize = true;
            this.grpDevelopment.Controls.Add(this.devLayout);
            this.grpDevelopment.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpDevelopment.Location = new System.Drawing.Point(3, 811);
            this.grpDevelopment.Name = "grpDevelopment";
            //this.grpDevelopment.Size = new System.Drawing.Size(1156, 125);
            this.grpDevelopment.TabIndex = 5;
            this.grpDevelopment.TabStop = false;
            this.grpDevelopment.Text = "Development Settings";
            // 
            // devLayout
            // 
            this.devLayout.AutoSize = true;
            this.devLayout.Controls.Add(this.chkShowDevelopmentSettings);
            this.devLayout.Controls.Add(this.chkShowThreadDebug);
            this.devLayout.Location = new System.Drawing.Point(0, 0);
            this.devLayout.Name = "devLayout";
            //this.devLayout.Size = new System.Drawing.Size(416, 100);
            this.devLayout.TabIndex = 0;
            // 
            // chkShowDevelopmentSettings
            // 
            this.chkShowDevelopmentSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkShowDevelopmentSettings.AutoSize = true;
            this.chkShowDevelopmentSettings.Location = new System.Drawing.Point(3, 3);
            this.chkShowDevelopmentSettings.Name = "chkShowDevelopmentSettings";
            //this.chkShowDevelopmentSettings.Size = new System.Drawing.Size(230, 24);
            this.chkShowDevelopmentSettings.TabIndex = 0;
            this.chkShowDevelopmentSettings.Text = "Show development settings";
            // 
            // chkShowThreadDebug
            // 
            this.chkShowThreadDebug.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.chkShowThreadDebug.AutoSize = true;
            this.chkShowThreadDebug.Location = new System.Drawing.Point(239, 3);
            this.chkShowThreadDebug.Name = "chkShowThreadDebug";
            //this.chkShowThreadDebug.Size = new System.Drawing.Size(174, 24);
            this.chkShowThreadDebug.TabIndex = 1;
            this.chkShowThreadDebug.Text = "Show thread debug";
            // 
            // btnLayout
            // 
            this.btnLayout.AutoSize = true;
            this.btnLayout.Controls.Add(this.btnSave);
            this.btnLayout.Controls.Add(this.btnCancel);
            this.btnLayout.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnLayout.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.btnLayout.Location = new System.Drawing.Point(3, 942);
            this.btnLayout.Name = "btnLayout";
            //this.btnLayout.Size = new System.Drawing.Size(1156, 29);
            this.btnLayout.TabIndex = 6;
            // 
            // btnSave
            // 
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(1078, 3);
            this.btnSave.Name = "btnSave";
            //this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.AutoSize = true;
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(997, 3);
            this.btnCancel.Name = "btnCancel";
            //this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.AutoSize = true;
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            // 
            // SettingsForm
            // 
            this.ClientSize = new System.Drawing.Size(1175, 840);
            this.Controls.Add(this.mainLayout);
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Obsidian Settings";
            this.mainLayout.ResumeLayout(false);
            this.mainLayout.PerformLayout();
            this.grpVault.ResumeLayout(false);
            this.grpVault.PerformLayout();
            this.vaultLayout.ResumeLayout(false);
            this.vaultLayout.PerformLayout();
            this.grpGeneral.ResumeLayout(false);
            this.grpGeneral.PerformLayout();
            this.generalLayout.ResumeLayout(false);
            this.generalLayout.PerformLayout();
            this.grpTiming.ResumeLayout(false);
            this.grpTiming.PerformLayout();
            this.timingLayout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numDelay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultDueDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefaultReminderHour)).EndInit();
            this.grpPatterns.ResumeLayout(false);
            this.grpPatterns.PerformLayout();
            this.patternsLayout.ResumeLayout(false);
            this.patternsLayout.PerformLayout();
            this.btnPanel.ResumeLayout(false);
            this.grpNoteCustomization.ResumeLayout(false);
            this.grpNoteCustomization.PerformLayout();
            this.noteTagLayout.ResumeLayout(false);
            this.noteTagLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numNoteTitleMaxLength)).EndInit();
            this.dateOptionsLayout.ResumeLayout(false);
            this.dateOptionsLayout.PerformLayout();
            this.grpDevelopment.ResumeLayout(false);
            this.grpDevelopment.PerformLayout();
            this.devLayout.ResumeLayout(false);
            this.devLayout.PerformLayout();
            this.btnLayout.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        private void chkNoteTitleIncludeDate_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkNoteTitleIncludeDate.Checked)
            {
                chkMoveDateToFrontInThread.Checked = false;
                chkMoveDateToFrontInThread.Enabled = false;
            }
            else
            {
                chkMoveDateToFrontInThread.Enabled = true;
            }
        }

        //private void InitializeComponent()
        //{
        //    //this.toolTip = new ToolTip();
        //    //System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));

        //    //// Main layout panel
        //    //this.mainLayout = new TableLayoutPanel();
        //    //this.mainLayout.Dock = DockStyle.Fill;
        //    //this.mainLayout.AutoScroll = true;
        //    //this.mainLayout.ColumnCount = 1;
        //    //this.mainLayout.RowCount = 1;
        //    //this.mainLayout.GrowStyle = TableLayoutPanelGrowStyle.AddRows;

        //    // Group: Vault/Paths
        //    //var grpVault = new GroupBox();
        //    //grpVault.Text = "Vault Settings";
        //    //grpVault.Dock = DockStyle.Top;
        //    //grpVault.AutoSize = true;
        //    //var vaultLayout = new TableLayoutPanel();
        //    //vaultLayout.Dock = DockStyle.Fill;
        //    //vaultLayout.ColumnCount = 2;
        //    //vaultLayout.AutoSize = true;
        //    //vaultLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
        //    //vaultLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
        //    //vaultLayout.Controls.Add(new Label { Text = "Vault Name:", Anchor = AnchorStyles.Left }, 0, 0);
        //    ////this.txtVaultName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
        //    //vaultLayout.Controls.Add(this.txtVaultName, 1, 0);
        //    //vaultLayout.Controls.Add(new Label { Text = "Vault Base Path:", Anchor = AnchorStyles.Left }, 0, 1);
        //    //this.txtVaultPath = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
        //    //vaultLayout.Controls.Add(this.txtVaultPath, 1, 1);
        //    //this.btnBrowse = new Button { Text = "Browse...", Anchor = AnchorStyles.Left };
        //    //vaultLayout.Controls.Add(this.btnBrowse, 1, 2);
        //    //vaultLayout.Controls.Add(new Label { Text = "Inbox Folder:", Anchor = AnchorStyles.Left }, 0, 3);
        //    //this.txtInboxFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
        //    //vaultLayout.Controls.Add(this.txtInboxFolder, 1, 3);
        //    //vaultLayout.Controls.Add(new Label { Text = "Contacts Folder:", Anchor = AnchorStyles.Left }, 0, 4);
        //    //this.txtContactsFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
        //    //vaultLayout.Controls.Add(this.txtContactsFolder, 1, 4);
        //    //grpVault.Controls.Add(vaultLayout);
        //    //this.mainLayout.Controls.Add(grpVault);

        //    //// Group: General Settings (labels and checkboxes grow with width)
        //    //var grpGeneral = new GroupBox();
        //    //grpGeneral.Text = "General Settings";
        //    //grpGeneral.Dock = DockStyle.Top;
        //    //grpGeneral.AutoSize = true;
        //    //var generalLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        //    //generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        //    //generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        //    //this.chkEnableContactSaving = new CheckBox { Text = "Enable Contact Saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkEnableContactSaving, 0, 0);
        //    //this.chkSearchEntireVaultForContacts = new CheckBox { Text = "Search entire vault for contacts", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkSearchEntireVaultForContacts, 1, 0);
        //    //this.chkLaunchObsidian = new CheckBox { Text = "Launch Obsidian after saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkLaunchObsidian, 0, 1);
        //    //this.chkShowCountdown = new CheckBox { Text = "Show countdown", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkShowCountdown, 1, 1);
        //    //this.chkCreateObsidianTask = new CheckBox { Text = "Create task in Obsidian note", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkCreateObsidianTask, 0, 2);
        //    //this.chkCreateOutlookTask = new CheckBox { Text = "Create task in Outlook", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkCreateOutlookTask, 1, 2);
        //    //this.chkAskForDates = new CheckBox { Text = "Ask for dates and times each time", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkAskForDates, 0, 3);
        //    //this.chkGroupEmailThreads = new CheckBox { Text = "Group email threads", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //generalLayout.Controls.Add(this.chkGroupEmailThreads, 1, 3);
        //    //grpGeneral.Controls.Add(generalLayout);
        //    //this.mainLayout.Controls.Add(grpGeneral);

        //    //// Group: Delay/Task timing
        //    //var grpTiming = new GroupBox();
        //    //grpTiming.Text = "Timing Settings";
        //    //grpTiming.Dock = DockStyle.Top;
        //    //grpTiming.AutoSize = true;
        //    //var timingLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        //    //timingLayout.Controls.Add(new Label { Text = "Delay (seconds):", Anchor = AnchorStyles.Left }, 0, 0);
        //    //this.numDelay = new NumericUpDown { Minimum = 0, Maximum = 10, Anchor = AnchorStyles.Left };
        //    //timingLayout.Controls.Add(this.numDelay, 1, 0);
        //    //timingLayout.Controls.Add(new Label { Text = "Due in Days:", Anchor = AnchorStyles.Left }, 0, 1);
        //    //this.numDefaultDueDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
        //    //timingLayout.Controls.Add(this.numDefaultDueDays, 1, 1);
        //    //timingLayout.Controls.Add(new Label { Text = "Reminder Days:", Anchor = AnchorStyles.Left }, 0, 2);
        //    //this.numDefaultReminderDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
        //    //timingLayout.Controls.Add(this.numDefaultReminderDays, 1, 2);
        //    //timingLayout.Controls.Add(new Label { Text = "Reminder Hour:", Anchor = AnchorStyles.Left }, 0, 3);
        //    //this.numDefaultReminderHour = new NumericUpDown { Minimum = 0, Maximum = 23, Anchor = AnchorStyles.Left };
        //    //timingLayout.Controls.Add(this.numDefaultReminderHour, 1, 3);
        //    //grpTiming.Controls.Add(timingLayout);
        //    //this.mainLayout.Controls.Add(grpTiming);

        //    //// Group: Patterns
        //    //var grpPatterns = new GroupBox();
        //    //grpPatterns.Text = "Subject Cleanup Patterns";
        //    //grpPatterns.Dock = DockStyle.Top;
        //    //grpPatterns.AutoSize = true;
        //    //var patternsLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        //    //this.lstPatterns = new ListBox { Height = 120, Width = 400 };
        //    //patternsLayout.Controls.Add(this.lstPatterns, 0, 0);
        //    //var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true };
        //    //this.btnAdd = new Button { Text = "Add" };
        //    //this.btnEdit = new Button { Text = "Edit" };
        //    //this.btnRemove = new Button { Text = "Remove" };
        //    //btnPanel.Controls.Add(this.btnAdd);
        //    //btnPanel.Controls.Add(this.btnEdit);
        //    //btnPanel.Controls.Add(this.btnRemove);
        //    //patternsLayout.Controls.Add(btnPanel, 1, 0);
        //    //grpPatterns.Controls.Add(patternsLayout);
        //    //this.mainLayout.Controls.Add(grpPatterns);

        //    //// Group: Note/Tag Customization
        //    //this.grpNoteCustomization = new GroupBox();
        //    //this.grpNoteCustomization.Text = "Note & Tag Customization";
        //    //this.grpNoteCustomization.Dock = DockStyle.Top;
        //    //this.grpNoteCustomization.AutoSize = true;
        //    //var noteTagLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        //    //noteTagLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
        //    //noteTagLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));
        //    //this.lblDefaultNoteTags = new Label { Text = "Default Note Tags (comma-separated):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
        //    //this.txtDefaultNoteTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
        //    //noteTagLayout.Controls.Add(this.lblDefaultNoteTags, 0, 0);
        //    //noteTagLayout.Controls.Add(this.txtDefaultNoteTags, 1, 0);
        //    //this.lblDefaultTaskTags = new Label { Text = "Default Task Tags (comma-separated, will be rendered as #tags):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
        //    //this.txtDefaultTaskTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
        //    //noteTagLayout.Controls.Add(this.lblDefaultTaskTags, 0, 1);
        //    //noteTagLayout.Controls.Add(this.txtDefaultTaskTags, 1, 1);
        //    //this.lblNoteTitleFormat = new Label { Text = "Note Title Format (use {Subject}, {Sender}, {Date}):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
        //    //this.txtNoteTitleFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
        //    //noteTagLayout.Controls.Add(this.lblNoteTitleFormat, 0, 2);
        //    //noteTagLayout.Controls.Add(this.txtNoteTitleFormat, 1, 2);
        //    //this.lblNoteTitleMaxLength = new Label { Text = "Max Title Length:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
        //    //this.numNoteTitleMaxLength = new NumericUpDown { Minimum = 10, Maximum = 200, Anchor = AnchorStyles.Left, Dock = DockStyle.Fill };
        //    //noteTagLayout.Controls.Add(this.lblNoteTitleMaxLength, 0, 3);
        //    //noteTagLayout.Controls.Add(this.numNoteTitleMaxLength, 1, 3);
        //    //// Add Note Title Include Date and Move Date checkboxes in a horizontal layout for alignment
        //    //var dateOptionsLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left };
        //    //this.chkNoteTitleIncludeDate = new CheckBox { Anchor = AnchorStyles.Left };
        //    //this.lblNoteTitleIncludeDate = new Label { Text = "Include Date in Title:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
        //    //dateOptionsLayout.Controls.Add(this.lblNoteTitleIncludeDate);
        //    //dateOptionsLayout.Controls.Add(this.chkNoteTitleIncludeDate);
        //    //this.chkMoveDateToFrontInThread = new CheckBox { Text = "Move date to front of filename when grouping threads", Anchor = AnchorStyles.Left, AutoSize = true };
        //    //dateOptionsLayout.Controls.Add(this.chkMoveDateToFrontInThread);
        //    //noteTagLayout.Controls.Add(dateOptionsLayout, 1, 4);

        //    //// Add event handler for enabling/disabling move date checkbox
        //    //this.chkNoteTitleIncludeDate.CheckedChanged += (s, e) => {
        //    //    if (!chkNoteTitleIncludeDate.Checked) {
        //    //        chkMoveDateToFrontInThread.Checked = false;
        //    //        chkMoveDateToFrontInThread.Enabled = false;
        //    //    } else {
        //    //        chkMoveDateToFrontInThread.Enabled = true;
        //    //    }
        //    //};

        //    //this.grpNoteCustomization.Controls.Add(noteTagLayout);
        //    //this.mainLayout.Controls.Add(this.grpNoteCustomization);

        //    //// Group: Development (always at bottom, always visible)
        //    //this.grpDevelopment = new GroupBox();
        //    //this.grpDevelopment.Text = "Development Settings";
        //    //this.grpDevelopment.Dock = DockStyle.Top;
        //    //this.grpDevelopment.AutoSize = true;
        //    //var devLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        //    //this.chkShowDevelopmentSettings = new CheckBox { Text = "Show development settings", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //this.chkShowThreadDebug = new CheckBox { Text = "Show thread debug", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        //    //devLayout.Controls.Add(this.chkShowDevelopmentSettings);
        //    //devLayout.Controls.Add(this.chkShowThreadDebug);
        //    //this.grpDevelopment.Controls.Add(devLayout);
        //    //this.mainLayout.Controls.Add(this.grpDevelopment);

        //    //// Save/Cancel buttons
        //    //var btnLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, AutoSize = true };
        //    //this.btnSave = new Button { Text = "Save", DialogResult = DialogResult.OK };
        //    //this.btnSave.Click += btnSave_Click;
        //    //this.btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
        //    //btnLayout.Controls.Add(this.btnSave);
        //    //btnLayout.Controls.Add(this.btnCancel);
        //    //this.mainLayout.Controls.Add(btnLayout);

        //    //// Set up the form
        //    //this.Controls.Add(this.mainLayout);
        //    //this.FormBorderStyle = FormBorderStyle.Sizable;
        //    //this.MaximizeBox = true;
        //    //this.MinimizeBox = true;
        //    //this.StartPosition = FormStartPosition.CenterScreen;
        //    //this.Text = "Obsidian Settings";
        //    //this.ClientSize = new System.Drawing.Size(700, 800);
        //}

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
            chkShowDevelopmentSettings.Checked = _settings.ShowDevelopmentSettings;
            chkShowThreadDebug.Checked = _settings.ShowThreadDebug;
            chkMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            chkMoveDateToFrontInThread.Enabled = chkNoteTitleIncludeDate.Checked;

            // Initialize development settings visibility
            grpDevelopment.Visible = _settings.ShowDevelopmentSettings;
            chkShowThreadDebug.Visible = _settings.ShowDevelopmentSettings;

            // Load patterns
            lstPatterns.Items.Clear();
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                lstPatterns.Items.Add(pattern);
            }

            txtDefaultNoteTags.Text = string.Join(", ", _settings.DefaultNoteTags ?? new List<string>());
            txtDefaultTaskTags.Text = string.Join(", ", _settings.DefaultTaskTags ?? new List<string>());
            txtNoteTitleFormat.Text = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
            numNoteTitleMaxLength.Value = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;
            chkNoteTitleIncludeDate.Checked = _settings.NoteTitleIncludeDate;
            chkMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            chkMoveDateToFrontInThread.Enabled = chkNoteTitleIncludeDate.Checked;
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
            _settings.ShowDevelopmentSettings = chkShowDevelopmentSettings.Checked;
            _settings.ShowThreadDebug = chkShowThreadDebug.Checked;
            _settings.MoveDateToFrontInThread = chkMoveDateToFrontInThread.Checked;

            // Save patterns
            _settings.SubjectCleanupPatterns.Clear();
            foreach (string pattern in lstPatterns.Items)
            {
                _settings.SubjectCleanupPatterns.Add(pattern);
            }

            _settings.DefaultNoteTags = txtDefaultNoteTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.DefaultTaskTags = txtDefaultTaskTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.NoteTitleFormat = txtNoteTitleFormat.Text.Trim();
            _settings.NoteTitleMaxLength = (int)numNoteTitleMaxLength.Value;
            _settings.NoteTitleIncludeDate = chkNoteTitleIncludeDate.Checked;

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

        private void chkShowDevelopmentSettings_CheckedChanged(object sender, EventArgs e)
        {
            grpDevelopment.Visible = chkShowDevelopmentSettings.Checked;
            chkShowThreadDebug.Visible = chkShowDevelopmentSettings.Checked;
        }
    }
} 