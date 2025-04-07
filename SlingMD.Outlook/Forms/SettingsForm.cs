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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.SuspendLayout();
            // 
            // SettingsForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingsForm";
            this.ResumeLayout(false);

        }

        private void LoadSettings()
        {
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            txtContactsFolder.Text = _settings.ContactsFolder;
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
            _settings.ContactsFolder = txtContactsFolder.Text;
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
    }
} 