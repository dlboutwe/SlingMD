using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Forms
{
    public partial class Settings : Form
    {
        private readonly ObsidianSettings _settings;

        public Settings(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void LoadSettings()
        {
            // Vault Settings
            txtVaultName.Text = _settings.VaultName;
            txtVaultBasePath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            txtContactsFolder.Text = _settings.ContactsFolder;
            txtAppointmentsFolder.Text = _settings.AppointmentsFolder;
            
            // General Settings
            cbLaunchObsidian.Checked = _settings.LaunchObsidian;
            cbEnableContactSaving.Checked = _settings.EnableContactSaving;
            cbSearchEntireVaultForContacts.Checked = _settings.SearchEntireVaultForContacts;
            txtContactsFolder.Enabled = _settings.EnableContactSaving;
            cbShowCountdown.Checked = _settings.ShowCountdown;

            // Timing Settings
            nbrObsidianDelaySeconds.Value = _settings.ObsidianDelaySeconds;
            nbrDefaultDueDays.Value = _settings.DefaultDueDays;
            nbrDefaultReminderDays.Value = _settings.DefaultReminderDays;
            nbrDefaultReminderHour.Value = _settings.DefaultReminderHour;

            // Load patterns
            lstSubjectCleanupPatterns.Items.Clear();
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                lstSubjectCleanupPatterns.Items.Add(pattern);
            }

            // Email Settings Tab
            lstDefaultNoteTags.Items.Clear();
            foreach (var pattern in _settings.DefaultNoteTags)
            {
                lstDefaultNoteTags.Items.Add(pattern);
            }
            txtNoteTitleFormat.Text = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
            nbrNoteTitleMaxLength.Value = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;
            cbGroupEmailThreads.Checked = _settings.GroupEmailThreads;
            cbMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            cbMoveDateToFrontInThread.Enabled = cbNoteTitleIncludeDate.Checked;
            cbNoteTitleIncludeDate.Checked = _settings.NoteTitleIncludeDate;
            cbMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            cbMoveDateToFrontInThread.Enabled = cbNoteTitleIncludeDate.Checked;

            // Appointment Settings Tab
            lstDefaultAppointmentTags.Items.Clear();
            foreach (var pattern in _settings.AppointmentDefaultNoteTags)
            {
                lstDefaultAppointmentTags.Items.Add(pattern);
            }
            txtAppointmentNoteTitleFormat.Text = _settings.AppointmentNoteTitleFormat ?? "{Date} - {Subject}";
            nbrAppointmentNoteTitleMaxLength.Value = _settings.AppointmentNoteTitleMaxLength > 0 ? _settings.AppointmentNoteTitleMaxLength : 50;
            cbAppointmentSaveAttachments.Checked = _settings.AppointmentSaveAttachments;

            //Task Settings Tab
            lstDefaultTaskTags.Items.Clear();
            foreach (var pattern in _settings.DefaultTaskTags)
            {
                lstDefaultTaskTags.Items.Add(pattern);
            }
            cbCreateObsidianTask.Checked = _settings.CreateObsidianTask;
            cbCreateOutlookTask.Checked = _settings.CreateOutlookTask;
            cbAskForDates.Checked = _settings.AskForDates;

            // Developer Settings Tab
            cbShowDevelopmentSettings.Checked = _settings.ShowDevelopmentSettings;
            cbShowThreadDebug.Checked = _settings.ShowThreadDebug;
            tabDeveloperSettings.Visible = _settings.ShowDevelopmentSettings;
            cbShowThreadDebug.Visible = _settings.ShowDevelopmentSettings;
        }

        private void btnBrowseVaultPath_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "Select Obsidian Vault Base Directory";
            folderBrowserDialog1.SelectedPath = txtVaultBasePath.Text;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtVaultBasePath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Vault Settings
            _settings.VaultName = txtVaultName.Text;
            _settings.VaultBasePath = txtVaultBasePath.Text;
            _settings.InboxFolder = txtInboxFolder.Text;
            _settings.ContactsFolder = txtContactsFolder.Text;
            _settings.AppointmentsFolder = txtAppointmentsFolder.Text;

            // General Settings
            _settings.LaunchObsidian = cbLaunchObsidian.Checked;
            _settings.EnableContactSaving = cbEnableContactSaving.Checked;
            _settings.SearchEntireVaultForContacts = cbSearchEntireVaultForContacts.Checked;
            _settings.EnableContactSaving = txtContactsFolder.Enabled;
            _settings.ShowCountdown = cbShowCountdown.Checked;

            // Timing Settings
            _settings.ObsidianDelaySeconds = (int)nbrObsidianDelaySeconds.Value;
            _settings.DefaultDueDays = (int)nbrDefaultDueDays.Value;
            _settings.DefaultReminderDays = (int)nbrDefaultReminderDays.Value;
            _settings.DefaultReminderHour = (int)nbrDefaultReminderHour.Value;

            

            // Load patterns
            _settings.SubjectCleanupPatterns.Clear();
            foreach (string pattern in lstSubjectCleanupPatterns.Items)
            {
                _settings.SubjectCleanupPatterns.Add(pattern);
            }

            // Email Settings Tab
            _settings.DefaultNoteTags = lstDefaultNoteTags.Items.Cast<string>().Select(s=> s.Trim()).Where((w=>!string.IsNullOrEmpty(w))).ToList();
            _settings.NoteTitleFormat = txtNoteTitleFormat.Text.Trim();
            _settings.NoteTitleMaxLength = (int)nbrNoteTitleMaxLength.Value;
            _settings.NoteTitleIncludeDate = cbNoteTitleIncludeDate.Checked;
            _settings.GroupEmailThreads = cbGroupEmailThreads.Checked;
            _settings.MoveDateToFrontInThread = cbMoveDateToFrontInThread.Checked;
            _settings.NoteTitleIncludeDate = cbNoteTitleIncludeDate.Checked;
            _settings.MoveDateToFrontInThread = cbMoveDateToFrontInThread.Checked;

            // Appointment Settings Tab
            _settings.AppointmentDefaultNoteTags = lstDefaultAppointmentTags.Items.Cast<string>().Select(s => s.Trim()).Where((w => !string.IsNullOrEmpty(w))).ToList();
            _settings.AppointmentNoteTitleFormat = txtAppointmentNoteTitleFormat.Text;
            _settings.AppointmentNoteTitleMaxLength = (int)nbrAppointmentNoteTitleMaxLength.Value;
            _settings.AppointmentSaveAttachments = cbAppointmentSaveAttachments.Checked;

            //Task Settings tab
            _settings.DefaultTaskTags = lstDefaultTaskTags.Items.Cast<string>().Select(s => s.Trim()).Where((w => !string.IsNullOrEmpty(w))).ToList();
            _settings.CreateObsidianTask = cbCreateObsidianTask.Checked;
            _settings.CreateOutlookTask = cbCreateOutlookTask.Checked;
            _settings.AskForDates = cbAskForDates.Checked;

            // Developer Settings Tab
            _settings.ShowDevelopmentSettings = cbShowDevelopmentSettings.Checked;
            _settings.ShowThreadDebug = cbShowThreadDebug.Checked;

            _settings.Save();
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cbEnableContactSaving_CheckedChanged(object sender, EventArgs e)
        {
            txtContactsFolder.Enabled = cbEnableContactSaving.Checked;
        }

        /* Subject Cleanup Patterns*/
        private void btnAddSubjectCleanupPattern_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Pattern", "Enter regex pattern:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    lstSubjectCleanupPatterns.Items.Add(form.InputText);
                }
            }
        }

        private void btnEditSubjectCleanupPattern_Click(object sender, EventArgs e)
        {
            if (lstSubjectCleanupPatterns.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Pattern", "Edit regex pattern:", lstSubjectCleanupPatterns.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstSubjectCleanupPatterns.SelectedIndex;
                        lstSubjectCleanupPatterns.Items[index] = form.InputText;
                    }
                }
            }
        }

        private void btnRemoveSubjectCleanupPattern_Click(object sender, EventArgs e)
        {
            if (lstSubjectCleanupPatterns.SelectedItem != null)
            {
                lstSubjectCleanupPatterns.Items.RemoveAt(lstSubjectCleanupPatterns.SelectedIndex);
            }
        }

        /* Default Note Tags */
        private void btnAddNoteTag_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Tag", "Enter tag:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    string tag = String.Join("", form.InputText.Split('#'));
                    lstDefaultNoteTags.Items.Add(tag);
                }
            }
        }

        private void btnEditNoteTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultNoteTags.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Tag", "Edit tag:", lstDefaultNoteTags.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstDefaultNoteTags.SelectedIndex;
                        string tag = String.Join("", form.InputText.Split('#'));
                        lstDefaultNoteTags.Items[index] = tag;
                    }
                }
            }
        }

        private void btnRemoveNoteTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultNoteTags.SelectedItem != null)
            {
                lstDefaultNoteTags.Items.RemoveAt(lstDefaultNoteTags.SelectedIndex);
            }
        }

        /* Default Appointment Tags */
        private void btnAddAppointmentTag_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Tag", "Enter tag:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    string tag = String.Join("", form.InputText.Split('#'));
                    lstDefaultAppointmentTags.Items.Add(tag);
                }
            }
        }

        private void btnEditAppointmentTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultAppointmentTags.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Tag", "Edit tag:", lstDefaultAppointmentTags.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstDefaultAppointmentTags.SelectedIndex;
                        string tag = String.Join("", form.InputText.Split('#'));
                        lstDefaultAppointmentTags.Items[index] = tag;
                    }
                }
            }
        }

        private void btnRemoveAppointmentTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultAppointmentTags.SelectedItem != null)
            {
                lstDefaultAppointmentTags.Items.RemoveAt(lstDefaultAppointmentTags.SelectedIndex);
            }
        }

        /* Default Task Tags */
        private void btnAddTaskTag_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Tag", "Enter tag:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    string tag = String.Join("", form.InputText.Split('#'));
                    lstDefaultTaskTags.Items.Add(tag);
                }
            }
        }

        private void btnEditTaskTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultTaskTags.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Tag", "Enter Tag:", lstDefaultTaskTags.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstDefaultTaskTags.SelectedIndex;
                        string tag = String.Join("", form.InputText.Split('#'));
                        lstDefaultTaskTags.Items[index] = tag;
                    }
                }
            }
        }

        private void btnRemoveTaskTag_Click(object sender, EventArgs e)
        {
            if (lstDefaultTaskTags.SelectedItem != null)
            {
                lstDefaultTaskTags.Items.RemoveAt(lstDefaultTaskTags.SelectedIndex);
            }
        }


        private void cbNoteTitleIncludeDate_CheckedChanged(object sender, EventArgs e)
        {
            if (!cbNoteTitleIncludeDate.Checked)
            {
                cbMoveDateToFrontInThread.Checked = false;
                cbMoveDateToFrontInThread.Enabled = false;
            }
            else
            {
                cbMoveDateToFrontInThread.Enabled = true;
            }
        }
    }
}
