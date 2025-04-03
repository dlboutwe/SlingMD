using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class EmailProcessor
    {
        private readonly ObsidianSettings _settings;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _createTasks = true;
        private bool _useRelativeReminder;

        public EmailProcessor(ObsidianSettings settings)
        {
            _settings = settings;
        }

        public async Task ProcessEmail(MailItem mail)
        {
            // Get task options first if needed
            if ((_settings.CreateOutlookTask || _settings.CreateObsidianTask) && _settings.AskForDates)
            {
                using (var form = new TaskOptionsForm(_settings.DefaultDueDays, _settings.DefaultReminderDays, _settings.DefaultReminderHour, _settings.UseRelativeReminder))
                {
                    if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        _taskDueDays = form.DueDays;
                        _taskReminderDays = form.ReminderDays;
                        _taskReminderHour = form.ReminderHour;
                        _useRelativeReminder = form.UseRelativeReminder;
                    }
                    else
                    {
                        _createTasks = false;
                    }
                }
            }
            else
            {
                _taskDueDays = _settings.DefaultDueDays;
                _taskReminderDays = _settings.DefaultReminderDays;
                _taskReminderHour = _settings.DefaultReminderHour;
                _useRelativeReminder = _settings.UseRelativeReminder;
            }

            using (var status = new StatusService())
            {
                try
                {
                    status.UpdateProgress("Processing email...", 0);

                    // Clean and prepare file name components
                    string subjectClean = CleanFileName(mail.Subject);
                    string senderClean = CleanFileName(mail.SenderName);
                    string fileDate = mail.ReceivedTime.ToString("yyyy-MM-dd");
                    
                    status.UpdateProgress("Creating note file", 25);

                    // Build file name and path
                    string fileName = $"{subjectClean} - {senderClean} - {fileDate}.md";
                    string filePath = Path.Combine(_settings.GetInboxPath(), fileName);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(fileName);

                    // Build YAML frontmatter
                    var frontmatter = new StringBuilder();
                    frontmatter.AppendLine("---");
                    frontmatter.AppendLine($"title: \"{mail.Subject}\"");
                    frontmatter.AppendLine($"from: \"[[{mail.SenderName}]]\"");
                    frontmatter.AppendLine($"fromEmail: \"{GetSenderEmail(mail)}\"");
                    frontmatter.AppendLine($"to: \"{BuildLinkedNames(mail.Recipients, OlMailRecipientType.olTo)}\"");
                    frontmatter.AppendLine($"toEmail: \"{BuildEmailList(mail.Recipients, OlMailRecipientType.olTo)}\"");

                    status.UpdateProgress("Processing email metadata", 50);

                    // Add CC if present
                    string ccLinked = BuildLinkedNames(mail.Recipients, OlMailRecipientType.olCC);
                    string ccEmails = BuildEmailList(mail.Recipients, OlMailRecipientType.olCC);
                    if (!string.IsNullOrEmpty(ccEmails))
                    {
                        frontmatter.AppendLine($"cc: \"{ccLinked}\"");
                        frontmatter.AppendLine($"ccEmail: \"{ccEmails}\"");
                    }

                    frontmatter.AppendLine($"date: {mail.ReceivedTime:yyyy-MM-dd HH:mm}");
                    frontmatter.AppendLine($"dailyNoteLink: \"[[{fileDate}]]\"");
                    frontmatter.AppendLine("tags: [email]");
                    frontmatter.AppendLine("---");
                    frontmatter.AppendLine();

                    // Add Obsidian task if enabled
                    if (_settings.CreateObsidianTask && _createTasks)
                    {
                        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                        string dueDate = DateTime.Now.Date.AddDays(_taskDueDays.Value).ToString("yyyy-MM-dd");
                        
                        // Calculate reminder date based on setting
                        DateTime reminderDateTime;
                        if (_useRelativeReminder)
                        {
                            // Relative: Calculate from due date
                            reminderDateTime = DateTime.Now.Date.AddDays(_taskDueDays.Value - _taskReminderDays.Value);
                        }
                        else
                        {
                            // Absolute: Calculate from today
                            reminderDateTime = DateTime.Now.Date.AddDays(_taskReminderDays.Value);
                        }
                        string reminderDate = reminderDateTime.ToString("yyyy-MM-dd");
                        
                        frontmatter.AppendLine($"- [ ] [[{fileNameNoExt}]] #FollowUp ➕ {currentDate} 🛫 {reminderDate} 📅 {dueDate}");
                        frontmatter.AppendLine();
                    }

                    status.UpdateProgress("Writing note file", 75);

                    // Combine content and write file
                    string content = frontmatter.ToString() + mail.Body;
                    WriteUtf8File(filePath, content);

                    // Create Outlook task if enabled
                    if (_settings.CreateOutlookTask && _createTasks)
                    {
                        status.UpdateProgress("Creating Outlook task", 80);
                        await CreateOutlookTaskAsync(mail);
                    }

                    status.UpdateProgress("Launching Obsidian", 90);

                    // Launch Obsidian if enabled
                    if (_settings.LaunchObsidian)
                    {
                        if (_settings.ObsidianDelaySeconds > 0)
                        {
                            if (_settings.ShowCountdown)
                            {
                                for (int i = _settings.ObsidianDelaySeconds; i > 0; i--)
                                {
                                    status.UpdateProgress($"Opening in Obsidian in {i} seconds...", 90);
                                    await Task.Delay(1000);
                                }
                            }
                            else
                            {
                                await Task.Delay(_settings.ObsidianDelaySeconds * 1000);
                            }
                        }
                        LaunchObsidian(_settings.VaultName, fileName);
                    }

                    status.ShowSuccess("Email saved to Obsidian successfully!", true);
                }
                catch (System.Exception ex)
                {
                    status.ShowError($"Error: {ex.Message}", false);
                    throw;
                }
            }
        }

        private string GetSenderEmail(MailItem mail)
        {
            try
            {
                // Try to get SMTP address using PR_SMTP_ADDRESS property
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                return mail.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
            }
            catch
            {
                // Fallback to SenderEmailAddress
                return mail.SenderEmailAddress;
            }
        }

        private string BuildLinkedNames(Recipients recipients, OlMailRecipientType type)
        {
            var names = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                if (recipient.Type == (int)type)
                {
                    names.Add($"[[{recipient.Name}]]");
                }
            }
            return string.Join(", ", names);
        }

        private string BuildEmailList(Recipients recipients, OlMailRecipientType type)
        {
            var emails = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                if (recipient.Type == (int)type)
                {
                    try
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        string email = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
                        if (!string.IsNullOrEmpty(email))
                        {
                            emails.Add(email);
                        }
                    }
                    catch
                    {
                        // Fallback to Address property
                        if (!string.IsNullOrEmpty(recipient.Address))
                        {
                            emails.Add(recipient.Address);
                        }
                    }
                }
            }
            return string.Join(", ", emails);
        }

        private string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Replace invalid characters with a dash
            var invalidChars = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(input);
            foreach (char c in invalidChars)
            {
                sb.Replace(c, '-');
            }
            return sb.ToString().Trim();
        }

        private void WriteUtf8File(string filePath, string content)
        {
            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            // Write the file with UTF-8 encoding
            using (var stream = new StreamWriter(filePath, false, new UTF8Encoding(false)))
            {
                stream.Write(content);
            }
        }

        private void LaunchObsidian(string vaultName, string filePath)
        {
            string encodedPath = Uri.EscapeDataString(filePath);
            string command = $"obsidian://open?vault={vaultName}&file={encodedPath}";
            
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = command,
                UseShellExecute = true
            });
        }

        private async Task CreateOutlookTaskAsync(MailItem mail)
        {
            try
            {
                var outlookApp = mail.Application;
                var task = outlookApp.CreateItem(OlItemType.olTaskItem);
                task.Subject = $"Follow up: {mail.Subject}";
                task.Body = $"Follow up on email from {mail.SenderName}\n\nOriginal email:\n{mail.Body}";
                
                // Set due date based on settings
                var dueDate = DateTime.Now.Date.AddDays(_taskDueDays.Value);
                task.DueDate = dueDate;
                task.ReminderSet = true;
                
                // Calculate reminder time based on setting
                DateTime reminderDate;
                if (_useRelativeReminder)
                {
                    // Relative: Calculate from due date
                    reminderDate = dueDate.AddDays(-_taskReminderDays.Value);
                }
                else
                {
                    // Absolute: Calculate from today
                    reminderDate = DateTime.Now.Date.AddDays(_taskReminderDays.Value);
                }
                var reminderTime = reminderDate.AddHours(_taskReminderHour.Value);
                
                // If reminder would be in the past, set it to the next possible time
                if (reminderTime < DateTime.Now)
                {
                    if (reminderTime.Date == DateTime.Now.Date)
                    {
                        // If it's today but earlier hour, set to next hour
                        reminderTime = DateTime.Now.AddHours(1);
                    }
                    else
                    {
                        // If it's a past day, set to tomorrow at the specified hour
                        reminderTime = DateTime.Now.Date.AddDays(1).AddHours(_taskReminderHour.Value);
                    }
                }
                
                task.ReminderTime = reminderTime;
                task.Save();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to create Outlook task: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
} 