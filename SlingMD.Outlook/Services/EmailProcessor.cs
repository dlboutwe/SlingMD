using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;
using System.Linq;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Outlook.Services
{
    public class EmailProcessor
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _createTasks = true;
        private bool _useRelativeReminder;

        public EmailProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService);
            _taskService = new TaskService(settings);
            _contactService = new ContactService(_fileService, _templateService);
        }

        public async Task ProcessEmail(MailItem mail)
        {
            // Get task options first if needed
            if ((_settings.CreateOutlookTask || _settings.CreateObsidianTask) && _settings.AskForDates)
            {
                using (var form = new TaskOptionsForm(_settings.DefaultDueDays, _settings.DefaultReminderDays, _settings.DefaultReminderHour, _settings.UseRelativeReminder))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        _taskService.InitializeTaskSettings(form.DueDays, form.ReminderDays, form.ReminderHour, form.UseRelativeReminder);
                    }
                    else
                    {
                        _taskService.DisableTaskCreation();
                    }
                }
            }
            else
            {
                _taskService.InitializeTaskSettings();
            }

            using (var status = new StatusService())
            {
                try
                {
                    status.UpdateProgress("Processing email...", 0);

                    // Clean and prepare file name components
                    string subjectClean = CleanSubject(mail.Subject);
                    if (subjectClean.Length > 50)  // Limit subject length
                    {
                        subjectClean = subjectClean.Substring(0, 47) + "...";
                    }
                    string senderClean = _contactService.GetShortName(mail.SenderName);
                    string fileDateTime = mail.ReceivedTime.ToString("yyyy-MM-dd-HHmm");
                    
                    status.UpdateProgress("Creating note file", 25);

                    // Build file name with date and time prepended
                    string fileName = $"{fileDateTime}-{subjectClean}-{senderClean}.md";
                    string filePath = Path.Combine(_settings.GetInboxPath(), fileName);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(fileName);

                    // Get thread info
                    string conversationId = _threadService.GetConversationId(mail);
                    string threadNoteName = _threadService.GetThreadNoteName(mail, subjectClean, senderClean, 
                        _contactService.GetShortName(GetFirstRecipient(mail)));
                    string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                    string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");

                    // Check for existing thread
                    var (hasExistingThread, earliestEmailThreadName, _) = 
                        _threadService.FindExistingThread(conversationId, _settings.GetInboxPath());

                    // If this is part of a thread and thread grouping is enabled, update paths
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        threadNoteName = earliestEmailThreadName ?? threadNoteName;
                        threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                        threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
                        filePath = Path.Combine(threadFolderPath, fileName);
                    }

                    status.UpdateProgress("Processing email metadata", 50);

                    // Build metadata for frontmatter
                    var metadata = new Dictionary<string, object>
                    {
                        { "title", mail.Subject },
                        { "from", $"[[{mail.SenderName}]]" },
                        { "fromEmail", _contactService.GetSenderEmail(mail) },
                        { "to", _contactService.BuildLinkedNames(mail.Recipients, OlMailRecipientType.olTo) },
                        { "toEmail", _contactService.BuildEmailList(mail.Recipients, OlMailRecipientType.olTo) },
                        { "threadId", conversationId },
                        { "date", mail.ReceivedTime },
                        { "dailyNoteLink", $"[[{mail.ReceivedTime:yyyy-MM-dd}]]" },
                        { "tags", "[email]" }
                    };

                    // Add CC if present
                    string ccLinked = _contactService.BuildLinkedNames(mail.Recipients, OlMailRecipientType.olCC);
                    string ccEmails = _contactService.BuildEmailList(mail.Recipients, OlMailRecipientType.olCC);
                    if (!string.IsNullOrEmpty(ccEmails))
                    {
                        metadata.Add("cc", ccLinked);
                        metadata.Add("ccEmail", ccEmails);
                    }

                    // Add threadNote if this is part of a thread and thread grouping is enabled
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        metadata.Add("threadNote", $"[[0-{threadNoteName}]]");
                    }

                    // Build content
                    var content = new System.Text.StringBuilder();
                    content.Append(_templateService.BuildFrontMatter(metadata));

                    // Add Obsidian task if enabled
                    if (_settings.CreateObsidianTask && _taskService.ShouldCreateTasks)
                    {
                        content.Append(_taskService.GenerateObsidianTask(fileNameNoExt));
                    }

                    content.Append(mail.Body);

                    status.UpdateProgress("Writing note file", 75);

                    // Write the file
                    _fileService.WriteUtf8File(filePath, content.ToString());

                    // If this is part of a thread and thread grouping is enabled
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        await _threadService.UpdateThreadNote(threadFolderPath, threadNotePath, conversationId, threadNoteName, mail);
                    }

                    // Create Outlook task if enabled
                    if (_settings.CreateOutlookTask && _taskService.ShouldCreateTasks)
                    {
                        status.UpdateProgress("Creating Outlook task", 80);
                        await _taskService.CreateOutlookTask(mail);
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
                        _fileService.LaunchObsidian(_settings.VaultName, fileName);
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

        private string CleanSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
                return string.Empty;

            string cleaned = subject;

            // Apply all cleanup patterns from settings
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                cleaned = Regex.Replace(cleaned, pattern, "", RegexOptions.IgnoreCase);
            }

            // Clean the remaining text for file name safety
            return _fileService.CleanFileName(cleaned);
        }

        private string GetFirstRecipient(MailItem mail)
        {
            foreach (Recipient recipient in mail.Recipients)
            {
                if (recipient.Type == (int)OlMailRecipientType.olTo)
                {
                    return recipient.Name;
                }
            }
            return "Unknown";
        }
    }
}