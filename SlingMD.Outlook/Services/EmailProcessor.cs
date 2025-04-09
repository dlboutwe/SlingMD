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
            // Declare variables at method level so they're accessible throughout the method
            List<string> contactNames = new List<string>();
            string fileName = string.Empty;
            string fileNameNoExt = string.Empty;
            string filePath = string.Empty;
            string obsidianLinkPath = string.Empty;  // Added to store the path to use for Obsidian launch

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

            // Collect all contact names - will be used later for contact creation
            contactNames.Add(mail.SenderName);
            foreach (Recipient recipient in mail.Recipients)
            {
                contactNames.Add(recipient.Name);
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

                    // Build file name - when not in a thread, date is at the end
                    fileName = $"{subjectClean}-{senderClean}-{fileDateTime}.md";
                    filePath = Path.Combine(_settings.GetInboxPath(), fileName);
                    fileNameNoExt = Path.GetFileNameWithoutExtension(fileName);
                    obsidianLinkPath = fileNameNoExt;  // Default to the regular file name for Obsidian link

                    // Get thread info
                    string conversationId = _threadService.GetConversationId(mail);
                    string threadNoteName = _threadService.GetThreadNoteName(mail, subjectClean, senderClean, 
                        _contactService.GetShortName(GetFirstRecipient(mail)));
                    string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                    string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");

                    // Check for existing thread and count emails with same thread ID
                    var threadInfo = _threadService.FindExistingThread(conversationId, _settings.GetInboxPath());
                    bool hasExistingThread = threadInfo.hasExistingThread;
                    string earliestEmailThreadName = threadInfo.earliestEmailThreadName;
                    int emailCount = threadInfo.emailCount;
                    
                    // Debug info to check email count - note that emailCount is how many emails are ALREADY in the thread
                    // We're NOT counting the current email
                    MessageBox.Show($"Thread ID: {conversationId}\nEmail count: {emailCount}\nExisting thread: {hasExistingThread}", "Thread Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Only group this email if there are at least 2 previous emails in the thread
                    // Changed from emailCount >= 1 to emailCount >= 2
                    bool shouldGroupThread = hasExistingThread && _settings.GroupEmailThreads && emailCount >= 2;
                    
                    // If this is part of a thread with at least one previous email and thread grouping is enabled, update paths
                    if (shouldGroupThread)
                    {
                        threadNoteName = earliestEmailThreadName ?? threadNoteName;
                        threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                        threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
                        
                        // For thread files, date goes at the front of the filename
                        fileName = $"{fileDateTime}-{subjectClean}-{senderClean}.md";
                        filePath = Path.Combine(threadFolderPath, fileName);
                        
                        // Update Obsidian link path to include the thread folder
                        // Use the folder name with forward slashes for Obsidian URI compatibility
                        obsidianLinkPath = $"{threadNoteName}/{fileNameNoExt}";
                    }
                    else
                    {
                        // Not grouping in a thread folder - use just the file name
                        obsidianLinkPath = fileNameNoExt;
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
                        { "tags", "email" }
                    };

                    // Add CC if present
                    var ccLinked = _contactService.BuildLinkedNames(mail.Recipients, OlMailRecipientType.olCC);
                    var ccEmails = _contactService.BuildEmailList(mail.Recipients, OlMailRecipientType.olCC);
                    if (ccEmails.Count > 0)
                    {
                        metadata.Add("cc", ccLinked);
                        metadata.Add("ccEmail", ccEmails);
                    }

                    // Add threadNote if this is part of a thread and thread grouping is enabled
                    if (shouldGroupThread)
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
                    if (shouldGroupThread)
                    {
                        await _threadService.UpdateThreadNote(threadFolderPath, threadNotePath, conversationId, threadNoteName, mail);
                    }

                    // Create Outlook task if enabled
                    if (_settings.CreateOutlookTask && _taskService.ShouldCreateTasks)
                    {
                        status.UpdateProgress("Creating Outlook task", 80);
                        await _taskService.CreateOutlookTask(mail);
                    }
                    
                    status.UpdateProgress("Completing email processing", 100);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error processing email: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
            // Process contacts outside the StatusService block
            // This ensures the progress window doesn't block the contact dialog
            if (_settings.EnableContactSaving && contactNames.Count > 0)
            {
                try
                {
                    // Remove duplicates and sort
                    contactNames = contactNames.Distinct().OrderBy(n => n).ToList();
                    
                    // Filter to only new contacts
                    var newContacts = new List<string>();
                    foreach (var contactName in contactNames)
                    {
                        if (!_contactService.ContactExists(contactName))
                        {
                            newContacts.Add(contactName);
                        }
                    }
                    
                    // Only show dialog if we have new contacts to create
                    if (newContacts.Count > 0)
                    {
                        // Show contact confirmation dialog
                        using (var dialog = new ContactConfirmationDialog(newContacts))
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                foreach (var contactName in dialog.SelectedContacts)
                                {
                                    // Create contact note for each selected contact
                                    _contactService.CreateContactNote(contactName);
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error processing contacts: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
            // Launch Obsidian if enabled
            if (_settings.LaunchObsidian)
            {
                try
                {
                    if (_settings.ShowCountdown && _settings.ObsidianDelaySeconds > 0)
                    {
                        using (var countdown = new CountdownForm(_settings.ObsidianDelaySeconds))
                        {
                            countdown.ShowDialog();
                        }
                    }
                    else if (_settings.ObsidianDelaySeconds > 0)
                    {
                        await Task.Delay(_settings.ObsidianDelaySeconds * 1000);
                    }

                    _fileService.LaunchObsidian(_settings.VaultName, obsidianLinkPath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error launching Obsidian: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            // Remove redundant Re_ RE_ prefixes - keep only one "Re_"
            cleaned = Regex.Replace(cleaned, @"(?:Re_\s*)+(?:RE_\s*)+", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:RE_\s*)+(?:Re_\s*)+", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:Re_\s*){2,}", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:RE_\s*){2,}", "Re_", RegexOptions.IgnoreCase);

            return _fileService.CleanFileName(cleaned.Trim());
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