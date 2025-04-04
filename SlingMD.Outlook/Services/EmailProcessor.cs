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
                    string subjectClean = CleanSubject(mail.Subject);
                    if (subjectClean.Length > 50)  // Limit subject length
                    {
                        subjectClean = subjectClean.Substring(0, 47) + "...";
                    }
                    string senderClean = GetShortName(mail.SenderName);
                    string fileDateTime = mail.ReceivedTime.ToString("yyyy-MM-dd-HHmm");
                    
                    status.UpdateProgress("Creating note file", 25);

                    // Build file name with date and time prepended
                    string fileName = $"{fileDateTime}-{subjectClean}-{senderClean}.md";
                    string filePath = Path.Combine(_settings.GetInboxPath(), fileName);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(fileName);

                    // Initialize StringBuilder for frontmatter
                    var frontmatter = new StringBuilder();

                    // Get thread info early
                    string conversationId = GetConversationId(mail);
                    string threadNoteName = GetThreadNoteName(mail);
                    string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                    string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");

                    // Check if there are other emails in this thread and get existing thread folder if any
                    bool hasExistingThread = false;
                    string existingThreadFolder = null;
                    DateTime? earliestEmailDate = null;
                    string earliestEmailThreadName = null;

                    // Only check for thread grouping if the setting is enabled
                    if (_settings.GroupEmailThreads)
                    {
                        var files = Directory.GetFiles(_settings.GetInboxPath(), "*.md", SearchOption.AllDirectories);
                        foreach (var file in files)
                        {
                            if (file.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                                continue;

                            string emailContent = File.ReadAllText(file);
                            var threadIdMatch = Regex.Match(emailContent, @"threadId: ""([^""]+)""");
                            if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                            {
                                hasExistingThread = true;

                                // Get the date from the file content
                                var dateMatch = Regex.Match(emailContent, @"date: (\d{4}-\d{2}-\d{2} \d{2}:\d{2})");
                                if (dateMatch.Success)
                                {
                                    var emailDate = DateTime.ParseExact(dateMatch.Groups[1].Value, "yyyy-MM-dd HH:mm", null);
                                    if (!earliestEmailDate.HasValue || emailDate < earliestEmailDate.Value)
                                    {
                                        earliestEmailDate = emailDate;
                                        // Get thread name from this email
                                        string directory = Path.GetDirectoryName(file);
                                        if (directory != _settings.GetInboxPath())
                                        {
                                            earliestEmailThreadName = Path.GetFileName(directory);
                                        }
                                        else
                                        {
                                            // If the earliest email is not in a thread folder yet,
                                            // generate its thread name using its subject and recipients
                                            var subjectMatch = Regex.Match(emailContent, @"title: ""([^""]+)""");
                                            var fromMatch = Regex.Match(emailContent, @"from: ""[^""]*\[\[([^""]+)\]\]""");
                                            var toMatch = Regex.Match(emailContent, @"to:.*?\n\s*- ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Singleline);
                                            
                                            if (subjectMatch.Success && fromMatch.Success && toMatch.Success)
                                            {
                                                string subject = CleanSubject(subjectMatch.Groups[1].Value);
                                                if (subject.Length > 50)
                                                {
                                                    subject = subject.Substring(0, 47) + "...";
                                                }
                                                string sender = GetShortName(fromMatch.Groups[1].Value);
                                                string recipient = GetShortName(toMatch.Groups[1].Value);
                                                earliestEmailThreadName = $"{subject}-{sender}-{recipient}";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // If we found an existing thread, use the earliest email's thread name
                        if (hasExistingThread && !string.IsNullOrEmpty(earliestEmailThreadName))
                        {
                            threadNoteName = earliestEmailThreadName;
                            threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                        }
                    }

                    // If this is part of a thread and thread grouping is enabled, update the file path to be in the thread folder
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        Directory.CreateDirectory(threadFolderPath);
                        filePath = Path.Combine(threadFolderPath, fileName);
                    }

                    status.UpdateProgress("Processing email metadata", 50);

                    // Build YAML frontmatter
                    frontmatter.AppendLine("---");
                    frontmatter.AppendLine($"title: \"{mail.Subject}\"");
                    frontmatter.AppendLine($"from: \"[[{mail.SenderName}]]\"");
                    frontmatter.AppendLine($"fromEmail: \"{GetSenderEmail(mail)}\"");
                    frontmatter.AppendLine("to:");
                    frontmatter.AppendLine(BuildLinkedNames(mail.Recipients, OlMailRecipientType.olTo));
                    frontmatter.AppendLine("toEmail:");
                    frontmatter.AppendLine(BuildEmailList(mail.Recipients, OlMailRecipientType.olTo));

                    // Always add the thread ID
                    frontmatter.AppendLine($"threadId: \"{conversationId}\"");

                    // Add CC if present
                    string ccLinked = BuildLinkedNames(mail.Recipients, OlMailRecipientType.olCC);
                    string ccEmails = BuildEmailList(mail.Recipients, OlMailRecipientType.olCC);
                    if (!string.IsNullOrEmpty(ccEmails))
                    {
                        frontmatter.AppendLine("cc:");
                        frontmatter.AppendLine(ccLinked);
                        frontmatter.AppendLine("ccEmail:");
                        frontmatter.AppendLine(ccEmails);
                    }

                    frontmatter.AppendLine($"date: {mail.ReceivedTime:yyyy-MM-dd HH:mm}");
                    frontmatter.AppendLine($"dailyNoteLink: \"[[{mail.ReceivedTime:yyyy-MM-dd}]]\"");
                    frontmatter.AppendLine("tags: [email]");
                    
                    // Add threadNote if this is part of a thread and thread grouping is enabled
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        frontmatter.AppendLine($"threadNote: \"[[0-{threadNoteName}]]\"");
                    }
                    
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

                    // If this is part of a thread and thread grouping is enabled, move any existing related emails into the thread folder
                    if (hasExistingThread && _settings.GroupEmailThreads)
                    {
                        var files = Directory.GetFiles(_settings.GetInboxPath(), "*.md", SearchOption.AllDirectories);
                        foreach (var file in files)
                        {
                            if (file.StartsWith(threadFolderPath, StringComparison.OrdinalIgnoreCase))
                                continue; // Skip files already in thread folder

                            // Read with UTF-8 encoding
                            string emailContent;
                            using (var reader = new StreamReader(file, encoding: new UTF8Encoding(false)))
                            {
                                emailContent = reader.ReadToEnd();
                            }
                            
                            var threadIdMatch = Regex.Match(emailContent, @"threadId: ""([^""]+)""");
                            if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                            {
                                // Get the date from the file content
                                var dateMatch = Regex.Match(emailContent, @"date: (\d{4}-\d{2}-\d{2} \d{2}:\d{2})");
                                string emailDateTime = dateMatch.Success 
                                    ? DateTime.ParseExact(dateMatch.Groups[1].Value, "yyyy-MM-dd HH:mm", null).ToString("yyyy-MM-dd-HHmm")
                                    : DateTime.Now.ToString("yyyy-MM-dd-HHmm");
                                
                                // Create new file name with date and time prepended
                                string oldFileName = Path.GetFileName(file);
                                string newFileName;
                                string newFilePath;
                                
                                // If filename already starts with a date-time pattern, use it as is
                                if (Regex.IsMatch(oldFileName, @"^\d{4}-\d{2}-\d{2}-\d{4}"))
                                {
                                    newFileName = oldFileName;
                                    newFilePath = Path.Combine(threadFolderPath, newFileName);
                                }
                                else
                                {
                                    newFileName = $"{emailDateTime}-{oldFileName.Substring(0, oldFileName.LastIndexOf(" - "))}.md";
                                    newFilePath = Path.Combine(threadFolderPath, newFileName);
                                }

                                // Add threadNote to frontmatter if not present
                                if (!emailContent.Contains($"threadNote: \"[[0-{threadNoteName}]]\""))
                                {
                                    var lines = emailContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).ToList();
                                    int frontmatterEnd = lines.FindIndex(1, line => line == "---"); // Find second ---
                                    if (frontmatterEnd > 0)
                                    {
                                        // Insert threadNote just before the closing ---
                                        lines.Insert(frontmatterEnd, $"threadNote: \"[[0-{threadNoteName}]]\"");
                                        emailContent = string.Join(Environment.NewLine, lines);
                                    }
                                    else
                                    {
                                        // If we can't find the end of frontmatter, try to add it after the tags line
                                        int tagsLine = lines.FindIndex(line => line.StartsWith("tags:"));
                                        if (tagsLine > 0)
                                        {
                                            lines.Insert(tagsLine + 1, $"threadNote: \"[[0-{threadNoteName}]]\"");
                                            emailContent = string.Join(Environment.NewLine, lines);
                                        }
                                    }
                                }

                                // Write the file with UTF-8 encoding
                                WriteUtf8File(newFilePath, emailContent);
                                
                                // Delete the old file
                                File.Delete(file);
                            }
                        }

                        // Create or update thread note
                        await UpdateThreadNote(mail, conversationId, threadNoteName, fileName);
                    }

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
                    names.Add($"  - \"[[{recipient.Name}]]\"");
                }
            }
            return $"\n{string.Join("\n", names)}";
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
                            emails.Add($"  - \"{email}\"");
                        }
                    }
                    catch
                    {
                        // Fallback to Address property
                        if (!string.IsNullOrEmpty(recipient.Address))
                        {
                            emails.Add($"  - \"{recipient.Address}\"");
                        }
                    }
                }
            }
            return $"\n{string.Join("\n", emails)}";
        }

        private string CleanFileName(string input)
        {
            return FileHelper.CleanFileName(input);
        }

        private void WriteUtf8File(string filePath, string content)
        {
            FileHelper.WriteUtf8File(filePath, content);
        }

        private void LaunchObsidian(string vaultName, string filePath)
        {
            FileHelper.LaunchObsidian(vaultName, filePath);
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
            return CleanFileName(cleaned);
        }

        private string GetThreadNoteName(MailItem mail)
        {
            // Get clean base subject
            string baseSubject = mail.ConversationTopic ?? mail.Subject;
            string cleanSubject = CleanSubject(baseSubject);
            
            if (cleanSubject.Length > 50)  // Limit subject length
            {
                cleanSubject = cleanSubject.Substring(0, 47) + "...";
            }
            
            // Get first sender and recipient initials or short names
            string firstSender = GetShortName(mail.SenderName);
            string firstRecipient = "";
            foreach (Recipient recipient in mail.Recipients)
            {
                if (recipient.Type == (int)OlMailRecipientType.olTo)
                {
                    firstRecipient = GetShortName(recipient.Name);
                    break;
                }
            }
            
            return $"{cleanSubject}-{firstSender}-{firstRecipient}";
        }

        private string GetShortName(string fullName)
        {
            // Clean the name first
            string cleanName = CleanFileName(fullName);
            
            // If name contains parentheses, take what's before them
            int parenIndex = cleanName.IndexOf('(');
            if (parenIndex > 0)
            {
                cleanName = cleanName.Substring(0, parenIndex).Trim();
            }

            // Split into parts
            string[] parts = cleanName.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (parts.Length == 0) return "Unknown";
            if (parts.Length == 1) return parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            
            // For multiple parts, use first name and last name initial
            string firstName = parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            string lastInitial = parts[parts.Length - 1].Substring(0, 1).ToUpper();
            return $"{firstName}{lastInitial}";
        }

        private string ProcessTemplate(string templateContent, Dictionary<string, string> replacements)
        {
            string result = templateContent;
            foreach (var replacement in replacements)
            {
                result = result.Replace($"{{{{{replacement.Key}}}}}", replacement.Value);
            }
            return result;
        }

        private async Task UpdateThreadNote(MailItem mail, string conversationId, string threadNoteName, string currentEmailFile)
        {
            // Initialize with default values
            string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
            string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");

            // First check if a thread note already exists for this conversation
            var threadNotes = Directory.GetFiles(_settings.GetInboxPath(), "0-*.md", SearchOption.AllDirectories);
            string existingThreadNote = null;
            
            foreach (var note in threadNotes)
            {
                string noteContent = File.ReadAllText(note);
                var threadIdMatch = Regex.Match(noteContent, @"threadId: ""([^""]+)""");
                if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                {
                    existingThreadNote = note;
                    // Update threadNoteName and paths to use the existing note's location
                    threadNoteName = Path.GetFileName(Path.GetDirectoryName(note));
                    threadFolderPath = Path.GetDirectoryName(note);
                    threadNotePath = note;
                    break;
                }
            }

            // If no existing thread note was found, create a new one
            if (existingThreadNote == null)
            {
                threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
                Directory.CreateDirectory(threadFolderPath);
            }
            
            // Try multiple locations for the template file
            string[] possibleTemplatePaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ThreadNoteTemplate.md"),
                Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Templates", "ThreadNoteTemplate.md"),
                Path.Combine(Directory.GetCurrentDirectory(), "Templates", "ThreadNoteTemplate.md"),
                Path.Combine(Environment.CurrentDirectory, "Templates", "ThreadNoteTemplate.md")
            };

            string templateContent = null;
            string foundPath = null;

            foreach (var path in possibleTemplatePaths)
            {
                if (File.Exists(path))
                {
                    templateContent = File.ReadAllText(path);
                    foundPath = path;
                    break;
                }
            }

            if (templateContent == null)
            {
                // If template file not found, use embedded template
                templateContent = @"---
title: ""{{title}}""
type: email-thread
threadId: ""{{threadId}}""
tags: [email-thread]
---

# {{title}}

```dataviewjs
// Get all emails with matching threadId from current folder
const threadId = ""{{threadId}}"";
const emails = dv.pages("""")
    .where(p => p.threadId === threadId && p.file.name !== dv.current().file.name)
    .sort(p => p.date, 'desc');

// Display thread summary
if (emails.length > 0) {
    const startDate = emails[emails.length-1].date;
    const latestDate = emails[0].date;
    const participants = new Set();
    emails.forEach(e => {
        // Handle from field
        if (e.from) {
            const fromName = String(e.from).match(/\[\[(.*?)\]\]/)?.[1];
            if (fromName) participants.add(fromName);
        }

        // Handle to field
        if (e.to) {
            const toList = Array.isArray(e.to) ? e.to : [e.to];
            toList.forEach(to => {
                const name = String(to).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }

        // Handle cc field
        if (e.cc) {
            const ccList = Array.isArray(e.cc) ? e.cc : [e.cc];
            ccList.forEach(cc => {
                const name = String(cc).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }
    });

    dv.header(2, 'Thread Summary');
    dv.list([
        `Started: ${startDate}`,
        `Latest: ${latestDate}`,
        `Messages: ${emails.length}`,
        `Participants: ${Array.from(participants).map(p => `[[${p}]]`).join(', ')}`
    ]);
}

// Display email timeline
dv.header(2, 'Email Timeline');
for (const email of emails) {
    dv.header(3, `${email.file.name} - ${email.date}`);
    dv.paragraph(`![[${email.file.name}]]`);
}
```";
            }
            
            // Prepare replacements
            var replacements = new Dictionary<string, string>
            {
                { "title", mail.ConversationTopic ?? mail.Subject },
                { "threadId", conversationId }
            };
            
            // Process the template
            string content = ProcessTemplate(templateContent, replacements);
            
            // Write thread note
            WriteUtf8File(threadNotePath, content);
        }

        private string MoveToThreadFolder(string emailPath, string threadFolderPath)
        {
            string fileName = Path.GetFileName(emailPath);
            string threadPath = Path.Combine(threadFolderPath, fileName);
            
            if (!Directory.Exists(threadFolderPath))
            {
                Directory.CreateDirectory(threadFolderPath);
            }

            if (File.Exists(threadPath))
            {
                File.Delete(threadPath);
            }

            File.Move(emailPath, threadPath);
            return threadPath;
        }

        private string BuildFrontMatter(string subject, string from, string fromEmail, List<string> to, List<string> toEmail, 
            string threadId, List<string> cc, List<string> ccEmail, DateTime date, string threadNote)
        {
            var frontMatter = new StringBuilder();
            frontMatter.AppendLine("---");
            frontMatter.AppendLine($"title: {subject}");
            frontMatter.AppendLine($"from: \"[[{from}]]\"");
            frontMatter.AppendLine($"fromEmail: \"{fromEmail}\"");

            // Handle 'to' fields
            frontMatter.AppendLine("to:");
            foreach (var person in to)
            {
                frontMatter.AppendLine($"  - [[{person}]]");
            }
            frontMatter.AppendLine("toEmail:");
            foreach (var email in toEmail)
            {
                frontMatter.AppendLine($"  - {email}");
            }

            frontMatter.AppendLine($"threadId: {threadId}");

            // Handle 'cc' fields if present
            if (cc != null && cc.Any())
            {
                frontMatter.AppendLine("cc:");
                foreach (var person in cc)
                {
                    frontMatter.AppendLine($"  - [[{person}]]");
                }
                frontMatter.AppendLine("ccEmail:");
                foreach (var email in ccEmail)
                {
                    frontMatter.AppendLine($"  - {email}");
                }
            }

            frontMatter.AppendLine($"date: {date:yyyy-MM-dd HH:mm}");
            frontMatter.AppendLine($"dailyNoteLink: [[{date:yyyy-MM-dd}]]");
            frontMatter.AppendLine("tags: [email]");
            if (!string.IsNullOrEmpty(threadNote))
            {
                frontMatter.AppendLine($"threadNote: [[{threadNote}]]");
            }
            frontMatter.AppendLine("---");
            frontMatter.AppendLine();

            return frontMatter.ToString();
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

        private string GetConversationId(MailItem mail)
        {
            try
            {
                // Try to get the conversation topic first as it's most reliable for threading
                if (!string.IsNullOrEmpty(mail.ConversationTopic))
                {
                    string normalizedSubject = mail.ConversationTopic;
                    // Remove all variations of Re, Fwd, etc. and [EXTERNAL] tags
                    normalizedSubject = Regex.Replace(normalizedSubject, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", "", RegexOptions.IgnoreCase);
                    // Also remove any "Re:" that might appear after [EXTERNAL]
                    normalizedSubject = Regex.Replace(normalizedSubject, @"^Re:\s+", "", RegexOptions.IgnoreCase);
                    return BitConverter.ToString(System.Security.Cryptography.MD5.Create()
                        .ComputeHash(Encoding.UTF8.GetBytes(normalizedSubject)))
                        .Replace("-", "").Substring(0, 16);
                }

                // Try to get the conversation index using PR_CONVERSATION_INDEX property
                const string PR_CONVERSATION_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x0071001F";
                byte[] conversationIndex = (byte[])mail.PropertyAccessor.GetProperty(PR_CONVERSATION_INDEX);
                
                if (conversationIndex != null && conversationIndex.Length >= 22)
                {
                    // The first 22 bytes of the conversation index identify the thread
                    // Convert to a readable string format
                    return BitConverter.ToString(conversationIndex.Take(22).ToArray())
                        .Replace("-", "").Substring(0, 16);
                }

                // If both methods fail, use the normalized subject as last resort
                string subject = mail.Subject;
                subject = Regex.Replace(subject, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", "", RegexOptions.IgnoreCase);
                subject = Regex.Replace(subject, @"^Re:\s+", "", RegexOptions.IgnoreCase);
                return BitConverter.ToString(System.Security.Cryptography.MD5.Create()
                    .ComputeHash(Encoding.UTF8.GetBytes(subject)))
                    .Replace("-", "").Substring(0, 16);
            }
            catch
            {
                return Guid.NewGuid().ToString("N").Substring(0, 16);
            }
        }
    }
}