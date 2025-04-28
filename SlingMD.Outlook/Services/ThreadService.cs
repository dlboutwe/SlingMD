using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class ThreadService
    {
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ObsidianSettings _settings;

        public ThreadService(FileService fileService, TemplateService templateService, ObsidianSettings settings)
        {
            _fileService = fileService;
            _templateService = templateService;
            _settings = settings;
        }

        public string GetConversationId(MailItem mail)
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
                    return BitConverter.ToString(MD5.Create()
                        .ComputeHash(Encoding.UTF8.GetBytes(normalizedSubject)))
                        .Replace("-", "").Substring(0, 16);
                }

                // Try to get the conversation index using PR_CONVERSATION_INDEX property
                const string PR_CONVERSATION_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x0071001F";
                byte[] conversationIndex = (byte[])mail.PropertyAccessor.GetProperty(PR_CONVERSATION_INDEX);
                
                if (conversationIndex != null && conversationIndex.Length >= 22)
                {
                    // The first 22 bytes of the conversation index identify the thread
                    return BitConverter.ToString(conversationIndex.Take(22).ToArray())
                        .Replace("-", "").Substring(0, 16);
                }

                // If both methods fail, use the normalized subject as last resort
                string subject = mail.Subject;
                subject = Regex.Replace(subject, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", "", RegexOptions.IgnoreCase);
                subject = Regex.Replace(subject, @"^Re:\s+", "", RegexOptions.IgnoreCase);
                return BitConverter.ToString(MD5.Create()
                    .ComputeHash(Encoding.UTF8.GetBytes(subject)))
                    .Replace("-", "").Substring(0, 16);
            }
            catch
            {
                return Guid.NewGuid().ToString("N").Substring(0, 16);
            }
        }

        public string GetThreadNoteName(MailItem mail, string cleanSubject, string firstSender, string firstRecipient)
        {
            string threadSubject;
            
            // Use ConversationTopic if available as it's typically cleaner
            threadSubject = !string.IsNullOrEmpty(mail.ConversationTopic) 
                ? mail.ConversationTopic 
                : mail.Subject;

            // Clean the subject using FileService which uses the patterns from settings
            threadSubject = _fileService.CleanFileName(threadSubject);

            // Truncate if too long
            if (threadSubject.Length > 50)
            {
                threadSubject = threadSubject.Substring(0, 47) + "...";
            }
            
            // Clean sender and recipient names
            firstSender = _fileService.CleanFileName(firstSender);
            firstRecipient = _fileService.CleanFileName(firstRecipient);
            
            // Ensure no space after 0- and consistent separators
            return $"0-{threadSubject.Trim()}-{firstSender}-{firstRecipient}".Replace("--", "-");
        }

        public async Task UpdateThreadNote(string threadFolderPath, string threadNotePath, string conversationId, string threadNoteName, MailItem mail)
        {
            var templateContent = _templateService.LoadTemplate("ThreadNoteTemplate.md") ?? 
                                _templateService.GetDefaultThreadNoteTemplate();

            // Clean the title using the same method as thread name
            string threadTitle = mail.ConversationTopic ?? mail.Subject;
            threadTitle = _fileService.CleanFileName(threadTitle);

            var replacements = new Dictionary<string, string>
            {
                { "title", threadTitle },
                { "threadId", conversationId }
            };

            string content = _templateService.ProcessTemplate(templateContent, replacements);
            _fileService.WriteUtf8File(threadNotePath, content);
        }

        public string MoveToThreadFolder(string emailPath, string threadFolderPath)
        {
            string fileName = Path.GetFileName(emailPath);
            string newFileName = fileName;
            // Add email id as a temporary suffix to avoid overwriting
            string emailId = null;
            bool inFrontMatter = false;
            foreach (var line in File.ReadLines(emailPath))
            {
                if (line.Trim() == "---")
                {
                    if (!inFrontMatter) { inFrontMatter = true; continue; }
                    else break;
                }
                if (inFrontMatter && line.Trim().StartsWith("internetMessageId:", StringComparison.OrdinalIgnoreCase))
                {
                    emailId = line.Trim().Substring("internetMessageId:".Length).Trim().Trim('"');
                    break;
                }
                if (inFrontMatter && line.Trim().StartsWith("entryId:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(emailId))
                {
                    emailId = line.Trim().Substring("entryId:".Length).Trim().Trim('"');
                }
            }
            if (!string.IsNullOrEmpty(emailId))
            {
                // Sanitize emailId for filename (alphanumeric only)
                string safeId = new string(emailId.Where(char.IsLetterOrDigit).ToArray());
                string nameNoExt = Path.GetFileNameWithoutExtension(fileName);
                string ext = Path.GetExtension(fileName);
                newFileName = $"{nameNoExt}-eid{safeId}{ext}";
            }
            string threadPath = Path.Combine(threadFolderPath, newFileName);
            _fileService.EnsureDirectoryExists(threadFolderPath);
            if (File.Exists(threadPath))
            {
                File.Delete(threadPath);
            }
            File.Move(emailPath, threadPath);
            return threadPath;
        }

        public (bool hasExistingThread, string earliestEmailThreadName, DateTime? earliestEmailDate, int emailCount) 
            FindExistingThread(string conversationId, string inboxPath)
        {
            bool hasExistingThread = false;
            DateTime? earliestEmailDate = null;
            string earliestEmailThreadName = null;
            int emailCount = 0;
            List<string> matchingFiles = new List<string>(); // Track matching files for debugging

            try 
            {
                // Get all markdown files from the inbox and subfolders
                var files = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
                
                // Search through each file for the thread ID
                foreach (var file in files)
                {
                    try
                    {
                        string emailContent = File.ReadAllText(file);
                        var threadIdMatch = Regex.Match(emailContent, @"threadId: ""([^""]+)""");
                        
                        // If this file belongs to the conversation thread
                        if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                        {
                            hasExistingThread = true;
                            emailCount++; // Increment email count for each matching email
                            matchingFiles.Add(file); // Add to our debugging list
                            
                            // Parse the date to find the earliest email
                            var dateMatch = Regex.Match(emailContent, @"date: (\d{4}-\d{2}-\d{2} \d{2}:\d{2})");
                            if (dateMatch.Success)
                            {
                                DateTime emailDate;
                                if (DateTime.TryParseExact(dateMatch.Groups[1].Value, "yyyy-MM-dd HH:mm", null, 
                                                         System.Globalization.DateTimeStyles.None, out emailDate))
                                {
                                    if (!earliestEmailDate.HasValue || emailDate < earliestEmailDate.Value)
                                    {
                                        earliestEmailDate = emailDate;
                                        
                                        // Check if this email is in a thread folder
                                        string directory = Path.GetDirectoryName(file);
                                        if (directory != inboxPath)
                                        {
                                            earliestEmailThreadName = Path.GetFileName(directory);
                                        }
                                        else
                                        {
                                            // Try to extract thread name components from frontmatter
                                            var subjectMatch = Regex.Match(emailContent, @"title: ""([^""]+)""");
                                            var fromMatch = Regex.Match(emailContent, @"from: ""[^""]*\[\[([^""]+)\]\]""");
                                            var toMatch = Regex.Match(emailContent, @"to:.*?\n\s*- ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Singleline);
                                            
                                            if (subjectMatch.Success && fromMatch.Success && toMatch.Success)
                                            {
                                                string subject = _fileService.CleanFileName(subjectMatch.Groups[1].Value);
                                                if (subject.Length > 50)
                                                {
                                                    subject = subject.Substring(0, 47) + "...";
                                                }
                                                string sender = fromMatch.Groups[1].Value;
                                                string recipient = toMatch.Groups[1].Value;
                                                earliestEmailThreadName = $"{subject}-{sender}-{recipient}";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (System.Exception)
                    {
                        // Skip files that can't be read
                        continue;
                    }
                }
                
                // If we found any matches, show more details for debugging
                if (emailCount > 0 && _settings.ShowThreadDebug)
                {
                    string filesList = string.Join("\n", matchingFiles);
                    System.Windows.Forms.MessageBox.Show(
                        $"Found {emailCount} existing emails with thread ID: {conversationId}\n\nFiles:\n{filesList}",
                        "Thread Match Details",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error searching for thread: {ex.Message}", "Thread Search Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }

            return (hasExistingThread, earliestEmailThreadName, earliestEmailDate, emailCount);
        }

        /// <summary>
        /// Renames all thread notes in the given thread folder with an incrementing suffix based on their date in front matter.
        /// Skips the thread summary note (0-...).
        /// Removes any email id suffix before applying the chronological suffix.
        /// </summary>
        public void ResuffixThreadNotes(string threadFolderPath, string baseName)
        {
            if (!Directory.Exists(threadFolderPath)) return;
            var files = Directory.GetFiles(threadFolderPath, baseName + "*.md", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("0-"))
                .ToList();
            var fileDates = new List<(string file, DateTime date)>();
            foreach (var file in files)
            {
                DateTime? date = null;
                bool inFrontMatter = false;
                foreach (var line in File.ReadLines(file))
                {
                    if (line.Trim() == "---")
                    {
                        if (!inFrontMatter) { inFrontMatter = true; continue; }
                        else break;
                    }
                    if (inFrontMatter && line.Trim().StartsWith("date:", StringComparison.OrdinalIgnoreCase))
                    {
                        var value = line.Trim().Substring("date:".Length).Trim().Trim('"');
                        if (DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out var parsed))
                            date = parsed;
                        else if (DateTime.TryParse(value, out var fallback))
                            date = fallback;
                        break;
                    }
                }
                if (date.HasValue)
                {
                    fileDates.Add((file, date.Value));
                }
            }
            fileDates = fileDates.OrderBy(fd => fd.date).ToList();
            int idx = 1;
            foreach (var fd in fileDates)
            {
                // Remove any email id suffix before applying the new suffix
                string nameNoExt = Path.GetFileNameWithoutExtension(fd.file);
                string ext = Path.GetExtension(fd.file);
                // Remove trailing -eid{id} if present (id is alphanumeric, usually long)
                nameNoExt = System.Text.RegularExpressions.Regex.Replace(nameNoExt, "-eid[a-zA-Z0-9]+$", "");
                string newName = $"{baseName}-{idx:D3}{ext}";
                string newPath = Path.Combine(threadFolderPath, newName);
                if (!fd.file.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (File.Exists(newPath)) File.Delete(newPath);
                    File.Move(fd.file, newPath);
                }
                idx++;
            }
        }
    }
} 