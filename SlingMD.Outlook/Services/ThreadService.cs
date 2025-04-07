using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SlingMD.Outlook.Services
{
    public class ThreadService
    {
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;

        public ThreadService(FileService fileService, TemplateService templateService)
        {
            _fileService = fileService;
            _templateService = templateService;
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
            if (cleanSubject.Length > 50)
            {
                cleanSubject = cleanSubject.Substring(0, 47) + "...";
            }
            
            return $"{cleanSubject}-{firstSender}-{firstRecipient}";
        }

        public async Task UpdateThreadNote(string threadFolderPath, string threadNotePath, string conversationId, string threadNoteName, MailItem mail)
        {
            var templateContent = _templateService.LoadTemplate("ThreadNoteTemplate.md") ?? 
                                _templateService.GetDefaultThreadNoteTemplate();

            var replacements = new Dictionary<string, string>
            {
                { "title", mail.ConversationTopic ?? mail.Subject },
                { "threadId", conversationId }
            };

            string content = _templateService.ProcessTemplate(templateContent, replacements);
            _fileService.WriteUtf8File(threadNotePath, content);
        }

        public string MoveToThreadFolder(string emailPath, string threadFolderPath)
        {
            string fileName = Path.GetFileName(emailPath);
            string threadPath = Path.Combine(threadFolderPath, fileName);
            
            _fileService.EnsureDirectoryExists(threadFolderPath);

            if (File.Exists(threadPath))
            {
                File.Delete(threadPath);
            }

            File.Move(emailPath, threadPath);
            return threadPath;
        }

        public (bool hasExistingThread, string earliestEmailThreadName, DateTime? earliestEmailDate) 
            FindExistingThread(string conversationId, string inboxPath)
        {
            bool hasExistingThread = false;
            DateTime? earliestEmailDate = null;
            string earliestEmailThreadName = null;

            var files = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                string emailContent = File.ReadAllText(file);
                var threadIdMatch = Regex.Match(emailContent, @"threadId: ""([^""]+)""");
                if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                {
                    hasExistingThread = true;

                    var dateMatch = Regex.Match(emailContent, @"date: (\d{4}-\d{2}-\d{2} \d{2}:\d{2})");
                    if (dateMatch.Success)
                    {
                        var emailDate = DateTime.ParseExact(dateMatch.Groups[1].Value, "yyyy-MM-dd HH:mm", null);
                        if (!earliestEmailDate.HasValue || emailDate < earliestEmailDate.Value)
                        {
                            earliestEmailDate = emailDate;
                            string directory = Path.GetDirectoryName(file);
                            if (directory != inboxPath)
                            {
                                earliestEmailThreadName = Path.GetFileName(directory);
                            }
                            else
                            {
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

            return (hasExistingThread, earliestEmailThreadName, earliestEmailDate);
        }
    }
} 