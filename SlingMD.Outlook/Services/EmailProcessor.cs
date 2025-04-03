using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class EmailProcessor
    {
        private readonly ObsidianSettings _settings;

        public EmailProcessor(ObsidianSettings settings)
        {
            _settings = settings;
        }

        public async Task ProcessEmail(MailItem mail)
        {
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

                    // Add task line
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                    frontmatter.AppendLine($"- [ ] [[{fileNameNoExt}]] #FollowUp ➕ {currentDate} 📅 {currentDate}");
                    frontmatter.AppendLine();

                    status.UpdateProgress("Writing note file", 75);

                    // Combine content and write file
                    string content = frontmatter.ToString() + mail.Body;
                    WriteUtf8File(filePath, content);

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
    }
} 