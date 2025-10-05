using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class AppointmentProcessor
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;

        public AppointmentProcessor() { }
        public AppointmentProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, settings);
            _taskService = new TaskService(settings);
            _contactService = new ContactService(_fileService, _templateService);
        }

        public async Task ProcessAppointment(AppointmentItem appointment)
        {
            List<string> contactNames = new List<string>();
            string fileName = string.Empty;
            string fileNameNoExt = string.Empty;
            string filePath = string.Empty;
            string obsidianLinkPath = string.Empty;  // Added to store the path to use for Obsidian
            string conversationId = string.Empty;
            string threadNoteName = string.Empty;
            string threadFolderPath = string.Empty;
            string threadNotePath = string.Empty;
            bool shouldGroupThread = false;

            contactNames.Add(appointment.GetOrganizer().Name);
            foreach( Recipient recipient in appointment.Recipients ) 
            {
                contactNames.Add(recipient.Name);
            }

            using (var status = new StatusService())
            {
                try
                {
                    status.UpdateProgress("Processing appointment...", 0);

                    // Build note title using settings
                    string noteTitle = appointment.Subject;
                    string senderClean = _contactService.GetShortName(appointment.GetOrganizer().Name);
                    string appointmentStartDateTime = appointment.Start.ToString("yyyy-MM-dd-HHmm");
                    string appointmentEndDateTime = appointment.End.ToString("yyyy-MM-dd-HHmm");
                    string dateStr = appointment.Start.ToString("yyyy-MM-dd");
                    string subjectClean = CleanSubject(appointment.Subject);

                    // Use settings for title format
                    string titleFormat = _settings.MeetingNoteTitleFormat ?? "{Date} - {Subject}";
                    int maxLength = _settings.MeetingNoteTitleMaxLength > 0 ? _settings.MeetingNoteTitleMaxLength : 50;

                    // Prepare replacements
                    string formattedTitle = titleFormat
                        .Replace("{Subject}", subjectClean)
                        .Replace("{Sender}", senderClean)
                        .Replace("{Date}", dateStr);

                    // Remove double spaces and trim
                    formattedTitle = Regex.Replace(formattedTitle, @"\s+", " ").Trim();

                    // Remove leading or trailing dashes and whitespace characters if date necessary
                    formattedTitle = Regex.Replace(formattedTitle, @"^[\s-]+|[\s-]+$", "").Trim();

                    // Trim to max length
                    if (formattedTitle.Length > maxLength)
                        formattedTitle = formattedTitle.Substring(0, maxLength - 3) + "...";
                    noteTitle = formattedTitle;

                    status.UpdateProgress("Processing appointment metadata", 50);

                    // Build metadata for frontmatter
                    var metadata = new Dictionary<string, object>
                    {
                        { "title", noteTitle },
                        { "organizer", $"[[{appointment.GetOrganizer().Name}]]" },
                        { "organizerEmail", appointment.GetOrganizer().Address },
                        { "attendees", _contactService.BuildLinkedNames(appointment.Recipients, new[] { OlMeetingRecipientType.olOptional, OlMeetingRecipientType.olRequired } ) },
                        { "attendeesEmail", _contactService.BuildEmailList(appointment.Recipients, new[] { OlMeetingRecipientType.olOptional, OlMeetingRecipientType.olRequired } ) },
                        { "Resources", _contactService.GetMeetingResourceData(appointment.Recipients) },
                        { "startDateTime", appointment.Start.ToString("yyyy-MM-dd HH:mm:ss") },
                        { "endDateTime", appointment.End.ToString("yyyy-MM-dd HH:mm:ss") },
                        { "dailyNoteLink", $"[[{appointment.Start:yyyy-MM-dd}]]" },
                        { "tags", (_settings.MeetingDefaultNoteTags != null && _settings.MeetingDefaultNoteTags.Count > 0) ? new List<string>(_settings.MeetingDefaultNoteTags) : new List<string> { "meeting" } }
                    };

                    //if attachments are to be processed, add metadata references
                    if(_settings.MeetingSaveAttachments)
                    {
                        metadata.Add("HasAttachments", appointment.Attachments.Count > 0 ? "Yes" : "No");
                        var attachmentLinks = new List<string>();
                        foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in appointment.Attachments)
                        {
                            attachmentLinks.Add($"[[{attachment.FileName}]]");
                        }
                        metadata.Add("Attachments", attachmentLinks);
                    }                    
                
                    // Build content
                    var content = new System.Text.StringBuilder();
                    content.Append(_templateService.BuildFrontMatter(metadata));

                    content.Append(appointment.Body);

                    status.UpdateProgress("Writing note file", 75);

                    // Check for duplicate email before writing the note
                    if (IsDuplicate(_settings.GetMeetingsPath(), formattedTitle))
                    {
                        status.UpdateProgress("Duplicate meeting detected. Skipping note creation.", 100);
                        return;
                    }


                    var noteFolder = _settings.GetMeetingsPath();                 
                    

                    fileName = Path.Combine(noteFolder, $"{formattedTitle}.md");
                    var fileNameNoExtResult = Path.GetFileNameWithoutExtension(fileName);
                    obsidianLinkPath = fileNameNoExtResult;

                    if (_settings.MeetingSaveAttachments)
                    {
                        noteFolder = Path.Combine(noteFolder, formattedTitle);
                        obsidianLinkPath = $"{noteFolder}/{fileNameNoExtResult}";
                    }


                    _fileService.EnsureDirectoryExists(noteFolder);
                    _fileService.WriteUtf8File(fileName, content.ToString());

                    status.UpdateProgress("Completing appointment processing", 100);


                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error processing appointment: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    // Ensure delay happens after all file operations
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

                    // Always use the latest obsidianLinkPath (updated after resuffixing)
                    _fileService.LaunchObsidian(_settings.VaultName, obsidianLinkPath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error launching Obsidian: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Checks if an email with the given filename already exists in the meetings folder or any subfolder.
        /// Returns true if a duplicate is found.
        /// </summary>
        private bool IsDuplicate(string meetingFolderPath, string fileName)
        {
            var mdFiles = Directory.GetFiles(meetingFolderPath, "*.md", SearchOption.AllDirectories);
            
            return mdFiles.Any(f=>Path.GetFileNameWithoutExtension(f) == fileName);
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

            // Replace colons (with or without spaces) with underscores
            cleaned = Regex.Replace(cleaned, @":\s*", "_", RegexOptions.IgnoreCase);

            // Handle Re_ (Reply) prefixes
            // Remove redundant Re_ RE_ prefixes - keep only one "Re_"
            cleaned = Regex.Replace(cleaned, @"(?:Re_\s*)+(?:RE_\s*)+", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:RE_\s*)+(?:Re_\s*)+", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:Re_\s*){2,}", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:RE_\s*){2,}", "Re_", RegexOptions.IgnoreCase);

            // Handle Fw_ (Forward) prefixes
            // Remove redundant Fw_ FW_ prefixes - keep only one "Fw_"
            cleaned = Regex.Replace(cleaned, @"(?:Fw_\s*)+(?:FW_\s*)+", "Fw_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:FW_\s*)+(?:Fw_\s*)+", "Fw_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:Fw_\s*){2,}", "Fw_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"(?:FW_\s*){2,}", "Fw_", RegexOptions.IgnoreCase);

            // Ensure there are no spaces after prefixes
            cleaned = Regex.Replace(cleaned, @"Re_\s+", "Re_", RegexOptions.IgnoreCase);
            cleaned = Regex.Replace(cleaned, @"Fw_\s+", "Fw_", RegexOptions.IgnoreCase);

            return _fileService.CleanFileName(cleaned.Trim());
        }
    }
}
