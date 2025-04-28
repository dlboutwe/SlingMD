using System;
using System.Configuration;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SlingMD.Outlook.Models
{
    public class ObsidianSettings
    {
        public string VaultName { get; set; } = "Logic";
        public string VaultBasePath { get; set; } = @"C:\Users\CalebBennett\Documents\Notes\";
        public string InboxFolder { get; set; } = "Inbox";
        public string ContactsFolder { get; set; } = "Contacts";
        public bool EnableContactSaving { get; set; } = true;
        public bool SearchEntireVaultForContacts { get; set; } = false;
        public bool LaunchObsidian { get; set; } = true;
        public int ObsidianDelaySeconds { get; set; } = 1;
        public bool ShowCountdown { get; set; } = true;
        public bool CreateObsidianTask { get; set; } = true;
        public bool CreateOutlookTask { get; set; } = false;
        public int DefaultDueDays { get; set; } = 1;  // Due tomorrow
        /// <summary>
        /// If true, DefaultReminderDays represents days before the due date.
        /// If false, DefaultReminderDays represents days from now (absolute).
        /// </summary>
        public bool UseRelativeReminder { get; set; } = false;
        /// <summary>
        /// Gets or sets the number of days for the reminder.
        /// If UseRelativeReminder is true: represents days before the due date
        /// If UseRelativeReminder is false: represents days from now (absolute)
        /// </summary>
        public int DefaultReminderDays { get; set; } = 0;  // Remind today
        public int DefaultReminderHour { get; set; } = 9;  // at 9am
        public bool AskForDates { get; set; } = false;
        public bool GroupEmailThreads { get; set; } = true;
        public bool ShowDevelopmentSettings { get; set; } = false;
        public bool ShowThreadDebug { get; set; } = false;
        /// <summary>
        /// Default tags to apply to the note's frontmatter.
        /// </summary>
        public List<string> DefaultNoteTags { get; set; } = new List<string> { "FollowUp" };
        /// <summary>
        /// Default tags to apply to the Obsidian task (in the note body).
        /// </summary>
        public List<string> DefaultTaskTags { get; set; } = new List<string> { "FollowUp" };
        /// <summary>
        /// Format for the note title. Use placeholders: {Subject}, {Sender}, {Date}.
        /// </summary>
        public string NoteTitleFormat { get; set; } = "{Subject} - {Date}";
        /// <summary>
        /// Maximum length for the note title. Titles longer than this will be trimmed with ellipsis.
        /// </summary>
        public int NoteTitleMaxLength { get; set; } = 50;
        /// <summary>
        /// Whether to include the date in the note title.
        /// </summary>
        public bool NoteTitleIncludeDate { get; set; } = true;

        public List<string> SubjectCleanupPatterns { get; set; } = new List<string>
        {
            // Remove all variations of Re/Fwd prefixes, including multiple occurrences
            @"^(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])*",  // Matches one or more prefixes at start
            @"(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])+",   // Matches prefixes anywhere in string
            // Common email tags
            @"\[EXTERNAL\]\s*",             // External email tags
            @"\[Internal\]\s*",             // Internal email tags
            @"\[Confidential\]\s*",         // Confidential tags
            @"\[Secure\]\s*",               // Secure email tags
            @"\[Sensitive\]\s*",            // Sensitive email tags
            @"\[Private\]\s*",              // Private email tags
            @"\[PHI\]\s*",                  // PHI email tags
            @"\[Encrypted\]\s*",            // Encrypted email tags
            @"\[SPAM\]\s*",                 // Spam tags
            // Cleanup
            @"^\s+|\s+$",                   // Leading/trailing whitespace
            @"[-_\s]{2,}",                  // Multiple separators into single hyphen
            @"^-+|-+$"                      // Leading/trailing hyphens
        };

        public string GetFullVaultPath()
        {
            return System.IO.Path.Combine(VaultBasePath, VaultName);
        }

        public string GetInboxPath()
        {
            return System.IO.Path.Combine(GetFullVaultPath(), InboxFolder);
        }

        public string GetContactsPath()
        {
            return Path.Combine(GetFullVaultPath(), ContactsFolder);
        }

        public void Save()
        {
            var settings = new Dictionary<string, object>
            {
                { "VaultName", VaultName },
                { "VaultBasePath", VaultBasePath },
                { "InboxFolder", InboxFolder },
                { "ContactsFolder", ContactsFolder },
                { "EnableContactSaving", EnableContactSaving },
                { "SearchEntireVaultForContacts", SearchEntireVaultForContacts },
                { "LaunchObsidian", LaunchObsidian },
                { "ObsidianDelaySeconds", ObsidianDelaySeconds },
                { "ShowCountdown", ShowCountdown },
                { "CreateObsidianTask", CreateObsidianTask },
                { "CreateOutlookTask", CreateOutlookTask },
                { "DefaultDueDays", DefaultDueDays },
                { "UseRelativeReminder", UseRelativeReminder },
                { "DefaultReminderDays", DefaultReminderDays },
                { "DefaultReminderHour", DefaultReminderHour },
                { "AskForDates", AskForDates },
                { "SubjectCleanupPatterns", SubjectCleanupPatterns },
                { "GroupEmailThreads", GroupEmailThreads },
                { "ShowDevelopmentSettings", ShowDevelopmentSettings },
                { "ShowThreadDebug", ShowThreadDebug },
                { "DefaultNoteTags", DefaultNoteTags },
                { "DefaultTaskTags", DefaultTaskTags },
                { "NoteTitleFormat", NoteTitleFormat },
                { "NoteTitleMaxLength", NoteTitleMaxLength },
                { "NoteTitleIncludeDate", NoteTitleIncludeDate }
            };

            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            
            // Ensure settings directory exists before saving
            string settingsPath = GetSettingsPath();
            string settingsDir = Path.GetDirectoryName(settingsPath);
            if (!Directory.Exists(settingsDir))
            {
                Directory.CreateDirectory(settingsDir);
            }
            
            File.WriteAllText(settingsPath, json);
        }

        public void Load()
        {
            if (File.Exists(GetSettingsPath()))
            {
                string json = File.ReadAllText(GetSettingsPath());
                var settings = JsonConvert.DeserializeObject<Dictionary<string, JToken>>(json);

                if (settings.ContainsKey("VaultName"))
                {
                    VaultName = settings["VaultName"].Value<string>();
                }
                if (settings.ContainsKey("VaultBasePath"))
                {
                    VaultBasePath = settings["VaultBasePath"].Value<string>();
                }
                if (settings.ContainsKey("InboxFolder"))
                {
                    InboxFolder = settings["InboxFolder"].Value<string>();
                }
                if (settings.ContainsKey("ContactsFolder"))
                {
                    ContactsFolder = settings["ContactsFolder"].Value<string>();
                }
                if (settings.ContainsKey("EnableContactSaving"))
                {
                    EnableContactSaving = settings["EnableContactSaving"].Value<bool>();
                }
                if (settings.ContainsKey("SearchEntireVaultForContacts"))
                {
                    SearchEntireVaultForContacts = settings["SearchEntireVaultForContacts"].Value<bool>();
                }
                if (settings.ContainsKey("LaunchObsidian"))
                {
                    LaunchObsidian = settings["LaunchObsidian"].Value<bool>();
                }
                if (settings.ContainsKey("ObsidianDelaySeconds"))
                {
                    ObsidianDelaySeconds = settings["ObsidianDelaySeconds"].Value<int>();
                }
                if (settings.ContainsKey("ShowCountdown"))
                {
                    ShowCountdown = settings["ShowCountdown"].Value<bool>();
                }
                if (settings.ContainsKey("CreateObsidianTask"))
                {
                    CreateObsidianTask = settings["CreateObsidianTask"].Value<bool>();
                }
                if (settings.ContainsKey("CreateOutlookTask"))
                {
                    CreateOutlookTask = settings["CreateOutlookTask"].Value<bool>();
                }
                if (settings.ContainsKey("DefaultDueDays"))
                {
                    DefaultDueDays = settings["DefaultDueDays"].Value<int>();
                }
                if (settings.ContainsKey("UseRelativeReminder"))
                {
                    UseRelativeReminder = settings["UseRelativeReminder"].Value<bool>();
                }
                if (settings.ContainsKey("DefaultReminderDays"))
                {
                    DefaultReminderDays = settings["DefaultReminderDays"].Value<int>();
                }
                if (settings.ContainsKey("DefaultReminderHour"))
                {
                    DefaultReminderHour = settings["DefaultReminderHour"].Value<int>();
                }
                if (settings.ContainsKey("AskForDates"))
                {
                    AskForDates = settings["AskForDates"].Value<bool>();
                }
                if (settings.ContainsKey("SubjectCleanupPatterns"))
                {
                    var patterns = settings["SubjectCleanupPatterns"].ToObject<List<string>>();
                    if (patterns != null && patterns.Count > 0)
                    {
                        SubjectCleanupPatterns = patterns;
                    }
                }
                if (settings.ContainsKey("GroupEmailThreads"))
                {
                    GroupEmailThreads = settings["GroupEmailThreads"].Value<bool>();
                }
                if (settings.ContainsKey("ShowDevelopmentSettings"))
                {
                    ShowDevelopmentSettings = settings["ShowDevelopmentSettings"].Value<bool>();
                }
                if (settings.ContainsKey("ShowThreadDebug"))
                {
                    ShowThreadDebug = settings["ShowThreadDebug"].Value<bool>();
                }
                if (settings.ContainsKey("DefaultNoteTags"))
                {
                    var tags = settings["DefaultNoteTags"].ToObject<List<string>>();
                    if (tags != null && tags.Count > 0)
                    {
                        DefaultNoteTags = tags;
                    }
                }
                if (settings.ContainsKey("DefaultTaskTags"))
                {
                    var tags = settings["DefaultTaskTags"].ToObject<List<string>>();
                    if (tags != null && tags.Count > 0)
                    {
                        DefaultTaskTags = tags;
                    }
                }
                if (settings.ContainsKey("NoteTitleFormat"))
                {
                    NoteTitleFormat = settings["NoteTitleFormat"].Value<string>();
                }
                if (settings.ContainsKey("NoteTitleMaxLength"))
                {
                    NoteTitleMaxLength = settings["NoteTitleMaxLength"].Value<int>();
                }
                if (settings.ContainsKey("NoteTitleIncludeDate"))
                {
                    NoteTitleIncludeDate = settings["NoteTitleIncludeDate"].Value<bool>();
                }
            }
        }

        protected virtual string GetSettingsPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlingMD.Outlook", "ObsidianSettings.json");
        }
    }
} 