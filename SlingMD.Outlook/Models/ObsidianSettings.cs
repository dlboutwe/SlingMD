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

        public List<string> SubjectCleanupPatterns { get; set; } = new List<string>
        {
            @"^(?:Re|Fwd|FW|RE|FWD):\s*",  // Reply and forward prefixes
            @"\[EXTERNAL\]\s*",             // External email tags
            @"\[Internal\]\s*",             // Internal email tags
            @"\[Confidential\]\s*",         // Confidential tags
            @"\[Secure\]\s*"                // Secure email tags
        };

        public string GetFullVaultPath()
        {
            return System.IO.Path.Combine(VaultBasePath, VaultName);
        }

        public string GetInboxPath()
        {
            return System.IO.Path.Combine(GetFullVaultPath(), InboxFolder);
        }

        public void Save()
        {
            var settings = new Dictionary<string, object>
            {
                { "VaultName", VaultName },
                { "VaultBasePath", VaultBasePath },
                { "InboxFolder", InboxFolder },
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
                { "SubjectCleanupPatterns", SubjectCleanupPatterns }
            };

            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText(GetSettingsPath(), json);
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
            }
        }

        private string GetSettingsPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlingMD.Outlook", "ObsidianSettings.json");
        }
    }
} 