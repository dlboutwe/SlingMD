using System;
using System.Configuration;

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

        public string GetFullVaultPath()
        {
            return System.IO.Path.Combine(VaultBasePath, VaultName);
        }

        public string GetInboxPath()
        {
            return System.IO.Path.Combine(GetFullVaultPath(), InboxFolder);
        }
    }
} 