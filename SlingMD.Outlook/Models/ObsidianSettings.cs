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