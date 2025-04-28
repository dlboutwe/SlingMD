using System;
using System.IO;
using Xunit;
using SlingMD.Outlook.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace SlingMD.Tests.Models
{
    public class ObsidianSettingsTests
    {
        private readonly string _testSettingsDir;
        private readonly string _testSettingsPath;

        public ObsidianSettingsTests()
        {
            // Setup test directory for settings
            _testSettingsDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Settings");
            _testSettingsPath = Path.Combine(_testSettingsDir, "ObsidianSettings.json");
            
            // Clean up any previous test data
            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }
            Directory.CreateDirectory(_testSettingsDir);
        }

        [Fact]
        public void Save_CreatesSettingsDirectoryIfNotExists()
        {
            // Arrange
            var settings = new ObsidianSettingsTestable();
            settings.TestSettingsPath = _testSettingsPath;

            // Delete the directory to test creation
            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }

            // Act
            settings.Save();

            // Assert
            Assert.True(Directory.Exists(_testSettingsDir), "Settings directory should be created");
            Assert.True(File.Exists(_testSettingsPath), "Settings file should be created");
        }

        [Fact]
        public void SearchEntireVaultForContacts_SavedAndLoaded_Correctly()
        {
            // Arrange - Create settings with SearchEntireVaultForContacts=true
            var settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                SearchEntireVaultForContacts = true
            };

            // Act - Save and load
            settings.Save();
            
            var loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            
            // Default value is false, so if it loads correctly it should be true
            Assert.False(loadedSettings.SearchEntireVaultForContacts, "Should start as false before loading");
            
            loadedSettings.Load();

            // Assert - Check that the setting was correctly loaded
            Assert.True(loadedSettings.SearchEntireVaultForContacts, "SearchEntireVaultForContacts should be true after loading");
        }
        
        [Fact]
        public void SaveAndLoad_PersistsAllSettings()
        {
            // Arrange
            var settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                VaultName = "TestVault",
                VaultBasePath = @"C:\Test\Path",
                InboxFolder = "TestInbox",
                ContactsFolder = "TestContacts",
                EnableContactSaving = true,
                SearchEntireVaultForContacts = true,
                LaunchObsidian = false,
                ObsidianDelaySeconds = 5,
                ShowCountdown = false,
                CreateObsidianTask = false,
                CreateOutlookTask = true,
                DefaultDueDays = 3,
                DefaultReminderDays = 2,
                DefaultReminderHour = 10,
                AskForDates = true,
                GroupEmailThreads = false,
                ShowDevelopmentSettings = true,
                ShowThreadDebug = true
            };

            // Add some test patterns
            settings.SubjectCleanupPatterns.Clear();
            settings.SubjectCleanupPatterns.Add("test-pattern-1");
            settings.SubjectCleanupPatterns.Add("test-pattern-2");

            // Act
            settings.Save();
            
            var loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            // Assert
            Assert.Equal(settings.VaultName, loadedSettings.VaultName);
            Assert.Equal(settings.VaultBasePath, loadedSettings.VaultBasePath);
            Assert.Equal(settings.InboxFolder, loadedSettings.InboxFolder);
            Assert.Equal(settings.ContactsFolder, loadedSettings.ContactsFolder);
            Assert.Equal(settings.EnableContactSaving, loadedSettings.EnableContactSaving);
            Assert.Equal(settings.SearchEntireVaultForContacts, loadedSettings.SearchEntireVaultForContacts);
            Assert.Equal(settings.LaunchObsidian, loadedSettings.LaunchObsidian);
            Assert.Equal(settings.ObsidianDelaySeconds, loadedSettings.ObsidianDelaySeconds);
            Assert.Equal(settings.ShowCountdown, loadedSettings.ShowCountdown);
            Assert.Equal(settings.CreateObsidianTask, loadedSettings.CreateObsidianTask);
            Assert.Equal(settings.CreateOutlookTask, loadedSettings.CreateOutlookTask);
            Assert.Equal(settings.DefaultDueDays, loadedSettings.DefaultDueDays);
            Assert.Equal(settings.DefaultReminderDays, loadedSettings.DefaultReminderDays);
            Assert.Equal(settings.DefaultReminderHour, loadedSettings.DefaultReminderHour);
            Assert.Equal(settings.AskForDates, loadedSettings.AskForDates);
            Assert.Equal(settings.GroupEmailThreads, loadedSettings.GroupEmailThreads);
            Assert.Equal(settings.ShowDevelopmentSettings, loadedSettings.ShowDevelopmentSettings);
            Assert.Equal(settings.ShowThreadDebug, loadedSettings.ShowThreadDebug);
            
            // Check patterns
            Assert.Equal(2, loadedSettings.SubjectCleanupPatterns.Count);
            Assert.Contains("test-pattern-1", loadedSettings.SubjectCleanupPatterns);
            Assert.Contains("test-pattern-2", loadedSettings.SubjectCleanupPatterns);
        }

        [Fact]
        public void LoadSettings_FileDoesNotExist_UsesDefaultValues()
        {
            // Arrange
            if (File.Exists(_testSettingsPath))
            {
                File.Delete(_testSettingsPath);
            }

            // Act
            var settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.Load();

            // Assert - Check that default values are used
            Assert.Equal("Logic", settings.VaultName);
            Assert.Contains("Documents\\Notes", settings.VaultBasePath);
            Assert.Equal("Inbox", settings.InboxFolder);
            Assert.Equal("Contacts", settings.ContactsFolder);
            Assert.True(settings.EnableContactSaving);
            Assert.False(settings.SearchEntireVaultForContacts);
            Assert.True(settings.LaunchObsidian);
        }

        [Fact]
        public void GetFullVaultPath_CombinesBasePathAndVaultName()
        {
            // Arrange
            var settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault"
            };

            // Act
            string fullPath = settings.GetFullVaultPath();

            // Assert
            Assert.Equal(@"C:\Test\Path\TestVault", fullPath);
        }

        [Fact]
        public void GetInboxPath_ReturnsCorrectPath()
        {
            // Arrange
            var settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault",
                InboxFolder = "TestInbox"
            };

            // Act
            string inboxPath = settings.GetInboxPath();

            // Assert
            Assert.Equal(@"C:\Test\Path\TestVault\TestInbox", inboxPath);
        }

        [Fact]
        public void GetContactsPath_ReturnsCorrectPath()
        {
            // Arrange
            var settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault",
                ContactsFolder = "TestContacts"
            };

            // Act
            string contactsPath = settings.GetContactsPath();

            // Assert
            Assert.Equal(@"C:\Test\Path\TestVault\TestContacts", contactsPath);
        }
    }

    // Testable version of ObsidianSettings that allows overriding the settings path
    public class ObsidianSettingsTestable : ObsidianSettings
    {
        public string TestSettingsPath { get; set; }

        protected override string GetSettingsPath()
        {
            return TestSettingsPath ?? base.GetSettingsPath();
        }
    }
}