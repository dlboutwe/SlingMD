using System;
using System.Collections.Generic;
using Xunit;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Models;

namespace SlingMD.Tests.Services
{
    public class TaskServiceTests
    {
        [Fact]
        public void GenerateObsidianTask_UsesTagsAndDates_SingleLine()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.InitializeTaskSettings(1, 1, 9, false); // due in 1 day, remind in 1 day

            var tags = new List<string> { "FollowUp", "ActionItem" };
            string result = service.GenerateObsidianTask("TestNote", tags);

            Assert.StartsWith("- [ ] [[TestNote]] #FollowUp #ActionItem", result);
            Assert.Contains("➕", result);
            Assert.Contains("🛫", result);
            Assert.Contains("📅", result);
            Assert.DoesNotContain("\n", result.TrimEnd()); // Should be single line
        }

        [Fact]
        public void GenerateObsidianTask_FormatsTagsWithHash()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.InitializeTaskSettings();

            var tags = new List<string> { "foo", "#bar", "baz" };
            string result = service.GenerateObsidianTask("Note", tags);

            Assert.Contains("#foo", result);
            Assert.Contains("#bar", result);
            Assert.Contains("#baz", result);
        }

        [Fact]
        public void GenerateObsidianTask_FallsBackToDefaultTag()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.InitializeTaskSettings();

            string result = service.GenerateObsidianTask("Note", null);
            Assert.Contains("#FollowUp", result);
        }

        [Fact]
        public void GenerateObsidianTask_EmptyTagList_FallsBackToDefault()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.InitializeTaskSettings();

            string result = service.GenerateObsidianTask("Note", new List<string>());
            Assert.Contains("#FollowUp", result);
        }

        [Fact]
        public void GenerateObsidianTask_DisabledTaskCreation_ReturnsEmpty()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.DisableTaskCreation();

            string result = service.GenerateObsidianTask("Note", new List<string> { "foo" });
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void GenerateObsidianTask_Dates_AreCorrectlyFormatted()
        {
            var settings = new ObsidianSettings();
            var service = new TaskService(settings);
            service.InitializeTaskSettings(2, 1, 9, false); // due in 2 days, remind in 1 day

            string result = service.GenerateObsidianTask("Note", new List<string> { "foo" });
            Assert.Matches(@"\d{4}-\d{2}-\d{2}", result); // Should contain dates in yyyy-MM-dd
        }
    }
} 