using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using System.Collections.Generic;
using System.Linq;

namespace SlingMD.Outlook.Services
{
    public class TaskService
    {
        private readonly ObsidianSettings _settings;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _useRelativeReminder;
        private bool _createTasks = true;

        public TaskService(ObsidianSettings settings)
        {
            _settings = settings;
        }

        public void InitializeTaskSettings(int? dueDays = null, int? reminderDays = null, int? reminderHour = null, bool? useRelativeReminder = null)
        {
            _taskDueDays = dueDays ?? _settings.DefaultDueDays;
            _taskReminderDays = reminderDays ?? _settings.DefaultReminderDays;
            _taskReminderHour = reminderHour ?? _settings.DefaultReminderHour;
            _useRelativeReminder = useRelativeReminder ?? _settings.UseRelativeReminder;
        }

        public bool ShouldCreateTasks => _createTasks;

        public void DisableTaskCreation()
        {
            _createTasks = false;
        }

        /// <summary>
        /// Generates a single-line Obsidian task with tags and dates.
        /// </summary>
        /// <param name="fileName">The note file name (without extension).</param>
        /// <param name="taskTags">A list of tags to include in the task line (e.g., ["FollowUp", "ActionItem"]).</param>
        /// <returns>The full Obsidian task line, including tags and dates, on a single line.</returns>
        public string GenerateObsidianTask(string fileName, List<string> taskTags = null)
        {
            if (!_createTasks) return string.Empty;

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            string dueDate = DateTime.Now.Date.AddDays(_taskDueDays.Value).ToString("yyyy-MM-dd");
            
            // Calculate reminder date based on setting
            DateTime reminderDateTime;
            if (_useRelativeReminder)
            {
                // Relative: Calculate from due date
                reminderDateTime = DateTime.Now.Date.AddDays(_taskDueDays.Value - _taskReminderDays.Value);
            }
            else
            {
                // Absolute: Calculate from today
                reminderDateTime = DateTime.Now.Date.AddDays(_taskReminderDays.Value);
            }
            string reminderDate = reminderDateTime.ToString("yyyy-MM-dd");

            // Format tags as #tag
            string tagsPart = (taskTags != null && taskTags.Count > 0)
                ? string.Join(" ", taskTags.Select(t => t.StartsWith("#") ? t : "#" + t))
                : "#FollowUp";

            // All on one line
            return $"- [ ] [[{fileName}]] {tagsPart} ➕ {currentDate} 🛫 {reminderDate} 📅 {dueDate}";
        }

        public async Task CreateOutlookTask(MailItem mail)
        {
            if (!_createTasks) return;

            try
            {
                var outlookApp = mail.Application;
                var task = outlookApp.CreateItem(OlItemType.olTaskItem);
                task.Subject = $"Follow up: {mail.Subject}";
                task.Body = $"Follow up on email from {mail.SenderName}\n\nOriginal email:\n{mail.Body}";
                
                // Set due date based on settings
                var dueDate = DateTime.Now.Date.AddDays(_taskDueDays.Value);
                task.DueDate = dueDate;
                task.ReminderSet = true;
                
                // Calculate reminder time based on setting
                DateTime reminderDate;
                if (_useRelativeReminder)
                {
                    // Relative: Calculate from due date
                    reminderDate = dueDate.AddDays(-_taskReminderDays.Value);
                }
                else
                {
                    // Absolute: Calculate from today
                    reminderDate = DateTime.Now.Date.AddDays(_taskReminderDays.Value);
                }
                var reminderTime = reminderDate.AddHours(_taskReminderHour.Value);
                
                // If reminder would be in the past, set it to the next possible time
                if (reminderTime < DateTime.Now)
                {
                    if (reminderTime.Date == DateTime.Now.Date)
                    {
                        // If it's today but earlier hour, set to next hour
                        reminderTime = DateTime.Now.AddHours(1);
                    }
                    else
                    {
                        // If it's a past day, set to tomorrow at the specified hour
                        reminderTime = DateTime.Now.Date.AddDays(1).AddHours(_taskReminderHour.Value);
                    }
                }
                
                task.ReminderTime = reminderTime;
                task.Save();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Failed to create Outlook task: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
} 