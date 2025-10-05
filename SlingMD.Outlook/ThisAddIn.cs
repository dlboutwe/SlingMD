using System;
using System.CodeDom;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Ribbon;

namespace SlingMD.Outlook
{
    public partial class ThisAddIn
    {
        private ObsidianSettings _settings;
        private EmailProcessor _emailProcessor;
        private AppointmentProcessor _appointmentProcessor;
        private SlingRibbon _ribbon;


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SlingRibbon(this);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _settings = LoadSettings();
            _emailProcessor = new EmailProcessor(_settings);
            _appointmentProcessor = new AppointmentProcessor(_settings);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Save settings when Outlook is closing
            if (_settings != null)
            {
                _settings.Save();
            }
        }

        private ObsidianSettings LoadSettings()
        {
            var settings = new ObsidianSettings();
            settings.Load(); // Load settings from file
            return settings;
        }

        public async void ProcessSelection()
        {
            try
            {
                // Get selected item
                var explorer = Application.ActiveExplorer();
                if (explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email or calendar appointment first.", "SlingMD",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }


                var mail = explorer.Selection[1] as MailItem;
                var appt = explorer.Selection[1] as AppointmentItem;
                if (mail == null && appt == null)
                {
                    MessageBox.Show("Selected item is not an email or calendar appointment.", "SlingMD",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (mail != null)
                {
                    // Process the email
                    await _emailProcessor.ProcessEmail(mail);
                }

                if (appt != null)
                {
                    await _appointmentProcessor.ProcessAppointment(appt);
                }


            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving email or appointment: {ex.Message}", "SlingMD", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void SaveTodaysAppointments()
        {
            try
            {
                

                // iterate through all of the accounts and pull the default calendars
                foreach (Account account in Application.Session.Accounts)
                {
                    // Get the default calendar folder
                    if (account.DeliveryStore.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar) is Folder calendarFolder)
                    {

                        // Define the time range for appointments
                        DateTime startTime = DateTime.Now.Date; // Start of today
                        DateTime endTime = startTime.AddDays(1); // Tomorrow

                        // Create a filter string for the time range
                        // Note: The date format 'g' (general date/time pattern) is suitable for Jet filters
                        string filter = "[Start] >= '" + startTime.ToString("g") + "' AND [End] <= '" + endTime.ToString("g") + "'";

                        // Get the items collection and apply the filter
                        Microsoft.Office.Interop.Outlook.Items calendarItems = calendarFolder.Items;
                        calendarItems.IncludeRecurrences = true; // Include recurring appointments
                        calendarItems.Sort("[Start]"); // Sort by start date

                        Microsoft.Office.Interop.Outlook.Items filteredAppointments = calendarItems.Restrict(filter);

                        // Iterate through the filtered appointments
                        foreach (object item in filteredAppointments)
                        {
                            if (item is Microsoft.Office.Interop.Outlook.AppointmentItem appointment)
                            {
                                // Process each appointment item (disable launching of obsidian, because that gets annoying.
                                await _appointmentProcessor.ProcessAppointment(appointment, true);
                            }
                        }
                    }
                    
                }

                
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error retrieving appointments: {ex.Message}");
            }
        }


    public void ShowSettings()
        {
            try
            {
                using (var form = new Settings(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Settings are automatically saved by the form
                        // Recreate email processor with new settings
                        _emailProcessor = new EmailProcessor(_settings);
                        // Recreate appointment processor with new settings
                        _appointmentProcessor = new AppointmentProcessor(_settings);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
