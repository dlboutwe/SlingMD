using System;
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
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup if needed
        }

        private ObsidianSettings LoadSettings()
        {
            var settings = new ObsidianSettings();
            settings.Load(); // Load settings from file
            return settings;
        }

        public async void ProcessSelectedEmail()
        {
            try
            {
                // Get selected email
                var explorer = Application.ActiveExplorer();
                if (explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var mail = explorer.Selection[1] as MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Selected item is not an email.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Process the email
                await _emailProcessor.ProcessEmail(mail);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving email: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ShowSettings()
        {
            try
            {
                using (var form = new SettingsForm(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Settings are automatically saved by the form
                        // Recreate email processor with new settings
                        _emailProcessor = new EmailProcessor(_settings);
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
