using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace SlingMD.Outlook.Ribbon
{
    [ComVisible(true)]
    public class SlingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly ThisAddIn _addIn;
        private Bitmap _slingLogo;

        public SlingRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
            LoadSlingLogo();
        }

        private void LoadSlingLogo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("SlingMD.Outlook.Resources.SlingMD_pixel.png"))
                {
                    if (stream != null)
                    {
                        _slingLogo = new Bitmap(stream);
                    }
                }
            }
            catch (Exception)
            {
                // If loading fails, we'll fall back to the default Office icon
                _slingLogo = null;
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SlingMD.Outlook.Ribbon.SlingRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Ribbon Callbacks

        public void OnSlingButtonClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.ProcessSelectedEmail();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing email: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSettingsButtonClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.ShowSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public Bitmap GetSlingButtonImage(Office.IRibbonControl control)
        {
            return _slingLogo;
        }

        #endregion

        #region Helpers

        private string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _slingLogo?.Dispose();
            }
        }

        #endregion
    }
} 