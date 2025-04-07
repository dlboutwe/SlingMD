using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Reflection;
using System.IO;

namespace SlingMD.Outlook.Forms
{
    public partial class BaseForm : Form
    {
        private bool _disposed = false;
        private Icon _customIcon = null;

        public BaseForm()
        {
            if (!DesignMode)
            {
                try
                {
                    using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SlingMD.Outlook.Resources.SlingMD.ico"))
                    {
                        if (stream != null)
                        {
                            _customIcon = new Icon(stream);
                            this.Icon = _customIcon;
                        }
                    }
                }
                catch (Exception)
                {
                    // Silently fail if icon cannot be loaded
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_customIcon != null)
                    {
                        _customIcon.Dispose();
                        _customIcon = null;
                    }
                }
                _disposed = true;
            }
            base.Dispose(disposing);
        }
    }
} 