using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ProgressForm : Form
    {
        private Label lblMessage;
        private ProgressBar progressBar;
        private Button btnClose;
        private Timer autoCloseTimer;

        public ProgressForm(string message = "Please wait...")
        {
            InitializeComponent();
            lblMessage.Text = message;

            autoCloseTimer = new Timer
            {
                Interval = 3000 // 3 seconds
            };
            autoCloseTimer.Tick += (s, e) =>
            {
                autoCloseTimer.Stop();
                this.Close();
            };
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgressForm));
            this.SuspendLayout();
            // 
            // ProgressForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ProgressForm";
            this.ResumeLayout(false);

        }

        public void UpdateProgress(string message, int percentage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(message, percentage)));
                return;
            }

            lblMessage.Text = message;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = Math.Min(100, Math.Max(0, percentage));
            this.Update();
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowSuccess(message, autoClose)));
                return;
            }

            lblMessage.Text = message;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 100;
            
            if (autoClose)
            {
                autoCloseTimer.Start();
            }
            else
            {
                btnClose.Visible = true;
            }
        }

        public void ShowError(string message, bool autoClose = false)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowError(message, autoClose)));
                return;
            }

            lblMessage.Text = message;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            btnClose.Visible = true;
            
            if (autoClose)
            {
                autoCloseTimer.Start();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            autoCloseTimer.Stop();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (lblMessage != null) lblMessage.Dispose();
                if (progressBar != null) progressBar.Dispose();
                if (btnClose != null) btnClose.Dispose();
                if (autoCloseTimer != null) autoCloseTimer.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 