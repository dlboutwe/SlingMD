using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;

        public ProgressForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Create progress bar
            this.progressBar = new ProgressBar();
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Step = 1;
            this.progressBar.Location = new Point(12, 50);
            this.progressBar.Size = new Size(350, 30);

            // Create status label
            this.lblStatus = new Label();
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new Point(12, 20);
            this.lblStatus.Size = new Size(350, 20);
            this.lblStatus.Text = "Processing...";

            // Configure form
            this.ClientSize = new Size(374, 100);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "SlingMD";
            this.TopMost = true;

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        public void UpdateProgress(string message, int percentage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, int>(UpdateProgress), new object[] { message, percentage });
                return;
            }

            this.lblStatus.Text = message;
            this.progressBar.Value = Math.Max(0, Math.Min(100, percentage));
            
            // Auto-close if we reach 100%
            if (percentage >= 100)
            {
                Timer closeTimer = new Timer();
                closeTimer.Interval = 1000; // 1 second delay
                closeTimer.Tick += (s, e) => 
                {
                    closeTimer.Stop();
                    this.Close();
                };
                closeTimer.Start();
            }
            
            this.Refresh();
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            UpdateProgress(message, 100);
            this.BackColor = Color.FromArgb(220, 255, 220);
            
            if (autoClose)
            {
                Timer closeTimer = new Timer();
                closeTimer.Interval = 2000; // 2 second delay
                closeTimer.Tick += (s, e) => 
                {
                    closeTimer.Stop();
                    this.Close();
                };
                closeTimer.Start();
            }
        }

        public void ShowError(string message, bool autoClose = false)
        {
            UpdateProgress(message, 100);
            this.BackColor = Color.FromArgb(255, 220, 220);
            
            if (autoClose)
            {
                Timer closeTimer = new Timer();
                closeTimer.Interval = 3000; // 3 second delay
                closeTimer.Tick += (s, e) => 
                {
                    closeTimer.Stop();
                    this.Close();
                };
                closeTimer.Start();
            }
        }
    }
} 