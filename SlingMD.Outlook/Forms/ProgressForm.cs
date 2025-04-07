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
            this.lblMessage = new Label();
            this.progressBar = new ProgressBar();
            this.btnClose = new Button();
            this.SuspendLayout();

            // Label
            this.lblMessage = new Label();
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(12, 9);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(35, 13);
            this.lblMessage.TabIndex = 0;
            this.lblMessage.Text = "Please wait...";

            // Progress Bar
            this.progressBar = new ProgressBar();
            this.progressBar.Location = new System.Drawing.Point(12, 34);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(360, 23);
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.MarqueeAnimationSpeed = 30;
            this.progressBar.TabIndex = 1;

            // Close Button (hidden by default)
            this.btnClose = new Button();
            this.btnClose.Text = "Close";
            this.btnClose.Size = new Size(75, 23);
            this.btnClose.Location = new Point(297, 63);
            this.btnClose.Visible = false;
            this.btnClose.Click += (s, e) => this.Close();

            // Form
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 101);
            this.ControlBox = false;
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblMessage);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Processing...";
            this.ResumeLayout(false);
            this.PerformLayout();
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