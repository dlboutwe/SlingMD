using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ProgressForm : Form
    {
        private readonly Timer _autoCloseTimer;
        private readonly Label _statusLabel;
        private readonly ProgressBar _progressBar;
        private readonly Button _closeButton;

        public ProgressForm()
        {
            // Form settings
            this.Text = "SlingMD Progress";
            this.Size = new Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // Status label
            _statusLabel = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(360, 20),
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Starting..."
            };
            this.Controls.Add(_statusLabel);

            // Progress bar
            _progressBar = new ProgressBar
            {
                Location = new Point(20, 50),
                Size = new Size(360, 23),
                Style = ProgressBarStyle.Continuous,
                Value = 0
            };
            this.Controls.Add(_progressBar);

            // Close button (hidden by default)
            _closeButton = new Button
            {
                Location = new Point(305, 85),
                Size = new Size(75, 23),
                Text = "Close",
                Visible = false
            };
            _closeButton.Click += (s, e) => this.Close();
            this.Controls.Add(_closeButton);

            // Auto-close timer
            _autoCloseTimer = new Timer
            {
                Interval = 3000 // 3 seconds
            };
            _autoCloseTimer.Tick += (s, e) =>
            {
                _autoCloseTimer.Stop();
                this.Close();
            };
        }

        public void UpdateProgress(string message, int percentage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(message, percentage)));
                return;
            }

            _statusLabel.Text = message;
            _progressBar.Value = Math.Min(100, Math.Max(0, percentage));
            this.Update();
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowSuccess(message, autoClose)));
                return;
            }

            _statusLabel.Text = message;
            _progressBar.Value = 100;
            
            if (autoClose)
            {
                _autoCloseTimer.Start();
            }
            else
            {
                _closeButton.Visible = true;
            }
        }

        public void ShowError(string message, bool autoClose = false)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowError(message, autoClose)));
                return;
            }

            _statusLabel.Text = message;
            _progressBar.Value = 0;
            _closeButton.Visible = true;
            
            if (autoClose)
            {
                _autoCloseTimer.Start();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _autoCloseTimer.Stop();
            base.OnFormClosing(e);
        }
    }
} 