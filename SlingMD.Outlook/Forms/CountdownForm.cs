using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class CountdownForm : Form
    {
        private int _secondsRemaining;
        private Timer _timer;
        private Label _lblCountdown;
        private Button _btnSkip;

        public CountdownForm(int seconds)
        {
            _secondsRemaining = seconds;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Create countdown label
            _lblCountdown = new Label();
            _lblCountdown.AutoSize = false;
            _lblCountdown.Dock = DockStyle.Fill;
            _lblCountdown.TextAlign = ContentAlignment.MiddleCenter;
            _lblCountdown.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            _lblCountdown.Text = $"Opening in Obsidian in {_secondsRemaining} seconds...";
            
            // Create skip button
            _btnSkip = new Button();
            _btnSkip.Text = "Skip";
            _btnSkip.Dock = DockStyle.Bottom;
            _btnSkip.Height = 40;
            _btnSkip.Click += BtnSkip_Click;
            
            // Add controls
            this.Controls.Add(_lblCountdown);
            this.Controls.Add(_btnSkip);
            
            // Form settings
            this.ClientSize = new System.Drawing.Size(300, 120);
            this.Text = "Opening in Obsidian";
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.TopMost = true;
            this.ShowIcon = false;
            this.BackColor = Color.White;
            
            // Initialize timer
            _timer = new Timer();
            _timer.Interval = 1000;
            _timer.Tick += Timer_Tick;
            
            this.ResumeLayout(false);
            this.Load += CountdownForm_Load;
        }

        private void CountdownForm_Load(object sender, EventArgs e)
        {
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _secondsRemaining--;
            
            if (_secondsRemaining <= 0)
            {
                _timer.Stop();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                _lblCountdown.Text = $"Opening in Obsidian in {_secondsRemaining} seconds...";
            }
        }

        private void BtnSkip_Click(object sender, EventArgs e)
        {
            _timer.Stop();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
            }
            
            base.Dispose(disposing);
        }
    }
} 