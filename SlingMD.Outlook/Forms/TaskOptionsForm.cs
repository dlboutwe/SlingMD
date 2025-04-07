using System;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace SlingMD.Outlook.Forms
{
    public partial class TaskOptionsForm : Form
    {
        private const int formWidth = 550;
        private const int formHeight = 250;
        private const int labelX = 20;
        private const int controlX = 180;
        private const int helpTextX = 350;
        private const int startY = 20;
        private const int lineHeight = 35;

        private Label lblDueDays;
        private NumericUpDown numDueDays;
        private DateTimePicker dtpDueDate;
        private Label lblDueDaysHelp;
        private Label lblReminderDays;
        private NumericUpDown numReminderDays;
        private DateTimePicker dtpReminderDate;
        private Label lblReminderDaysHelp;
        private Label lblReminderHour;
        private NumericUpDown numReminderHour;
        private Label lblReminderHourHelp;
        private CheckBox chkUseRelativeReminder;
        private Button btnOK;
        private Button btnCancel;

        public DateTime DueDate { get; private set; }
        public DateTime ReminderDate { get; private set; }

        public int DueDays => chkUseRelativeReminder.Checked ? (int)numDueDays.Value : (dtpDueDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderDays => chkUseRelativeReminder.Checked ? (int)numReminderDays.Value : (dtpReminderDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderHour => (int)numReminderHour.Value;
        public bool UseRelativeReminder => chkUseRelativeReminder.Checked;

        public TaskOptionsForm(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder = false)
        {
            InitializeComponent();
            DueDate = DateTime.Now.Date.AddDays(defaultDueDays);
            ReminderDate = DateTime.Now.Date.AddDays(defaultReminderDays);
            chkUseRelativeReminder.Checked = useRelativeReminder;
            numDueDays.Value = defaultDueDays;
            numReminderDays.Value = defaultReminderDays;
            numReminderHour.Value = defaultReminderHour;
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        public TaskOptionsForm(DateTime defaultDueDate, DateTime defaultReminderDate)
        {
            InitializeComponent();
            DueDate = defaultDueDate;
            ReminderDate = defaultReminderDate;
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void InitializeControls()
        {
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskOptionsForm));
            this.SuspendLayout();
            // 
            // TaskOptionsForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TaskOptionsForm";
            this.ResumeLayout(false);

        }

        private void ChkUseRelativeReminder_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void UpdateControlsVisibility()
        {
            bool useRelative = chkUseRelativeReminder.Checked;
            
            numDueDays.Visible = useRelative;
            dtpDueDate.Visible = !useRelative;
            
            numReminderDays.Visible = useRelative;
            dtpReminderDate.Visible = !useRelative;

            // When switching to absolute dates, update the date pickers based on current numeric values
            if (!useRelative)
            {
                dtpDueDate.Value = DateTime.Now.Date.AddDays((double)numDueDays.Value);
                dtpReminderDate.Value = DateTime.Now.Date.AddDays((double)numReminderDays.Value);
            }
            // When switching to relative dates, update the numeric values based on current date pickers
            else
            {
                numDueDays.Value = Math.Max(0, (dtpDueDate.Value.Date - DateTime.Now.Date).Days);
                numReminderDays.Value = Math.Max(0, (dtpReminderDate.Value.Date - DateTime.Now.Date).Days);
            }
        }

        private void UpdateHelpText()
        {
            if (chkUseRelativeReminder.Checked)
            {
                lblDueDaysHelp.Text = "(Days from today)";
                lblReminderDaysHelp.Text = "(Days before due date)";
            }
            else
            {
                lblDueDaysHelp.Text = "(Select date)";
                lblReminderDaysHelp.Text = "(Select reminder date)";
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (chkUseRelativeReminder.Checked)
            {
                if (numReminderDays.Value > numDueDays.Value)
                {
                    MessageBox.Show(
                        "Reminder days cannot be greater than due days when using relative dates.",
                        "Invalid Reminder",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    this.DialogResult = DialogResult.None;
                }
            }
            else
            {
                if (dtpReminderDate.Value.Date > dtpDueDate.Value.Date)
                {
                    MessageBox.Show(
                        "Reminder date cannot be after the due date.",
                        "Invalid Reminder",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    this.DialogResult = DialogResult.None;
                }
            }
        }
    }
} 