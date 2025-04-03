using System;
using System.Windows.Forms;
using System.Drawing;

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

        public int DueDays => chkUseRelativeReminder.Checked ? (int)numDueDays.Value : (dtpDueDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderDays => chkUseRelativeReminder.Checked ? (int)numReminderDays.Value : (dtpReminderDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderHour => (int)numReminderHour.Value;
        public bool UseRelativeReminder => chkUseRelativeReminder.Checked;

        public TaskOptionsForm(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder = false)
        {
            InitializeComponent(defaultDueDays, defaultReminderDays, defaultReminderHour, useRelativeReminder);
        }

        private void InitializeComponent(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder)
        {
            this.Text = "Task Options";
            this.Size = new Size(formWidth, formHeight);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // Due Days/Date
            this.lblDueDays = new Label();
            this.lblDueDays.Text = "Due:";
            this.lblDueDays.Location = new Point(labelX, startY);
            this.lblDueDays.Size = new Size(100, 25);

            this.numDueDays = new NumericUpDown();
            this.numDueDays.Location = new Point(controlX, startY);
            this.numDueDays.Size = new Size(60, 25);
            this.numDueDays.Font = new Font("Segoe UI", 10F);
            this.numDueDays.Minimum = 0;
            this.numDueDays.Maximum = 365;
            this.numDueDays.Value = defaultDueDays;

            this.dtpDueDate = new DateTimePicker();
            this.dtpDueDate.Location = new Point(controlX, startY);
            this.dtpDueDate.Size = new Size(150, 25);
            this.dtpDueDate.Font = new Font("Segoe UI", 10F);
            this.dtpDueDate.Format = DateTimePickerFormat.Short;
            this.dtpDueDate.Value = DateTime.Now.Date.AddDays(defaultDueDays);
            this.dtpDueDate.MinDate = DateTime.Now.Date;

            this.lblDueDaysHelp = new Label();
            this.lblDueDaysHelp.Location = new Point(helpTextX, startY + 3);
            this.lblDueDaysHelp.AutoSize = true;

            // Relative Reminder Checkbox
            this.chkUseRelativeReminder = new CheckBox();
            this.chkUseRelativeReminder.Text = "Use Relative Dates";
            this.chkUseRelativeReminder.Location = new Point(labelX, startY + lineHeight);
            this.chkUseRelativeReminder.Size = new Size(200, 25);
            this.chkUseRelativeReminder.Checked = useRelativeReminder;
            this.chkUseRelativeReminder.CheckedChanged += ChkUseRelativeReminder_CheckedChanged;

            // Reminder Days/Date
            this.lblReminderDays = new Label();
            this.lblReminderDays.Text = "Reminder:";
            this.lblReminderDays.Location = new Point(labelX, startY + lineHeight * 2);
            this.lblReminderDays.Size = new Size(150, 25);

            this.numReminderDays = new NumericUpDown();
            this.numReminderDays.Location = new Point(controlX, startY + lineHeight * 2);
            this.numReminderDays.Size = new Size(60, 25);
            this.numReminderDays.Font = new Font("Segoe UI", 10F);
            this.numReminderDays.Minimum = 0;
            this.numReminderDays.Maximum = 365;
            this.numReminderDays.Value = defaultReminderDays;

            this.dtpReminderDate = new DateTimePicker();
            this.dtpReminderDate.Location = new Point(controlX, startY + lineHeight * 2);
            this.dtpReminderDate.Size = new Size(150, 25);
            this.dtpReminderDate.Font = new Font("Segoe UI", 10F);
            this.dtpReminderDate.Format = DateTimePickerFormat.Short;
            this.dtpReminderDate.Value = DateTime.Now.Date.AddDays(defaultReminderDays);
            this.dtpReminderDate.MinDate = DateTime.Now.Date;

            this.lblReminderDaysHelp = new Label();
            this.lblReminderDaysHelp.Location = new Point(helpTextX, startY + lineHeight * 2 + 3);
            this.lblReminderDaysHelp.AutoSize = true;

            // Reminder Hour
            this.lblReminderHour = new Label();
            this.lblReminderHour.Text = "Reminder Hour:";
            this.lblReminderHour.Location = new Point(labelX, startY + lineHeight * 3);
            this.lblReminderHour.Size = new Size(150, 25);

            this.numReminderHour = new NumericUpDown();
            this.numReminderHour.Location = new Point(controlX, startY + lineHeight * 3);
            this.numReminderHour.Size = new Size(60, 25);
            this.numReminderHour.Font = new Font("Segoe UI", 10F);
            this.numReminderHour.Minimum = 0;
            this.numReminderHour.Maximum = 23;
            this.numReminderHour.Value = defaultReminderHour;

            this.lblReminderHourHelp = new Label();
            this.lblReminderHourHelp.Text = "(24-hour format)";
            this.lblReminderHourHelp.Location = new Point(helpTextX, startY + lineHeight * 3 + 3);
            this.lblReminderHourHelp.AutoSize = true;

            // Buttons
            this.btnOK = new Button();
            this.btnOK.Text = "OK";
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.Location = new Point(formWidth - 180, formHeight - 80);
            this.btnOK.Size = new Size(75, 25);
            this.btnOK.Click += BtnOK_Click;

            this.btnCancel = new Button();
            this.btnCancel.Text = "Cancel";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new Point(formWidth - 90, formHeight - 80);
            this.btnCancel.Size = new Size(75, 25);

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                lblDueDays, numDueDays, dtpDueDate, lblDueDaysHelp,
                chkUseRelativeReminder,
                lblReminderDays, numReminderDays, dtpReminderDate, lblReminderDaysHelp,
                lblReminderHour, numReminderHour, lblReminderHourHelp,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // Initialize visibility based on mode
            UpdateControlsVisibility();
            UpdateHelpText();
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