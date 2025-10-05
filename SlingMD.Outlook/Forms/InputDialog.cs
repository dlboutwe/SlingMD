using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Reflection;

namespace SlingMD.Outlook.Forms
{
    public partial class InputDialog : Form
    {
        private TextBox txtInput;
        private Button btnOK;
        private Button btnCancel;

        public string InputText => txtInput.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent();
            Text = title;

            // Create and configure controls
            var lblPrompt = new Label
            {
                Text = prompt,
                AutoSize = true,
                Location = new System.Drawing.Point(12, 12)
            };

            txtInput = new TextBox
            {
                Location = new System.Drawing.Point(12, lblPrompt.Bottom + 6),
                Width = 360,
                Text = defaultValue
            };

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(txtInput.Right - 160, txtInput.Bottom + 12),
                Width = 75
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(btnOK.Right + 10, btnOK.Top),
                Width = 75
            };

            // Add controls to form
            Controls.AddRange(new Control[] { lblPrompt, txtInput, btnOK, btnCancel });

            // Set form properties
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            AcceptButton = btnOK;
            CancelButton = btnCancel;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            Padding = new Padding(12);
        }
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InputDialog));
            this.SuspendLayout();
            // 
            // InputDialog
            // 
            this.ClientSize = new System.Drawing.Size(384, 141);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "InputDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.TopMost = true;
            this.ResumeLayout(false);

        }
    }
} 