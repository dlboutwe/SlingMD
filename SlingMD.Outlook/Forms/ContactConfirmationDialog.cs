using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ContactConfirmationDialog : Form
    {
        private CheckedListBox chkListContacts;
        private Label lblInfo;
        private Panel btnPanel;
        private Button btnOk;
        private Button btnCancel;
        private Button btnSelectAll;

        public List<string> SelectedContacts { get; private set; } = new List<string>();
        private List<string> _contacts;

        public ContactConfirmationDialog(List<string> contacts)
        {
            InitializeComponent();
            _contacts = contacts;
            PopulateContactList();
        }

        private void InitializeComponent()
        {
            this.chkListContacts = new System.Windows.Forms.CheckedListBox();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnPanel = new System.Windows.Forms.Panel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkListContacts
            // 
            this.chkListContacts.CheckOnClick = true;
            this.chkListContacts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chkListContacts.FormattingEnabled = true;
            this.chkListContacts.Location = new System.Drawing.Point(0, 23);
            this.chkListContacts.Name = "chkListContacts";
            this.chkListContacts.Size = new System.Drawing.Size(400, 277);
            this.chkListContacts.TabIndex = 0;
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblInfo.Location = new System.Drawing.Point(0, 0);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Padding = new System.Windows.Forms.Padding(5);
            this.lblInfo.Size = new System.Drawing.Size(183, 23);
            this.lblInfo.TabIndex = 1;
            this.lblInfo.Text = "Select contacts to create notes for:";
            // 
            // btnPanel
            // 
            this.btnPanel.Controls.Add(this.btnOk);
            this.btnPanel.Controls.Add(this.btnCancel);
            this.btnPanel.Controls.Add(this.btnSelectAll);
            this.btnPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnPanel.Location = new System.Drawing.Point(0, 300);
            this.btnPanel.Name = "btnPanel";
            this.btnPanel.Size = new System.Drawing.Size(400, 50);
            this.btnPanel.TabIndex = 2;
            // 
            // btnOk
            // 
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOk.Location = new System.Drawing.Point(120, 10);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(120, 30);
            this.btnOk.TabIndex = 0;
            this.btnOk.Text = "Create Selected";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(250, 10);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 30);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(10, 10);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(100, 30);
            this.btnSelectAll.TabIndex = 2;
            this.btnSelectAll.Text = "Select All";
            // 
            // ContactConfirmationDialog
            // 
            this.AcceptButton = this.btnOk;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(400, 350);
            this.Controls.Add(this.chkListContacts);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.btnPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ContactConfirmationDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create Contact Notes";
            this.TopMost = true;
            this.btnPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void PopulateContactList()
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                chkListContacts.Items.Clear();
                foreach (var contact in _contacts)
                {
                    chkListContacts.Items.Add(contact, true);
                }
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                SelectedContacts.Clear();
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    if (chkListContacts.GetItemChecked(i))
                    {
                        SelectedContacts.Add(chkListContacts.Items[i].ToString());
                    }
                }
            }
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                bool anyUnchecked = false;
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    if (!chkListContacts.GetItemChecked(i))
                    {
                        anyUnchecked = true;
                        break;
                    }
                }

                // Toggle all items
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    chkListContacts.SetItemChecked(i, anyUnchecked);
                }
            }
        }
    }
} 