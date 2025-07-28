using ost2pst;

namespace ost2pst
{
    partial class ost2pst
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            folderView = new TreeView();
            openOST = new Button();
            fileDetails = new Label();
            exportPST = new Button();
            pstDetails = new Label();
            label1 = new Label();
            statusList = new ListBox();
            pstGroup = new GroupBox();
            xbRemovePassword = new CheckBox();
            pstGroup.SuspendLayout();
            SuspendLayout();
            // 
            // folderView
            // 
            folderView.Location = new Point(41, 189);
            folderView.Name = "folderView";
            folderView.Size = new Size(255, 358);
            folderView.TabIndex = 0;
            folderView.AfterSelect += treeViewFolders_AfterSelect;
            // 
            // openOST
            // 
            openOST.BackColor = SystemColors.Control;
            openOST.FlatAppearance.BorderColor = SystemColors.ActiveBorder;
            openOST.FlatAppearance.BorderSize = 4;
            openOST.FlatStyle = FlatStyle.Popup;
            openOST.Location = new Point(41, 12);
            openOST.Name = "openOST";
            openOST.Size = new Size(254, 33);
            openOST.TabIndex = 1;
            openOST.Text = "Select OST (or PST) file";
            openOST.UseVisualStyleBackColor = false;
            openOST.Click += openOST_Click;
            // 
            // fileDetails
            // 
            fileDetails.AutoSize = true;
            fileDetails.Font = new Font("Courier New", 9F);
            fileDetails.Location = new Point(41, 58);
            fileDetails.Name = "fileDetails";
            fileDetails.Size = new Size(133, 15);
            fileDetails.TabIndex = 2;
            fileDetails.Text = "input file details";
            // 
            // exportPST
            // 
            exportPST.BackColor = SystemColors.Control;
            exportPST.FlatAppearance.BorderColor = SystemColors.ActiveBorder;
            exportPST.FlatAppearance.BorderSize = 4;
            exportPST.FlatStyle = FlatStyle.Popup;
            exportPST.Location = new Point(39, 601);
            exportPST.Name = "exportPST";
            exportPST.Size = new Size(256, 30);
            exportPST.TabIndex = 3;
            exportPST.Text = "Export to PST";
            exportPST.UseVisualStyleBackColor = false;
            exportPST.Click += exportPST_Click;
            // 
            // pstDetails
            // 
            pstDetails.AutoSize = true;
            pstDetails.Font = new Font("Courier New", 9F);
            pstDetails.Location = new Point(40, 634);
            pstDetails.Name = "pstDetails";
            pstDetails.Size = new Size(140, 15);
            pstDetails.TabIndex = 2;
            pstDetails.Text = "output file details";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label1.ForeColor = SystemColors.ActiveCaptionText;
            label1.Location = new Point(41, 171);
            label1.Name = "label1";
            label1.Size = new Size(100, 15);
            label1.TabIndex = 5;
            label1.Text = "input folder tree";
            // 
            // statusList
            // 
            statusList.BackColor = Color.Black;
            statusList.ForeColor = Color.Lime;
            statusList.FormattingEnabled = true;
            statusList.ItemHeight = 15;
            statusList.Location = new Point(327, 12);
            statusList.Name = "statusList";
            statusList.Size = new Size(442, 784);
            statusList.TabIndex = 5;
            // 
            // pstGroup
            // 
            pstGroup.Controls.Add(xbRemovePassword);
            pstGroup.Location = new Point(43, 547);
            pstGroup.Name = "pstGroup";
            pstGroup.Size = new Size(253, 48);
            pstGroup.TabIndex = 6;
            pstGroup.TabStop = false;
            pstGroup.Text = "only for pst input file";
            // 
            // xbRemovePassword
            // 
            xbRemovePassword.AutoSize = true;
            xbRemovePassword.Location = new Point(6, 22);
            xbRemovePassword.Name = "xbRemovePassword";
            xbRemovePassword.Size = new Size(104, 19);
            xbRemovePassword.TabIndex = 1;
            xbRemovePassword.Text = "reset password";
            xbRemovePassword.UseVisualStyleBackColor = true;
            // 
            // ost2pst
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Info;
            ClientSize = new Size(820, 817);
            Controls.Add(pstGroup);
            Controls.Add(statusList);
            Controls.Add(label1);
            Controls.Add(exportPST);
            Controls.Add(pstDetails);
            Controls.Add(fileDetails);
            Controls.Add(openOST);
            Controls.Add(folderView);
            Name = "ost2pst";
            Text = "read";
            pstGroup.ResumeLayout(false);
            pstGroup.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TreeView folderView;
        private Button openOST;


        private Label fileDetails;
        private Button exportPST;
        private Label pstDetails;
        private Label label1;
        private ListBox statusList;
        private GroupBox pstGroup;
        private CheckBox xbRemovePassword;
    }
}
