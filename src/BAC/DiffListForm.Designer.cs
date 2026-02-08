namespace BAC
{
    partial class DiffListForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            panelInfo = new Panel();
            lblDate = new Label();
            lblEmail = new Label();
            lblFrom = new Label();
            label3 = new Label();
            label2 = new Label();
            label1 = new Label();
            lvGames = new ListView();
            columnName = new ColumnHeader();
            statusStrip1 = new StatusStrip();
            statusLabel = new ToolStripStatusLabel();
            panelInfo.SuspendLayout();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // panelInfo
            // 
            panelInfo.Controls.Add(lblDate);
            panelInfo.Controls.Add(lblEmail);
            panelInfo.Controls.Add(lblFrom);
            panelInfo.Controls.Add(label3);
            panelInfo.Controls.Add(label2);
            panelInfo.Controls.Add(label1);
            panelInfo.Dock = DockStyle.Top;
            panelInfo.Location = new Point(0, 0);
            panelInfo.Name = "panelInfo";
            panelInfo.Size = new Size(517, 96);
            panelInfo.TabIndex = 0;
            // 
            // lblDate
            // 
            lblDate.AutoSize = true;
            lblDate.Location = new Point(64, 64);
            lblDate.Name = "lblDate";
            lblDate.Size = new Size(0, 15);
            lblDate.TabIndex = 5;
            // 
            // lblEmail
            // 
            lblEmail.AutoSize = true;
            lblEmail.Cursor = Cursors.Hand;
            lblEmail.Font = new Font("Segoe UI", 9F, FontStyle.Underline);
            lblEmail.ForeColor = Color.Maroon;
            lblEmail.Location = new Point(64, 32);
            lblEmail.Name = "lblEmail";
            lblEmail.Size = new Size(0, 15);
            lblEmail.TabIndex = 4;
            lblEmail.Click += lblEmail_Click;
            // 
            // lblFrom
            // 
            lblFrom.AutoSize = true;
            lblFrom.Location = new Point(64, 0);
            lblFrom.Name = "lblFrom";
            lblFrom.Size = new Size(0, 15);
            lblFrom.TabIndex = 3;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(0, 64);
            label3.Name = "label3";
            label3.Size = new Size(34, 15);
            label3.TabIndex = 2;
            label3.Text = "Date:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(0, 32);
            label2.Name = "label2";
            label2.Size = new Size(39, 15);
            label2.TabIndex = 1;
            label2.Text = "Email:";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(0, 0);
            label1.Name = "label1";
            label1.Size = new Size(38, 15);
            label1.TabIndex = 0;
            label1.Text = "From:";
            // 
            // lvGames
            // 
            lvGames.Columns.AddRange(new ColumnHeader[] { columnName });
            lvGames.Dock = DockStyle.Fill;
            lvGames.FullRowSelect = true;
            lvGames.Location = new Point(0, 96);
            lvGames.Name = "lvGames";
            lvGames.Size = new Size(517, 235);
            lvGames.TabIndex = 1;
            lvGames.UseCompatibleStateImageBehavior = false;
            lvGames.View = View.Details;
            lvGames.SelectedIndexChanged += lvGames_SelectedIndexChanged;
            // 
            // columnName
            // 
            columnName.Text = "Name";
            columnName.Width = 467;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { statusLabel });
            statusStrip1.Location = new Point(0, 331);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(517, 22);
            statusStrip1.TabIndex = 2;
            statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(0, 17);
            // 
            // DiffListForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(517, 353);
            Controls.Add(lvGames);
            Controls.Add(statusStrip1);
            Controls.Add(panelInfo);
            MaximizeBox = false;
            Name = "DiffListForm";
            Text = "DIFF";
            FormClosing += DiffListForm_FormClosing;
            Load += DiffListForm_Load;
            Resize += DiffListForm_Resize;
            panelInfo.ResumeLayout(false);
            panelInfo.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Panel panelInfo;
        private Label lblDate;
        private Label lblEmail;
        private Label lblFrom;
        private Label label3;
        private Label label2;
        private Label label1;
        private ListView lvGames;
        private ColumnHeader columnName;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel statusLabel;
    }
}
