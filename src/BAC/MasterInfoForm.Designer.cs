namespace BAC
{
    partial class MasterInfoForm
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
            picInfo = new Panel();
            lblName = new Label();
            tabControl1 = new TabControl();
            tabGame = new TabPage();
            lblInfo = new Label();
            tabVersion = new TabPage();
            tabNotes = new TabPage();
            picNoInfo = new Panel();
            label1 = new Label();
            picInfo.SuspendLayout();
            tabControl1.SuspendLayout();
            tabGame.SuspendLayout();
            picNoInfo.SuspendLayout();
            SuspendLayout();
            // 
            // picInfo
            // 
            picInfo.Controls.Add(lblName);
            picInfo.Controls.Add(tabControl1);
            picInfo.Dock = DockStyle.Fill;
            picInfo.Location = new Point(0, 0);
            picInfo.Name = "picInfo";
            picInfo.Size = new Size(621, 586);
            picInfo.TabIndex = 0;
            // 
            // lblName
            // 
            lblName.Dock = DockStyle.Top;
            lblName.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblName.Location = new Point(0, 0);
            lblName.Name = "lblName";
            lblName.Size = new Size(621, 32);
            lblName.TabIndex = 1;
            lblName.Text = "NAME GOES HERE";
            lblName.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabGame);
            tabControl1.Controls.Add(tabVersion);
            tabControl1.Controls.Add(tabNotes);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(621, 586);
            tabControl1.TabIndex = 0;
            // 
            // tabGame
            // 
            tabGame.Controls.Add(lblInfo);
            tabGame.Location = new Point(4, 24);
            tabGame.Name = "tabGame";
            tabGame.Padding = new Padding(3);
            tabGame.Size = new Size(613, 558);
            tabGame.TabIndex = 0;
            tabGame.Text = "&Game";
            tabGame.UseVisualStyleBackColor = true;
            // 
            // lblInfo
            // 
            lblInfo.AutoSize = true;
            lblInfo.Location = new Point(6, 3);
            lblInfo.Name = "lblInfo";
            lblInfo.Size = new Size(392, 90);
            lblInfo.TabIndex = 0;
            lblInfo.Text = "Game details would be displayed here.\r\n\r\nFields: ID, Publisher, Year, Programmer, Musician, Genre,\r\nPlayers, Control, Comment, Language, Prequel, Sequel, Related,\r\nAdult\r\n\r\n";
            // 
            // tabVersion
            // 
            tabVersion.Location = new Point(4, 24);
            tabVersion.Name = "tabVersion";
            tabVersion.Padding = new Padding(3);
            tabVersion.Size = new Size(613, 558);
            tabVersion.TabIndex = 1;
            tabVersion.Text = "&Version";
            tabVersion.UseVisualStyleBackColor = true;
            // 
            // tabNotes
            // 
            tabNotes.Location = new Point(4, 24);
            tabNotes.Name = "tabNotes";
            tabNotes.Padding = new Padding(3);
            tabNotes.Size = new Size(613, 558);
            tabNotes.TabIndex = 2;
            tabNotes.Text = "&Notes";
            tabNotes.UseVisualStyleBackColor = true;
            // 
            // picNoInfo
            // 
            picNoInfo.Controls.Add(label1);
            picNoInfo.Dock = DockStyle.Fill;
            picNoInfo.Location = new Point(0, 0);
            picNoInfo.Name = "picNoInfo";
            picNoInfo.Size = new Size(621, 586);
            picNoInfo.TabIndex = 1;
            // 
            // label1
            // 
            label1.Dock = DockStyle.Fill;
            label1.Font = new Font("Segoe UI", 12F);
            label1.Location = new Point(0, 0);
            label1.Name = "label1";
            label1.Size = new Size(621, 586);
            label1.TabIndex = 0;
            label1.Text = "NO GAME SELECTED OR\r\nCHOSEN DIFF GAME ID DOES NOT EXIST\r\nIN THE MASTER DATABASE!";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // MasterInfoForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(621, 586);
            Controls.Add(picNoInfo);
            Controls.Add(picInfo);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "MasterInfoForm";
            Text = "MASTER GAME INFO";
            picInfo.ResumeLayout(false);
            tabControl1.ResumeLayout(false);
            tabGame.ResumeLayout(false);
            tabGame.PerformLayout();
            picNoInfo.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel picInfo;
        private TabControl tabControl1;
        private TabPage tabGame;
        private TabPage tabVersion;
        private TabPage tabNotes;
        private Panel picNoInfo;
        private Label label1;
        private Label lblName;
        private Label lblInfo;
    }
}
