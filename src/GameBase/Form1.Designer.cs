namespace GameBase;

partial class Form1
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
    private Button btnOptions;
    private TableLayoutPanel tableLayoutPanel1;
    private SplitContainer splitContainer1;
    private ListBox lstGames;
    private ListBox lstExtras;
    private Label lblStatus;

    private void InitializeComponent()
    {
        btnOptions = new Button();
        tableLayoutPanel1 = new TableLayoutPanel();
        splitContainer1 = new SplitContainer();
        lstGames = new ListBox();
        lstExtras = new ListBox();
        lblStatus = new Label();
        ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
        splitContainer1.Panel1.SuspendLayout();
        splitContainer1.Panel2.SuspendLayout();
        splitContainer1.SuspendLayout();
        tableLayoutPanel1.SuspendLayout();
        SuspendLayout();
        //
        // btnOptions
        //
        btnOptions.Location = new Point(3, 3);
        btnOptions.Name = "btnOptions";
        btnOptions.Size = new Size(120, 23);
        btnOptions.TabIndex = 0;
        btnOptions.Text = "Options";
        btnOptions.UseVisualStyleBackColor = true;
        btnOptions.Click += btnOptions_Click;
        //
        // tableLayoutPanel1
        //
        tableLayoutPanel1.ColumnCount = 1;
        tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel1.Controls.Add(btnOptions, 0, 0);
        tableLayoutPanel1.Controls.Add(splitContainer1, 0, 1);
        tableLayoutPanel1.Controls.Add(lblStatus, 0, 2);
        tableLayoutPanel1.Dock = DockStyle.Fill;
        tableLayoutPanel1.Location = new Point(0, 0);
        tableLayoutPanel1.Name = "tableLayoutPanel1";
        tableLayoutPanel1.RowCount = 3;
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        tableLayoutPanel1.Size = new Size(800, 450);
        tableLayoutPanel1.TabIndex = 1;
        //
        // splitContainer1
        //
        splitContainer1.Dock = DockStyle.Fill;
        splitContainer1.Location = new Point(3, 32);
        splitContainer1.Name = "splitContainer1";
        //
        // splitContainer1.Panel1
        //
        splitContainer1.Panel1.Controls.Add(lstGames);
        //
        // splitContainer1.Panel2
        //
        splitContainer1.Panel2.Controls.Add(lstExtras);
        splitContainer1.Size = new Size(794, 386);
        splitContainer1.SplitterDistance = 397;
        splitContainer1.TabIndex = 1;
        //
        // lstGames
        //
        lstGames.Dock = DockStyle.Fill;
        lstGames.FormattingEnabled = true;
        lstGames.ItemHeight = 15;
        lstGames.Location = new Point(0, 0);
        lstGames.Name = "lstGames";
        lstGames.Size = new Size(397, 386);
        lstGames.TabIndex = 0;
        //
        // lstExtras
        //
        lstExtras.Dock = DockStyle.Fill;
        lstExtras.FormattingEnabled = true;
        lstExtras.ItemHeight = 15;
        lstExtras.Location = new Point(0, 0);
        lstExtras.Name = "lstExtras";
        lstExtras.Size = new Size(393, 386);
        lstExtras.TabIndex = 0;
        //
        // lblStatus
        //
        lblStatus.AutoSize = true;
        lblStatus.Dock = DockStyle.Fill;
        lblStatus.Location = new Point(3, 421);
        lblStatus.Name = "lblStatus";
        lblStatus.Padding = new Padding(3);
        lblStatus.Size = new Size(794, 29);
        lblStatus.TabIndex = 2;
        lblStatus.Text = "Ready";
        //
        // Form1
        //
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(800, 450);
        Controls.Add(tableLayoutPanel1);
        Name = "Form1";
        Text = "GameBase";
        splitContainer1.Panel1.ResumeLayout(false);
        splitContainer1.Panel2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
        splitContainer1.ResumeLayout(false);
        tableLayoutPanel1.ResumeLayout(false);
        tableLayoutPanel1.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
}
