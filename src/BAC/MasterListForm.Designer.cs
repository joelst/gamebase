namespace BAC
{
  partial class MasterListForm
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
      lvGames = new ListView();
      columnName = new ColumnHeader();
      statusStrip1 = new StatusStrip();
      statusLabel = new ToolStripStatusLabel();
      statusStrip1.SuspendLayout();
      SuspendLayout();
      //
      // lvGames
      //
      lvGames.Columns.AddRange(new ColumnHeader[] { columnName });
      lvGames.Dock = DockStyle.Fill;
      lvGames.FullRowSelect = true;
      lvGames.Location = new Point(0, 0);
      lvGames.Name = "lvGames";
      lvGames.Size = new Size(517, 331);
      lvGames.TabIndex = 0;
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
      statusStrip1.TabIndex = 1;
      statusStrip1.Text = "statusStrip1";
      //
      // statusLabel
      //
      statusLabel.Name = "statusLabel";
      statusLabel.Size = new Size(0, 17);
      //
      // MasterListForm
      //
      AutoScaleDimensions = new SizeF(7F, 15F);
      AutoScaleMode = AutoScaleMode.Font;
      ClientSize = new Size(517, 353);
      Controls.Add(lvGames);
      Controls.Add(statusStrip1);
      MaximizeBox = false;
      Name = "MasterListForm";
      Text = "Master DB List";
      FormClosing += MasterListForm_FormClosing;
      Load += MasterListForm_Load;
      Resize += MasterListForm_Resize;
      statusStrip1.ResumeLayout(false);
      statusStrip1.PerformLayout();
      ResumeLayout(false);
      PerformLayout();
    }

    #endregion

    private ListView lvGames;
    private ColumnHeader columnName;
    private StatusStrip statusStrip1;
    private ToolStripStatusLabel statusLabel;
  }
}
