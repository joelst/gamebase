namespace BAC
{
  partial class MainForm
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
      toolStrip1 = new ToolStrip();
      btnMaster = new ToolStripButton();
      btnDiff = new ToolStripButton();
      toolStrip1.SuspendLayout();
      SuspendLayout();
      //
      // toolStrip1
      //
      toolStrip1.ImageScalingSize = new Size(32, 32);
      toolStrip1.Items.AddRange(new ToolStripItem[] { btnMaster, btnDiff });
      toolStrip1.Location = new Point(0, 0);
      toolStrip1.Name = "toolStrip1";
      toolStrip1.Size = new Size(800, 39);
      toolStrip1.TabIndex = 1;
      toolStrip1.Text = "toolStrip1";
      //
      // btnMaster
      //
      btnMaster.Image = Properties.Resources.master;
      btnMaster.ImageTransparentColor = Color.Lime;
      btnMaster.Name = "btnMaster";
      btnMaster.Size = new Size(81, 36);
      btnMaster.Text = "Master";
      btnMaster.Click += btnMaster_Click;
      //
      // btnDiff
      //
      btnDiff.Image = Properties.Resources.diff;
      btnDiff.ImageTransparentColor = Color.Lime;
      btnDiff.Name = "btnDiff";
      btnDiff.Size = new Size(60, 36);
      btnDiff.Text = "Diff";
      btnDiff.Click += btnDiff_Click;
      //
      // MainForm
      //
      AutoScaleDimensions = new SizeF(7F, 15F);
      AutoScaleMode = AutoScaleMode.Font;
      ClientSize = new Size(800, 450);
      Controls.Add(toolStrip1);
      IsMdiContainer = true;
      Name = "MainForm";
      Text = "Borg Assimilation Console";
      FormClosing += MainForm_FormClosing;
      Load += MainForm_Load;
      toolStrip1.ResumeLayout(false);
      toolStrip1.PerformLayout();
      ResumeLayout(false);
      PerformLayout();
    }

    #endregion

    private ToolStrip toolStrip1;
    private ToolStripButton btnMaster;
    private ToolStripButton btnDiff;
  }
}
