namespace BAC
{
  partial class SplashForm
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
      lblAppName = new Label();
      lblVersion = new Label();
      SuspendLayout();
      //
      // lblAppName
      //
      lblAppName.AutoSize = true;
      lblAppName.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
      lblAppName.Location = new Point(12, 80);
      lblAppName.Name = "lblAppName";
      lblAppName.Size = new Size(376, 32);
      lblAppName.TabIndex = 0;
      lblAppName.Text = "Borg Assimilation Console";
      //
      // lblVersion
      //
      lblVersion.AutoSize = true;
      lblVersion.Location = new Point(150, 120);
      lblVersion.Name = "lblVersion";
      lblVersion.Size = new Size(97, 15);
      lblVersion.TabIndex = 1;
      lblVersion.Text = "GameBase Utility";
      //
      // SplashForm
      //
      AutoScaleDimensions = new SizeF(7F, 15F);
      AutoScaleMode = AutoScaleMode.Font;
      BackColor = SystemColors.Window;
      ClientSize = new Size(400, 200);
      Controls.Add(lblVersion);
      Controls.Add(lblAppName);
      FormBorderStyle = FormBorderStyle.None;
      Name = "SplashForm";
      StartPosition = FormStartPosition.CenterScreen;
      Text = "BAC";
      Load += SplashForm_Load;
      ResumeLayout(false);
      PerformLayout();
    }

    #endregion

    private Label lblAppName;
    private Label lblVersion;
  }
}
