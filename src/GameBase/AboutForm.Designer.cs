namespace GameBase;

partial class AboutForm
{
  private System.ComponentModel.IContainer components = null;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  #region Windows Form Designer generated code

  private void InitializeComponent()
  {
    lblTitle = new Label();
    lblVersion = new Label();
    lblCopyright = new Label();
    txtAbout = new TextBox();
    btnOk = new Button();
    SuspendLayout();
    //
    // lblTitle
    //
    lblTitle.AutoSize = true;
    lblTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
    lblTitle.Location = new Point(12, 15);
    lblTitle.Name = "lblTitle";
    lblTitle.Size = new Size(105, 25);
    lblTitle.TabIndex = 0;
    lblTitle.Text = "GameBase";
    //
    // lblVersion
    //
    lblVersion.AutoSize = true;
    lblVersion.Location = new Point(12, 50);
    lblVersion.Name = "lblVersion";
    lblVersion.Size = new Size(70, 15);
    lblVersion.TabIndex = 1;
    lblVersion.Text = "Version 1.0.0";
    //
    // lblCopyright
    //
    lblCopyright.AutoSize = true;
    lblCopyright.Location = new Point(12, 75);
    lblCopyright.Name = "lblCopyright";
    lblCopyright.Size = new Size(200, 15);
    lblCopyright.TabIndex = 2;
    lblCopyright.Text = "Copyright © 1999-2010 James Burrows";
    //
    // txtAbout
    //
    txtAbout.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
    txtAbout.Location = new Point(12, 110);
    txtAbout.Multiline = true;
    txtAbout.Name = "txtAbout";
    txtAbout.ReadOnly = true;
    txtAbout.ScrollBars = ScrollBars.Vertical;
    txtAbout.Size = new Size(560, 280);
    txtAbout.TabIndex = 3;
    //
    // btnOk
    //
    btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnOk.Location = new Point(497, 407);
    btnOk.Name = "btnOk";
    btnOk.Size = new Size(75, 23);
    btnOk.TabIndex = 4;
    btnOk.Text = "OK";
    btnOk.UseVisualStyleBackColor = true;
    btnOk.Click += btnOk_Click;
    //
    // AboutForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(584, 442);
    Controls.Add(btnOk);
    Controls.Add(txtAbout);
    Controls.Add(lblCopyright);
    Controls.Add(lblVersion);
    Controls.Add(lblTitle);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "AboutForm";
    ShowIcon = false;
    ShowInTaskbar = false;
    StartPosition = FormStartPosition.CenterParent;
    Text = "About GameBase";
    ResumeLayout(false);
    PerformLayout();
  }

  #endregion

  private Label lblTitle;
  private Label lblVersion;
  private Label lblCopyright;
  private TextBox txtAbout;
  private Button btnOk;
}
