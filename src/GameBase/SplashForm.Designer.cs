namespace GameBase;

partial class SplashForm
{
  private System.ComponentModel.IContainer components = null;
  private Label lblTitle;
  private Label lblStatus;
  private ProgressBar progressBar;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  private void InitializeComponent()
  {
    lblTitle = new Label();
    lblStatus = new Label();
    progressBar = new ProgressBar();
    SuspendLayout();
    //
    // lblTitle
    //
    lblTitle.AutoSize = true;
    lblTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
    lblTitle.Location = new Point(12, 12);
    lblTitle.Name = "lblTitle";
    lblTitle.Size = new Size(105, 25);
    lblTitle.TabIndex = 0;
    lblTitle.Text = "GameBase";
    //
    // lblStatus
    //
    lblStatus.AutoSize = true;
    lblStatus.Location = new Point(12, 45);
    lblStatus.Name = "lblStatus";
    lblStatus.Size = new Size(83, 15);
    lblStatus.TabIndex = 1;
    lblStatus.Text = "Loading...";
    //
    // progressBar
    //
    progressBar.Location = new Point(12, 70);
    progressBar.MarqueeAnimationSpeed = 30;
    progressBar.Name = "progressBar";
    progressBar.Size = new Size(260, 10);
    progressBar.Style = ProgressBarStyle.Marquee;
    progressBar.TabIndex = 2;
    //
    // SplashForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(284, 100);
    ControlBox = false;
    Controls.Add(progressBar);
    Controls.Add(lblStatus);
    Controls.Add(lblTitle);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "SplashForm";
    ShowInTaskbar = false;
    StartPosition = FormStartPosition.CenterScreen;
    Text = "GameBase";
    ResumeLayout(false);
    PerformLayout();
  }
}
