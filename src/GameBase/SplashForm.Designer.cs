namespace GameBase;

partial class SplashForm
{
  private System.ComponentModel.IContainer components = null;
  private PictureBox picLogo;
  private Panel pnlBottom;
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
    picLogo = new PictureBox();
    pnlBottom = new Panel();
    lblStatus = new Label();
    progressBar = new ProgressBar();
    ((System.ComponentModel.ISupportInitialize)picLogo).BeginInit();
    pnlBottom.SuspendLayout();
    SuspendLayout();
    //
    // picLogo
    //
    picLogo.BackColor = Color.Magenta;
    picLogo.Dock = DockStyle.Top;
    picLogo.Location = new Point(0, 0);
    picLogo.Name = "picLogo";
    picLogo.Size = new Size(800, 117);
    picLogo.SizeMode = PictureBoxSizeMode.CenterImage;
    picLogo.TabIndex = 0;
    picLogo.TabStop = false;
    //
    // pnlBottom
    //
    pnlBottom.Controls.Add(lblStatus);
    pnlBottom.Controls.Add(progressBar);
    pnlBottom.Dock = DockStyle.Fill;
    pnlBottom.Location = new Point(0, 117);
    pnlBottom.Name = "pnlBottom";
    pnlBottom.Size = new Size(800, 63);
    pnlBottom.TabIndex = 3;
    //
    // lblStatus
    //
    lblStatus.AutoSize = true;
    lblStatus.Location = new Point(12, 10);
    lblStatus.Name = "lblStatus";
    lblStatus.Size = new Size(83, 15);
    lblStatus.TabIndex = 1;
    lblStatus.Text = "Loading...";
    //
    // progressBar
    //
    progressBar.Location = new Point(12, 35);
    progressBar.MarqueeAnimationSpeed = 30;
    progressBar.Name = "progressBar";
    progressBar.Size = new Size(776, 10);
    progressBar.Style = ProgressBarStyle.Marquee;
    progressBar.TabIndex = 2;
    //
    // SplashForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(800, 180);
    ControlBox = false;
    Controls.Add(pnlBottom);
    Controls.Add(picLogo);
    FormBorderStyle = FormBorderStyle.None;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "SplashForm";
    ShowInTaskbar = false;
    StartPosition = FormStartPosition.CenterScreen;
    Text = "GameBase";
    TransparencyKey = Color.Magenta;
    ((System.ComponentModel.ISupportInitialize)picLogo).EndInit();
    pnlBottom.ResumeLayout(false);
    pnlBottom.PerformLayout();
    ResumeLayout(false);
  }
}
