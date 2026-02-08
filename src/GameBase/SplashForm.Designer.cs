namespace GameBase;

partial class SplashForm
{
  private System.ComponentModel.IContainer components = null;
  private PictureBox picLogo;
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
    lblStatus = new Label();
    progressBar = new ProgressBar();
    ((System.ComponentModel.ISupportInitialize)picLogo).BeginInit();
    SuspendLayout();
    //
    // picLogo
    //
    picLogo.Dock = DockStyle.Top;
    picLogo.Location = new Point(0, 0);
    picLogo.Name = "picLogo";
    picLogo.Size = new Size(800, 117);
    picLogo.SizeMode = PictureBoxSizeMode.CenterImage;
    picLogo.TabIndex = 0;
    picLogo.TabStop = false;
    
    // Load embedded GIF resource
    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
    var resourceName = "GameBase.GameBase.gif";
    using (var stream = assembly.GetManifestResourceStream(resourceName))
    {
      if (stream != null)
      {
        picLogo.Image = Image.FromStream(stream);
      }
    }
    //
    // lblStatus
    //
    lblStatus.AutoSize = true;
    lblStatus.Location = new Point(12, 127);
    lblStatus.Name = "lblStatus";
    lblStatus.Size = new Size(83, 15);
    lblStatus.TabIndex = 1;
    lblStatus.Text = "Loading...";
    //
    // progressBar
    //
    progressBar.Location = new Point(12, 152);
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
    Controls.Add(progressBar);
    Controls.Add(lblStatus);
    Controls.Add(picLogo);
    FormBorderStyle = FormBorderStyle.None;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "SplashForm";
    ShowInTaskbar = false;
    StartPosition = FormStartPosition.CenterScreen;
    Text = "GameBase";
    ((System.ComponentModel.ISupportInitialize)picLogo).EndInit();
    ResumeLayout(false);
    PerformLayout();
  }
}
