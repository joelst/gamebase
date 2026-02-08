namespace BAC;

/// <summary>
/// Splash screen shown during application startup.
/// </summary>
public partial class SplashForm : Form
{
  private Timer? _timer;
  private readonly DateTime _startTime;

  public SplashForm()
  {
    InitializeComponent();
    _startTime = DateTime.Now;
  }

  private void SplashForm_Load(object? sender, EventArgs e)
  {
    // Show splash for 1.75 seconds minimum
    _timer = new Timer
    {
      Interval = 100 // Check every 100ms
    };
    _timer.Tick += Timer_Tick;
    _timer.Start();
  }

  private void Timer_Tick(object? sender, EventArgs e)
  {
    var elapsed = (DateTime.Now - _startTime).TotalSeconds;
    if (elapsed >= 1.75)
    {
      _timer?.Stop();
      _timer?.Dispose();
      Close();
    }
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    _timer?.Stop();
    _timer?.Dispose();
    base.OnFormClosing(e);
  }
}
