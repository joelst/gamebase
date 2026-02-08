namespace GameBase;

public partial class SplashForm : Form
{
  public SplashForm()
  {
    InitializeComponent();
    ApplySystemTheme();
    LoadLogo();
  }

  private void ApplySystemTheme()
  {
    // Detect if system is using dark mode
    bool isDarkMode = IsDarkModeEnabled();

    if (isDarkMode)
    {
      // Dark theme
      pnlBottom.BackColor = Color.FromArgb(32, 32, 32);
      lblStatus.ForeColor = Color.White;
    }
    else
    {
      // Light theme
      pnlBottom.BackColor = Color.FromArgb(240, 240, 240);
      lblStatus.ForeColor = Color.Black;
    }
  }

  private bool IsDarkModeEnabled()
  {
    try
    {
      // Check Windows Registry for dark mode setting
      using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
        @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");

      if (key != null)
      {
        var appsUseLightTheme = key.GetValue("AppsUseLightTheme");
        if (appsUseLightTheme != null)
        {
          return (int)appsUseLightTheme == 0; // 0 = dark mode, 1 = light mode
        }
      }
    }
    catch
    {
      // If registry access fails, fall back to system colors
    }

    // Fallback: check if system colors suggest dark theme
    var backColor = SystemColors.Window;
    return (backColor.R + backColor.G + backColor.B) < 382; // Average < 127
  }

  private void LoadLogo()
  {
    try
    {
      var assembly = System.Reflection.Assembly.GetExecutingAssembly();
      var resourceName = "GameBase.GameBase.gif";
      using var stream = assembly.GetManifestResourceStream(resourceName);
      if (stream != null)
      {
        var image = Image.FromStream(stream);

        // Make the image background transparent if it's a GIF
        if (image is Bitmap bitmap)
        {
          bitmap.MakeTransparent(Color.White); // Make white background transparent
        }

        picLogo.Image = image;
      }
    }
    catch
    {
      // If logo fails to load, continue without it
    }
  }
}
