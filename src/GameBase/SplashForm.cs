namespace GameBase;

public partial class SplashForm : Form
{
  public SplashForm()
  {
    InitializeComponent();
    LoadLogo();
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
        picLogo.Image = Image.FromStream(stream);
      }
    }
    catch
    {
      // If logo fails to load, continue without it
    }
  }
}
