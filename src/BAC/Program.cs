namespace BAC;

internal static class Program
{
  /// <summary>
  ///  The main entry point for the application.
  /// </summary>
  [STAThread]
  static void Main()
  {
    ApplicationConfiguration.Initialize();

    // Show splash screen
    using (var splash = new SplashForm())
    {
      splash.ShowDialog();
    }

    // Run main form
    Application.Run(new MainForm());
  }
}
