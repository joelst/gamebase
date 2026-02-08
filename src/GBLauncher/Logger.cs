namespace GBLauncher;

/// <summary>
/// Logging utility for GBLauncher.
/// </summary>
internal class Logger : IDisposable
{
  private readonly StreamWriter _writer;
  private readonly string _logFilePath;

  public Logger(string appPath)
  {
    _logFilePath = Path.Combine(appPath, "GBLauncher.log");

    // Delete existing log file (always start with empty log)
    if (File.Exists(_logFilePath))
    {
      try { File.Delete(_logFilePath); } catch { }
    }

    _writer = new StreamWriter(_logFilePath, false);
    _writer.WriteLine($"***** GameBase Launcher v1.0 *****");
    _writer.WriteLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    _writer.Flush();
  }

  public void Log(string message, bool forceQuit = false)
  {
    _writer.WriteLine($"--> {message}");
    _writer.Flush();

    if (forceQuit)
    {
      Console.Error.WriteLine($"ERROR: {message}");
      Dispose();
      Environment.Exit(1);
    }
  }

  public void Dispose()
  {
    _writer?.Dispose();
  }
}
