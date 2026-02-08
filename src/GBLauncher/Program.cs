namespace GBLauncher;

internal static class Program
{
  [STAThread]
  static void Main(string[] args)
  {
    Logger? logger = null;

    try
    {
      // Initialize logger
      var appPath = AppDomain.CurrentDomain.BaseDirectory;
      logger = new Logger(appPath);

      // Parse command-line arguments
      if (args.Length == 0)
      {
        CommandLineParser.ShowUsage();
        logger.Log("Empty command line - usage shown", true);
        return;
      }

      var commandLine = string.Join(" ", args);
      logger.Log($"Command line: [{commandLine}]");

      var options = CommandLineParser.Parse(args);

      // Validate options
      if (!options.IsValid())
      {
        if (string.IsNullOrEmpty(options.DatabaseName))
        {
          logger.Log("Missing or invalid database in command line.", true);
        }
        else if (string.IsNullOrEmpty(options.GameId) && string.IsNullOrEmpty(options.GameFilename))
        {
          logger.Log("Missing or invalid id/filename param in command line.", true);
        }
        else if (!string.IsNullOrEmpty(options.GameId) && !string.IsNullOrEmpty(options.GameFilename))
        {
          logger.Log("Please specify only one of id/filename params in command line.", true);
        }
        return;
      }

      logger.Log($"db param: [{options.DatabaseName}]");
      if (!string.IsNullOrEmpty(options.GameId))
        logger.Log($"id param: [{options.GameId}]");
      if (!string.IsNullOrEmpty(options.GameFilename))
        logger.Log($"filename param: [{options.GameFilename}]");
      logger.Log($"emu param: [{options.EmulatorIndex}]" + (options.EmulatorIndex == -1 ? " (default)" : ""));
      logger.Log($"wait param: [{(options.WaitForExit ? "1" : "0")}]");

      // Launch the game
      var launcher = new GameLauncher(options, logger);
      launcher.Launch();

      logger.Log("GBLauncher completed successfully");
    }
    catch (Exception ex)
    {
      logger?.Log($"Unhandled exception: {ex.Message}", true);
      Console.Error.WriteLine($"FATAL ERROR: {ex.Message}");
      Console.Error.WriteLine(ex.StackTrace);
      Environment.Exit(1);
    }
    finally
    {
      logger?.Dispose();
    }
  }
}
