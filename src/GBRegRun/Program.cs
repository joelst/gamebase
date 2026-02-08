// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: Console application entry point

namespace GBRegRun;

internal class Program
{
  private const string AppName = "GBRegRun";

  static int Main(string[] args)
  {
    try
    {
      // Get application path
      string appPath = AppContext.BaseDirectory;

      // Load configuration
      var config = new ConfigurationReader(appPath);
      if (!config.LoadConfiguration(out string? errorMessage))
      {
        if (!config.SilentMode)
        {
          Console.WriteLine($"{AppName}: {errorMessage}");
        }
        return 1;
      }

      if (config.Jobs.Count == 0)
      {
        if (!config.SilentMode)
        {
          Console.WriteLine($"{AppName}: No jobs to execute");
        }
        return 1;
      }

      // Execute jobs
      var runner = new JobRunner(appPath, config.SilentMode);
      foreach (var job in config.Jobs)
      {
        runner.ExecuteJob(job);
      }

      return 0;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"{AppName}: Unexpected error - {ex.Message}");
      return 1;
    }
  }
}
