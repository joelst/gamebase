// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: INI configuration file reader

namespace GBRegRun;

/// <summary>
/// Reads and parses the GBRegRun.ini configuration file.
/// </summary>
internal class ConfigurationReader
{
  private const string IniFileName = "GBRegRun.ini";
  private readonly string _appPath;

  public List<Job> Jobs { get; private set; } = new();
  public bool SilentMode { get; private set; }

  public ConfigurationReader(string appPath)
  {
    _appPath = appPath.EndsWith("\\") ? appPath : appPath + "\\";
  }

  /// <summary>
  /// Reads the configuration from GBRegRun.ini file.
  /// </summary>
  /// <returns>True if configuration loaded successfully, false otherwise.</returns>
  public bool LoadConfiguration(out string? errorMessage)
  {
    errorMessage = null;
    Jobs.Clear();
    SilentMode = false;

    string iniPath = Path.Combine(_appPath, IniFileName);

    if (!File.Exists(iniPath))
    {
      errorMessage = $"Configuration file not found: {IniFileName}";
      return false;
    }

    try
    {
      string[] lines = File.ReadAllLines(iniPath);
      string currentSection = string.Empty;

      foreach (string rawLine in lines)
      {
        string line = rawLine.Trim();

        // Skip empty lines and comments
        if (string.IsNullOrEmpty(line) || line.StartsWith(';'))
          continue;

        // Check for section header
        if (line.StartsWith('[') && line.EndsWith(']'))
        {
          currentSection = line.Length > 2 ? line.Substring(1, line.Length - 2) : string.Empty;
          continue;
        }

        // Parse key=value pairs in relevant sections
        if (currentSection == "Files" || currentSection == "Settings")
        {
          int equalPos = line.IndexOf('=');
          if (equalPos <= 0)
            continue;

          string key = line.Substring(0, equalPos).Trim();
          string value = line.Substring(equalPos + 1).Trim();

          if (string.IsNullOrEmpty(value))
            continue;

          if (currentSection == "Files")
          {
            JobType? jobType = key.ToLowerInvariant() switch
            {
              "reg" => JobType.Register,
              "unreg" => JobType.Unregister,
              "run" => JobType.RunNoWait,
              "wait" => JobType.RunAndWait,
              _ => null
            };

            if (jobType.HasValue)
            {
              Jobs.Add(new Job
              {
                Type = jobType.Value,
                RelativePathFile = value
              });
            }
          }
          else if (currentSection == "Settings")
          {
            if (key.Equals("mode", StringComparison.OrdinalIgnoreCase))
            {
              SilentMode = value.Equals("silent", StringComparison.OrdinalIgnoreCase);
            }
          }
        }
      }

      if (Jobs.Count == 0)
      {
        errorMessage = $"No jobs found in {IniFileName}";
        return false;
      }

      // Validate all files exist
      foreach (var job in Jobs)
      {
        string fullPath = Path.Combine(_appPath, job.RelativePathFile);
        if (!File.Exists(fullPath))
        {
          errorMessage = $"File not found: {job.RelativePathFile}";
          return false;
        }
      }

      return true;
    }
    catch (Exception ex)
    {
      errorMessage = $"Error reading configuration: {ex.Message}";
      return false;
    }
  }
}
