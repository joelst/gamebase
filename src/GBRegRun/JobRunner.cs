// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: Job execution logic

using System.Diagnostics;

namespace GBRegRun;

/// <summary>
/// Executes jobs: DLL registration/unregistration and process launching.
/// </summary>
internal class JobRunner
{
  private readonly string _appPath;
  private readonly bool _silentMode;

  public JobRunner(string appPath, bool silentMode)
  {
    _appPath = appPath.EndsWith("\\") ? appPath : appPath + "\\";
    _silentMode = silentMode;
  }

  /// <summary>
  /// Executes a single job.
  /// </summary>
  public void ExecuteJob(Job job)
  {
    string fullPath = Path.Combine(_appPath, job.RelativePathFile);

    switch (job.Type)
    {
      case JobType.Register:
        RegisterDll(fullPath);
        break;

      case JobType.Unregister:
        UnregisterDll(fullPath);
        break;

      case JobType.RunNoWait:
        RunProcess(fullPath, waitForExit: false);
        break;

      case JobType.RunAndWait:
        RunProcess(fullPath, waitForExit: true);
        break;
    }
  }

  /// <summary>
  /// Registers a DLL using regsvr32.
  /// </summary>
  private void RegisterDll(string dllPath)
  {
    string arguments = _silentMode ? $"/s \"{dllPath}\"" : $"\"{dllPath}\"";
    RunRegSvr32(arguments);
  }

  /// <summary>
  /// Unregisters a DLL using regsvr32.
  /// </summary>
  private void UnregisterDll(string dllPath)
  {
    string arguments = _silentMode ? $"/s /u \"{dllPath}\"" : $"/u \"{dllPath}\"";
    RunRegSvr32(arguments);
  }

  /// <summary>
  /// Runs regsvr32.exe with the specified arguments.
  /// </summary>
  private void RunRegSvr32(string arguments)
  {
    try
    {
      var startInfo = new ProcessStartInfo
      {
        FileName = "regsvr32.exe",
        Arguments = arguments,
        UseShellExecute = false,
        CreateNoWindow = true,
        WindowStyle = ProcessWindowStyle.Hidden
      };

      using var process = Process.Start(startInfo);
      process?.WaitForExit();
    }
    catch (Exception ex)
    {
      if (!_silentMode)
      {
        Console.WriteLine($"Error running regsvr32: {ex.Message}");
      }
    }
  }

  /// <summary>
  /// Launches a process, optionally waiting for it to complete.
  /// </summary>
  private void RunProcess(string executablePath, bool waitForExit)
  {
    try
    {
      var startInfo = new ProcessStartInfo
      {
        FileName = executablePath,
        UseShellExecute = true,
        WindowStyle = ProcessWindowStyle.Normal
      };

      using var process = Process.Start(startInfo);

      if (waitForExit && process != null)
      {
        process.WaitForExit();
      }
    }
    catch (Exception ex)
    {
      if (!_silentMode)
      {
        Console.WriteLine($"Error running process: {ex.Message}");
      }
    }
  }
}
