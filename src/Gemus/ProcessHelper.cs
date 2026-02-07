// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Process management helper using Win32 APIs
// Converted from clsGemus process-related API declarations and methods

using Gemus.Models;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace Gemus
{
  /// <summary>
  /// Provides process management, shell execution, and drive type detection.
  /// Converted from VB6 clsGemus Win32 API usage.
  /// </summary>
  public static class ProcessHelper
  {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern uint GetDriveType(string lpRootPathName);

    private const uint DRIVE_REMOVABLE = 2;
    private const uint DRIVE_FIXED = 3;
    private const uint DRIVE_REMOTE = 4;
    private const uint DRIVE_CDROM = 5;
    private const uint DRIVE_RAMDISK = 6;

    /// <summary>
    /// Determines the type of drive for the given root path.
    /// </summary>
    public static PathLocationType GetDriveLocationType(string rootPath)
    {
      if (string.IsNullOrEmpty(rootPath) || rootPath.Length < 2)
        return PathLocationType.Invalid;

      string root = rootPath.Length >= 3 ? rootPath[..3] : rootPath + @"\";
      uint driveType = GetDriveType(root);

      return driveType switch
      {
        DRIVE_REMOVABLE => PathLocationType.Removable,
        DRIVE_FIXED => PathLocationType.FixedHD,
        DRIVE_CDROM => PathLocationType.CD,
        DRIVE_REMOTE => PathLocationType.FixedHD, // treat network as fixed
        DRIVE_RAMDISK => PathLocationType.FixedHD, // treat ramdisk as fixed
        _ => PathLocationType.Invalid
      };
    }

    /// <summary>
    /// Checks if a drive is a CD/DVD-ROM.
    /// </summary>
    public static bool IsCdRomDrive(string rootPath)
    {
      return GetDriveLocationType(rootPath) == PathLocationType.CD;
    }

    /// <summary>
    /// Launches a process and optionally waits for it to exit.
    /// </summary>
    public static (bool success, Process? process) LaunchProcess(
        string fileName, string arguments = "", string? workingDirectory = null,
        bool waitForExit = false, ProcessWindowStyle windowStyle = ProcessWindowStyle.Normal)
    {
      try
      {
        var startInfo = new ProcessStartInfo
        {
          FileName = fileName,
          Arguments = arguments,
          UseShellExecute = true,
          WindowStyle = windowStyle
        };

        if (!string.IsNullOrEmpty(workingDirectory))
        {
          startInfo.WorkingDirectory = workingDirectory;
        }

        var process = Process.Start(startInfo);

        if (waitForExit && process != null)
        {
          process.WaitForExit();
        }

        return (true, process);
      }
      catch
      {
        return (false, null);
      }
    }

    /// <summary>
    /// Launches a file with its default shell association (e.g., open a .txt in Notepad).
    /// </summary>
    public static (bool success, Process? process) ShellExecuteFile(
        string filePath, string arguments = "", bool waitForExit = false)
    {
      return LaunchProcess(filePath, arguments, null, waitForExit);
    }

    /// <summary>
    /// Kills all running processes matching the given filename.
    /// </summary>
    public static int KillProcessesByFilename(string filename)
    {
      int killed = 0;
      string upperFilename = filename.ToUpperInvariant();

      try
      {
        var processes = Process.GetProcesses();
        foreach (var proc in processes)
        {
          try
          {
            string procName = proc.ProcessName.ToUpperInvariant();
            if (procName == upperFilename ||
                procName == Path.GetFileNameWithoutExtension(upperFilename))
            {
              proc.Kill();
              killed++;
            }
          }
          catch
          {
            // Process may have exited or access denied
          }
          finally
          {
            proc.Dispose();
          }
        }
      }
      catch
      {
        // Ignore enumeration errors
      }

      return killed;
    }

    /// <summary>
    /// Waits for a process to exit, with a timeout in milliseconds.
    /// Returns true if the process exited, false on timeout.
    /// </summary>
    public static bool WaitForProcessExit(Process process, int timeoutMs = -1)
    {
      try
      {
        if (timeoutMs < 0)
        {
          process.WaitForExit();
          return true;
        }
        return process.WaitForExit(timeoutMs);
      }
      catch
      {
        return true; // Process already exited
      }
    }

    /// <summary>
    /// Waits for a specified number of milliseconds while checking if a process is still running.
    /// </summary>
    public static void WaitWhileProcessRunning(int milliseconds, Process process)
    {
      var stopwatch = System.Diagnostics.Stopwatch.StartNew();
      while (stopwatch.ElapsedMilliseconds < milliseconds)
      {
        try
        {
          if (process.HasExited)
            return;
        }
        catch
        {
          return; // Process no longer accessible
        }
        Thread.Sleep(50);
      }
    }

    /// <summary>
    /// Gets file path and directory from a full path.
    /// </summary>
    public static (string directory, string fileName) SplitPath(string fullPath)
    {
      string dir = Path.GetDirectoryName(fullPath) ?? string.Empty;
      string file = Path.GetFileName(fullPath);
      return (dir, file);
    }
  }
}
