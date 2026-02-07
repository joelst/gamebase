//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBFuncs:  General Functions used by GB
//==============================================

using System.Runtime.InteropServices;
using System.Text;

namespace GBFuncs
{
  /// <summary>
  /// Provides general utility functions for GameBase.
  /// Converted from VB6 clsFunctions class.
  /// </summary>
  public class FunctionsHelper
  {
    private string _appPathWithBackslash = string.Empty;
    private string _messagesFile = string.Empty;

    // Event for file finding
    public event Action<string>? FoundFile;

    #region INI File Operations

    // P/Invoke declarations for INI file operations (maintains compatibility with existing .ini files)
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern int GetPrivateProfileString(
        string lpAppName,
        string lpKeyName,
        string lpDefault,
        StringBuilder lpReturnedString,
        int nSize,
        string lpFileName);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern int WritePrivateProfileString(
        string lpAppName,
        string? lpKeyName,
        string? lpString,
        string lpFileName);

    /// <summary>
    /// Gets a setting from an INI file.
    /// </summary>
    public string INI_GetSetting(string section, string key, string defaultValue, string iniFile)
    {
      StringBuilder result = new StringBuilder(2048);
      int length = GetPrivateProfileString(section, key, defaultValue, result, result.Capacity, iniFile);
      return result.ToString();
    }

    /// <summary>
    /// Writes a setting to an INI file.
    /// </summary>
    public bool INI_WriteSetting(string section, string key, string value, string iniFile)
    {
      return WritePrivateProfileString(section, key, value, iniFile) != 0;
    }

    /// <summary>
    /// Deletes a key from an INI file.
    /// </summary>
    public void INI_DeleteKey(string section, string key, string iniFile)
    {
      WritePrivateProfileString(section, key, null, iniFile);
    }

    /// <summary>
    /// Deletes an entire section from an INI file.
    /// </summary>
    public void INI_DeleteSection(string section, string iniFile)
    {
      WritePrivateProfileString(section, null, null, iniFile);
    }

    /// <summary>
    /// Gets a message from the GBText INI file.
    /// </summary>
    public string GetINIMessage(string section, string key, string defaultValue)
    {
      if (string.IsNullOrEmpty(_messagesFile))
        return defaultValue;

      string message = INI_GetSetting(section, key, defaultValue, _messagesFile);
      return RTrimWhitespace(message);
    }

    #endregion

    #region Path and File Operations

    /// <summary>
    /// Separates a full path and filename into its components.
    /// </summary>
    public void GetFileAndPath(string pathAndFile, out string path, out string filename)
    {
      if (string.IsNullOrEmpty(pathAndFile))
      {
        path = string.Empty;
        filename = string.Empty;
        return;
      }

      int lastSlash = Math.Max(pathAndFile.LastIndexOf('\\'), pathAndFile.LastIndexOf('/'));

      if (lastSlash >= 0)
      {
        path = pathAndFile.Substring(0, lastSlash + 1);
        filename = pathAndFile.Substring(lastSlash + 1);
      }
      else
      {
        path = string.Empty;
        filename = pathAndFile;
      }
    }

    /// <summary>
    /// Ensures a folder path ends with a backslash.
    /// </summary>
    private string GetFolderWithBackslash(string path)
    {
      if (string.IsNullOrEmpty(path))
        return string.Empty;

      if (!path.EndsWith("\\") && !path.EndsWith("/"))
        return path + "\\";

      return path;
    }

    /// <summary>
    /// Gets the size of a file in bytes.
    /// </summary>
    public long GetBigFileSize(string filePath)
    {
      try
      {
        FileInfo fileInfo = new FileInfo(filePath);
        return fileInfo.Exists ? fileInfo.Length : 0;
      }
      catch
      {
        return 0;
      }
    }

    /// <summary>
    /// Creates a file with specified contents.
    /// </summary>
    public bool CreateFileWithContents(string filePath, string contents, out string error)
    {
      error = string.Empty;
      try
      {
        File.WriteAllText(filePath, contents);
        return true;
      }
      catch (Exception ex)
      {
        error = ex.Message;
        return false;
      }
    }

    /// <summary>
    /// Finds files matching a pattern in a directory.
    /// </summary>
    public void FindFiles(string path, string searchPattern)
    {
      try
      {
        string searchPath = GetFolderWithBackslash(path);

        // Search for files
        foreach (string file in Directory.GetFiles(searchPath, searchPattern))
        {
          FoundFile?.Invoke(file);
        }

        // Recursively search subdirectories
        foreach (string directory in Directory.GetDirectories(searchPath))
        {
          FindFiles(directory, searchPattern);
        }
      }
      catch
      {
        // Silently ignore errors (matching VB6 behavior)
      }
    }

    /// <summary>
    /// Finds a specific file in a directory tree.
    /// </summary>
    public string FindFile(string filename, string searchPath)
    {
      try
      {
        // Search current directory
        string fullPath = Path.Combine(searchPath, filename);
        if (File.Exists(fullPath))
          return fullPath;

        // Search subdirectories
        foreach (string directory in Directory.GetDirectories(searchPath))
        {
          string result = FindFile(filename, directory);
          if (!string.IsNullOrEmpty(result))
            return result;
        }
      }
      catch
      {
        // Silently ignore errors
      }

      return string.Empty;
    }

    #endregion

    #region String Utilities

    /// <summary>
    /// Removes whitespace from the right side of a string.
    /// </summary>
    public string RTrimWhitespace(string input)
    {
      if (string.IsNullOrEmpty(input))
        return string.Empty;

      return input.TrimEnd();
    }

    /// <summary>
    /// Strips null characters from a string.
    /// </summary>
    private string StripNull(string input)
    {
      if (string.IsNullOrEmpty(input))
        return string.Empty;

      int nullIndex = input.IndexOf('\0');
      return nullIndex >= 0 ? input.Substring(0, nullIndex) : input;
    }

    #endregion

    #region Configuration

    /// <summary>
    /// Sets the path to the GBText INI file for messages.
    /// </summary>
    public void SetGBTextINIFile(string iniFilePath)
    {
      _messagesFile = iniFilePath;
    }

    /// <summary>
    /// Sets the application path.
    /// </summary>
    public void SetAppPath(string appPath)
    {
      _appPathWithBackslash = GetFolderWithBackslash(appPath);
    }

    /// <summary>
    /// Gets the application path with trailing backslash.
    /// </summary>
    public string AppPath => _appPathWithBackslash;

    #endregion

    #region Shell Operations

    /// <summary>
    /// Executes a file using the shell.
    /// </summary>
    public bool ShellExecuteFile(string filePath, string commandLineParams, bool wait)
    {
      try
      {
        System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
        {
          FileName = filePath,
          Arguments = commandLineParams,
          UseShellExecute = true
        };

        using (var process = System.Diagnostics.Process.Start(psi))
        {
          if (process != null && wait)
          {
            process.WaitForExit();
          }
          return true;
        }
      }
      catch
      {
        return false;
      }
    }

    #endregion

    #region Debug

    /// <summary>
    /// Writes a message to the debug log.
    /// </summary>
    public void WriteToDebugLog(string message)
    {
      try
      {
        string logFile = Path.Combine(_appPathWithBackslash, "debug.log");
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string logEntry = $"[{timestamp}] {message}{Environment.NewLine}";
        File.AppendAllText(logFile, logEntry);
      }
      catch
      {
        // Silently ignore errors in logging
      }
    }

    #endregion
  }
}
