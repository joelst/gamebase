// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: Windows API P/Invoke declarations

using System.Runtime.InteropServices;

namespace GBRegRun;

/// <summary>
/// Native Windows API methods via P/Invoke.
/// </summary>
internal static class NativeMethods
{
  [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
  public static extern int GetShortPathName(
    string lpszLongPath,
    char[] lpszShortPath,
    int cchBuffer);

  /// <summary>
  /// Converts a long path to 8.3 short path format.
  /// </summary>
  public static string GetShortPath(string longPath)
  {
    if (string.IsNullOrEmpty(longPath))
      return string.Empty;

    var shortPath = new char[256];
    int result = GetShortPathName(longPath, shortPath, shortPath.Length);

    if (result > 0 && result < shortPath.Length)
    {
      return new string(shortPath, 0, result);
    }

    return longPath; // Return original if conversion fails
  }
}
