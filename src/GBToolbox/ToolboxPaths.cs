//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Path helpers
//==============================================

namespace GBToolbox;

internal static class ToolboxPaths
{
  public static string AppBasePath => AppContext.BaseDirectory;

  public static string EnsureTrailingSeparator(string path)
  {
    if (string.IsNullOrWhiteSpace(path))
    {
      return string.Empty;
    }

    if (path.EndsWith(Path.DirectorySeparatorChar) || path.EndsWith(Path.AltDirectorySeparatorChar))
    {
      return path;
    }

    return path + Path.DirectorySeparatorChar;
  }

  public static string GetGameBaseRoot()
  {
    return EnsureTrailingSeparator(AppBasePath);
  }

  public static string GetDatabaseFile(string baseFolder)
  {
    string folderName = new DirectoryInfo(baseFolder).Name;
    return Path.Combine(baseFolder, folderName + ".db");
  }

  public static string GetIniFolder(string baseFolder)
  {
    return baseFolder;
  }

  public static string GetScriptsFolder(string baseFolder)
  {
    return Path.Combine(baseFolder, "Scripts");
  }

  public static string GetTextFolder(string baseFolder)
  {
    return baseFolder;
  }
}
