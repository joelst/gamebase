//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Resource loader
//==============================================

namespace GBToolbox;

internal static class ToolboxResourceLoader
{
  private static readonly string[] ResourceFiles =
  {
    "Text.ini",
    "Credits.ini",
    "help.htm",
    "Gfx.txt"
  };

  public static string? TryReadResourceFile(string fileName)
  {
    string sourcePath = Path.Combine(ToolboxPaths.AppBasePath, "Source", "GBToolbox", fileName);
    if (File.Exists(sourcePath))
    {
      return File.ReadAllText(sourcePath);
    }

    string repoRelative = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Source", "GBToolbox", fileName);
    string normalized = Path.GetFullPath(repoRelative);
    if (File.Exists(normalized))
    {
      return File.ReadAllText(normalized);
    }

    return null;
  }

  public static void EnsureDefaultFiles(string targetFolder)
  {
    Directory.CreateDirectory(targetFolder);

    foreach (string file in ResourceFiles)
    {
      string targetPath = Path.Combine(targetFolder, file);
      if (File.Exists(targetPath))
      {
        continue;
      }

      string? content = TryReadResourceFile(file);
      if (content == null)
      {
        content = GetFallbackContent(file);
      }

      File.WriteAllText(targetPath, content);
    }
  }

  private static string GetFallbackContent(string fileName)
  {
    return fileName switch
    {
      "help.htm" => "<html><head><title>GameBase Database Help</title></head><body>Help not found.</body></html>",
      "Credits.ini" => "[Help Menu]\nHelpWebSiteText=Database Web Site\nHelpWebsite=http://www.bu22.com/\nHelpEmailText=Email the Database Author\nHelpEmail=putyourown@emailaddresshere.com\nHelpFileText=Database Help\nHelpFile=help.htm\n",
      "Gfx.txt" => "Refer to the GameBase documentation for graphics customization details.\n",
      _ => ""
    };
  }
}
