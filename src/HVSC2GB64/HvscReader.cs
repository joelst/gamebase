// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1997-2004 James Burrows - Migrated to .NET 2026
// HVSC2GB64: HVSC directory and file reader

namespace HVSC2GB64;

/// <summary>
/// Reads HVSC (High Voltage SID Collection) directory structure and SID files.
/// </summary>
internal class HvscReader
{
  private readonly string _hvscPath;

  public HvscReader(string hvscPath)
  {
    _hvscPath = hvscPath.EndsWith("\\") ? hvscPath : hvscPath + "\\";
  }

  /// <summary>
  /// Scans HVSC directory for SID files.
  /// </summary>
  public List<SidData> ReadAllSidFiles()
  {
    var sidFiles = new List<SidData>();

    if (!Directory.Exists(_hvscPath))
    {
      Console.WriteLine($"Error: HVSC path not found: {_hvscPath}");
      return sidFiles;
    }

    Console.WriteLine($"Scanning HVSC directory: {_hvscPath}");

    try
    {
      var files = Directory.GetFiles(_hvscPath, "*.sid", SearchOption.AllDirectories);
      Console.WriteLine($"Found {files.Length} SID files");

      foreach (var file in files)
      {
        try
        {
          var sidData = ParseSidFile(file);
          if (sidData != null)
          {
            sidFiles.Add(sidData);
          }
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Warning: Could not parse {file}: {ex.Message}");
        }
      }
    }
    catch (Exception ex)
    {
      Console.WriteLine($"Error scanning HVSC directory: {ex.Message}");
    }

    return sidFiles;
  }

  /// <summary>
  /// Parses a SID file to extract metadata.
  /// Note: This is a simplified parser. Full SID format parsing would require binary reading.
  /// </summary>
  private SidData? ParseSidFile(string filePath)
  {
    try
    {
      // Get relative path from HVSC root
      string relativePath = filePath.Substring(_hvscPath.Length);
      string filename = Path.GetFileName(filePath);

      // For simplicity, create basic SID data
      // A full implementation would parse the SID file header to extract:
      // - Name, Author, Copyright fields
      // - Number of songs, default song
      var sidData = new SidData
      {
        RelativePath = relativePath,
        Filename = filename,
        Name = Path.GetFileNameWithoutExtension(filename),
        Author = "Unknown", // Would parse from SID header
        Copyright = string.Empty,
        Songs = 1,
        DefaultSong = 1,
        IsGame = false
      };

      return sidData;
    }
    catch
    {
      return null;
    }
  }
}
