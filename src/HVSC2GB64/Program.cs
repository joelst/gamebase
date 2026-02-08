// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1997-2004 James Burrows - Migrated to .NET 2026
// HVSC2GB64: Console application entry point

namespace HVSC2GB64;

internal class Program
{
  private const string AppName = "HVSC2GB64";
  private const string Copyright = "(c) 2003-2026 James Burrows";

  static int Main(string[] args)
  {
    Console.WriteLine($"{AppName} - High Voltage SID Collection to GameBase64 Updater");
    Console.WriteLine(Copyright);
    Console.WriteLine();

    try
    {
      // Parse command-line arguments
      if (args.Length < 2)
      {
        ShowUsage();
        return 1;
      }

      string hvscPath = args[0];
      string databasePath = args[1];

      // Validate paths
      if (!Directory.Exists(hvscPath))
      {
        Console.WriteLine($"Error: HVSC path not found: {hvscPath}");
        return 1;
      }

      if (!File.Exists(databasePath))
      {
        Console.WriteLine($"Error: Database not found: {databasePath}");
        return 1;
      }

      Console.WriteLine($"HVSC Path: {hvscPath}");
      Console.WriteLine($"Database: {databasePath}");
      Console.WriteLine();

      // Check database version
      var synchronizer = new DatabaseSynchronizer(databasePath);
      if (!synchronizer.CheckDatabaseVersion())
      {
        return 1;
      }

      Console.WriteLine();

      // Read HVSC files
      var reader = new HvscReader(hvscPath);
      var sidFiles = reader.ReadAllSidFiles();

      if (sidFiles.Count == 0)
      {
        Console.WriteLine("No SID files found in HVSC directory");
        return 1;
      }

      Console.WriteLine();

      // Synchronize to database
      synchronizer.SynchronizeSidData(sidFiles);

      Console.WriteLine();
      Console.WriteLine("HVSC import completed successfully");
      return 0;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"Fatal error: {ex.Message}");
      Console.WriteLine(ex.StackTrace);
      return 1;
    }
  }

  static void ShowUsage()
  {
    Console.WriteLine("Usage:");
    Console.WriteLine("  HVSC2GB64.exe <hvsc_path> <database_path>");
    Console.WriteLine();
    Console.WriteLine("Example:");
    Console.WriteLine("  HVSC2GB64.exe \"C:\\HVSC\" \"C:\\GB64\\GB64.db\"");
    Console.WriteLine();
    Console.WriteLine("Parameters:");
    Console.WriteLine("  hvsc_path      Path to HVSC root directory");
    Console.WriteLine("  database_path  Path to GameBase64 database file");
  }
}
