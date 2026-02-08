//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBMigrator: CLI Entry Point
//==============================================

using System.Runtime.InteropServices;
using GBMigrator;

if (args.Length < 2)
{
  Console.WriteLine("GameBase Access to SQLite Migrator");
  Console.WriteLine();
  Console.WriteLine("Usage: GBMigrator <access-db-path> <sqlite-db-path>");
  Console.WriteLine();
  Console.WriteLine("Example:");
  Console.WriteLine("  GBMigrator C:\\GameBase\\GameBase.mdb C:\\GameBase\\GameBase.db");
  return 1;
}

// Check platform first
if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
{
  Console.Error.WriteLine("Error: GBMigrator requires Windows to access Microsoft Access databases.");
  Console.Error.WriteLine("The OleDb provider for Access (.mdb) files is only available on Windows.");
  Console.Error.WriteLine();
  Console.Error.WriteLine("To migrate on non-Windows systems:");
  Console.Error.WriteLine("  1. Run GBMigrator on a Windows machine first");
  Console.Error.WriteLine("  2. Copy the resulting .db file to your target system");
  return 1;
}

string accessPath = args[0];
string sqlitePath = args[1];

if (!File.Exists(accessPath))
{
  Console.Error.WriteLine($"Error: Access database not found: {accessPath}");
  return 1;
}

if (File.Exists(sqlitePath))
{
  Console.Write($"Target database already exists: {sqlitePath}");
  Console.Write("\nOverwrite? (y/N): ");
  string? response = Console.ReadLine();
  if (response?.Trim().ToLowerInvariant() != "y")
  {
    Console.WriteLine("Migration cancelled.");
    return 0;
  }
  File.Delete(sqlitePath);
}

var progress = new Progress<string>(message => Console.WriteLine(message));
var migrator = new AccessToSqliteMigrator(accessPath, sqlitePath, progress);

try
{
  await migrator.MigrateAsync();
  Console.WriteLine();
  Console.WriteLine("Migration completed successfully!");
  return 0;
}
catch (InvalidOperationException ex) when (ex.Message.Contains("provider") || ex.Message.Contains("Provider"))
{
  Console.Error.WriteLine();
  Console.Error.WriteLine("Error: Microsoft Access Database Engine not found.");
  Console.Error.WriteLine();
  Console.Error.WriteLine("Please install the Microsoft Access Database Engine:");
  Console.Error.WriteLine("  - Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920");
  Console.Error.WriteLine("  - Install the version matching your .NET runtime (x64 or x86)");
  Console.Error.WriteLine();
  Console.Error.WriteLine($"Technical details: {ex.Message}");
  return 1;
}
catch (Exception ex)
{
  Console.Error.WriteLine();
  Console.Error.WriteLine($"Migration failed: {ex.Message}");
  if (ex.InnerException != null)
  {
    Console.Error.WriteLine($"Inner error: {ex.InnerException.Message}");
  }
  Console.Error.WriteLine();
  Console.Error.WriteLine("Stack trace:");
  Console.Error.WriteLine(ex.StackTrace);
  return 1;
}
