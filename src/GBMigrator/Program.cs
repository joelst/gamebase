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
catch (Exception ex)
{
  Console.Error.WriteLine();
  Console.Error.WriteLine($"Migration failed: {ex.Message}");
  Console.Error.WriteLine(ex.StackTrace);
  return 1;
}
