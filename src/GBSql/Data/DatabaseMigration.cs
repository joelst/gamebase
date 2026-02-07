//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  Database Migration Helper
//==============================================

namespace GBSql.Data;

/// <summary>
/// Helper class for database initialization and migration tasks.
/// </summary>
public static class DatabaseMigration
{
  /// <summary>
  /// Creates a new GameBase SQLite database with schema initialization.
  /// </summary>
  /// <param name="databasePath">Path where the database file should be created.</param>
  /// <param name="overwriteExisting">If true, deletes existing database before creating. Default is false.</param>
  /// <returns>A configured GameBaseSqliteRepository connected to the new database.</returns>
  public static async Task<GameBaseSqliteRepository> CreateNewDatabaseAsync(
    string databasePath,
    bool overwriteExisting = false)
  {
    if (overwriteExisting && File.Exists(databasePath))
    {
      File.Delete(databasePath);
    }

    var context = new GameBaseSqliteContext(databasePath);
    await context.InitializeDatabaseAsync();
    return new GameBaseSqliteRepository(context);
  }

  /// <summary>
  /// Seeds the database with sample data for testing.
  /// </summary>
  public static async Task SeedSampleDataAsync(GameBaseSqliteContext context)
  {
    // Add sample Publishers
    await context.ExecuteNonQueryAsync(
      "INSERT INTO Publishers (Name) VALUES (@Name)",
      new { Name = "Electronic Arts" });
    await context.ExecuteNonQueryAsync(
      "INSERT INTO Publishers (Name) VALUES (@Name)",
      new { Name = "Activision" });

    // Add sample Years
    await context.ExecuteNonQueryAsync(
      "INSERT INTO Years (Name) VALUES (@Name)",
      new { Name = "1985" });
    await context.ExecuteNonQueryAsync(
      "INSERT INTO Years (Name) VALUES (@Name)",
      new { Name = "1990" });

    // Add sample Programmers
    await context.ExecuteNonQueryAsync(
      "INSERT INTO Programmers (Name) VALUES (@Name)",
      new { Name = "John Doe" });

    // Add sample Games
    await context.ExecuteNonQueryAsync(@"
      INSERT INTO Games (Name, CCode, PU_Id, YE_Id, PR_Id, Fav, Classic, Adult, PlayersFrom, PlayersTo)
      VALUES (@Name, @CCode, @PU_Id, @YE_Id, @PR_Id, @Fav, @Classic, @Adult, @PlayersFrom, @PlayersTo)",
      new
      {
        Name = "Test Game 1",
        CCode = "TG1",
        PU_Id = 1,
        YE_Id = 1,
        PR_Id = 1,
        Fav = 1,
        Classic = 0,
        Adult = 0,
        PlayersFrom = 1,
        PlayersTo = 2
      });

    await context.ExecuteNonQueryAsync(@"
      INSERT INTO Games (Name, CCode, PU_Id, YE_Id, Fav, Classic, Adult, PlayersFrom, PlayersTo)
      VALUES (@Name, @CCode, @PU_Id, @YE_Id, @Fav, @Classic, @Adult, @PlayersFrom, @PlayersTo)",
      new
      {
        Name = "Test Game 2",
        CCode = "TG2",
        PU_Id = 2,
        YE_Id = 2,
        Fav = 0,
        Classic = 1,
        Adult = 0,
        PlayersFrom = 1,
        PlayersTo = 1
      });

    // Add sample Extras
    await context.ExecuteNonQueryAsync(@"
      INSERT INTO Extras (GA_Id, Name, Filename, EX_Type)
      VALUES (@GA_Id, @Name, @Filename, @EX_Type)",
      new
      {
        GA_Id = 1,
        Name = "Manual",
        Filename = "manual.pdf",
        EX_Type = 1
      });
  }
}
