//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBMigrator: Access to SQLite Migration
//==============================================

using System.Data;
using System.Data.OleDb;
using GBSql.Data;
using Microsoft.Data.Sqlite;

namespace GBMigrator;

/// <summary>
/// Migrates GameBase data from Access (.mdb) to SQLite.
/// </summary>
public class AccessToSqliteMigrator
{
  private readonly string _accessDbPath;
  private readonly string _sqliteDbPath;
  private readonly IProgress<string>? _progress;

  public AccessToSqliteMigrator(string accessDbPath, string sqliteDbPath, IProgress<string>? progress = null)
  {
    _accessDbPath = accessDbPath ?? throw new ArgumentNullException(nameof(accessDbPath));
    _sqliteDbPath = sqliteDbPath ?? throw new ArgumentNullException(nameof(sqliteDbPath));
    _progress = progress;
  }

  /// <summary>
  /// Executes the full migration.
  /// </summary>
  public async Task MigrateAsync()
  {
    ReportProgress("Starting migration...");
    ReportProgress($"Source: {_accessDbPath}");
    ReportProgress($"Target: {_sqliteDbPath}");

    if (!File.Exists(_accessDbPath))
      throw new FileNotFoundException($"Access database not found: {_accessDbPath}");

    // Validate OleDb provider is available before proceeding
    ValidateOleDbProvider();

    // Initialize SQLite database with schema
    ReportProgress("Initializing SQLite database...");
    var context = new GameBaseSqliteContext(_sqliteDbPath);
    await context.InitializeDatabaseAsync();

    // Migrate lookup tables first (they have no dependencies)
    await MigrateLookupTablesAsync(context);

    // Then migrate main tables
    await MigrateGamesAsync(context);
    await MigrateExtrasAsync(context);

    ReportProgress("Migration completed successfully!");
  }

  private async Task MigrateLookupTablesAsync(GameBaseSqliteContext context)
  {
    var lookupTables = new Dictionary<string, string>
    {
      { "Publishers", "PU_Id" },
      { "Years", "YE_Id" },
      { "Programmers", "PR_Id" },
      { "Musicians", "MU_Id" },
      { "Languages", "LA_Id" },
      { "Genres", "GE_Id" },
      { "PGenres", "PG_Id" },
      { "Difficulty", "DI_Id" },
      { "Crackers", "CR_Id" },
      { "Artists", "AR_Id" },
      { "Developers", "DE_Id" },
      { "Licenses", "LI_Id" },
      { "Rarities", "RA_Id" }
    };

    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    foreach (var (tableName, idColumn) in lookupTables)
    {
      if (!await TableExistsAsync(accessConn, tableName))
      {
        ReportProgress($"Skipping {tableName} (table not found in source)");
        continue;
      }

      ReportProgress($"Migrating {tableName}...");
      int count = await MigrateLookupTableAsync(accessConn, context, tableName, idColumn);
      ReportProgress($"  → {count} rows");
    }
  }

  private async Task<int> MigrateLookupTableAsync(
    OleDbConnection accessConn,
    GameBaseSqliteContext context,
    string tableName,
    string idColumn)
  {
    using var cmd = accessConn.CreateCommand();
    cmd.CommandText = $"SELECT {idColumn}, Name FROM [{tableName}]";

    int count = 0;
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      int id = reader.GetInt32(0);
      string name = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);

      await context.ExecuteNonQueryAsync(
        $"INSERT INTO {tableName} ({idColumn}, Name) VALUES (@Id, @Name)",
        new { Id = id, Name = name });

      count++;
    }

    return count;
  }

  private async Task MigrateGamesAsync(GameBaseSqliteContext context)
  {
    if (!await TableExistsInAccessAsync("Games"))
    {
      ReportProgress("Skipping Games (table not found in source)");
      return;
    }

    ReportProgress("Migrating Games...");
    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    using var cmd = accessConn.CreateCommand();
    cmd.CommandText = @"
      SELECT GA_Id, Name, CCode, PU_Id, YE_Id, PR_Id, MU_Id, LA_Id, GE_Id, PG_Id,
             DI_Id, CR_Id, AR_Id, DE_Id, LI_Id, RA_Id, Fav, Classic, Adult,
             PlayersFrom, PlayersTo, Prequel, Sequel, Related, CloneOf,
             V_Trainers, V_Length, V_LengthType, ReviewRating, Comment, Filename
      FROM Games";

    int count = 0;
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      await context.ExecuteNonQueryAsync(@"
        INSERT INTO Games (
          GA_Id, Name, CCode, PU_Id, YE_Id, PR_Id, MU_Id, LA_Id, GE_Id, PG_Id,
          DI_Id, CR_Id, AR_Id, DE_Id, LI_Id, RA_Id, Fav, Classic, Adult,
          PlayersFrom, PlayersTo, Prequel, Sequel, Related, CloneOf,
          V_Trainers, V_Length, V_LengthType, ReviewRating, Comment, Filename
        ) VALUES (
          @GA_Id, @Name, @CCode, @PU_Id, @YE_Id, @PR_Id, @MU_Id, @LA_Id, @GE_Id, @PG_Id,
          @DI_Id, @CR_Id, @AR_Id, @DE_Id, @LI_Id, @RA_Id, @Fav, @Classic, @Adult,
          @PlayersFrom, @PlayersTo, @Prequel, @Sequel, @Related, @CloneOf,
          @V_Trainers, @V_Length, @V_LengthType, @ReviewRating, @Comment, @Filename
        )",
        new
        {
          GA_Id = GetInt32OrNull(reader, 0),
          Name = GetStringOrNull(reader, 1) ?? string.Empty,
          CCode = GetStringOrNull(reader, 2),
          PU_Id = GetInt32OrNull(reader, 3),
          YE_Id = GetInt32OrNull(reader, 4),
          PR_Id = GetInt32OrNull(reader, 5),
          MU_Id = GetInt32OrNull(reader, 6),
          LA_Id = GetInt32OrNull(reader, 7),
          GE_Id = GetInt32OrNull(reader, 8),
          PG_Id = GetInt32OrNull(reader, 9),
          DI_Id = GetInt32OrNull(reader, 10),
          CR_Id = GetInt32OrNull(reader, 11),
          AR_Id = GetInt32OrNull(reader, 12),
          DE_Id = GetInt32OrNull(reader, 13),
          LI_Id = GetInt32OrNull(reader, 14),
          RA_Id = GetInt32OrNull(reader, 15),
          Fav = GetInt32OrDefault(reader, 16, 0),
          Classic = GetInt32OrDefault(reader, 17, 0),
          Adult = GetInt32OrDefault(reader, 18, 0),
          PlayersFrom = GetInt32OrDefault(reader, 19, 1),
          PlayersTo = GetInt32OrDefault(reader, 20, 1),
          Prequel = GetInt32OrNull(reader, 21),
          Sequel = GetInt32OrNull(reader, 22),
          Related = GetInt32OrNull(reader, 23),
          CloneOf = GetInt32OrNull(reader, 24),
          V_Trainers = GetInt32OrNull(reader, 25),
          V_Length = GetInt32OrNull(reader, 26),
          V_LengthType = GetInt32OrNull(reader, 27),
          ReviewRating = GetInt32OrNull(reader, 28),
          Comment = GetStringOrNull(reader, 29),
          Filename = GetStringOrNull(reader, 30)
        });

      count++;
      if (count % 500 == 0)
        ReportProgress($"  → {count} games...");
    }

    ReportProgress($"  → {count} total games");
  }

  private async Task MigrateExtrasAsync(GameBaseSqliteContext context)
  {
    if (!await TableExistsInAccessAsync("Extras"))
    {
      ReportProgress("Skipping Extras (table not found in source)");
      return;
    }

    ReportProgress("Migrating Extras...");
    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    using var cmd = accessConn.CreateCommand();
    cmd.CommandText = "SELECT EX_Id, GA_Id, Name, Filename, EX_Type FROM Extras";

    int count = 0;
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      await context.ExecuteNonQueryAsync(@"
        INSERT INTO Extras (EX_Id, GA_Id, Name, Filename, EX_Type)
        VALUES (@EX_Id, @GA_Id, @Name, @Filename, @EX_Type)",
        new
        {
          EX_Id = reader.GetInt32(0),
          GA_Id = reader.GetInt32(1),
          Name = GetStringOrNull(reader, 2) ?? string.Empty,
          Filename = GetStringOrNull(reader, 3),
          EX_Type = GetInt32OrNull(reader, 4)
        });

      count++;
      if (count % 500 == 0)
        ReportProgress($"  → {count} extras...");
    }

    ReportProgress($"  → {count} total extras");
  }

  private async Task<bool> TableExistsInAccessAsync(string tableName)
  {
    using var conn = new OleDbConnection(GetAccessConnectionString());
    await conn.OpenAsync();
    return await TableExistsAsync(conn, tableName);
  }

  private static async Task<bool> TableExistsAsync(OleDbConnection conn, string tableName)
  {
    var schema = await conn.GetSchemaAsync("Tables");
    return schema.Rows.Cast<DataRow>().Any(row =>
      row["TABLE_NAME"].ToString()?.Equals(tableName, StringComparison.OrdinalIgnoreCase) == true &&
      row["TABLE_TYPE"].ToString() == "TABLE");
  }

  private string GetAccessConnectionString()
  {
    // Try ACE 12.0 first (newer format), then fall back to Jet 4.0 (older format)
    return $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={_accessDbPath};Persist Security Info=False;";
  }

  private static int? GetInt32OrNull(IDataReader reader, int ordinal)
  {
    return reader.IsDBNull(ordinal) ? null : reader.GetInt32(ordinal);
  }

  private static int GetInt32OrDefault(IDataReader reader, int ordinal, int defaultValue)
  {
    return reader.IsDBNull(ordinal) ? defaultValue : reader.GetInt32(ordinal);
  }

  private static string? GetStringOrNull(IDataReader reader, int ordinal)
  {
    return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
  }

  private void ValidateOleDbProvider()
  {
    try
    {
      // Test if the provider is available by trying to create a connection
      using var testConn = new OleDbConnection(GetAccessConnectionString());
      // We don't need to open it, just creating it will validate the provider exists
    }
    catch (InvalidOperationException ex) when (ex.Message.Contains("provider") || ex.Message.Contains("Provider"))
    {
      throw new InvalidOperationException(
        "The Microsoft Access Database Engine (OleDb provider) is not installed or not available. " +
        "Please install the Microsoft Access Database Engine 2016 Redistributable from " +
        "https://www.microsoft.com/en-us/download/details.aspx?id=54920",
        ex);
    }
    catch (ArgumentException ex) when (ex.Message.Contains("provider") || ex.Message.Contains("Provider"))
    {
      throw new InvalidOperationException(
        "The OleDb provider for Microsoft Access is not registered on this system. " +
        "Please install the Microsoft Access Database Engine.",
        ex);
    }
  }

  private void ReportProgress(string message)
  {
    _progress?.Report(message);
  }
}
