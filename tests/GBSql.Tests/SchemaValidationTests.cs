//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql.Tests:  Schema Validation Tests
//==============================================

using GBSql.Data;
using Microsoft.Data.Sqlite;

namespace GBSql.Tests;

public class SchemaValidationTests : IDisposable
{
  private readonly string _testDbPath;
  private readonly GameBaseSqliteContext _context;

  public SchemaValidationTests()
  {
    _testDbPath = Path.Combine(Path.GetTempPath(), $"gamebase_schema_{Guid.NewGuid()}.db");
    _context = new GameBaseSqliteContext(_testDbPath);
    _context.InitializeDatabaseAsync().GetAwaiter().GetResult();
  }

  [Fact]
  public async Task DatabaseSchema_IncludesExpectedTables()
  {
    var conn = await _context.GetConnectionAsync();
    var tables = await GetTableNamesAsync(conn);

    var expectedTables = new[]
    {
      "Games",
      "Extras",
      "Publishers",
      "Years",
      "Programmers",
      "Musicians",
      "Languages",
      "Genres",
      "PGenres",
      "Difficulty",
      "Crackers",
      "Artists",
      "Developers",
      "Licenses",
      "Rarities"
    };

    foreach (var table in expectedTables)
    {
      Assert.Contains(table, tables, StringComparer.OrdinalIgnoreCase);
    }
  }

  [Fact]
  public async Task GamesTable_IncludesScreenshotFilename()
  {
    var conn = await _context.GetConnectionAsync();
    var columns = await GetColumnNamesAsync(conn, "Games");

    Assert.Contains("ScreenshotFilename", columns, StringComparer.OrdinalIgnoreCase);
  }

  [Fact]
  public async Task ExtrasTable_IncludesExpectedColumns()
  {
    var conn = await _context.GetConnectionAsync();
    var columns = await GetColumnNamesAsync(conn, "Extras");

    var expectedColumns = new[] { "EX_Id", "GA_Id", "Name", "Filename", "EX_Type" };
    foreach (var column in expectedColumns)
    {
      Assert.Contains(column, columns, StringComparer.OrdinalIgnoreCase);
    }
  }

  [Fact]
  public async Task LookupTables_HaveIdAndName()
  {
    var conn = await _context.GetConnectionAsync();

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

    foreach (var (table, idColumn) in lookupTables)
    {
      var columns = await GetColumnNamesAsync(conn, table);
      Assert.Contains(idColumn, columns, StringComparer.OrdinalIgnoreCase);
      Assert.Contains("Name", columns, StringComparer.OrdinalIgnoreCase);
    }
  }

  private static async Task<List<string>> GetTableNamesAsync(SqliteConnection conn)
  {
    var results = new List<string>();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table'";
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      results.Add(reader.GetString(0));
    }
    return results;
  }

  private static async Task<List<string>> GetColumnNamesAsync(SqliteConnection conn, string tableName)
  {
    var results = new List<string>();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"PRAGMA table_info({tableName});";
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      results.Add(reader.GetString(1));
    }
    return results;
  }

  public void Dispose()
  {
    _context?.Dispose();

    GC.Collect();
    GC.WaitForPendingFinalizers();

    if (File.Exists(_testDbPath))
    {
      try
      {
        File.Delete(_testDbPath);
      }
      catch (IOException)
      {
        // Ignore cleanup failure for temp DB
      }
    }
  }
}
