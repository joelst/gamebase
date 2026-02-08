//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  SQLite Database Context
//==============================================

using Microsoft.Data.Sqlite;
using System.Data;

namespace GBSql.Data;

/// <summary>
/// SQLite database context for GameBase.
/// Manages connections and provides database initialization.
/// </summary>
public class GameBaseSqliteContext : IDisposable
{
  private readonly string _connectionString;
  private SqliteConnection? _connection;
  private bool _disposed;

  public GameBaseSqliteContext(string databasePath)
  {
    _connectionString = $"Data Source={databasePath}";
  }

  /// <summary>
  /// Gets or creates the database connection.
  /// </summary>
  public async Task<SqliteConnection> GetConnectionAsync()
  {
    if (_connection == null)
    {
      _connection = new SqliteConnection(_connectionString);
      await _connection.OpenAsync();
    }
    else if (_connection.State != ConnectionState.Open)
    {
      await _connection.OpenAsync();
    }
    return _connection;
  }

  /// <summary>
  /// Initializes the database schema if it doesn't exist.
  /// </summary>
  public async Task InitializeDatabaseAsync()
  {
    var conn = await GetConnectionAsync();

    // Create Games table
    await ExecuteNonQueryAsync(@"
      CREATE TABLE IF NOT EXISTS Games (
        GA_Id INTEGER PRIMARY KEY AUTOINCREMENT,
        Name TEXT NOT NULL,
        CCode TEXT,
        PU_Id INTEGER,
        YE_Id INTEGER,
        PR_Id INTEGER,
        MU_Id INTEGER,
        LA_Id INTEGER,
        GE_Id INTEGER,
        PG_Id INTEGER,
        DI_Id INTEGER,
        CR_Id INTEGER,
        AR_Id INTEGER,
        DE_Id INTEGER,
        LI_Id INTEGER,
        RA_Id INTEGER,
        Fav INTEGER DEFAULT 0,
        Classic INTEGER DEFAULT 0,
        Adult INTEGER DEFAULT 0,
        PlayersFrom INTEGER DEFAULT 1,
        PlayersTo INTEGER DEFAULT 1,
        Prequel INTEGER,
        Sequel INTEGER,
        Related INTEGER,
        CloneOf INTEGER,
        V_Trainers INTEGER,
        V_Length INTEGER,
        V_LengthType INTEGER,
        ReviewRating INTEGER,
        Comment TEXT,
        Filename TEXT,
        ScreenshotFilename TEXT
      )");

    // Add ScreenshotFilename if missing
    await EnsureColumnExists(conn, "Games", "ScreenshotFilename", "TEXT");

    // Create lookup tables
    var lookupTables = new[]
    {
      "Publishers", "Years", "Programmers", "Musicians", "Languages",
      "Genres", "PGenres", "Difficulty", "Crackers", "Artists",
      "Developers", "Licenses", "Rarities"
    };

    foreach (var table in lookupTables)
    {
      string idColumn = table switch
      {
        "Publishers" => "PU_Id",
        "Years" => "YE_Id",
        "Programmers" => "PR_Id",
        "Musicians" => "MU_Id",
        "Languages" => "LA_Id",
        "Genres" => "GE_Id",
        "PGenres" => "PG_Id",
        "Difficulty" => "DI_Id",
        "Crackers" => "CR_Id",
        "Artists" => "AR_Id",
        "Developers" => "DE_Id",
        "Licenses" => "LI_Id",
        "Rarities" => "RA_Id",
        _ => "Id"
      };

      await ExecuteNonQueryAsync($@"
        CREATE TABLE IF NOT EXISTS {table} (
          {idColumn} INTEGER PRIMARY KEY AUTOINCREMENT,
          Name TEXT NOT NULL
        )");
    }

    // Create Extras table
    await ExecuteNonQueryAsync(@"
      CREATE TABLE IF NOT EXISTS Extras (
        EX_Id INTEGER PRIMARY KEY AUTOINCREMENT,
        GA_Id INTEGER NOT NULL,
        Name TEXT NOT NULL,
        Filename TEXT,
        EX_Type INTEGER,
        FOREIGN KEY (GA_Id) REFERENCES Games(GA_Id)
      )");

    // Create indexes for common queries
    await ExecuteNonQueryAsync("CREATE INDEX IF NOT EXISTS idx_games_name ON Games(Name)");
    await ExecuteNonQueryAsync("CREATE INDEX IF NOT EXISTS idx_extras_ga_id ON Extras(GA_Id)");
  }

  private static async Task EnsureColumnExists(SqliteConnection conn, string table, string column, string type)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"PRAGMA table_info({table});";
    var exists = false;
    using (var reader = await cmd.ExecuteReaderAsync())
    {
      while (await reader.ReadAsync())
      {
        if (reader.GetString(1).Equals(column, StringComparison.OrdinalIgnoreCase))
        {
          exists = true;
          break;
        }
      }
    }
    if (!exists)
    {
      using var alter = conn.CreateCommand();
      alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {type};";
      await alter.ExecuteNonQueryAsync();
    }
  }

  /// <summary>
  /// Executes a non-query SQL command.
  /// </summary>
  public async Task<int> ExecuteNonQueryAsync(string sql, object? parameters = null)
  {
    var conn = await GetConnectionAsync();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;

    if (parameters != null)
    {
      AddParameters(cmd, parameters);
    }

    return await cmd.ExecuteNonQueryAsync();
  }

  /// <summary>
  /// Adds parameters to a command from an anonymous object.
  /// </summary>
  private static void AddParameters(SqliteCommand cmd, object parameters)
  {
    var props = parameters.GetType().GetProperties();
    foreach (var prop in props)
    {
      var value = prop.GetValue(parameters);
      cmd.Parameters.AddWithValue($"@{prop.Name}", value ?? DBNull.Value);
    }
  }

  public void Dispose()
  {
    if (!_disposed)
    {
      _connection?.Dispose();
      _disposed = true;
    }
    GC.SuppressFinalize(this);
  }
}
