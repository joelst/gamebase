//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  SQLite Repository Implementation
//==============================================

using GBSql.Models;
using Microsoft.Data.Sqlite;
using System.Dynamic;

namespace GBSql.Data;

/// <summary>
/// SQLite implementation of GameBase repository.
/// </summary>
public class GameBaseSqliteRepository : IGameBaseRepository
{
  private readonly GameBaseSqliteContext _context;
  private bool _disposed;

  public GameBaseSqliteRepository(GameBaseSqliteContext context)
  {
    _context = context ?? throw new ArgumentNullException(nameof(context));
  }

  /// <summary>
  /// Creates a new repository with a database path.
  /// </summary>
  public static async Task<GameBaseSqliteRepository> CreateAsync(string databasePath, bool initializeSchema = true)
  {
    var context = new GameBaseSqliteContext(databasePath);
    if (initializeSchema)
    {
      await context.InitializeDatabaseAsync();
    }
    return new GameBaseSqliteRepository(context);
  }

  /// <summary>
  /// Executes a raw SQL query and returns dynamic results.
  /// Used by QueryBuilder for custom view queries.
  /// </summary>
  public async Task<IEnumerable<dynamic>> ExecuteQueryAsync(string sql)
  {
    var conn = await _context.GetConnectionAsync();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;

    var results = new List<dynamic>();
    using var reader = await cmd.ExecuteReaderAsync();

    while (await reader.ReadAsync())
    {
      var row = new ExpandoObject() as IDictionary<string, object?>;
      for (int i = 0; i < reader.FieldCount; i++)
      {
        row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
      }
      results.Add(row!);
    }

    return results;
  }

  /// <summary>
  /// Executes a typed query with optional parameters.
  /// </summary>
  public async Task<IEnumerable<T>> QueryAsync<T>(string sql, object? parameters = null) where T : new()
  {
    var conn = await _context.GetConnectionAsync();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;

    if (parameters != null)
    {
      AddParameters(cmd, parameters);
    }

    var results = new List<T>();
    using var reader = await cmd.ExecuteReaderAsync();

    while (await reader.ReadAsync())
    {
      var item = MapToObject<T>(reader);
      results.Add(item);
    }

    return results;
  }

  public async Task<GameEntity?> GetGameByIdAsync(int gaId)
  {
    var results = await QueryAsync<GameEntity>(
      "SELECT * FROM Games WHERE GA_Id = @GA_Id",
      new { GA_Id = gaId });
    return results.FirstOrDefault();
  }

  public async Task<IEnumerable<GameEntity>> GetGamesAsync(int? limit = null)
  {
    string sql = limit.HasValue
      ? $"SELECT * FROM Games LIMIT {limit.Value}"
      : "SELECT * FROM Games";
    return await QueryAsync<GameEntity>(sql);
  }

  public async Task<IEnumerable<GameEntity>> SearchGamesByNameAsync(string searchTerm)
  {
    return await QueryAsync<GameEntity>(
      "SELECT * FROM Games WHERE Name LIKE @SearchTerm",
      new { SearchTerm = $"%{searchTerm}%" });
  }

  public async Task<IEnumerable<ExtraEntity>> GetExtrasForGameAsync(int gaId)
  {
    return await QueryAsync<ExtraEntity>(
      "SELECT * FROM Extras WHERE GA_Id = @GA_Id",
      new { GA_Id = gaId });
  }

  public async Task<int> UpsertExtraAsync(ExtraEntity extra)
  {
    var conn = await _context.GetConnectionAsync();
    using var cmd = conn.CreateCommand();

    if (extra.EX_Id > 0)
    {
      cmd.CommandText = @"UPDATE Extras SET Name=@Name, Filename=@Filename, EX_Type=@EX_Type WHERE EX_Id=@EX_Id";
      AddParameters(cmd, new { extra.Name, extra.Filename, extra.EX_Type, extra.EX_Id });
      await cmd.ExecuteNonQueryAsync();
      return extra.EX_Id;
    }
    else
    {
      cmd.CommandText = @"INSERT INTO Extras (GA_Id, Name, Filename, EX_Type) VALUES (@GA_Id, @Name, @Filename, @EX_Type); SELECT last_insert_rowid();";
      AddParameters(cmd, new { extra.GA_Id, extra.Name, extra.Filename, extra.EX_Type });
      var result = await cmd.ExecuteScalarAsync();
      return Convert.ToInt32(result);
    }
  }

  public async Task DeleteExtraAsync(int exId)
  {
    await _context.ExecuteNonQueryAsync("DELETE FROM Extras WHERE EX_Id = @EX_Id", new { EX_Id = exId });
  }

  public async Task SaveExtrasForGameAsync(int gaId, IEnumerable<ExtraEntity> extras)
  {
    var conn = await _context.GetConnectionAsync();
    var existing = (await GetExtrasForGameAsync(gaId)).ToList();
    var incomingList = extras.ToList();
    var incomingIds = incomingList.Where(e => e.EX_Id > 0).Select(e => e.EX_Id).ToHashSet();

    // Delete removed extras
    foreach (var ex in existing)
    {
      if (!incomingIds.Contains(ex.EX_Id))
      {
        await DeleteExtraAsync(ex.EX_Id);
      }
    }

    // Upsert extras
    foreach (var extra in incomingList)
    {
      if (extra.GA_Id == 0) extra.GA_Id = gaId;
      var id = await UpsertExtraAsync(extra);
      extra.EX_Id = id;
    }
  }

  public async Task<IEnumerable<LookupEntity>> GetLookupTableAsync(string tableName)
  {
    // Validate table name to prevent SQL injection
    var validTables = new[]
    {
      "Publishers", "Years", "Programmers", "Musicians", "Languages",
      "Genres", "PGenres", "Difficulty", "Crackers", "Artists",
      "Developers", "Licenses", "Rarities"
    };

    if (!validTables.Contains(tableName, StringComparer.OrdinalIgnoreCase))
    {
      throw new ArgumentException($"Invalid table name: {tableName}", nameof(tableName));
    }

    string idColumn = tableName switch
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

    var conn = await _context.GetConnectionAsync();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"SELECT {idColumn} as Id, Name FROM {tableName}";

    var results = new List<LookupEntity>();
    using var reader = await cmd.ExecuteReaderAsync();

    while (await reader.ReadAsync())
    {
      results.Add(new LookupEntity
      {
        Id = reader.GetInt32(0),
        Name = reader.GetString(1)
      });
    }

    return results;
  }

  private static void AddParameters(SqliteCommand cmd, object parameters)
  {
    var props = parameters.GetType().GetProperties();
    foreach (var prop in props)
    {
      var value = prop.GetValue(parameters);
      cmd.Parameters.AddWithValue($"@{prop.Name}", value ?? DBNull.Value);
    }
  }

  private static T MapToObject<T>(SqliteDataReader reader) where T : new()
  {
    var obj = new T();
    var props = typeof(T).GetProperties();

    for (int i = 0; i < reader.FieldCount; i++)
    {
      string columnName = reader.GetName(i);
      var prop = props.FirstOrDefault(p =>
        p.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

      if (prop != null && !reader.IsDBNull(i))
      {
        var value = reader.GetValue(i);
        prop.SetValue(obj, Convert.ChangeType(value, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType));
      }
    }

    return obj;
  }

  public void Dispose()
  {
    if (!_disposed)
    {
      _context?.Dispose();
      _disposed = true;
    }
    GC.SuppressFinalize(this);
  }
}
