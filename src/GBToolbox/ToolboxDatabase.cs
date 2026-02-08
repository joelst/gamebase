//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: SQLite helpers
//==============================================

using Microsoft.Data.Sqlite;

namespace GBToolbox;

internal static class ToolboxDatabase
{
  public static SqliteConnection Open(string databasePath)
  {
    var conn = new SqliteConnection($"Data Source={databasePath}");
    conn.Open();
    return conn;
  }

  public static bool TableExists(SqliteConnection conn, string tableName)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name=@name;";
    cmd.Parameters.AddWithValue("@name", tableName);
    var result = cmd.ExecuteScalar();
    return result != null;
  }

  public static bool ColumnExists(SqliteConnection conn, string tableName, string columnName)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"PRAGMA table_info([{tableName}]);";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
    {
      string name = reader.GetString(1);
      if (name.Equals(columnName, StringComparison.OrdinalIgnoreCase))
      {
        return true;
      }
    }

    return false;
  }

  public static int ExecuteNonQuery(SqliteConnection conn, string sql, params SqliteParameter[] parameters)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;
    foreach (var param in parameters)
    {
      cmd.Parameters.Add(param);
    }
    return cmd.ExecuteNonQuery();
  }

  public static object? ExecuteScalar(SqliteConnection conn, string sql, params SqliteParameter[] parameters)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;
    foreach (var param in parameters)
    {
      cmd.Parameters.Add(param);
    }
    return cmd.ExecuteScalar();
  }
}
