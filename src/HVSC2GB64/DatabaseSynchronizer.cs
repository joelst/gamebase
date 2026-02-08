// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1997-2004 James Burrows - Migrated to .NET 2026
// HVSC2GB64: Database synchronization

using Microsoft.Data.Sqlite;

namespace HVSC2GB64;

/// <summary>
/// Synchronizes HVSC data with GB64 database.
/// </summary>
internal class DatabaseSynchronizer
{
  private readonly string _databasePath;
  private const double RequiredDbVersion = 2.5;

  public DatabaseSynchronizer(string databasePath)
  {
    _databasePath = databasePath;
  }

  /// <summary>
  /// Verifies database version compatibility.
  /// </summary>
  public bool CheckDatabaseVersion()
  {
    try
    {
      using var connection = new SqliteConnection($"Data Source={_databasePath}");
      connection.Open();

      using var command = connection.CreateCommand();
      command.CommandText = "SELECT MajorVersion, MinorVersion FROM Config LIMIT 1";

      using var reader = command.ExecuteReader();
      if (reader.Read())
      {
        int major = reader.GetInt32(0);
        int minor = reader.GetInt32(1);
        double version = major + (minor / 10.0);

        Console.WriteLine($"Database version: {version}");

        if (version < RequiredDbVersion)
        {
          Console.WriteLine($"Error: Database version {version} is too old. Required: {RequiredDbVersion}");
          return false;
        }

        return true;
      }

      Console.WriteLine("Error: Could not read database version");
      return false;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"Error checking database: {ex.Message}");
      return false;
    }
  }

  /// <summary>
  /// Synchronizes HVSC SID data to database.
  /// </summary>
  public void SynchronizeSidData(List<SidData> sidFiles)
  {
    Console.WriteLine($"Synchronizing {sidFiles.Count} SID files to database...");

    try
    {
      using var connection = new SqliteConnection($"Data Source={_databasePath}");
      connection.Open();

      int added = 0;
      int updated = 0;
      int skipped = 0;

      foreach (var sid in sidFiles)
      {
        try
        {
          if (SidExistsInDatabase(connection, sid.RelativePath))
          {
            // Could update existing record here if needed
            skipped++;
          }
          else
          {
            AddSidToDatabase(connection, sid);
            added++;
          }
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Warning: Could not sync {sid.Filename}: {ex.Message}");
        }
      }

      Console.WriteLine($"Sync complete: {added} added, {updated} updated, {skipped} skipped");
    }
    catch (Exception ex)
    {
      Console.WriteLine($"Error during synchronization: {ex.Message}");
    }
  }

  private bool SidExistsInDatabase(SqliteConnection connection, string relativePath)
  {
    using var command = connection.CreateCommand();
    command.CommandText = "SELECT COUNT(*) FROM Music WHERE Filename = @path";
    command.Parameters.AddWithValue("@path", relativePath);

    var count = Convert.ToInt32(command.ExecuteScalar());
    return count > 0;
  }

  private void AddSidToDatabase(SqliteConnection connection, SidData sid)
  {
    using var command = connection.CreateCommand();
    command.CommandText = @"
      INSERT INTO Music (Filename, Name, Author, Copyright, Songs, DefaultSong)
      VALUES (@filename, @name, @author, @copyright, @songs, @defaultSong)";

    command.Parameters.AddWithValue("@filename", sid.RelativePath);
    command.Parameters.AddWithValue("@name", sid.Name);
    command.Parameters.AddWithValue("@author", sid.Author);
    command.Parameters.AddWithValue("@copyright", sid.Copyright);
    command.Parameters.AddWithValue("@songs", sid.Songs);
    command.Parameters.AddWithValue("@defaultSong", sid.DefaultSong);

    command.ExecuteNonQuery();
  }
}
