using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Provides SQLite database helper methods for BAC.
/// </summary>
internal static class BacDatabase
{
  /// <summary>
  /// Required database version (major.minor).
  /// </summary>
  public const double RequiredVersion = 2.1;

  /// <summary>
  /// Opens a SQLite database connection.
  /// </summary>
  /// <param name="databasePath">Path to the .db file.</param>
  /// <returns>An open SqliteConnection.</returns>
  public static SqliteConnection OpenDatabase(string databasePath)
  {
    var connectionString = new SqliteConnectionStringBuilder
    {
      DataSource = databasePath,
      Mode = SqliteOpenMode.ReadWrite
    }.ToString();

    var connection = new SqliteConnection(connectionString);
    connection.Open();
    return connection;
  }

  /// <summary>
  /// Gets the database version from a GameBase database.
  /// </summary>
  /// <param name="databasePath">Path to the .db file.</param>
  /// <returns>The database version (major.minor), or 1.3 if not found.</returns>
  public static double GetDatabaseVersion(string databasePath)
  {
    try
    {
      using var connection = OpenDatabase(databasePath);
      using var command = connection.CreateCommand();
      command.CommandText = "SELECT MajorVersion, MinorVersion FROM Config LIMIT 1";

      using var reader = command.ExecuteReader();
      if (reader.Read())
      {
        var major = reader.GetInt32(0);
        var minor = reader.GetInt32(1);
        return double.Parse($"{major}.{minor}");
      }

      // Default to 1.3 if Config table is empty
      return 1.3;
    }
    catch
    {
      // If any error occurs (table doesn't exist, etc.), default to 1.3
      return 1.3;
    }
  }

  /// <summary>
  /// Validates that a database has the required version.
  /// </summary>
  /// <param name="databasePath">Path to the .db file.</param>
  /// <param name="actualVersion">Outputs the actual version found.</param>
  /// <returns>True if version is compatible, false otherwise.</returns>
  public static bool ValidateDatabaseVersion(string databasePath, out double actualVersion)
  {
    actualVersion = GetDatabaseVersion(databasePath);
    return Math.Abs(actualVersion - RequiredVersion) < 0.01;
  }
}
