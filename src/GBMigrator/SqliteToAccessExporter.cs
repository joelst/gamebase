//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBMigrator: SQLite to Access Export
//==============================================

using System.Data;
using System.Data.OleDb;
using GBSql.Data;
using Microsoft.Data.Sqlite;

namespace GBMigrator;

/// <summary>
/// Exports GameBase data from SQLite to Access (.mdb) format.
/// Preserves backward compatibility with VB6 GameBase.
/// </summary>
public class SqliteToAccessExporter
{
  private readonly string _sqliteDbPath;
  private readonly string _accessDbPath;
  private readonly IProgress<string>? _progress;

  public SqliteToAccessExporter(string sqliteDbPath, string accessDbPath, IProgress<string>? progress = null)
  {
    _sqliteDbPath = sqliteDbPath ?? throw new ArgumentNullException(nameof(sqliteDbPath));
    _accessDbPath = accessDbPath ?? throw new ArgumentNullException(nameof(accessDbPath));
    _progress = progress;
  }

  /// <summary>
  /// Executes the full export.
  /// </summary>
  public async Task ExportAsync()
  {
    ReportProgress("Starting export...");
    ReportProgress($"Source: {_sqliteDbPath}");
    ReportProgress($"Target: {_accessDbPath}");

    if (!File.Exists(_sqliteDbPath))
      throw new FileNotFoundException($"SQLite database not found: {_sqliteDbPath}");

    // Validate OleDb provider is available
    ValidateOleDbProvider();

    // Create Access database with schema
    ReportProgress("Creating Access database...");
    CreateAccessDatabase();

    // Export lookup tables first
    await ExportLookupTablesAsync();

    // Then export main tables
    await ExportGamesAsync();
    await ExportExtrasAsync();

    ReportProgress("Export completed successfully!");
  }

  private void CreateAccessDatabase()
  {
    if (File.Exists(_accessDbPath))
      File.Delete(_accessDbPath);

    // Create new Access database using ADOX
    var catalog = Activator.CreateInstance(Type.GetTypeFromProgID("ADOX.Catalog")!);
    if (catalog == null)
      throw new InvalidOperationException("Failed to create ADOX.Catalog");

    try
    {
      var method = catalog.GetType().GetMethod("Create");
      var connectionString = GetAccessConnectionString();
      method?.Invoke(catalog, new object[] { connectionString });
      ReportProgress("Access database created");
    }
    finally
    {
      System.Runtime.InteropServices.Marshal.ReleaseComObject(catalog);
    }

    // Create schema
    using var conn = new OleDbConnection(GetAccessConnectionString());
    conn.Open();

    // Create lookup tables
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

    foreach (var (tableName, idColumn) in lookupTables)
    {
      using var cmd = conn.CreateCommand();
      cmd.CommandText = $@"
        CREATE TABLE [{tableName}] (
          [{idColumn}] LONG NOT NULL PRIMARY KEY,
          [Name] TEXT(255)
        )";
      cmd.ExecuteNonQuery();
    }

    // Create Games table
    using (var cmd = conn.CreateCommand())
    {
      cmd.CommandText = @"
        CREATE TABLE [Games] (
          [GA_Id] LONG NOT NULL PRIMARY KEY,
          [Name] TEXT(255) NOT NULL,
          [CCode] TEXT(255),
          [PU_Id] LONG,
          [YE_Id] LONG,
          [PR_Id] LONG,
          [MU_Id] LONG,
          [LA_Id] LONG,
          [GE_Id] LONG,
          [PG_Id] LONG,
          [DI_Id] LONG,
          [CR_Id] LONG,
          [AR_Id] LONG,
          [DE_Id] LONG,
          [LI_Id] LONG,
          [RA_Id] LONG,
          [Fav] LONG DEFAULT 0,
          [Classic] LONG DEFAULT 0,
          [Adult] LONG DEFAULT 0,
          [PlayersFrom] LONG DEFAULT 1,
          [PlayersTo] LONG DEFAULT 1,
          [Prequel] LONG,
          [Sequel] LONG,
          [Related] LONG,
          [CloneOf] LONG,
          [V_Trainers] LONG,
          [V_Length] LONG,
          [V_LengthType] LONG,
          [ReviewRating] LONG,
          [Comment] MEMO,
          [Filename] TEXT(255)
        )";
      cmd.ExecuteNonQuery();
    }

    // Create Extras table
    using (var cmd = conn.CreateCommand())
    {
      cmd.CommandText = @"
        CREATE TABLE [Extras] (
          [EX_Id] LONG NOT NULL PRIMARY KEY,
          [GA_Id] LONG,
          [Name] TEXT(255),
          [Filename] TEXT(255),
          [Path] TEXT(255),
          [Type] TEXT(50)
        )";
      cmd.ExecuteNonQuery();
    }

    // Create Config table
    using (var cmd = conn.CreateCommand())
    {
      cmd.CommandText = @"
        CREATE TABLE [Config] (
          [Id] LONG NOT NULL PRIMARY KEY,
          [MajorVersion] LONG DEFAULT 2,
          [MinorVersion] LONG DEFAULT 5
        )";
      cmd.ExecuteNonQuery();

      cmd.CommandText = "INSERT INTO [Config] (Id, MajorVersion, MinorVersion) VALUES (1, 2, 5)";
      cmd.ExecuteNonQuery();
    }

    ReportProgress("Access database schema created");
  }

  private async Task ExportLookupTablesAsync()
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

    using var sqliteConn = new SqliteConnection($"Data Source={_sqliteDbPath}");
    await sqliteConn.OpenAsync();

    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    foreach (var (tableName, idColumn) in lookupTables)
    {
      ReportProgress($"Exporting {tableName}...");
      int count = await ExportLookupTableAsync(sqliteConn, accessConn, tableName, idColumn);
      ReportProgress($"  → {count} rows");
    }
  }

  private async Task<int> ExportLookupTableAsync(
    SqliteConnection sqliteConn,
    OleDbConnection accessConn,
    string tableName,
    string idColumn)
  {
    using var selectCmd = sqliteConn.CreateCommand();
    selectCmd.CommandText = $"SELECT {idColumn}, Name FROM {tableName}";

    int count = 0;
    using var reader = await selectCmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      int id = reader.GetInt32(0);
      string name = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);

      using var insertCmd = accessConn.CreateCommand();
      insertCmd.CommandText = $"INSERT INTO [{tableName}] ([{idColumn}], [Name]) VALUES (?, ?)";
      insertCmd.Parameters.Add(new OleDbParameter("@Id", id));
      insertCmd.Parameters.Add(new OleDbParameter("@Name", name));
      await insertCmd.ExecuteNonQueryAsync();

      count++;
    }

    return count;
  }

  private async Task ExportGamesAsync()
  {
    ReportProgress("Exporting Games...");

    using var sqliteConn = new SqliteConnection($"Data Source={_sqliteDbPath}");
    await sqliteConn.OpenAsync();

    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    using var selectCmd = sqliteConn.CreateCommand();
    selectCmd.CommandText = @"
      SELECT GA_Id, Name, CCode, PU_Id, YE_Id, PR_Id, MU_Id, LA_Id, GE_Id, PG_Id,
             DI_Id, CR_Id, AR_Id, DE_Id, LI_Id, RA_Id, Fav, Classic, Adult,
             PlayersFrom, PlayersTo, Prequel, Sequel, Related, CloneOf,
             V_Trainers, V_Length, V_LengthType, ReviewRating, Comment, Filename
      FROM Games";

    int count = 0;
    using var reader = await selectCmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      using var insertCmd = accessConn.CreateCommand();
      insertCmd.CommandText = @"
        INSERT INTO [Games] (
          [GA_Id], [Name], [CCode], [PU_Id], [YE_Id], [PR_Id], [MU_Id], [LA_Id], [GE_Id], [PG_Id],
          [DI_Id], [CR_Id], [AR_Id], [DE_Id], [LI_Id], [RA_Id], [Fav], [Classic], [Adult],
          [PlayersFrom], [PlayersTo], [Prequel], [Sequel], [Related], [CloneOf],
          [V_Trainers], [V_Length], [V_LengthType], [ReviewRating], [Comment], [Filename]
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

      for (int i = 0; i < 31; i++)
      {
        insertCmd.Parameters.Add(new OleDbParameter($"@p{i}", reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i)));
      }

      await insertCmd.ExecuteNonQueryAsync();

      count++;
      if (count % 500 == 0)
        ReportProgress($"  → {count} games...");
    }

    ReportProgress($"  → {count} total games");
  }

  private async Task ExportExtrasAsync()
  {
    ReportProgress("Exporting Extras...");

    using var sqliteConn = new SqliteConnection($"Data Source={_sqliteDbPath}");
    await sqliteConn.OpenAsync();

    using var accessConn = new OleDbConnection(GetAccessConnectionString());
    await accessConn.OpenAsync();

    using var selectCmd = sqliteConn.CreateCommand();
    selectCmd.CommandText = "SELECT EX_Id, GA_Id, Name, Filename, Path, Type FROM Extras";

    int count = 0;
    using var reader = await selectCmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      using var insertCmd = accessConn.CreateCommand();
      insertCmd.CommandText = @"
        INSERT INTO [Extras] ([EX_Id], [GA_Id], [Name], [Filename], [Path], [Type])
        VALUES (?, ?, ?, ?, ?, ?)";

      insertCmd.Parameters.Add(new OleDbParameter("@EX_Id", reader.IsDBNull(0) ? DBNull.Value : reader.GetInt32(0)));
      insertCmd.Parameters.Add(new OleDbParameter("@GA_Id", reader.IsDBNull(1) ? DBNull.Value : reader.GetInt32(1)));
      insertCmd.Parameters.Add(new OleDbParameter("@Name", reader.IsDBNull(2) ? DBNull.Value : reader.GetString(2)));
      insertCmd.Parameters.Add(new OleDbParameter("@Filename", reader.IsDBNull(3) ? DBNull.Value : reader.GetString(3)));
      insertCmd.Parameters.Add(new OleDbParameter("@Path", reader.IsDBNull(4) ? DBNull.Value : reader.GetString(4)));
      insertCmd.Parameters.Add(new OleDbParameter("@Type", reader.IsDBNull(5) ? DBNull.Value : reader.GetString(5)));

      await insertCmd.ExecuteNonQueryAsync();

      count++;
      if (count % 500 == 0)
        ReportProgress($"  → {count} extras...");
    }

    ReportProgress($"  → {count} total extras");
  }

  private string GetAccessConnectionString()
  {
    return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_accessDbPath};";
  }

  private void ValidateOleDbProvider()
  {
    try
    {
      using var conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;");
    }
    catch (Exception ex)
    {
      throw new InvalidOperationException(
        "Microsoft Access Database Engine (ACE) provider not found. " +
        "Please install the Microsoft Access Database Engine 2016 Redistributable. " +
        "Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920",
        ex);
    }
  }

  private void ReportProgress(string message)
  {
    _progress?.Report(message);
  }
}
