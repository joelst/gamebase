//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBMigrator: SQLite Diff Utility
//==============================================

using System.Globalization;
using GBSql.Data;
using Microsoft.Data.Sqlite;

namespace GBMigrator;

public sealed class SqliteDiffTool
{
  private readonly string _currentPath;
  private readonly string _backupPath;
  private readonly string _outputPath;
  private readonly DiffLetterSelection _letters;

  private SqliteDiffTool(string currentPath, string backupPath, string outputPath, DiffLetterSelection letters)
  {
    _currentPath = currentPath;
    _backupPath = backupPath;
    _outputPath = outputPath;
    _letters = letters;
  }

  public static async Task<int> RunAsync(string[] args)
  {
    if (!TryParseArgs(args, out var tool, out var error))
    {
      Console.Error.WriteLine(error);
      Console.WriteLine();
      PrintUsage();
      return 1;
    }

    if (!File.Exists(tool._currentPath))
    {
      Console.Error.WriteLine($"Error: Current database not found: {tool._currentPath}");
      return 1;
    }

    if (!File.Exists(tool._backupPath))
    {
      Console.Error.WriteLine($"Error: Backup database not found: {tool._backupPath}");
      return 1;
    }

    if (!tool._letters.HasSelection)
    {
      Console.Error.WriteLine("Error: At least one letter (or 0-9) must be selected.");
      return 1;
    }

    if (File.Exists(tool._outputPath))
    {
      Console.Write($"Target update database already exists: {tool._outputPath}");
      Console.Write("\nOverwrite? (y/N): ");
      string? response = Console.ReadLine();
      if (!string.Equals(response?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
      {
        Console.WriteLine("Diff cancelled.");
        return 0;
      }
      File.Delete(tool._outputPath);
    }

    using var currentConn = new SqliteConnection($"Data Source={tool._currentPath}");
    using var backupConn = new SqliteConnection($"Data Source={tool._backupPath}");
    await currentConn.OpenAsync();
    await backupConn.OpenAsync();

    var columns = await GetGameColumnsAsync(currentConn);
    if (columns.Count == 0)
    {
      Console.Error.WriteLine("Error: Games table not found in current database.");
      return 1;
    }

    Console.WriteLine("Scanning for changes...");
    var changes = await FindChangesAsync(currentConn, backupConn, columns, tool._letters);

    if (changes.Count == 0)
    {
      Console.WriteLine("No changes found.");
      return 0;
    }

    await InitializeUpdateDatabaseAsync(tool._outputPath);
    await WriteChangesAsync(tool._outputPath, columns, changes);

    Console.WriteLine();
    Console.WriteLine("Changed games:");
    foreach (var name in changes.Select(c => NormalizeForDisplay(GetRowValue(c, "Name"))).Where(n => !string.IsNullOrWhiteSpace(n)))
    {
      Console.WriteLine($"- {name}");
    }

    var info = PromptForUpdateInfo();
    await WriteInfoAsync(tool._outputPath, info);

    Console.WriteLine();
    Console.WriteLine($"Update database created: {tool._outputPath}");
    Console.WriteLine("You should ZIP and email this file to the database maintainer.");

    Console.WriteLine();
    Console.Write("Replace the backup database with your current database now? (y/N): ");
    string? replaceResponse = Console.ReadLine();
    if (string.Equals(replaceResponse?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
    {
      File.Copy(tool._currentPath, tool._backupPath, true);
      Console.WriteLine("Backup database successfully replaced.");
    }
    else
    {
      Console.WriteLine("Remember to replace your backup database manually.");
    }

    return 0;
  }

  private static async Task<List<string>> GetGameColumnsAsync(SqliteConnection conn)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "PRAGMA table_info(Games);";

    var columns = new List<string>();
    using var reader = await cmd.ExecuteReaderAsync();
    while (await reader.ReadAsync())
    {
      var name = reader.IsDBNull(1) ? null : reader.GetString(1);
      if (!string.IsNullOrWhiteSpace(name))
      {
        columns.Add(name);
      }
    }

    return columns;
  }

  private static async Task<List<Dictionary<string, object?>>> FindChangesAsync(
    SqliteConnection currentConn,
    SqliteConnection backupConn,
    IReadOnlyList<string> columns,
    DiffLetterSelection letters)
  {
    var changes = new List<Dictionary<string, object?>>();
    string columnList = string.Join(", ", columns.Select(c => $"[{c}]"));

    using var currentCmd = currentConn.CreateCommand();
    currentCmd.CommandText = $"SELECT {columnList} FROM Games WHERE {letters.WhereClause} ORDER BY Name";
    letters.ApplyParameters(currentCmd);

    using var currentReader = await currentCmd.ExecuteReaderAsync();
    int scanned = 0;
    while (await currentReader.ReadAsync())
    {
      var currentRow = ReadRow(currentReader, columns);
      int? gameId = GetInt32Value(currentRow, "GA_Id");

      var backupRow = gameId.HasValue
        ? await LoadBackupRowAsync(backupConn, columns, gameId.Value)
        : null;

      if (backupRow == null || RowsDiffer(currentRow, backupRow, columns))
      {
        changes.Add(currentRow);
      }

      scanned++;
      if (scanned % 500 == 0)
      {
        Console.WriteLine($"  Scanned {scanned} games...");
      }
    }

    return changes;
  }

  private static async Task<Dictionary<string, object?>?> LoadBackupRowAsync(
    SqliteConnection backupConn,
    IReadOnlyList<string> columns,
    int gaId)
  {
    string columnList = string.Join(", ", columns.Select(c => $"[{c}]"));

    using var cmd = backupConn.CreateCommand();
    cmd.CommandText = $"SELECT {columnList} FROM Games WHERE GA_Id = @GA_Id";
    cmd.Parameters.AddWithValue("@GA_Id", gaId);

    using var reader = await cmd.ExecuteReaderAsync();
    if (!await reader.ReadAsync())
    {
      return null;
    }

    return ReadRow(reader, columns);
  }

  private static Dictionary<string, object?> ReadRow(SqliteDataReader reader, IReadOnlyList<string> columns)
  {
    var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    for (int i = 0; i < columns.Count; i++)
    {
      row[columns[i]] = reader.IsDBNull(i) ? null : reader.GetValue(i);
    }
    return row;
  }

  private static bool RowsDiffer(
    IReadOnlyDictionary<string, object?> currentRow,
    IReadOnlyDictionary<string, object?> backupRow,
    IReadOnlyList<string> columns)
  {
    foreach (var column in columns)
    {
      string? currentValue = NormalizeForCompare(GetRowValue(currentRow, column));
      string? backupValue = NormalizeForCompare(GetRowValue(backupRow, column));
      if (!string.Equals(currentValue, backupValue, StringComparison.Ordinal))
      {
        return true;
      }
    }

    return false;
  }

  private static async Task InitializeUpdateDatabaseAsync(string outputPath)
  {
    var context = new GameBaseSqliteContext(outputPath);
    await context.InitializeDatabaseAsync();

    await context.ExecuteNonQueryAsync(@"
      CREATE TABLE IF NOT EXISTS Info (
        Name TEXT,
        Email TEXT,
        GameBase TEXT,
        Notes TEXT,
        Date TEXT
      )");
  }

  private static async Task WriteChangesAsync(
    string outputPath,
    IReadOnlyList<string> columns,
    IReadOnlyList<Dictionary<string, object?>> changes)
  {
    using var conn = new SqliteConnection($"Data Source={outputPath}");
    await conn.OpenAsync();

    using var transaction = conn.BeginTransaction();
    string columnList = string.Join(", ", columns.Select(c => $"[{c}]"));
    string parameterList = string.Join(", ", columns.Select(c => $"@{c}"));

    foreach (var row in changes)
    {
      using var cmd = conn.CreateCommand();
      cmd.Transaction = transaction;
      cmd.CommandText = $"INSERT INTO Games ({columnList}) VALUES ({parameterList})";

      foreach (var column in columns)
      {
        cmd.Parameters.AddWithValue($"@{column}", GetRowValue(row, column) ?? DBNull.Value);
      }

      await cmd.ExecuteNonQueryAsync();
    }

    await transaction.CommitAsync();
  }

  private static UpdateInfo PromptForUpdateInfo()
  {
    string name = PromptRequired("Type your name so the database maintainer knows who this update is from:", 255);
    string email = PromptRequired("Type your email address so the database maintainer knows how to contact you:", 255);
    string gameBase = PromptRequired("Type the name of your GameBase database:", 255);
    string notes = PromptOptional("Type any notes you wish the database maintainer to see:");
    string date = DateTime.Now.ToString("dd MMM yyyy", CultureInfo.InvariantCulture);

    return new UpdateInfo(name, email, gameBase, notes, date);
  }

  private static async Task WriteInfoAsync(string outputPath, UpdateInfo info)
  {
    using var conn = new SqliteConnection($"Data Source={outputPath}");
    await conn.OpenAsync();

    using var cmd = conn.CreateCommand();
    cmd.CommandText = "INSERT INTO Info (Name, Email, GameBase, Notes, Date) VALUES (@Name, @Email, @GameBase, @Notes, @Date)";
    cmd.Parameters.AddWithValue("@Name", info.Name);
    cmd.Parameters.AddWithValue("@Email", info.Email);
    cmd.Parameters.AddWithValue("@GameBase", info.GameBase);
    cmd.Parameters.AddWithValue("@Notes", info.Notes);
    cmd.Parameters.AddWithValue("@Date", info.Date);

    await cmd.ExecuteNonQueryAsync();
  }

  private static string PromptRequired(string prompt, int maxLength)
  {
    while (true)
    {
      Console.WriteLine();
      Console.WriteLine(prompt);
      string? input = Console.ReadLine();
      if (string.IsNullOrWhiteSpace(input))
      {
        Console.WriteLine("Value is required.");
        continue;
      }
      if (input.Length > maxLength)
      {
        Console.WriteLine($"Value must not exceed {maxLength} characters.");
        continue;
      }
      return input.Trim();
    }
  }

  private static string PromptOptional(string prompt)
  {
    Console.WriteLine();
    Console.WriteLine(prompt);
    return Console.ReadLine()?.Trim() ?? string.Empty;
  }

  private static string? NormalizeForCompare(object? value)
  {
    if (value == null || value is DBNull)
    {
      return null;
    }

    if (value is byte[] bytes)
    {
      return Convert.ToBase64String(bytes);
    }

    return Convert.ToString(value, CultureInfo.InvariantCulture);
  }

  private static string NormalizeForDisplay(object? value)
  {
    return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
  }

  private static object? GetRowValue(IReadOnlyDictionary<string, object?> row, string column)
  {
    return row.TryGetValue(column, out var value) ? value : null;
  }

  private static int? GetInt32Value(IReadOnlyDictionary<string, object?> row, string column)
  {
    var value = GetRowValue(row, column);
    if (value == null || value is DBNull)
    {
      return null;
    }

    try
    {
      return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }
    catch
    {
      return null;
    }
  }

  private static bool TryParseArgs(string[] args, out SqliteDiffTool tool, out string error)
  {
    tool = null!;
    error = string.Empty;

    if (args.Length < 2)
    {
      error = "Error: Missing required arguments for diff.";
      return false;
    }

    var positional = new List<string>();
    string? lettersArg = null;

    for (int i = 0; i < args.Length; i++)
    {
      string arg = args[i];
      if (arg.Equals("--letters", StringComparison.OrdinalIgnoreCase))
      {
        if (i + 1 >= args.Length)
        {
          error = "Error: --letters requires a value.";
          return false;
        }
        lettersArg = args[i + 1];
        i++;
      }
      else if (arg.StartsWith("--", StringComparison.Ordinal))
      {
        error = $"Error: Unknown option {arg}.";
        return false;
      }
      else
      {
        positional.Add(arg);
      }
    }

    if (positional.Count < 2)
    {
      error = "Error: diff requires <current-db> and <backup-db>.";
      return false;
    }

    string currentPath = positional[0];
    string backupPath = positional[1];
    string outputPath = positional.Count >= 3
      ? positional[2]
      : Path.Combine(Path.GetDirectoryName(currentPath) ?? ".", "update.db");

    var letters = DiffLetterSelection.Parse(lettersArg);
    tool = new SqliteDiffTool(currentPath, backupPath, outputPath, letters);
    return true;
  }

  private static void PrintUsage()
  {
    Console.WriteLine("Diff usage:");
    Console.WriteLine("  GBMigrator diff <current-db> <backup-db> [output-db] [--letters <letters>]");
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine("  GBMigrator diff C:\\GameBase\\Current.db C:\\GameBase\\Backup.db");
    Console.WriteLine("  GBMigrator diff Current.db Backup.db update.db --letters A,B,C,0-9");
    Console.WriteLine();
    Console.WriteLine("Letters:");
    Console.WriteLine("  Use A-Z to restrict by starting letter. Use 0-9 to include non-letter names.");
    Console.WriteLine("  If omitted, all letters and 0-9 are included.");
  }

  private sealed record UpdateInfo(string Name, string Email, string GameBase, string Notes, string Date);

  private sealed class DiffLetterSelection
  {
    private readonly HashSet<char> _letters;
    private readonly bool _includeNonLetters;

    private DiffLetterSelection(HashSet<char> letters, bool includeNonLetters)
    {
      _letters = letters;
      _includeNonLetters = includeNonLetters;
    }

    public bool HasSelection => _letters.Count > 0 || _includeNonLetters;

    public string WhereClause
    {
      get
      {
        var clauses = new List<string>();
        int index = 0;
        foreach (var letter in _letters.OrderBy(c => c))
        {
          clauses.Add($"lower(Name) LIKE @letter{index}");
          index++;
        }

        if (_includeNonLetters)
        {
          clauses.Add("NOT (" + string.Join(" OR ", Enumerable.Range('a', 26).Select(i => $"lower(Name) LIKE '{(char)i}%'")) + ")");
        }

        return clauses.Count == 0 ? "1=0" : string.Join(" OR ", clauses);
      }
    }

    public void ApplyParameters(SqliteCommand command)
    {
      int index = 0;
      foreach (var letter in _letters.OrderBy(c => c))
      {
        command.Parameters.AddWithValue($"@letter{index}", $"{char.ToLowerInvariant(letter)}%");
        index++;
      }
    }

    public static DiffLetterSelection Parse(string? lettersArg)
    {
      if (string.IsNullOrWhiteSpace(lettersArg))
      {
        return new DiffLetterSelection(new HashSet<char>(Enumerable.Range('A', 26).Select(i => (char)i)), true);
      }

      var letters = new HashSet<char>();
      bool includeNonLetters = false;

      foreach (var token in lettersArg.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
      {
        string trimmed = token.Trim();
        if (trimmed.Equals("all", StringComparison.OrdinalIgnoreCase))
        {
          foreach (var ch in Enumerable.Range('A', 26).Select(i => (char)i))
          {
            letters.Add(ch);
          }
          includeNonLetters = true;
          continue;
        }

        if (trimmed.Equals("0-9", StringComparison.OrdinalIgnoreCase))
        {
          includeNonLetters = true;
          continue;
        }

        if (trimmed.Length == 3 && trimmed[1] == '-')
        {
          char start = char.ToUpperInvariant(trimmed[0]);
          char end = char.ToUpperInvariant(trimmed[2]);
          if (start >= 'A' && end <= 'Z' && start <= end)
          {
            for (char c = start; c <= end; c++)
            {
              letters.Add(c);
            }
          }
          continue;
        }

        foreach (char c in trimmed)
        {
          if (char.IsLetter(c))
          {
            letters.Add(char.ToUpperInvariant(c));
          }
        }
      }

      return new DiffLetterSelection(letters, includeNonLetters);
    }
  }
}
