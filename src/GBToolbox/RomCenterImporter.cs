//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: RomCenter 2.00 importer
//==============================================

using GBSql.Data;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace GBToolbox;

internal sealed class RomCenterImporter
{
  private const char Delimiter = '\u00FF';
  private static readonly uint[] Crc32Table = BuildCrc32Table();

  public async Task<RomCenterImportResult> ImportAsync(RomCenterImportOptions options, ProgressForm progress, CancellationToken token)
  {
    var entries = options.IgnoreDuplicates
      ? ReadEntriesDedup(options, progress, token)
      : ReadEntries(options, progress, token);

    if (token.IsCancellationRequested)
    {
      return RomCenterImportResult.Cancelled();
    }

    if (entries.Count == 0)
    {
      return RomCenterImportResult.Failed("There were no games to import.");
    }

    if (options.ScanForRoms && !string.IsNullOrWhiteSpace(options.RomPath))
    {
      await ScanForRomsAsync(options, entries, progress, token);
      if (token.IsCancellationRequested)
      {
        return RomCenterImportResult.Cancelled();
      }
    }

    if (options.UseTosec)
    {
      ApplyTosecParsing(entries, progress, token);
      if (token.IsCancellationRequested)
      {
        return RomCenterImportResult.Cancelled();
      }
    }

    string dbPath = options.ExistingDatabasePath ?? await CreateNewDatabaseAsync(options, progress, token);
    if (string.IsNullOrWhiteSpace(dbPath))
    {
      return RomCenterImportResult.Failed("Failed to create database.");
    }

    await ImportEntriesAsync(dbPath, entries, progress, token);

    return RomCenterImportResult.Success(dbPath);
  }

  private List<RomCenterEntry> ReadEntries(RomCenterImportOptions options, ProgressForm progress, CancellationToken token)
  {
    var entries = new List<RomCenterEntry>();
    int count = 0;

    foreach (string line in File.ReadLines(options.DataFilePath))
    {
      if (token.IsCancellationRequested)
      {
        break;
      }

      if (line.Length == 0 || line[0] != Delimiter)
      {
        continue;
      }

      var fields = line.Split(Delimiter);
      if (fields.Length != 11)
      {
        continue;
      }

      var entry = RomCenterEntry.FromFields(fields, options);
      entries.Add(entry);
      count++;
      if (count % 250 == 0)
      {
        progress.UpdateStatus("Loading datafile...", count);
        Application.DoEvents();
      }
    }

    return entries;
  }

  private List<RomCenterEntry> ReadEntriesDedup(RomCenterImportOptions options, ProgressForm progress, CancellationToken token)
  {
    var entries = new List<RomCenterEntry>();
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    int count = 0;

    foreach (string line in File.ReadLines(options.DataFilePath))
    {
      if (token.IsCancellationRequested)
      {
        break;
      }

      if (line.Length == 0 || line[0] != Delimiter)
      {
        continue;
      }

      var fields = line.Split(Delimiter);
      if (fields.Length != 11)
      {
        continue;
      }

      string name = fields[options.NameColumnIndex];
      if (!seen.Add(name))
      {
        continue;
      }

      var entry = RomCenterEntry.FromFields(fields, options);
      entries.Add(entry);
      count++;
      if (count % 250 == 0)
      {
        progress.UpdateStatus("Loading datafile (skipping duplicates)...", count);
        Application.DoEvents();
      }
    }

    return entries;
  }

  private async Task ScanForRomsAsync(RomCenterImportOptions options, List<RomCenterEntry> entries, ProgressForm progress, CancellationToken token)
  {
    var extensions = options.MatchExtensions ? entries.Select(e => Path.GetExtension(e.FileName)).Where(e => !string.IsNullOrWhiteSpace(e)).Distinct(StringComparer.OrdinalIgnoreCase).ToList() : new List<string>();

    progress.UpdateStatus("Preparing file list...");
    Application.DoEvents();

    var files = new List<string>();
    foreach (string file in Directory.EnumerateFiles(options.RomPathWithSlash, "*", SearchOption.AllDirectories))
    {
      if (token.IsCancellationRequested)
      {
        return;
      }

      if (options.MatchExtensions && extensions.Count > 0)
      {
        string ext = Path.GetExtension(file);
        if (!extensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
        {
          continue;
        }
      }

      files.Add(file);
    }

    var hashMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    int index = 0;
    progress.UpdateStatus($"Calculating CRC-32 for {files.Count} files...");
    progress.UpdateProgress(0, files.Count);

    foreach (string file in files)
    {
      if (token.IsCancellationRequested)
      {
        return;
      }

      string crc = ComputeCrc32(file);
      long size = new FileInfo(file).Length;
      string key = crc + ":" + size.ToString(CultureInfo.InvariantCulture);
      if (!hashMap.ContainsKey(key))
      {
        hashMap[key] = file;
      }

      index++;
      progress.UpdateProgress(index, files.Count);
      if (index % 100 == 0)
      {
        progress.UpdateStatus("Calculating CRC-32...", index);
      }
      Application.DoEvents();
    }

    int matched = 0;
    progress.UpdateStatus("Matching CRC-32/Size...");
    progress.UpdateProgress(0, entries.Count);
    for (int i = 0; i < entries.Count; i++)
    {
      if (token.IsCancellationRequested)
      {
        return;
      }

      var entry = entries[i];
      string key = entry.Crc + ":" + entry.Size.ToString(CultureInfo.InvariantCulture);
      if (hashMap.TryGetValue(key, out string? filePath) && filePath != null)
      {
        entry.Found = true;
        entry.DbFile = filePath.Substring(options.RomPathWithSlash.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        matched++;
      }

      progress.UpdateProgress(i + 1, entries.Count);
      if (i % 250 == 0)
      {
        progress.UpdateStatus("Matching CRC-32/Size...", matched);
      }
      Application.DoEvents();
    }
  }

  private static string ComputeCrc32(string filePath)
  {
    using var stream = File.OpenRead(filePath);
    uint crc = 0xFFFFFFFF;
    byte[] buffer = new byte[8192];
    int bytesRead;
    while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
    {
      for (int i = 0; i < bytesRead; i++)
      {
        byte index = (byte)((crc ^ buffer[i]) & 0xFF);
        crc = (crc >> 8) ^ Crc32Table[index];
      }
    }

    crc ^= 0xFFFFFFFF;
    return crc.ToString("X8", CultureInfo.InvariantCulture);
  }

  private static uint[] BuildCrc32Table()
  {
    var table = new uint[256];
    for (uint i = 0; i < table.Length; i++)
    {
      uint value = i;
      for (int j = 0; j < 8; j++)
      {
        value = (value & 1) == 1 ? (value >> 1) ^ 0xEDB88320 : value >> 1;
      }
      table[i] = value;
    }
    return table;
  }

  private static void ApplyTosecParsing(List<RomCenterEntry> entries, ProgressForm progress, CancellationToken token)
  {
    progress.UpdateStatus("Obtaining DB info from TOSEC filenames...");
    progress.UpdateProgress(0, entries.Count);

    for (int i = 0; i < entries.Count; i++)
    {
      if (token.IsCancellationRequested)
      {
        return;
      }

      entries[i].ApplyTosecMetadata();
      progress.UpdateProgress(i + 1, entries.Count);
      if (i % 250 == 0)
      {
        progress.UpdateStatus("Parsing TOSEC metadata...", i + 1);
      }
      Application.DoEvents();
    }
  }

  private static async Task<string> CreateNewDatabaseAsync(RomCenterImportOptions options, ProgressForm progress, CancellationToken token)
  {
    string baseFolder = options.GameBaseFolder;
    Directory.CreateDirectory(baseFolder);
    Directory.CreateDirectory(Path.Combine(baseFolder, "Gfx"));
    Directory.CreateDirectory(Path.Combine(baseFolder, "Scripts"));
    Directory.CreateDirectory(Path.Combine(baseFolder, "Sounds"));
    Directory.CreateDirectory(Path.Combine(baseFolder, "Splash"));

    ToolboxResourceLoader.EnsureDefaultFiles(baseFolder);
    ToolboxResourceLoader.EnsureDefaultFiles(Path.Combine(baseFolder, "Gfx"));

    File.WriteAllText(Path.Combine(baseFolder, "Credits.txt"), $"Created by GBToolbox on {DateTime.Now:dddd dd MMMM yyyy}.\n");
    File.WriteAllText(Path.Combine(baseFolder, "Sounds", "sounds.txt"), "Place your Classic Game WAV files here.");
    File.WriteAllText(Path.Combine(baseFolder, "Splash", "Splash.txt"), "Place your splash screens here.");

    string dbPath = Path.Combine(baseFolder, options.GameBaseName + ".db");
    progress.UpdateStatus("Creating new database...");
    Application.DoEvents();

    var context = new GameBaseSqliteContext(dbPath);
    await context.InitializeDatabaseAsync();
    context.Dispose();

    if (options.ScanForRoms && !string.IsNullOrWhiteSpace(options.RomPath))
    {
      string pathsIni = Path.Combine(baseFolder, "Paths.ini");
      var helper = new GBFuncs.FunctionsHelper();
      helper.INI_WriteSetting("Games", "1", options.RomPath, pathsIni);

      string emusIni = Path.Combine(baseFolder, "Emulators.ini");
      helper.INI_WriteSetting("File Types", "SupportTypes", options.DefaultTypes, emusIni);
    }

    return dbPath;
  }

  private static async Task ImportEntriesAsync(string dbPath, List<RomCenterEntry> entries, ProgressForm progress, CancellationToken token)
  {
    progress.UpdateStatus("Adding games to database...");
    Application.DoEvents();

    using var conn = ToolboxDatabase.Open(dbPath);
    EnsureLookupDefaults(conn);

    using var tx = conn.BeginTransaction();
    int count = 0;
    foreach (var entry in entries)
    {
      if (token.IsCancellationRequested)
      {
        break;
      }

      if (string.IsNullOrWhiteSpace(entry.Name))
      {
        continue;
      }

      int publisherId = GetOrCreateLookupId(conn, "Publishers", "PU_Id", entry.Publisher ?? "Unknown Publisher");
      int languageId = GetOrCreateLookupId(conn, "Languages", "LA_Id", entry.Language ?? "English");
      int crackerId = GetOrCreateLookupId(conn, "Crackers", "CR_Id", entry.Cracker ?? "Unknown");
      int yearId = GetOrCreateLookupId(conn, "Years", "YE_Id", entry.YearName ?? "????");

      string filename = entry.ResolveFileName();

      using var cmd = conn.CreateCommand();
      cmd.Transaction = tx;
      cmd.CommandText = @"INSERT INTO Games (Name, YE_Id, PU_Id, LA_Id, CR_Id, Comment, Filename, CCode)
VALUES (@Name, @YE_Id, @PU_Id, @LA_Id, @CR_Id, @Comment, @Filename, @CCode);";
      cmd.Parameters.AddWithValue("@Name", entry.Name);
      cmd.Parameters.AddWithValue("@YE_Id", yearId);
      cmd.Parameters.AddWithValue("@PU_Id", publisherId);
      cmd.Parameters.AddWithValue("@LA_Id", languageId);
      cmd.Parameters.AddWithValue("@CR_Id", crackerId);
      cmd.Parameters.AddWithValue("@Comment", entry.Comment ?? string.Empty);
      cmd.Parameters.AddWithValue("@Filename", filename);
      cmd.Parameters.AddWithValue("@CCode", entry.Found ? "2" : "4");
      cmd.ExecuteNonQuery();

      count++;
      if (count % 250 == 0)
      {
        progress.UpdateStatus("Adding games to database...", count);
        Application.DoEvents();
      }
    }

    await tx.CommitAsync(token);
  }

  private static void EnsureLookupDefaults(SqliteConnection conn)
  {
    EnsureLookupRow(conn, "Publishers", "PU_Id", 1, "Unknown Publisher");
    EnsureLookupRow(conn, "Musicians", "MU_Id", 1, "Unknown Musician");
    EnsureLookupRow(conn, "Programmers", "PR_Id", 1, "Unknown Programmer");
    EnsureLookupRow(conn, "Languages", "LA_Id", 1, "English");
    EnsureLookupRow(conn, "Genres", "GE_Id", 1, "Uncategorized");
    EnsureLookupRow(conn, "Crackers", "CR_Id", 1, "Unknown");
    EnsureLookupRow(conn, "Years", "YE_Id", 1, "????");
  }

  private static void EnsureLookupRow(SqliteConnection conn, string table, string idColumn, int id, string name)
  {
    if (!ToolboxDatabase.TableExists(conn, table))
    {
      return;
    }

    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"INSERT OR IGNORE INTO {table} ({idColumn}, Name) VALUES (@Id, @Name);";
    cmd.Parameters.AddWithValue("@Id", id);
    cmd.Parameters.AddWithValue("@Name", name);
    cmd.ExecuteNonQuery();
  }

  private static int GetOrCreateLookupId(SqliteConnection conn, string table, string idColumn, string name)
  {
    using var select = conn.CreateCommand();
    select.CommandText = $"SELECT {idColumn} FROM {table} WHERE Name = @Name LIMIT 1;";
    select.Parameters.AddWithValue("@Name", name);
    var result = select.ExecuteScalar();
    if (result != null)
    {
      return Convert.ToInt32(result, CultureInfo.InvariantCulture);
    }

    using var insert = conn.CreateCommand();
    insert.CommandText = $"INSERT INTO {table} (Name) VALUES (@Name); SELECT last_insert_rowid();";
    insert.Parameters.AddWithValue("@Name", name);
    object? id = insert.ExecuteScalar();
    return Convert.ToInt32(id, CultureInfo.InvariantCulture);
  }
}

internal sealed class RomCenterImportOptions
{
  public string DataFilePath { get; set; } = string.Empty;
  public int NameColumnIndex { get; set; } = 2;
  public int FileColumnIndex { get; set; } = 4;
  public bool IgnoreDuplicates { get; set; }
  public bool UseTosec { get; set; }
  public bool AddZipExtension { get; set; }
  public bool ScanForRoms { get; set; }
  public bool MatchExtensions { get; set; } = true;
  public string RomPath { get; set; } = string.Empty;
  public string RomPathWithSlash { get; set; } = string.Empty;
  public string GameBaseName { get; set; } = string.Empty;
  public string GameBaseFolder { get; set; } = string.Empty;
  public string GameBaseFolderWithSlash { get; set; } = string.Empty;
  public string DefaultTypes { get; set; } = string.Empty;
  public string? ExistingDatabasePath { get; set; }
}

internal sealed class RomCenterEntry
{
  public string Name { get; private set; } = string.Empty;
  public string FileName { get; private set; } = string.Empty;
  public string DbFile { get; set; } = string.Empty;
  public string Crc { get; private set; } = string.Empty;
  public long Size { get; private set; }
  public bool Found { get; set; }
  public string? Comment { get; private set; }
  public string? Publisher { get; private set; }
  public string? Language { get; private set; }
  public string? Cracker { get; private set; }
  public string? YearName { get; private set; }

  public static RomCenterEntry FromFields(string[] fields, RomCenterImportOptions options)
  {
    var entry = new RomCenterEntry
    {
      Size = fields[7] != "N" ? ParseLong(fields[7]) : ParseLong(fields[9]),
      Crc = fields[6].ToUpperInvariant(),
      FileName = fields[options.FileColumnIndex],
      Name = fields[options.NameColumnIndex]
    };

    if (options.AddZipExtension)
    {
      entry.FileName += ".zip";
    }

    return entry;
  }

  public void ApplyTosecMetadata()
  {
    string name = Name;
    string file = FileName;

    if (name.StartsWith("ZZZ-UNK-", StringComparison.OrdinalIgnoreCase))
    {
      name = name.Substring(8);
    }
    if (file.StartsWith("ZZZ-UNK-", StringComparison.OrdinalIgnoreCase))
    {
      file = file.Substring(8);
    }

    int paren = name.IndexOf('(');
    if (paren > 0)
    {
      Name = name.Substring(0, paren).Trim();
    }

    string lower = file.ToLowerInvariant();
    if (lower.Contains("(pd)"))
    {
      Comment = "Public Domain";
    }

    if (lower.Contains("(pal)"))
    {
      YearName = ExtractYear(lower);
    }
    else
    {
      YearName = ExtractYear(lower);
    }

    Language = ExtractLanguage(lower) ?? "English";

    if (lower.Contains("[cr"))
    {
      Cracker = ExtractBracketed(lower, "cr") ?? "Unknown";
    }

    Publisher = ExtractPublisher(file);
  }

  public string ResolveFileName()
  {
    return Found && !string.IsNullOrWhiteSpace(DbFile) ? DbFile : FileName;
  }

  private static long ParseLong(string value)
  {
    return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
  }

  private static string? ExtractYear(string input)
  {
    int start = input.IndexOf('(');
    while (start >= 0)
    {
      int end = input.IndexOf(')', start + 1);
      if (end > start)
      {
        string token = input.Substring(start + 1, end - start - 1);
        if (token.Length == 4 && token.All(char.IsDigit))
        {
          return token;
        }
      }
      start = input.IndexOf('(', start + 1);
    }

    return null;
  }

  private static string? ExtractLanguage(string input)
  {
    if (input.Contains("(us)")) return "United States (English)";
    if (input.Contains("(uk)")) return "United Kingdom (English)";
    if (input.Contains("(en)")) return "English";
    if (input.Contains("(de)")) return "German";
    if (input.Contains("(fr)")) return "French";
    if (input.Contains("(es)")) return "Spanish";
    if (input.Contains("(it)")) return "Italian";
    if (input.Contains("(nl)")) return "Dutch";
    return null;
  }

  private static string? ExtractBracketed(string input, string prefix)
  {
    int start = input.IndexOf("[" + prefix, StringComparison.OrdinalIgnoreCase);
    if (start < 0)
    {
      return null;
    }
    int end = input.IndexOf(']', start + 1);
    if (end < 0)
    {
      return null;
    }
    string token = input.Substring(start + 1, end - start - 1);
    if (token.Length <= prefix.Length + 1)
    {
      return null;
    }
    return token.Substring(prefix.Length + 1).Trim();
  }

  private static string? ExtractPublisher(string file)
  {
    int first = file.IndexOf('(');
    if (first < 0)
    {
      return null;
    }
    int second = file.IndexOf('(', first + 1);
    if (second < 0)
    {
      return null;
    }
    int end = file.IndexOf(')', second + 1);
    if (end < 0)
    {
      return null;
    }
    string publisher = file.Substring(second + 1, end - second - 1).Trim();
    return publisher == "-" ? null : publisher;
  }
}

internal sealed class RomCenterImportResult
{
  public bool IsSuccess { get; private set; }
  public bool CancelledByUser { get; private set; }
  public string Message { get; private set; } = string.Empty;
  public string? DatabasePath { get; private set; }

  public static RomCenterImportResult Success(string dbPath)
  {
    return new RomCenterImportResult { IsSuccess = true, DatabasePath = dbPath };
  }

  public static RomCenterImportResult Failed(string message)
  {
    return new RomCenterImportResult { IsSuccess = false, Message = message };
  }

  public static RomCenterImportResult Cancelled()
  {
    return new RomCenterImportResult { IsSuccess = false, CancelledByUser = true };
  }
}
