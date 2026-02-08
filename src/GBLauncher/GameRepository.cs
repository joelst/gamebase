using Microsoft.Data.Sqlite;

namespace GBLauncher;

/// <summary>
/// Repository for accessing game data from SQLite database.
/// </summary>
internal class GameRepository : IDisposable
{
    private readonly SqliteConnection _connection;
    private readonly Logger _logger;

    public GameRepository(string databasePath, Logger logger)
    {
        _logger = logger;
        
        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = databasePath,
            Mode = SqliteOpenMode.ReadOnly
        }.ToString();

        _connection = new SqliteConnection(connectionString);
        _connection.Open();
        _logger.Log("Opened Database");
    }

    public (int Major, int Minor) GetDatabaseVersion()
    {
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT MajorVersion, MinorVersion FROM Config LIMIT 1";
        
        using var reader = command.ExecuteReader();
        if (reader.Read())
        {
            return (reader.GetInt32(0), reader.GetInt32(1));
        }

        _logger.Log("Missing Version/Update info in Config Table!", true);
        return (0, 0); // Won't reach here due to forceQuit
    }

    public GameData? GetGameById(int gameId)
    {
        var query = @"SELECT GA_Id, Comment, Control, Filename, FileToRun, FilenameIndex, FA, V_Comment, 
                      PlayersTo, PlayersFrom, V_PalNTSC, Gemus, V_TrueDriveEmu 
                      FROM Games WHERE GA_Id = @id";

        using var command = _connection.CreateCommand();
        command.CommandText = query;
        command.Parameters.AddWithValue("@id", gameId);

        _logger.Log($"SQL used to read game information [{query}]");

        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            return null;
        }

        return new GameData
        {
            Id = reader.GetInt32(0),
            Comment = reader.IsDBNull(1) ? "" : reader.GetString(1),
            Control = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
            Filename = reader.IsDBNull(3) ? "" : reader.GetString(3),
            FileToRun = reader.IsDBNull(4) ? "" : reader.GetString(4),
            FilenameIndex = reader.IsDBNull(5) ? 0 : reader.GetInt32(5),
            FileAvailable = reader.IsDBNull(6) ? false : reader.GetBoolean(6),
            VersionComment = reader.IsDBNull(7) ? "" : reader.GetString(7),
            PlayersTo = reader.IsDBNull(8) ? 0 : reader.GetInt32(8),
            PlayersFrom = reader.IsDBNull(9) ? 0 : reader.GetInt32(9),
            PalNtsc = reader.IsDBNull(10) ? 0 : reader.GetInt32(10),
            GemusPairs = reader.IsDBNull(11) ? "" : reader.GetString(11),
            TrueDriveEmulation = reader.IsDBNull(12) ? false : reader.GetBoolean(12)
        };
    }

    public int? GetGameIdByFilename(string filename)
    {
        // Try multiple query strategies (VB6 had 4 queries with increasing fuzziness)
        var queries = new[]
        {
            $"SELECT GA_Id FROM Games WHERE Filename = '{EscapeWildcards(filename)}'",
            $"SELECT GA_Id FROM Games WHERE Filename LIKE '%\\{EscapeWildcards(filename)}'",
            $"SELECT GA_Id FROM Games WHERE Filename LIKE '{EscapeWildcards(filename)}.%'",
            $"SELECT GA_Id FROM Games WHERE Filename LIKE '%\\{EscapeWildcards(filename)}.%'"
        };

        foreach (var query in queries)
        {
            using var command = _connection.CreateCommand();
            command.CommandText = query;

            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                _logger.Log($"Game ID determined [{reader.GetInt32(0)}] from Game Filename [{filename}]");
                _logger.Log($"SQL used to determine Game ID [{query}]");
                return reader.GetInt32(0);
            }
            else
            {
                _logger.Log($"SQL [{query}] returned no rows");
            }
        }

        return null;
    }

    private static string EscapeWildcards(string input)
    {
        // Escape SQL LIKE wildcards for literal matching
        return input
            .Replace("[", "[[]")
            .Replace("?", "[?]")
            .Replace("*", "[*]")
            .Replace("!", "[!]")
            .Replace("#", "[#]");
    }

    public void Dispose()
    {
        _connection?.Dispose();
    }
}
