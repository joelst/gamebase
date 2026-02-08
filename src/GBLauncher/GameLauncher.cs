using GBFuncs;

namespace GBLauncher;

/// <summary>
/// Main orchestrator for game launching.
/// </summary>
internal class GameLauncher
{
    private readonly LauncherOptions _options;
    private readonly Logger _logger;
    private readonly string _appPath;
    private readonly FunctionsHelper _functions;

    public GameLauncher(LauncherOptions options, Logger logger)
    {
        _options = options;
        _logger = logger;
        _appPath = AppDomain.CurrentDomain.BaseDirectory;
        if (!_appPath.EndsWith(Path.DirectorySeparatorChar))
            _appPath += Path.DirectorySeparatorChar;

        _functions = new FunctionsHelper();
        _logger.Log($"App Path: [{_appPath}]");
    }

    public void Launch()
    {
        // Construct database path
        var dbPath = Path.Combine(_appPath, _options.DatabaseName, $"{_options.DatabaseName}.db");
        if (!File.Exists(dbPath))
        {
            _logger.Log($"Database: [{dbPath}] not found", true);
            return;
        }

        _logger.Log($"Database name: [{_options.DatabaseName}]");
        _logger.Log($"Database file: [{dbPath}]");

        var pathsFile = Path.Combine(_appPath, _options.DatabaseName, "Paths.ini");
        var settingsFile = Path.Combine(_appPath, _options.DatabaseName, "Settings.ini");
        _logger.Log($"Paths file: [{pathsFile}]");
        _logger.Log($"Settings file: [{settingsFile}]");

        // Load settings
        var reZip = _functions.INI_GetSetting("General", "ReZipGames", "0", settingsFile) == "1";
        var reZipAsk = _functions.INI_GetSetting("General", "ReZipAsk", "1", settingsFile) == "1";
        var zipNotify = _functions.INI_GetSetting("General", "ReZipNotify", "1", settingsFile) == "1";
        var reZipLegacy = _functions.INI_GetSetting("General", "ReZipLegacy", "0", settingsFile) == "1";
        var reZipAdditional = int.Parse(_functions.INI_GetSetting("General", "ReZipAdditional", "2", settingsFile));

        _logger.Log($"ReZipGames: [{reZip}]");
        _logger.Log($"ReZipAsk: [{reZipAsk}]");
        _logger.Log($"ReZipNotify: [{zipNotify}]");
        _logger.Log($"ReZipLegacy: [{reZipLegacy}]");
        _logger.Log($"ReZipAdditional: [{(reZipAdditional == 0 ? "Ignore" : reZipAdditional == 1 ? "Zip" : "Ask")}]");

        // Open database and validate version
        using var repository = new GameRepository(dbPath, _logger);
        var (major, minor) = repository.GetDatabaseVersion();
        var dbVersion = major + (minor / 10.0);
        
        const double requiredVersion = 2.5;
        if (dbVersion < requiredVersion)
        {
            _logger.Log($"Minimum required GameBase database version: [{requiredVersion}], you have: [{dbVersion}]", true);
            return;
        }
        _logger.Log($"Database version validated [{dbVersion}]");

        // Load game paths
        var gamePaths = LoadGamePaths(pathsFile);
        if (gamePaths.Count == 0)
        {
            _logger.Log($"No valid game paths exist as specified in [{pathsFile}]", true);
            return;
        }

        // Get game ID
        int gameId;
        if (!string.IsNullOrEmpty(_options.GameId))
        {
            gameId = int.Parse(_options.GameId);
            _logger.Log($"Using provided game ID: [{gameId}]");
        }
        else if (!string.IsNullOrEmpty(_options.GameFilename))
        {
            var foundId = repository.GetGameIdByFilename(_options.GameFilename);
            if (!foundId.HasValue)
            {
                _logger.Log($"Game ID unable to be determined from Game Filename [{_options.GameFilename}]", true);
                return;
            }
            gameId = foundId.Value;
        }
        else
        {
            _logger.Log("No game ID or filename provided", true);
            return;
        }

        // Load game info
        var game = repository.GetGameById(gameId);
        if (game == null)
        {
            _logger.Log($"Game ID [{gameId}] not found", true);
            return;
        }
        _logger.Log($"Game information loaded for game ID [{gameId}]");

        // Find game file
        string? gameFilePath = null;
        if (game.FileAvailable && gamePaths.Count > 0)
        {
            gameFilePath = FindGameFile(game.Filename, gamePaths);
            if (gameFilePath == null)
            {
                _logger.Log($"Gamefile [{game.Filename}] not located", true);
                return;
            }
            _logger.Log($"Game file found: [{gameFilePath}]");
        }

        // Launch via Gemus
        var emulatorLauncher = new EmulatorLauncher(_logger);
        emulatorLauncher.Launch(
            _appPath,
            _options.DatabaseName,
            gameFilePath,
            game,
            _options.EmulatorIndex,
            _options.WaitForExit
        );
    }

    private List<string> LoadGamePaths(string pathsFile)
    {
        var paths = new List<string>();
        var pathIdx = 1;

        while (true)
        {
            var path = _functions.INI_GetSetting("Games", pathIdx.ToString(), "", pathsFile);
            if (string.IsNullOrEmpty(path))
                break;

            var validatedPath = PathHelper.ValidatePath(path, _appPath);
            if (validatedPath != null)
            {
                paths.Add(validatedPath);
                _logger.Log($"Game path found: [{validatedPath}]");
            }

            pathIdx++;
        }

        return paths;
    }

    private string? FindGameFile(string filename, List<string> gamePaths)
    {
        foreach (var basePath in gamePaths)
        {
            var fullPath = Path.Combine(basePath, filename);
            if (File.Exists(fullPath))
            {
                return fullPath;
            }
        }
        return null;
    }
}
