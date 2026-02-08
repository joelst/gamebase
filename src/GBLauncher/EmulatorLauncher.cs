using Gemus;

namespace GBLauncher;

/// <summary>
/// Launches emulators via Gemus library.
/// </summary>
internal class EmulatorLauncher
{
    private readonly Logger _logger;

    public EmulatorLauncher(Logger logger)
    {
        _logger = logger;
    }

    public void Launch(
        string appPath,
        string databaseName,
        string? gameFilePath,
        GameData game,
        int emulatorIndex,
        bool waitForExit)
    {
        try
        {
            _logger.Log($"Launching game ID [{game.Id}] with emulator index [{emulatorIndex}]");

            var gemus = new GemusEngine();
            
            // Set up Gemus parameters
            var result = gemus.LaunchGame(
                appPath,
                databaseName,
                gameFilePath ?? "",
                game.GemusPairs,
                emulatorIndex,
                game.TrueDriveEmulation,
                waitForExit
            );

            if (result)
            {
                _logger.Log("Game launched successfully");
            }
            else
            {
                _logger.Log("Failed to launch game", true);
            }
        }
        catch (Exception ex)
        {
            _logger.Log($"Error launching game: {ex.Message}", true);
        }
    }
}
