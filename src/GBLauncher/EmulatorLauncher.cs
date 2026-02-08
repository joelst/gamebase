// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBLauncher: Emulator launcher using Gemus engine

using GBFuncs;
using Gemus;
using Gemus.Models;

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

      // Initialize Gemus engine dependencies
      var funcs = new FunctionsHelper();
      var registry = new RegistryService();
      var gemusLogger = new GemusLogger();
      var unpacker = new ArchiveUnpacker();
      var gemus = new GemusEngine(funcs, registry, gemusLogger, unpacker);

      // Configure Gemus
      gemus.SetAppContext(appPath, string.Empty, "GBLauncher");
      gemus.Init(databaseName);
      gemus.SetSilentMode(true);

      // Configure rezip settings from INI
      string gemusIniPath = Path.Combine(appPath, databaseName, "Gemus.ini");
      bool rezip = funcs.INI_GetSetting("ReZip", "ReZip", "0", gemusIniPath) == "1";
      bool rezipAsk = funcs.INI_GetSetting("ReZip", "AskBeforeRepack", "1", gemusIniPath) == "1";
      bool rezipNotify = funcs.INI_GetSetting("ReZip", "NotifyRepackSuccess", "0", gemusIniPath) == "1";
      bool rezipLegacy = funcs.INI_GetSetting("ReZip", "UseLegacyMethod", "0", gemusIniPath) == "1";

      gemus.SetRezip(rezip, rezipAsk, RezipAdditionalFiles.Ignore);
      gemus.SetRezipNotify(rezipNotify);
      gemus.SetRezipLegacy(rezipLegacy);

      if (waitForExit)
      {
        gemus.SetWaitIfNotRezipping(true);
      }

      // Prepare default file (may be modified by RunEmulator)
      string defaultFile = game.FileToRun ?? string.Empty;

      // Get extract path from database
      string extractPath = Path.Combine(appPath, databaseName, "GB_Game");
      Directory.CreateDirectory(extractPath);

      _logger.Log($"Launching with file: {gameFilePath}");

      // Launch the game
      string error = gemus.RunEmulator(
        emulatorIndex,
        gameFilePath ?? string.Empty,
        game.FilenameIndex,
        ref defaultFile,
        extractPath,
        fromDb: true,
        numPlayersFrom: game.PlayersFrom > 0 ? game.PlayersFrom : null,
        numPlayersTo: game.PlayersTo > 0 ? game.PlayersTo : null,
        control: game.Control > 0 ? game.Control : null,
        gameComment: game.Comment,
        versionComment: game.VersionComment,
        palNtsc: game.PalNtsc >= 0 ? game.PalNtsc : null,
        trueDriveEmu: game.TrueDriveEmulation,
        notes: game.GemusPairs
      );

      if (!string.IsNullOrEmpty(error))
      {
        _logger.Log($"GEMUS error: {error}", true);
      }
      else
      {
        _logger.Log("Game launched successfully by GEMUS");
      }
    }
    catch (Exception ex)
    {
      _logger.Log($"Error launching game: {ex.Message}", true);
    }
  }
}
