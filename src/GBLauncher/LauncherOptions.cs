namespace GBLauncher;

/// <summary>
/// Launcher options parsed from command-line arguments.
/// </summary>
internal class LauncherOptions
{
  public string DatabaseName { get; set; } = string.Empty;
  public string? GameId { get; set; }
  public string? GameFilename { get; set; }
  public int EmulatorIndex { get; set; } = -1; // -1 means default
  public bool WaitForExit { get; set; }

  public bool IsValid()
  {
    if (string.IsNullOrEmpty(DatabaseName))
      return false;

    // Must have either GameId or GameFilename, but not both
    var hasId = !string.IsNullOrEmpty(GameId);
    var hasFilename = !string.IsNullOrEmpty(GameFilename);
    return hasId != hasFilename; // XOR: exactly one must be true
  }
}
