namespace GBLauncher;

/// <summary>
/// Game data model.
/// </summary>
internal class GameData
{
    public int Id { get; set; }
    public string Filename { get; set; } = string.Empty;
    public string FileToRun { get; set; } = string.Empty;
    public int FilenameIndex { get; set; }
    public bool FileAvailable { get; set; }
    public string Comment { get; set; } = string.Empty;
    public string VersionComment { get; set; } = string.Empty;
    public int PlayersFrom { get; set; }
    public int PlayersTo { get; set; }
    public int PalNtsc { get; set; }
    public int Control { get; set; }
    public bool TrueDriveEmulation { get; set; }
    public string GemusPairs { get; set; } = string.Empty;
}
