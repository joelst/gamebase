//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  Database Functions used in GB
//==============================================

using System.Runtime.InteropServices;
using System.Text;

namespace GBSql
{
  /// <summary>
  /// Types of second-column display handling.
  /// Converted from VB6 SecondColTypeEnum.
  /// </summary>
  public enum SecondColType
  {
    Normal = 0,
    YesNo = 1,
    Publisher = 2,
    Year = 3,
    Trainers = 4,
    Players = 5,
    Genres = 6,
    Rating = 7,
    PalNtsc = 8,
    Control = 9,
    Length = 10,
    ReviewRating = 11
  }

  /// <summary>
  /// Represents a single second-column field definition.
  /// Converted from VB6 SecondColumnType.
  /// </summary>
  public class SecondColumnItem
  {
    public string MenuText { get; set; } = string.Empty;
    public string ColumnTitle { get; set; } = string.Empty;
    public string Fields { get; set; } = string.Empty;
    public string MusicFields { get; set; } = string.Empty;
    public SecondColType Type { get; set; } = SecondColType.Normal;
    public bool GameOnlyTable { get; set; } = true;
  }

  /// <summary>
  /// Represents a menu array of second-column items.
  /// </summary>
  public class SecondColumnArray
  {
    public string MenuText { get; set; } = string.Empty;
    public List<SecondColumnItem> Items { get; set; } = new();
    public int MaxColumns => Items.Count - 1;
  }

  /// <summary>
  /// Manages the second column information for the menus in GameBase.
  /// Converted from VB6 clsSecCol class.
  /// </summary>
  public class SecondColumnManager
  {
    private const int MaxMenuArrays = 2;
    private const string MenuBar = "-";

    private readonly List<SecondColumnArray> _arrays = new();
    private string _iniFile = string.Empty;

    // P/Invoke for INI file operations (maintaining compatibility)
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern int GetPrivateProfileString(
        string lpAppName,
        string lpKeyName,
        string lpDefault,
        StringBuilder lpReturnedString,
        int nSize,
        string lpFileName);

    public SecondColumnManager()
    {
      InitColumns();
    }

    /// <summary>
    /// Sets the INI file path for reading localised text.
    /// </summary>
    public void SetIniFile(string iniFile)
    {
      _iniFile = iniFile;
    }

    /// <summary>
    /// Gets a message from the INI file.
    /// </summary>
    private string GetINIMessage(string section, string key, string defaultValue)
    {
      if (string.IsNullOrEmpty(_iniFile))
        return defaultValue;

      var sb = new StringBuilder(2096);
      int len = GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, _iniFile);
      return sb.ToString(0, len);
    }

    /// <summary>
    /// Returns the data for a specific second-column field.
    /// </summary>
    public void GetSecondColField(
        int menuArrayIndex,
        int fieldArrayIndex,
        out string menuText,
        out string columnTitle,
        out string fields,
        out string musicFields,
        out int type,
        out bool gameOnlyTable)
    {
      var item = _arrays[menuArrayIndex].Items[fieldArrayIndex];
      menuText = item.MenuText;
      columnTitle = item.ColumnTitle;
      fields = item.Fields;
      musicFields = item.MusicFields;
      type = (int)item.Type;
      gameOnlyTable = item.GameOnlyTable;
    }

    /// <summary>
    /// Gets the total number of menu arrays.
    /// </summary>
    public int MaxArrays => MaxMenuArrays;

    /// <summary>
    /// Gets the maximum field index for a given menu array.
    /// </summary>
    public int GetMaxFields(int arrayIndex) => _arrays[arrayIndex].MaxColumns;

    /// <summary>
    /// Gets the menu text for a menu array.
    /// </summary>
    public string GetMenuText(int arrayIndex) => _arrays[arrayIndex].MenuText;

    /// <summary>
    /// Removes ampersand accelerator markers from menu text.
    /// </summary>
    private static string RemoveAmpersand(string menuText)
    {
      string result = menuText.Replace("&&", "~~~~");
      result = result.Replace("&", "");
      result = result.Replace("~~~~", "&");
      return result;
    }

    /// <summary>
    /// Initializes all second-column arrays with their field definitions.
    /// </summary>
    public void InitColumns()
    {
      _arrays.Clear();

      // Menu array 0: Game Info
      var gameInfo = new SecondColumnArray
      {
        MenuText = GetINIMessage("MainMenus", "Menu_View_SC_GameInfo", "Game Info")
      };
      gameInfo.Items.AddRange(new[]
      {
                CreateItem(SecondColType.Publisher, "Menu_View_SC_Publisher", "Software House/Publisher",
                    "Publishers.Publisher AS ColX", "", true),
                CreateItem(SecondColType.Year, "Menu_View_SC_Year", "Year",
                    "Games.YE_Id AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Developer", "Developer",
                    "Developers.Developer AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Programmer", "Programmer",
                    "Programmers.Programmer AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Musician", "Musician",
                    "Musicians.Musician AS ColX", "Musicians.Musician AS ColX", false),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Artist", "Artist",
                    "Artists.Artist AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_License", "License",
                    "Licenses.License AS ColX", "", true),
                CreateItem(SecondColType.Genres, "Menu_View_SC_Genre", "Genre",
                    "Genres.Genre AS ColX, PGenres.PG_Id as ColXExtra", "", true),
                CreateItem(SecondColType.Players, "Menu_View_SC_Players", "No. of Players",
                    "Games.PlayersFrom AS ColX, Games.PlayersTo AS ColXExtra, Games.PlayersSim AS ColXExtra2", "", true),
                CreateItem(SecondColType.Control, "Menu_View_SC_Control", "Control",
                    "Games.Control as ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Language", "Language",
                    "Languages.Language AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Rarity", "Rarity",
                    "Rarities.Rarity AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_WebLinkName", "More Info Name",
                    "Games.WebLink_Name AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_WebLinkURL", "More Info URL",
                    "Games.WebLink_URL AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Comment", "Comment",
                    "Games.Comment AS ColX", "", true),
            });
      _arrays.Add(gameInfo);

      // Menu array 1: Version Info
      var versionInfo = new SecondColumnArray
      {
        MenuText = GetINIMessage("MainMenus", "Menu_View_SC_VersionInfo", "Version Info")
      };
      versionInfo.Items.AddRange(new[]
      {
                CreateItem(SecondColType.Normal, "Menu_View_SC_GameFilename", "Game Filename",
                    "Games.Filename AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_CrackingGroup", "Cracking Group",
                    "Crackers.Cracker AS ColX", "", true),
                CreateItem(SecondColType.Trainers, "Menu_View_SC_Trainers", "No. of Trainers",
                    "Games.V_Trainers AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_Playable", "Playable",
                    "Games.V_Playable AS ColX", "", true),
                CreateItem(SecondColType.Length, "Menu_View_SC_GameLength", "Game Length",
                    "Games.V_Length AS ColX, Games.V_LengthType AS ColXExtra", "", true),
                CreateItem(SecondColType.PalNtsc, "Menu_View_SC_NTSC", "PAL/NTSC",
                    "Games.V_PalNTSC AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_VWebLinkName", "Download Name",
                    "Games.V_WebLink_Name AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_VWebLinkURL", "Download URL",
                    "Games.V_WebLink_URL AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_Original", "Original",
                    "Games.V_Original AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_LoadingScreen", "Loading Screen",
                    "Games.V_LoadingScreen AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_TitleScreen", "Title Screen",
                    "Games.V_TitleScreen AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_IncludedDocs", "Included Docs",
                    "Games.V_IncludedDocs AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_HighScoreSaver", "High Score Saver",
                    "Games.V_HighScoreSaver AS ColX", "", true),
                CreateItem(SecondColType.YesNo, "Menu_View_SC_TrueDriveEmu", "True Drive Emulation",
                    "Games.V_TrueDriveEmu AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_VersionComment", "Version Comment",
                    "Games.V_Comment AS ColX", "", true),
                // Menu bar separator
                new SecondColumnItem { MenuText = MenuBar, GameOnlyTable = true },
                CreateItem(SecondColType.Normal, "Menu_View_SC_SidFilename", "Music Filename",
                    "Games.SidFilename AS ColX", "Music.Filename AS ColX", false),
                CreateItem(SecondColType.Normal, "Menu_View_SC_PictureFilename", "Picture Filename",
                    "Games.ScrnShotFilename AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_MusicianNick", "Musician Nickname",
                    "Musicians.Nick AS ColX", "Musicians.Nick AS ColX", false),
                CreateItem(SecondColType.Normal, "Menu_View_SC_MusicianGroup", "Musician Group",
                    "Musicians.Grp AS ColX", "Musicians.Grp AS ColX", false),
                CreateItem(SecondColType.Normal, "Menu_View_SC_MusicianPhotoFilename", "Musician Photo Filename",
                    "Musicians.Photo AS ColX", "Musicians.Photo AS ColX", false),
            });
      _arrays.Add(versionInfo);

      // Menu array 2: Personal Info
      var personalInfo = new SecondColumnArray
      {
        MenuText = GetINIMessage("MainMenus", "Menu_View_SC_Personal", "Personal Info")
      };
      personalInfo.Items.AddRange(new[]
      {
                CreateItem(SecondColType.Normal, "Menu_View_SC_HighScore", "High Score",
                    "Games.HighScore AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_Difficulty", "Difficulty Level",
                    "Difficulty.Difficulty AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_TimesPlayed", "Times Played",
                    "Games.TimesPlayed AS ColX", "", true),
                CreateItem(SecondColType.Normal, "Menu_View_SC_LastPlayed", "Date Last Played",
                    "Games.DateLastPlayed AS ColX", "", true),
                CreateItem(SecondColType.Rating, "Menu_View_SC_Rating", "Rating",
                    "Games.Rating as ColX", "", true),
                CreateItem(SecondColType.ReviewRating, "Menu_View_SC_ReviewRating", "Review Rating",
                    "Games.ReviewRating AS ColX", "", true),
            });
      _arrays.Add(personalInfo);
    }

    /// <summary>
    /// Helper to create a SecondColumnItem from INI-backed menu text.
    /// </summary>
    private SecondColumnItem CreateItem(
        SecondColType type,
        string iniKey,
        string defaultText,
        string fields,
        string musicFields,
        bool gameOnlyTable)
    {
      string menuText = GetINIMessage("MainMenus", iniKey, defaultText);
      return new SecondColumnItem
      {
        Type = type,
        MenuText = menuText,
        ColumnTitle = RemoveAmpersand(menuText),
        Fields = fields,
        MusicFields = musicFields,
        GameOnlyTable = gameOnlyTable
      };
    }
  }
}
