using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Form showing detailed information about a game in the master database.
/// </summary>
public partial class MasterInfoForm : Form
{
  public MasterInfoForm()
  {
    InitializeComponent();
  }

  public void ShowGame(int gameId, SqliteConnection connection)
  {
    if (WindowState == FormWindowState.Minimized)
    {
      WindowState = FormWindowState.Normal;
    }

    // For now, just show the panel indicating we have a game
    // In the VB6 version, this form is mostly empty - it just switches panels
    // A full implementation would load all game details from the database
    picInfo.BringToFront();
    picInfo.Enabled = true;

    lblName.Text = $"Game ID: {gameId}";

    // TODO: Load full game details when merge functionality is implemented
    // This would query the Games table and populate all the tab controls
  }

  public void ShowNoGame()
  {
    if (WindowState == FormWindowState.Minimized)
    {
      WindowState = FormWindowState.Normal;
    }

    picNoInfo.BringToFront();
    picInfo.Enabled = false;
  }
}
