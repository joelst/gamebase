using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Form showing the list of games in the master database.
/// </summary>
public partial class MasterListForm : Form
{
  private SqliteConnection? _connection;
  private string? _databasePath;
  private readonly RegistryHelper _registry;
  private readonly MasterInfoForm _masterInfoForm;

  public MasterListForm(SqliteConnection connection, string databasePath, RegistryHelper registry, MasterInfoForm masterInfoForm)
  {
    InitializeComponent();
    _connection = connection;
    _databasePath = databasePath;
    _registry = registry;
    _masterInfoForm = masterInfoForm;
  }

  public void SetDatabase(SqliteConnection connection, string databasePath)
  {
    _connection = connection;
    _databasePath = databasePath;
  }

  private void MasterListForm_Load(object sender, EventArgs e)
  {
    RestoreWindowSettings();
  }

  private void MasterListForm_FormClosing(object sender, FormClosingEventArgs e)
  {
    SaveWindowSettings();
  }

  public void LoadList()
  {
    if (_connection == null || _databasePath == null) return;

    statusLabel.Text = "Loading Game List...";
    Cursor = Cursors.WaitCursor;
    Application.DoEvents();

    try
    {
      // Update caption with database name
      var fileName = Path.GetFileNameWithoutExtension(_databasePath);
      Text = $"MASTER [{fileName}]";

      // Clear existing items
      lvGames.BeginUpdate();
      lvGames.Items.Clear();

      // Load games from database
      using var command = _connection.CreateCommand();
      command.CommandText = "SELECT GA_Id, Name FROM Games ORDER BY Name";

      using var reader = command.ExecuteReader();
      while (reader.Read())
      {
        var id = reader.GetInt32(0);
        var name = reader.GetString(1);
        var item = new ListViewItem(name)
        {
          Tag = id
        };
        lvGames.Items.Add(item);
      }

      lvGames.EndUpdate();
      statusLabel.Text = $"{lvGames.Items.Count} Games";
      Show();
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Error loading game list: {ex.Message}", "BAC", MessageBoxButtons.OK, MessageBoxIcon.Error);
      statusLabel.Text = "Error loading games";
    }
    finally
    {
      Cursor = Cursors.Default;
    }
  }

  private void lvGames_SelectedIndexChanged(object sender, EventArgs e)
  {
    if (lvGames.SelectedItems.Count > 0)
    {
      var item = lvGames.SelectedItems[0];
      var gameId = (int)item.Tag!;
      _masterInfoForm.ShowGame(gameId, _connection!);
    }
  }

  public void SelectGame(int gameId)
  {
    // Find and select the game in the list
    foreach (ListViewItem item in lvGames.Items)
    {
      if ((int)item.Tag! == gameId)
      {
        item.Selected = true;
        item.EnsureVisible();
        _masterInfoForm.ShowGame(gameId, _connection!);
        return;
      }
    }

    // Game not found - clear selection
    SelectNoGame();
  }

  public void SelectNoGame()
  {
    lvGames.SelectedItems.Clear();
    _masterInfoForm.ShowNoGame();
  }

  private void RestoreWindowSettings()
  {
    var width = _registry.GetIntValue(BacPaths.RegistryWindow, "MasterListWidth", -1);
    var height = _registry.GetIntValue(BacPaths.RegistryWindow, "MasterListHeight", -1);
    var top = _registry.GetIntValue(BacPaths.RegistryWindow, "MasterListTop", -1);
    var left = _registry.GetIntValue(BacPaths.RegistryWindow, "MasterListLeft", -1);
    var maximized = _registry.GetBoolValue(BacPaths.RegistryWindow, "MasterListMaximized", false);
    var col1Width = _registry.GetIntValue(BacPaths.RegistryWindow, "MasterListCol1Width", 467);

    if (width < 266) width = 532;
    else if (width > Screen.PrimaryScreen!.Bounds.Width) width = Screen.PrimaryScreen.Bounds.Width;

    if (height < 66) height = 381;
    else if (height > Screen.PrimaryScreen.Bounds.Height) height = Screen.PrimaryScreen.Bounds.Height;

    if (top < 0 || top > Screen.PrimaryScreen.Bounds.Height ||
        left < 0 || left > Screen.PrimaryScreen.Bounds.Width)
    {
      top = 0;
      left = 0;
    }

    Location = new Point(left, top);
    Size = new Size(width, height);

    if (maximized)
    {
      WindowState = FormWindowState.Maximized;
    }

    columnName.Width = col1Width;
  }

  private void SaveWindowSettings()
  {
    if (WindowState == FormWindowState.Normal)
    {
      _registry.SetIntValue(BacPaths.RegistryWindow, "MasterListWidth", Width);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MasterListHeight", Height);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MasterListTop", Top);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MasterListLeft", Left);
    }
    _registry.SetBoolValue(BacPaths.RegistryWindow, "MasterListMaximized", WindowState == FormWindowState.Maximized);
    _registry.SetIntValue(BacPaths.RegistryWindow, "MasterListCol1Width", columnName.Width);
  }

  private void MasterListForm_Resize(object sender, EventArgs e)
  {
    if (ClientSize.Width > 0)
    {
      lvGames.Width = ClientSize.Width;
    }
    if (ClientSize.Height > statusStrip1.Height)
    {
      lvGames.Height = ClientSize.Height - statusStrip1.Height;
    }
  }
}
