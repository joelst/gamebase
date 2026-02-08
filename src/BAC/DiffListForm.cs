using GBFuncs;
using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Form showing the list of games in the diff database.
/// </summary>
public partial class DiffListForm : Form
{
    private SqliteConnection? _connection;
    private string? _databasePath;
    private readonly RegistryHelper _registry;
    private readonly MasterListForm? _masterListForm;
    private readonly MasterInfoForm _masterInfoForm;
    private string _dbName = "";

    public DiffListForm(SqliteConnection connection, string databasePath, RegistryHelper registry, 
                        MasterListForm? masterListForm, MasterInfoForm masterInfoForm)
    {
        InitializeComponent();
        _connection = connection;
        _databasePath = databasePath;
        _registry = registry;
        _masterListForm = masterListForm;
        _masterInfoForm = masterInfoForm;
    }

    public void SetDatabase(SqliteConnection connection, string databasePath)
    {
        _connection = connection;
        _databasePath = databasePath;
    }

    private void DiffListForm_Load(object sender, EventArgs e)
    {
        RestoreWindowSettings();
    }

    private void DiffListForm_FormClosing(object sender, FormClosingEventArgs e)
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
            // Get diff info
            using (var command = _connection.CreateCommand())
            {
                command.CommandText = "SELECT GameBase, Name, Email, Date FROM Info LIMIT 1";
                using var reader = command.ExecuteReader();
                if (reader.Read())
                {
                    _dbName = reader.IsDBNull(0) ? "" : reader.GetString(0);
                    lblFrom.Text = reader.IsDBNull(1) ? "" : reader.GetString(1);
                    lblEmail.Text = reader.IsDBNull(2) ? "" : reader.GetString(2);
                    lblDate.Text = reader.IsDBNull(3) ? "" : reader.GetString(3);
                    Text = $"DIFF [{_dbName}]";
                }
            }

            // Clear existing items
            lvGames.BeginUpdate();
            lvGames.Items.Clear();

            // Load games from database
            using (var command = _connection.CreateCommand())
            {
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
        if (lvGames.SelectedItems.Count > 0 && _masterListForm != null)
        {
            var item = lvGames.SelectedItems[0];
            var gameId = (int)item.Tag!;

            // Restore master form if minimized
            if (_masterListForm.WindowState == FormWindowState.Minimized)
            {
                _masterListForm.WindowState = FormWindowState.Normal;
            }

            // Try to select the same game in master list
            _masterListForm.SelectGame(gameId);
        }
    }

    private void lblEmail_Click(object sender, EventArgs e)
    {
        if (!string.IsNullOrEmpty(lblEmail.Text) && !string.IsNullOrEmpty(_dbName) && !string.IsNullOrEmpty(lblDate.Text))
        {
            var subject = Uri.EscapeDataString($"Re: Your submitted DIFF [{_dbName}] dated [{lblDate.Text}]");
            var mailto = $"mailto:{lblEmail.Text}?subject={subject}";
            
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = mailto,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error launching email client: {ex.Message}", "BAC", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void RestoreWindowSettings()
    {
        var width = _registry.GetIntValue(BacPaths.RegistryWindow, "DiffListWidth", -1);
        var height = _registry.GetIntValue(BacPaths.RegistryWindow, "DiffListHeight", -1);
        var top = _registry.GetIntValue(BacPaths.RegistryWindow, "DiffListTop", -1);
        var left = _registry.GetIntValue(BacPaths.RegistryWindow, "DiffListLeft", -1);
        var maximized = _registry.GetBoolValue(BacPaths.RegistryWindow, "DiffListMaximized", false);
        var col1Width = _registry.GetIntValue(BacPaths.RegistryWindow, "DiffListCol1Width", 467);

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
            _registry.SetIntValue(BacPaths.RegistryWindow, "DiffListWidth", Width);
            _registry.SetIntValue(BacPaths.RegistryWindow, "DiffListHeight", Height);
            _registry.SetIntValue(BacPaths.RegistryWindow, "DiffListTop", Top);
            _registry.SetIntValue(BacPaths.RegistryWindow, "DiffListLeft", Left);
        }
        _registry.SetBoolValue(BacPaths.RegistryWindow, "DiffListMaximized", WindowState == FormWindowState.Maximized);
        _registry.SetIntValue(BacPaths.RegistryWindow, "DiffListCol1Width", columnName.Width);
    }

    private void DiffListForm_Resize(object sender, EventArgs e)
    {
        if (ClientSize.Width > 0)
        {
            lvGames.Width = ClientSize.Width;
        }
        if (ClientSize.Height > (panelInfo.Height + statusStrip1.Height))
        {
            lvGames.Height = ClientSize.Height - panelInfo.Height - statusStrip1.Height;
        }
    }
}
