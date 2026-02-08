using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Main MDI container form for BAC application.
/// </summary>
public partial class MainForm : Form
{
  private readonly RegistryHelper _registry;

  private SqliteConnection? _masterConnection;
  private SqliteConnection? _diffConnection;
  private string? _masterPath;
  private string? _diffPath;

  private MasterListForm? _masterListForm;
  private DiffListForm? _diffListForm;
  private MasterInfoForm? _masterInfoForm;

  public MainForm()
  {
    InitializeComponent();
    _registry = new RegistryHelper();
  }

  private void MainForm_Load(object sender, EventArgs e)
  {
    // Restore window position and size
    RestoreWindowSettings();

    // Create child forms but don't show them yet
    _masterInfoForm = new MasterInfoForm
    {
      MdiParent = this
    };
    _masterInfoForm.Show();
    _masterInfoForm.ShowNoGame();
  }

  private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
  {
    // Save window settings
    SaveWindowSettings();

    // Close databases
    CloseMasterDatabase();
    CloseDiffDatabase();

    // Clean up child forms
    _masterListForm?.Close();
    _diffListForm?.Close();
    _masterInfoForm?.Close();
  }

  private void RestoreWindowSettings()
  {
    var width = _registry.GetIntValue(BacPaths.RegistryWindow, "MainWidth", -1);
    var height = _registry.GetIntValue(BacPaths.RegistryWindow, "MainHeight", -1);
    var top = _registry.GetIntValue(BacPaths.RegistryWindow, "MainTop", -1);
    var left = _registry.GetIntValue(BacPaths.RegistryWindow, "MainLeft", -1);
    var maximized = _registry.GetBoolValue(BacPaths.RegistryWindow, "MainMaximized", false);

    // Validate and set size
    if (width < 500) width = 635;
    else if (width > Screen.PrimaryScreen!.Bounds.Width) width = Screen.PrimaryScreen.Bounds.Width;

    if (height < 500) height = 563;
    else if (height > Screen.PrimaryScreen.Bounds.Height) height = Screen.PrimaryScreen.Bounds.Height;

    // Validate and set position
    if (top < 0 || top > Screen.PrimaryScreen.Bounds.Height ||
        left < 0 || left > Screen.PrimaryScreen.Bounds.Width)
    {
      StartPosition = FormStartPosition.CenterScreen;
    }
    else
    {
      StartPosition = FormStartPosition.Manual;
      Location = new Point(left, top);
    }

    Size = new Size(width, height);

    if (maximized)
    {
      WindowState = FormWindowState.Maximized;
    }
  }

  private void SaveWindowSettings()
  {
    if (WindowState == FormWindowState.Normal)
    {
      _registry.SetIntValue(BacPaths.RegistryWindow, "MainWidth", Width);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MainHeight", Height);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MainTop", Top);
      _registry.SetIntValue(BacPaths.RegistryWindow, "MainLeft", Left);
    }
    _registry.SetBoolValue(BacPaths.RegistryWindow, "MainMaximized", WindowState == FormWindowState.Maximized);
  }

  private void btnMaster_Click(object sender, EventArgs e)
  {
    OpenMasterDatabase();
  }

  private void btnDiff_Click(object sender, EventArgs e)
  {
    OpenDiffDatabase();
  }

  private void OpenMasterDatabase()
  {
    var lastDir = _registry.GetStringValue(BacPaths.RegistryLastDirs, "Master", Application.StartupPath);
    if (!Directory.Exists(lastDir))
    {
      lastDir = Application.StartupPath;
    }

    using var dialog = new OpenFileDialog
    {
      Title = "Select the MASTER Database",
      Filter = "SQLite Database|*.db",
      InitialDirectory = lastDir,
      CheckFileExists = true
    };

    if (dialog.ShowDialog() == DialogResult.OK)
    {
      var filePath = dialog.FileName;

      // Validate database version
      if (!BacDatabase.ValidateDatabaseVersion(filePath, out var actualVersion))
      {
        MessageBox.Show(
            $"Incompatible database version (v{actualVersion:F1}). Expected v{BacDatabase.RequiredVersion:F1}.",
            "BAC",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        return;
      }

      // Close existing database
      CloseMasterDatabase();

      // Open new database
      try
      {
        _masterConnection = BacDatabase.OpenDatabase(filePath);
        _masterPath = filePath;

        // Save last directory
        var directory = Path.GetDirectoryName(filePath);
        if (directory != null)
        {
          _registry.SetStringValue(BacPaths.RegistryLastDirs, "Master", directory);
        }

        // Show master list form
        if (_masterListForm == null || _masterListForm.IsDisposed)
        {
          _masterListForm = new MasterListForm(_masterConnection, _masterPath, _registry, _masterInfoForm!)
          {
            MdiParent = this
          };
        }
        else
        {
          _masterListForm.SetDatabase(_masterConnection, _masterPath);
        }

        _masterListForm.LoadList();
        _masterListForm.Show();
      }
      catch (Exception ex)
      {
        MessageBox.Show($"Error opening database: {ex.Message}", "BAC", MessageBoxButtons.OK, MessageBoxIcon.Error);
        CloseMasterDatabase();
      }
    }
  }

  private void OpenDiffDatabase()
  {
    var lastDir = _registry.GetStringValue(BacPaths.RegistryLastDirs, "Diff", Application.StartupPath);
    if (!Directory.Exists(lastDir))
    {
      lastDir = Application.StartupPath;
    }

    using var dialog = new OpenFileDialog
    {
      Title = "Select the DIFF Database",
      Filter = "SQLite Database|*.db",
      InitialDirectory = lastDir,
      CheckFileExists = true
    };

    if (dialog.ShowDialog() == DialogResult.OK)
    {
      var filePath = dialog.FileName;

      // Validate database version
      if (!BacDatabase.ValidateDatabaseVersion(filePath, out var actualVersion))
      {
        MessageBox.Show(
            $"Incompatible database version (v{actualVersion:F1}). Expected v{BacDatabase.RequiredVersion:F1}.",
            "BAC",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        return;
      }

      // Close existing database
      CloseDiffDatabase();

      // Open new database
      try
      {
        _diffConnection = BacDatabase.OpenDatabase(filePath);
        _diffPath = filePath;

        // Save last directory
        var directory = Path.GetDirectoryName(filePath);
        if (directory != null)
        {
          _registry.SetStringValue(BacPaths.RegistryLastDirs, "Diff", directory);
        }

        // Show diff list form
        if (_diffListForm == null || _diffListForm.IsDisposed)
        {
          _diffListForm = new DiffListForm(_diffConnection, _diffPath, _registry, _masterListForm, _masterInfoForm!)
          {
            MdiParent = this
          };
        }
        else
        {
          _diffListForm.SetDatabase(_diffConnection, _diffPath);
        }

        _diffListForm.LoadList();
        _diffListForm.Show();
      }
      catch (Exception ex)
      {
        MessageBox.Show($"Error opening database: {ex.Message}", "BAC", MessageBoxButtons.OK, MessageBoxIcon.Error);
        CloseDiffDatabase();
      }
    }
  }

  private void CloseMasterDatabase()
  {
    if (_masterConnection != null)
    {
      _masterConnection.Close();
      _masterConnection.Dispose();
      _masterConnection = null;
      _masterPath = null;
    }
  }

  private void CloseDiffDatabase()
  {
    if (_diffConnection != null)
    {
      _diffConnection.Close();
      _diffConnection.Dispose();
      _diffConnection = null;
      _diffPath = null;
    }
  }

  public bool IsMasterOpen => _masterConnection != null;
  public bool IsDiffOpen => _diffConnection != null;
}
