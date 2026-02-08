// GameBase - Import Access Database Form

using GBMigrator;

namespace GameBase;

/// <summary>
/// Form for importing Access databases into SQLite format.
/// </summary>
public partial class ImportDatabaseForm : Form
{
  private string _accessPath = string.Empty;
  private string _sqlitePath = string.Empty;
  private bool _isImporting = false;

  public string SqliteDatabasePath => _sqlitePath;

  public ImportDatabaseForm()
  {
    InitializeComponent();
  }

  private void InitializeComponent()
  {
    Text = "Import Access Database";
    Size = new Size(700, 500);
    StartPosition = FormStartPosition.CenterParent;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    FormClosing += ImportDatabaseForm_FormClosing;

    // Title
    var lblTitle = new Label
    {
      Text = "Import GameBase Access Database",
      Location = new Point(20, 20),
      Size = new Size(660, 30),
      Font = new Font("Segoe UI", 14, FontStyle.Bold)
    };
    Controls.Add(lblTitle);

    var lblDescription = new Label
    {
      Text = "Convert your existing Access (.mdb) database to SQLite format.\n" +
             "This will create a new SQLite database file that can be used with GameBase.",
      Location = new Point(20, 60),
      Size = new Size(660, 40),
      Font = new Font("Segoe UI", 9, FontStyle.Regular)
    };
    Controls.Add(lblDescription);

    // Access database path
    var lblAccess = new Label
    {
      Text = "Access Database File (.mdb):",
      Location = new Point(20, 120),
      Size = new Size(200, 20)
    };
    Controls.Add(lblAccess);

    var txtAccess = new TextBox
    {
      Name = "txtAccessPath",
      Location = new Point(20, 145),
      Size = new Size(560, 25),
      ReadOnly = true
    };
    Controls.Add(txtAccess);

    var btnBrowseAccess = new Button
    {
      Text = "Browse...",
      Location = new Point(590, 143),
      Size = new Size(90, 27)
    };
    btnBrowseAccess.Click += (s, e) =>
    {
      using var dialog = new OpenFileDialog
      {
        Filter = "Access Database Files (*.mdb)|*.mdb|All Files (*.*)|*.*",
        Title = "Select Access Database"
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtAccess.Text = dialog.FileName;
        _accessPath = dialog.FileName;
        
        // Auto-suggest SQLite path
        string suggestedPath = Path.ChangeExtension(dialog.FileName, ".db");
        var txtSqlite = Controls.Find("txtSqlitePath", false).FirstOrDefault() as TextBox;
        if (txtSqlite != null && string.IsNullOrEmpty(txtSqlite.Text))
        {
          txtSqlite.Text = suggestedPath;
          _sqlitePath = suggestedPath;
        }
        
        ValidatePaths();
      }
    };
    Controls.Add(btnBrowseAccess);

    // SQLite database path
    var lblSqlite = new Label
    {
      Text = "Target SQLite Database File (.db):",
      Location = new Point(20, 190),
      Size = new Size(250, 20)
    };
    Controls.Add(lblSqlite);

    var txtSqlite = new TextBox
    {
      Name = "txtSqlitePath",
      Location = new Point(20, 215),
      Size = new Size(560, 25)
    };
    txtSqlite.TextChanged += (s, e) =>
    {
      _sqlitePath = txtSqlite.Text;
      ValidatePaths();
    };
    Controls.Add(txtSqlite);

    var btnBrowseSqlite = new Button
    {
      Text = "Browse...",
      Location = new Point(590, 213),
      Size = new Size(90, 27)
    };
    btnBrowseSqlite.Click += (s, e) =>
    {
      using var dialog = new SaveFileDialog
      {
        Filter = "SQLite Database Files (*.db)|*.db|All Files (*.*)|*.*",
        Title = "Save SQLite Database As",
        DefaultExt = "db"
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtSqlite.Text = dialog.FileName;
        _sqlitePath = dialog.FileName;
        ValidatePaths();
      }
    };
    Controls.Add(btnBrowseSqlite);

    // Status label
    var lblStatus = new Label
    {
      Name = "lblStatus",
      Text = "Please select the Access database file to import.",
      Location = new Point(20, 260),
      Size = new Size(660, 20),
      Font = new Font("Segoe UI", 9, FontStyle.Italic),
      ForeColor = Color.Gray
    };
    Controls.Add(lblStatus);

    // Progress section
    var lblProgress = new Label
    {
      Name = "lblProgress",
      Text = "",
      Location = new Point(20, 290),
      Size = new Size(660, 20),
      Visible = false
    };
    Controls.Add(lblProgress);

    var progressBar = new ProgressBar
    {
      Name = "progressBar",
      Location = new Point(20, 315),
      Size = new Size(660, 25),
      Style = ProgressBarStyle.Marquee,
      Visible = false
    };
    Controls.Add(progressBar);

    // Log output
    var txtLog = new TextBox
    {
      Name = "txtLog",
      Location = new Point(20, 350),
      Size = new Size(660, 60),
      Multiline = true,
      ReadOnly = true,
      ScrollBars = ScrollBars.Vertical,
      Font = new Font("Consolas", 8),
      Visible = false
    };
    Controls.Add(txtLog);

    // Buttons
    var btnCancel = new Button
    {
      Name = "btnCancel",
      Text = "Cancel",
      Location = new Point(500, 425),
      Size = new Size(85, 30),
      DialogResult = DialogResult.Cancel
    };
    Controls.Add(btnCancel);

    var btnImport = new Button
    {
      Name = "btnImport",
      Text = "Import",
      Location = new Point(595, 425),
      Size = new Size(85, 30),
      Enabled = false
    };
    btnImport.Click += BtnImport_Click;
    Controls.Add(btnImport);

    CancelButton = btnCancel;
  }

  private void ValidatePaths()
  {
    var lblStatus = Controls.Find("lblStatus", false).FirstOrDefault() as Label;
    var btnImport = Controls.Find("btnImport", false).FirstOrDefault() as Button;

    if (lblStatus == null || btnImport == null) return;

    bool accessValid = !string.IsNullOrEmpty(_accessPath) && File.Exists(_accessPath);
    bool sqliteValid = !string.IsNullOrEmpty(_sqlitePath);

    if (accessValid && sqliteValid)
    {
      if (File.Exists(_sqlitePath))
      {
        lblStatus.Text = "⚠ Target file exists and will be overwritten.";
        lblStatus.ForeColor = Color.DarkOrange;
      }
      else
      {
        lblStatus.Text = "✓ Ready to import. Click 'Import' to begin.";
        lblStatus.ForeColor = Color.Green;
      }
      btnImport.Enabled = true;
    }
    else
    {
      string error = "";
      if (!accessValid) error += "Please select a valid Access database file. ";
      if (!sqliteValid) error += "Please specify a target SQLite database path.";
      lblStatus.Text = error;
      lblStatus.ForeColor = Color.Red;
      btnImport.Enabled = false;
    }
  }

  private async void BtnImport_Click(object? sender, EventArgs e)
  {
    if (_isImporting) return;

    // Confirm overwrite if file exists
    if (File.Exists(_sqlitePath))
    {
      var result = MessageBox.Show(
        $"The target file already exists:\n{_sqlitePath}\n\nDo you want to overwrite it?",
        "Confirm Overwrite",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning);

      if (result != DialogResult.Yes)
        return;

      try
      {
        File.Delete(_sqlitePath);
      }
      catch (Exception ex)
      {
        MessageBox.Show($"Cannot delete existing file: {ex.Message}", "Error",
          MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }
    }

    // Start import
    _isImporting = true;
    
    var btnImport = Controls.Find("btnImport", false).FirstOrDefault() as Button;
    var btnCancel = Controls.Find("btnCancel", false).FirstOrDefault() as Button;
    var lblProgress = Controls.Find("lblProgress", false).FirstOrDefault() as Label;
    var progressBar = Controls.Find("progressBar", false).FirstOrDefault() as ProgressBar;
    var txtLog = Controls.Find("txtLog", false).FirstOrDefault() as TextBox;
    var lblStatus = Controls.Find("lblStatus", false).FirstOrDefault() as Label;

    if (btnImport != null) btnImport.Enabled = false;
    if (btnCancel != null) btnCancel.Enabled = false;
    if (lblProgress != null)
    {
      lblProgress.Text = "Importing database...";
      lblProgress.Visible = true;
    }
    if (progressBar != null) progressBar.Visible = true;
    if (txtLog != null) txtLog.Visible = true;
    if (lblStatus != null) lblStatus.Visible = false;

    void Log(string message)
    {
      if (txtLog != null)
      {
        if (txtLog.InvokeRequired)
        {
          txtLog.Invoke(() =>
          {
            txtLog.AppendText(message + Environment.NewLine);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
          });
        }
        else
        {
          txtLog.AppendText(message + Environment.NewLine);
          txtLog.SelectionStart = txtLog.Text.Length;
          txtLog.ScrollToCaret();
        }
      }
    }

    try
    {
      Log($"Starting migration...");
      Log($"Source: {_accessPath}");
      Log($"Target: {_sqlitePath}");
      Log("");

      var progress = new Progress<string>(message => Log(message));
      var migrator = new AccessToSqliteMigrator(_accessPath, _sqlitePath, progress);

      await Task.Run(() => migrator.MigrateAsync());

      Log("");
      Log("✓ Import completed successfully!");

      if (lblProgress != null) lblProgress.Text = "Import complete!";
      if (progressBar != null) progressBar.Style = ProgressBarStyle.Continuous;

      MessageBox.Show(
        "Database imported successfully!\n\n" +
        "The SQLite database is ready to use.",
        "Import Complete",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information);

      DialogResult = DialogResult.OK;
    }
    catch (InvalidOperationException ex) when (ex.Message.Contains("provider") || ex.Message.Contains("Provider"))
    {
      Log("");
      Log($"ERROR: {ex.Message}");
      Log("");
      Log("Microsoft Access Database Engine not found.");
      Log("Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920");

      MessageBox.Show(
        "Microsoft Access Database Engine not found.\n\n" +
        "Please install the Microsoft Access Database Engine:\n" +
        "https://www.microsoft.com/en-us/download/details.aspx?id=54920\n\n" +
        "Install the version matching your system (x64 or x86).",
        "Import Failed",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);

      if (btnImport != null) btnImport.Enabled = true;
      if (btnCancel != null) btnCancel.Enabled = true;
    }
    catch (Exception ex)
    {
      Log("");
      Log($"ERROR: {ex.Message}");

      MessageBox.Show(
        $"Import failed: {ex.Message}\n\n" +
        "Please check the log for details.",
        "Import Failed",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);

      if (btnImport != null) btnImport.Enabled = true;
      if (btnCancel != null) btnCancel.Enabled = true;
    }
    finally
    {
      _isImporting = false;
    }
  }

  private void ImportDatabaseForm_FormClosing(object? sender, FormClosingEventArgs e)
  {
    if (_isImporting)
    {
      var result = MessageBox.Show(
        "Import is in progress. Are you sure you want to cancel?",
        "Confirm Cancel",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning);

      if (result != DialogResult.Yes)
      {
        e.Cancel = true;
      }
    }
  }
}
