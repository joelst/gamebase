// GameBase - Export to Access Database Form

using GBMigrator;

namespace GameBase;

/// <summary>
/// Form for exporting SQLite databases to Access format for backward compatibility.
/// </summary>
public partial class ExportDatabaseForm : Form
{
  private string _sqlitePath = string.Empty;
  private string _accessPath = string.Empty;
  private bool _isExporting = false;

  public ExportDatabaseForm(string currentDatabasePath = "")
  {
    _sqlitePath = currentDatabasePath;
    InitializeComponent();
  }

  private void InitializeComponent()
  {
    Text = "Export to Access Database";
    Size = new Size(700, 500);
    StartPosition = FormStartPosition.CenterParent;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    FormClosing += ExportDatabaseForm_FormClosing;

    // Title
    var lblTitle = new Label
    {
      Text = "Export to Access Database (.mdb)",
      Location = new Point(20, 20),
      Size = new Size(660, 30),
      Font = new Font("Segoe UI", 14, FontStyle.Bold)
    };
    Controls.Add(lblTitle);

    var lblDescription = new Label
    {
      Text = "Export your SQLite database to Access format for backward compatibility.\n" +
             "This allows you to use the database with older VB6 GameBase versions.",
      Location = new Point(20, 60),
      Size = new Size(660, 40),
      Font = new Font("Segoe UI", 9, FontStyle.Regular)
    };
    Controls.Add(lblDescription);

    // SQLite database path
    var lblSqlite = new Label
    {
      Text = "SQLite Database File (.db):",
      Location = new Point(20, 120),
      Size = new Size(200, 20)
    };
    Controls.Add(lblSqlite);

    var txtSqlite = new TextBox
    {
      Name = "txtSqlitePath",
      Location = new Point(20, 145),
      Size = new Size(560, 25),
      Text = _sqlitePath
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
      Location = new Point(590, 143),
      Size = new Size(90, 27)
    };
    btnBrowseSqlite.Click += (s, e) =>
    {
      using var dialog = new OpenFileDialog
      {
        Filter = "SQLite Database Files (*.db)|*.db|All Files (*.*)|*.*",
        Title = "Select SQLite Database"
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtSqlite.Text = dialog.FileName;
        _sqlitePath = dialog.FileName;

        // Auto-suggest Access path
        string suggestedPath = Path.ChangeExtension(dialog.FileName, ".mdb");
        var txtAccess = Controls.Find("txtAccessPath", false).FirstOrDefault() as TextBox;
        if (txtAccess != null && string.IsNullOrEmpty(txtAccess.Text))
        {
          txtAccess.Text = suggestedPath;
          _accessPath = suggestedPath;
        }

        ValidatePaths();
      }
    };
    Controls.Add(btnBrowseSqlite);

    // Access database path
    var lblAccess = new Label
    {
      Text = "Target Access Database File (.mdb):",
      Location = new Point(20, 190),
      Size = new Size(250, 20)
    };
    Controls.Add(lblAccess);

    var txtAccess = new TextBox
    {
      Name = "txtAccessPath",
      Location = new Point(20, 215),
      Size = new Size(560, 25)
    };
    txtAccess.TextChanged += (s, e) =>
    {
      _accessPath = txtAccess.Text;
      ValidatePaths();
    };
    Controls.Add(txtAccess);

    var btnBrowseAccess = new Button
    {
      Text = "Browse...",
      Location = new Point(590, 213),
      Size = new Size(90, 27)
    };
    btnBrowseAccess.Click += (s, e) =>
    {
      using var dialog = new SaveFileDialog
      {
        Filter = "Access Database Files (*.mdb)|*.mdb|All Files (*.*)|*.*",
        Title = "Save Access Database As",
        DefaultExt = "mdb"
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtAccess.Text = dialog.FileName;
        _accessPath = dialog.FileName;
        ValidatePaths();
      }
    };
    Controls.Add(btnBrowseAccess);

    // Status label
    var lblStatus = new Label
    {
      Name = "lblStatus",
      Text = "Please select the SQLite database file to export.",
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

    var btnExport = new Button
    {
      Name = "btnExport",
      Text = "Export",
      Location = new Point(595, 425),
      Size = new Size(85, 30),
      Enabled = false
    };
    btnExport.Click += BtnExport_Click;
    Controls.Add(btnExport);

    CancelButton = btnCancel;

    // Initial validation
    ValidatePaths();
  }

  private void ValidatePaths()
  {
    var lblStatus = Controls.Find("lblStatus", false).FirstOrDefault() as Label;
    var btnExport = Controls.Find("btnExport", false).FirstOrDefault() as Button;

    if (lblStatus == null || btnExport == null) return;

    bool sqliteValid = !string.IsNullOrEmpty(_sqlitePath) && File.Exists(_sqlitePath);
    bool accessValid = !string.IsNullOrEmpty(_accessPath);

    if (sqliteValid && accessValid)
    {
      if (File.Exists(_accessPath))
      {
        lblStatus.Text = "⚠ Target file exists and will be overwritten.";
        lblStatus.ForeColor = Color.DarkOrange;
      }
      else
      {
        lblStatus.Text = "✓ Ready to export. Click 'Export' to begin.";
        lblStatus.ForeColor = Color.Green;
      }
      btnExport.Enabled = true;
    }
    else
    {
      string error = "";
      if (!sqliteValid) error += "Please select a valid SQLite database file. ";
      if (!accessValid) error += "Please specify a target Access database path.";
      lblStatus.Text = error;
      lblStatus.ForeColor = Color.Red;
      btnExport.Enabled = false;
    }
  }

  private async void BtnExport_Click(object? sender, EventArgs e)
  {
    if (_isExporting) return;

    // Confirm overwrite if file exists
    if (File.Exists(_accessPath))
    {
      var result = MessageBox.Show(
        $"The target file already exists:\n{_accessPath}\n\nDo you want to overwrite it?",
        "Confirm Overwrite",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning);

      if (result != DialogResult.Yes)
        return;
    }

    // Start export
    _isExporting = true;

    var btnExport = Controls.Find("btnExport", false).FirstOrDefault() as Button;
    var btnCancel = Controls.Find("btnCancel", false).FirstOrDefault() as Button;
    var lblProgress = Controls.Find("lblProgress", false).FirstOrDefault() as Label;
    var progressBar = Controls.Find("progressBar", false).FirstOrDefault() as ProgressBar;
    var txtLog = Controls.Find("txtLog", false).FirstOrDefault() as TextBox;
    var lblStatus = Controls.Find("lblStatus", false).FirstOrDefault() as Label;

    if (btnExport != null) btnExport.Enabled = false;
    if (btnCancel != null) btnCancel.Enabled = false;
    if (lblProgress != null)
    {
      lblProgress.Text = "Exporting database...";
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
      Log($"Starting export...");
      Log($"Source: {_sqlitePath}");
      Log($"Target: {_accessPath}");
      Log("");

      var progress = new Progress<string>(message => Log(message));
      var exporter = new SqliteToAccessExporter(_sqlitePath, _accessPath, progress);

      await Task.Run(() => exporter.ExportAsync());

      Log("");
      Log("✓ Export completed successfully!");

      if (lblProgress != null) lblProgress.Text = "Export complete!";
      if (progressBar != null) progressBar.Style = ProgressBarStyle.Continuous;

      MessageBox.Show(
        "Database exported successfully!\n\n" +
        "The Access database is compatible with VB6 GameBase versions.",
        "Export Complete",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information);

      DialogResult = DialogResult.OK;
    }
    catch (InvalidOperationException ex) when (ex.Message.Contains("provider") || ex.Message.Contains("Provider") || ex.Message.Contains("ACE") || ex.Message.Contains("ADOX"))
    {
      Log("");
      Log($"ERROR: {ex.Message}");
      Log("");
      Log("Microsoft Access Database Engine (ACE) or ADOX not found.");
      Log("Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920");

      MessageBox.Show(
        "Microsoft Access Database Engine (ACE) or ADOX not found.\n\n" +
        "Please install the Microsoft Access Database Engine:\n" +
        "https://www.microsoft.com/en-us/download/details.aspx?id=54920\n\n" +
        "Install the version matching your system (x64 or x86).",
        "Export Failed",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);

      if (btnExport != null) btnExport.Enabled = true;
      if (btnCancel != null) btnCancel.Enabled = true;
    }
    catch (Exception ex)
    {
      Log("");
      Log($"ERROR: {ex.Message}");

      MessageBox.Show(
        $"Export failed: {ex.Message}\n\n" +
        "Please check the log for details.",
        "Export Failed",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);

      if (btnExport != null) btnExport.Enabled = true;
      if (btnCancel != null) btnCancel.Enabled = true;
    }
    finally
    {
      _isExporting = false;
    }
  }

  private void ExportDatabaseForm_FormClosing(object? sender, FormClosingEventArgs e)
  {
    if (_isExporting)
    {
      var result = MessageBox.Show(
        "Export is in progress. Are you sure you want to cancel?",
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
