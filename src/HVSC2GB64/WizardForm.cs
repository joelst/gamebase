// HVSC2GB64 Wizard Form - Main wizard interface

namespace HVSC2GB64;

public partial class WizardForm : Form
{
  private const string AppTitle = "GB64's HVSC Update Wizard";
  private const string Copyright = "(c) 2003-2026 James Burrows";
  private const double RequiredDbVersion = 2.5;

  private int _currentStep = 0;
  private Panel[] _steps = null!;
  private Button _btnNext = null!;
  private Button _btnBack = null!;
  private Button _btnAbort = null!;
  private ProgressBar _progressBar = null!;

  // Configuration data
  private string _hvscPath = string.Empty;
  private string _databasePath = string.Empty;
  private List<SidData> _allSidFiles = new();
  private DatabaseSynchronizer? _synchronizer;
  private HvscReader? _reader;

  public WizardForm()
  {
    InitializeComponent();
  }

  private void InitializeComponent()
  {
    Text = AppTitle;
    Size = new Size(800, 600);
    StartPosition = FormStartPosition.CenterScreen;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = true;

    // Create wizard steps
    _steps = new Panel[]
    {
      CreateWelcomeStep(),
      CreatePathSelectionStep(),
      CreateScanningStep(),
      CreateCompletionStep()
    };

    // Add all steps to form (only first will be visible)
    foreach (var step in _steps)
    {
      step.Visible = false;
      Controls.Add(step);
    }

    // Create navigation buttons
    CreateNavigationButtons();

    // Show first step
    ShowStep(0);
  }

  private Panel CreateWelcomeStep()
  {
    var panel = new Panel
    {
      Location = new Point(20, 20),
      Size = new Size(760, 480),
      BorderStyle = BorderStyle.None
    };

    var lblWelcome = new Label
    {
      Text = $"Welcome to {AppTitle}",
      Location = new Point(20, 20),
      Size = new Size(720, 40),
      Font = new Font("Segoe UI", 16, FontStyle.Bold)
    };
    panel.Controls.Add(lblWelcome);

    var lblCopyright = new Label
    {
      Text = Copyright,
      Location = new Point(20, 70),
      Size = new Size(720, 20),
      Font = new Font("Segoe UI", 9, FontStyle.Regular)
    };
    panel.Controls.Add(lblCopyright);

    var lblDescription = new Label
    {
      Text = "This wizard will help you synchronize your GameBase64 database with " +
             "the High Voltage SID Collection (HVSC).\n\n" +
             "The wizard will:\n" +
             "  • Scan your HVSC directory for SID files\n" +
             "  • Compare them with your GB64 database\n" +
             "  • Add new SID files to your database\n" +
             "  • Update existing entries with latest information\n\n" +
             "Click 'Next' to begin.",
      Location = new Point(20, 110),
      Size = new Size(720, 300),
      Font = new Font("Segoe UI", 10, FontStyle.Regular)
    };
    panel.Controls.Add(lblDescription);

    return panel;
  }

  private Panel CreatePathSelectionStep()
  {
    var panel = new Panel
    {
      Location = new Point(20, 20),
      Size = new Size(760, 480),
      BorderStyle = BorderStyle.None
    };

    var lblTitle = new Label
    {
      Text = "Select HVSC and Database Paths",
      Location = new Point(20, 20),
      Size = new Size(720, 30),
      Font = new Font("Segoe UI", 14, FontStyle.Bold)
    };
    panel.Controls.Add(lblTitle);

    // HVSC Path
    var lblHvsc = new Label
    {
      Text = "HVSC Root Directory:",
      Location = new Point(20, 70),
      Size = new Size(200, 20)
    };
    panel.Controls.Add(lblHvsc);

    var txtHvsc = new TextBox
    {
      Name = "txtHvscPath",
      Location = new Point(20, 95),
      Size = new Size(600, 25)
    };
    panel.Controls.Add(txtHvsc);

    var btnBrowseHvsc = new Button
    {
      Text = "Browse...",
      Location = new Point(630, 93),
      Size = new Size(100, 27)
    };
    btnBrowseHvsc.Click += (s, e) =>
    {
      using var dialog = new FolderBrowserDialog
      {
        Description = "Select HVSC root directory",
        ShowNewFolderButton = false
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtHvsc.Text = dialog.SelectedPath;
        _hvscPath = dialog.SelectedPath;
        ValidatePathStep();
      }
    };
    panel.Controls.Add(btnBrowseHvsc);

    // Database Path
    var lblDatabase = new Label
    {
      Text = "GB64 Database File:",
      Location = new Point(20, 140),
      Size = new Size(200, 20)
    };
    panel.Controls.Add(lblDatabase);

    var txtDatabase = new TextBox
    {
      Name = "txtDatabasePath",
      Location = new Point(20, 165),
      Size = new Size(600, 25)
    };
    panel.Controls.Add(txtDatabase);

    var btnBrowseDb = new Button
    {
      Text = "Browse...",
      Location = new Point(630, 163),
      Size = new Size(100, 27)
    };
    btnBrowseDb.Click += (s, e) =>
    {
      using var dialog = new OpenFileDialog
      {
        Filter = "Database Files (*.mdb;*.db)|*.mdb;*.db|All Files (*.*)|*.*",
        Title = "Select GB64 Database"
      };
      if (dialog.ShowDialog() == DialogResult.OK)
      {
        txtDatabase.Text = dialog.FileName;
        _databasePath = dialog.FileName;
        ValidatePathStep();
      }
    };
    panel.Controls.Add(btnBrowseDb);

    // Status label
    var lblStatus = new Label
    {
      Name = "lblPathStatus",
      Text = "Please select both paths to continue.",
      Location = new Point(20, 220),
      Size = new Size(720, 60),
      Font = new Font("Segoe UI", 9, FontStyle.Italic),
      ForeColor = Color.Gray
    };
    panel.Controls.Add(lblStatus);

    return panel;
  }

  private Panel CreateScanningStep()
  {
    var panel = new Panel
    {
      Location = new Point(20, 20),
      Size = new Size(760, 480),
      BorderStyle = BorderStyle.None
    };

    var lblTitle = new Label
    {
      Text = "Scanning and Synchronizing",
      Location = new Point(20, 20),
      Size = new Size(720, 30),
      Font = new Font("Segoe UI", 14, FontStyle.Bold)
    };
    panel.Controls.Add(lblTitle);

    var lblStatus = new Label
    {
      Name = "lblScanStatus",
      Text = "Preparing to scan HVSC directory...",
      Location = new Point(20, 70),
      Size = new Size(720, 40),
      Font = new Font("Segoe UI", 10, FontStyle.Regular)
    };
    panel.Controls.Add(lblStatus);

    var progressBar = new ProgressBar
    {
      Name = "progressScan",
      Location = new Point(20, 120),
      Size = new Size(720, 30),
      Style = ProgressBarStyle.Marquee
    };
    panel.Controls.Add(progressBar);

    var txtLog = new TextBox
    {
      Name = "txtScanLog",
      Location = new Point(20, 160),
      Size = new Size(720, 300),
      Multiline = true,
      ReadOnly = true,
      ScrollBars = ScrollBars.Vertical,
      Font = new Font("Consolas", 9, FontStyle.Regular)
    };
    panel.Controls.Add(txtLog);

    return panel;
  }

  private Panel CreateCompletionStep()
  {
    var panel = new Panel
    {
      Location = new Point(20, 20),
      Size = new Size(760, 480),
      BorderStyle = BorderStyle.None
    };

    var lblTitle = new Label
    {
      Text = "Synchronization Complete",
      Location = new Point(20, 20),
      Size = new Size(720, 30),
      Font = new Font("Segoe UI", 14, FontStyle.Bold),
      ForeColor = Color.Green
    };
    panel.Controls.Add(lblTitle);

    var lblSummary = new Label
    {
      Name = "lblCompletionSummary",
      Text = "Processing complete.",
      Location = new Point(20, 70),
      Size = new Size(720, 200),
      Font = new Font("Segoe UI", 10, FontStyle.Regular)
    };
    panel.Controls.Add(lblSummary);

    return panel;
  }

  private void CreateNavigationButtons()
  {
    // Separator line
    var separator = new Panel
    {
      Location = new Point(20, 510),
      Size = new Size(760, 2),
      BackColor = Color.LightGray
    };
    Controls.Add(separator);

    // Abort button
    _btnAbort = new Button
    {
      Text = "&Abort",
      Location = new Point(20, 525),
      Size = new Size(100, 30)
    };
    _btnAbort.Click += BtnAbort_Click;
    Controls.Add(_btnAbort);

    // Back button
    _btnBack = new Button
    {
      Text = "< &Back",
      Location = new Point(570, 525),
      Size = new Size(100, 30),
      Enabled = false
    };
    _btnBack.Click += BtnBack_Click;
    Controls.Add(_btnBack);

    // Next button
    _btnNext = new Button
    {
      Text = "&Next >",
      Location = new Point(680, 525),
      Size = new Size(100, 30)
    };
    _btnNext.Click += BtnNext_Click;
    Controls.Add(_btnNext);

    // Progress bar (initially hidden)
    _progressBar = new ProgressBar
    {
      Location = new Point(130, 530),
      Size = new Size(430, 20),
      Visible = false
    };
    Controls.Add(_progressBar);
  }

  private void ShowStep(int stepIndex)
  {
    // Hide all steps
    foreach (var step in _steps)
    {
      step.Visible = false;
    }

    // Show current step
    _currentStep = stepIndex;
    _steps[_currentStep].Visible = true;

    // Update button states
    _btnBack.Enabled = _currentStep > 0 && _currentStep < _steps.Length - 1;
    _btnNext.Enabled = _currentStep < _steps.Length - 1;

    // Update button text on last step
    if (_currentStep == _steps.Length - 1)
    {
      _btnNext.Text = "&Finish";
      _btnNext.Enabled = true;
    }
    else if (_currentStep == 0)
    {
      _btnNext.Text = "&Next >";
    }
    else if (_currentStep == 1)
    {
      _btnNext.Text = "&Start >";
      ValidatePathStep();
    }
    else
    {
      _btnNext.Text = "&Next >";
    }
  }

  private void ValidatePathStep()
  {
    if (_currentStep != 1) return;

    var txtHvsc = _steps[1].Controls.Find("txtHvscPath", false).FirstOrDefault() as TextBox;
    var txtDb = _steps[1].Controls.Find("txtDatabasePath", false).FirstOrDefault() as TextBox;
    var lblStatus = _steps[1].Controls.Find("lblPathStatus", false).FirstOrDefault() as Label;

    if (txtHvsc == null || txtDb == null || lblStatus == null) return;

    _hvscPath = txtHvsc.Text;
    _databasePath = txtDb.Text;

    bool hvscValid = Directory.Exists(_hvscPath);
    bool dbValid = File.Exists(_databasePath);

    if (hvscValid && dbValid)
    {
      lblStatus.Text = "✓ Paths validated. Click 'Start' to begin synchronization.";
      lblStatus.ForeColor = Color.Green;
      _btnNext.Enabled = true;
    }
    else
    {
      string error = "";
      if (!hvscValid) error += "HVSC directory not found. ";
      if (!dbValid) error += "Database file not found.";
      lblStatus.Text = error;
      lblStatus.ForeColor = Color.Red;
      _btnNext.Enabled = false;
    }
  }

  private async void BtnNext_Click(object? sender, EventArgs e)
  {
    if (_currentStep == _steps.Length - 1)
    {
      // Finish button clicked
      Close();
      return;
    }

    if (_currentStep == 1)
    {
      // Starting scan - move to scanning step and begin processing
      ShowStep(2);
      await PerformScanAndSync();
      ShowStep(3);
    }
    else
    {
      ShowStep(_currentStep + 1);
    }
  }

  private void BtnBack_Click(object? sender, EventArgs e)
  {
    if (_currentStep > 0)
    {
      ShowStep(_currentStep - 1);
    }
  }

  private void BtnAbort_Click(object? sender, EventArgs e)
  {
    if (MessageBox.Show("Are you sure you want to abort the wizard?", "Confirm Abort",
        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
    {
      Close();
    }
  }

  private async Task PerformScanAndSync()
  {
    var lblStatus = _steps[2].Controls.Find("lblScanStatus", false).FirstOrDefault() as Label;
    var txtLog = _steps[2].Controls.Find("txtScanLog", false).FirstOrDefault() as TextBox;
    var lblSummary = _steps[3].Controls.Find("lblCompletionSummary", false).FirstOrDefault() as Label;

    if (lblStatus == null || txtLog == null || lblSummary == null) return;

    void Log(string message)
    {
      if (txtLog.InvokeRequired)
      {
        txtLog.Invoke(() => txtLog.AppendText(message + Environment.NewLine));
      }
      else
      {
        txtLog.AppendText(message + Environment.NewLine);
      }
    }

    try
    {
      // Disable navigation during processing
      _btnNext.Enabled = false;
      _btnBack.Enabled = false;
      _btnAbort.Enabled = false;

      // Check database version
      lblStatus.Text = "Checking database version...";
      Log("Checking database version...");
      await Task.Run(() =>
      {
        _synchronizer = new DatabaseSynchronizer(_databasePath);
        if (!_synchronizer.CheckDatabaseVersion())
        {
          throw new Exception($"Database version is too old. Required: {RequiredDbVersion}");
        }
      });
      Log($"Database version OK (>= {RequiredDbVersion})");
      Log("");

      // Scan HVSC directory
      lblStatus.Text = "Scanning HVSC directory...";
      Log($"Scanning HVSC directory: {_hvscPath}");
      await Task.Run(() =>
      {
        _reader = new HvscReader(_hvscPath);
        _allSidFiles = _reader.ReadAllSidFiles();
      });
      Log($"Found {_allSidFiles.Count} SID files");
      Log("");

      // Synchronize with database
      lblStatus.Text = "Synchronizing with database...";
      Log("Synchronizing with GB64 database...");
      int added = 0;
      await Task.Run(() =>
      {
        if (_synchronizer != null)
        {
          added = _synchronizer.SynchronizeSidData(_allSidFiles);
        }
      });
      Log($"Added/Updated {added} entries");
      Log("");
      Log("Synchronization complete!");

      // Update completion summary
      lblSummary.Text = $"Synchronization completed successfully!\n\n" +
                       $"Statistics:\n" +
                       $"  • Total SID files found: {_allSidFiles.Count}\n" +
                       $"  • Entries added/updated: {added}\n\n" +
                       $"Your GB64 database has been updated with the latest HVSC data.";

      lblStatus.Text = "Synchronization complete!";
    }
    catch (Exception ex)
    {
      Log($"ERROR: {ex.Message}");
      lblStatus.Text = "Error occurred during synchronization";
      lblSummary.Text = $"An error occurred:\n\n{ex.Message}\n\n" +
                       "Please check the log for details.";
      MessageBox.Show($"Error: {ex.Message}", "Synchronization Error",
          MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
    finally
    {
      _btnNext.Enabled = true;
      _btnAbort.Enabled = true;
    }
  }
}
