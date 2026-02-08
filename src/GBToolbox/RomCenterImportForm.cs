//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: RomCenter 2.00 import dialog
//==============================================

namespace GBToolbox;

internal sealed class RomCenterImportForm : Form
{
  private readonly string _dataFilePath;
  private readonly string? _existingDbPath;
  private readonly TextBox _txtGameBaseName;
  private readonly ComboBox _comboNameField;
  private readonly ComboBox _comboFileField;
  private readonly CheckBox _chkIgnoreDuplicates;
  private readonly CheckBox _chkTosec;
  private readonly CheckBox _chkAddZip;
  private readonly CheckBox _chkScan;
  private readonly CheckBox _chkMatch;
  private readonly TextBox _txtRomPath;
  private readonly Button _btnStart;

  public RomCenterImportForm(string dataFilePath, string? existingDbPath)
  {
    _dataFilePath = dataFilePath;
    _existingDbPath = existingDbPath;

    Text = "Import RomCenter 2.00 Data";
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(560, 420);

    var lblDataFile = new Label
    {
      Text = "Datafile:",
      AutoSize = false,
      Left = 12,
      Top = 16,
      Width = 90
    };

    var txtDataFile = new TextBox
    {
      Left = 110,
      Top = 12,
      Width = 430,
      ReadOnly = true,
      Text = _dataFilePath
    };

    var lblGameBase = new Label
    {
      Text = "GameBase name:",
      AutoSize = false,
      Left = 12,
      Top = 48,
      Width = 90
    };

    _txtGameBaseName = new TextBox
    {
      Left = 110,
      Top = 44,
      Width = 250
    };

    var lblNameField = new Label
    {
      Text = "Name field:",
      AutoSize = false,
      Left = 12,
      Top = 80,
      Width = 90
    };

    _comboNameField = new ComboBox
    {
      Left = 110,
      Top = 76,
      Width = 250,
      DropDownStyle = ComboBoxStyle.DropDownList
    };

    var lblFileField = new Label
    {
      Text = "File field:",
      AutoSize = false,
      Left = 12,
      Top = 112,
      Width = 90
    };

    _comboFileField = new ComboBox
    {
      Left = 110,
      Top = 108,
      Width = 250,
      DropDownStyle = ComboBoxStyle.DropDownList
    };

    _chkIgnoreDuplicates = new CheckBox
    {
      Left = 380,
      Top = 46,
      Width = 160,
      Text = "Ignore duplicates"
    };

    _chkTosec = new CheckBox
    {
      Left = 380,
      Top = 76,
      Width = 160,
      Text = "Use TOSEC"
    };

    _chkAddZip = new CheckBox
    {
      Left = 380,
      Top = 106,
      Width = 160,
      Text = "Add .zip"
    };

    _chkScan = new CheckBox
    {
      Left = 380,
      Top = 136,
      Width = 160,
      Text = "Scan for ROMs"
    };

    _chkMatch = new CheckBox
    {
      Left = 380,
      Top = 166,
      Width = 160,
      Text = "Match extensions",
      Checked = true
    };

    var lblRomPath = new Label
    {
      Text = "ROM path:",
      AutoSize = false,
      Left = 12,
      Top = 156,
      Width = 90
    };

    _txtRomPath = new TextBox
    {
      Left = 110,
      Top = 152,
      Width = 350
    };

    var btnBrowse = new Button
    {
      Text = "Browse...",
      Left = 470,
      Top = 150,
      Width = 70
    };

    var lblHelp = new Label
    {
      Text = "Select which fields contain the game name and filename.\nScanning will calculate CRC-32 values in your ROM folder.",
      AutoSize = false,
      Left = 12,
      Top = 200,
      Width = 528,
      Height = 50
    };

    _btnStart = new Button
    {
      Text = "Start",
      Left = 380,
      Top = 356,
      Width = 75
    };

    var btnCancel = new Button
    {
      Text = "Cancel",
      Left = 465,
      Top = 356,
      Width = 75,
      DialogResult = DialogResult.Cancel
    };

    Controls.Add(lblDataFile);
    Controls.Add(txtDataFile);
    Controls.Add(lblGameBase);
    Controls.Add(_txtGameBaseName);
    Controls.Add(lblNameField);
    Controls.Add(_comboNameField);
    Controls.Add(lblFileField);
    Controls.Add(_comboFileField);
    Controls.Add(_chkIgnoreDuplicates);
    Controls.Add(_chkTosec);
    Controls.Add(_chkAddZip);
    Controls.Add(_chkScan);
    Controls.Add(_chkMatch);
    Controls.Add(lblRomPath);
    Controls.Add(_txtRomPath);
    Controls.Add(btnBrowse);
    Controls.Add(lblHelp);
    Controls.Add(_btnStart);
    Controls.Add(btnCancel);

    AcceptButton = _btnStart;
    CancelButton = btnCancel;

    btnBrowse.Click += BrowseRomPath;
    _btnStart.Click += StartImport;
    _chkScan.CheckedChanged += (_, _) => _txtRomPath.Enabled = _chkScan.Checked;

    Load += RomCenterImportForm_Load;
  }

  private void RomCenterImportForm_Load(object? sender, EventArgs e)
  {
    _comboNameField.Items.AddRange(new object[]
    {
      new FieldOption(1, "1 - Full Name"),
      new FieldOption(2, "2 - Short Name"),
      new FieldOption(3, "3 - Alternate Name"),
      new FieldOption(4, "4 - Name / Group"),
      new FieldOption(5, "5 - Name / Comment"),
      new FieldOption(8, "8 - Name / Group / Comment")
    });

    _comboFileField.Items.AddRange(new object[]
    {
      new FieldOption(1, "1 - Full Name"),
      new FieldOption(3, "3 - File Name"),
      new FieldOption(5, "5 - File / Comment"),
      new FieldOption(7, "7 - File / Group / Comment")
    });

    _comboNameField.SelectedIndex = 0;
    _comboFileField.SelectedIndex = 1;
    _chkScan.Checked = false;
    _txtRomPath.Enabled = false;

    if (!string.IsNullOrWhiteSpace(_existingDbPath))
    {
      _txtGameBaseName.Text = Path.GetFileNameWithoutExtension(_existingDbPath);
      _txtGameBaseName.Enabled = false;
      _btnStart.Text = "Import";
    }
    else
    {
      _txtGameBaseName.Text = Path.GetFileNameWithoutExtension(_dataFilePath);
    }
  }

  private void BrowseRomPath(object? sender, EventArgs e)
  {
    using var dialog = new FolderBrowserDialog
    {
      Description = "Select ROM folder"
    };

    if (dialog.ShowDialog(this) == DialogResult.OK)
    {
      _txtRomPath.Text = dialog.SelectedPath;
    }
  }

  private async void StartImport(object? sender, EventArgs e)
  {
    if (string.IsNullOrWhiteSpace(_txtGameBaseName.Text))
    {
      MessageBox.Show("Please enter a GameBase name.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    if (_chkScan.Checked && string.IsNullOrWhiteSpace(_txtRomPath.Text))
    {
      MessageBox.Show("Please select a ROM path to scan.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    var nameField = (FieldOption)_comboNameField.SelectedItem!;
    var fileField = (FieldOption)_comboFileField.SelectedItem!;

    var options = new RomCenterImportOptions
    {
      DataFilePath = _dataFilePath,
      NameColumnIndex = nameField.Index,
      FileColumnIndex = fileField.Index,
      IgnoreDuplicates = _chkIgnoreDuplicates.Checked,
      UseTosec = _chkTosec.Checked,
      AddZipExtension = _chkAddZip.Checked,
      ScanForRoms = _chkScan.Checked,
      MatchExtensions = _chkMatch.Checked,
      RomPath = _txtRomPath.Text.Trim(),
      GameBaseName = _txtGameBaseName.Text.Trim(),
      ExistingDatabasePath = _existingDbPath
    };

    options.RomPathWithSlash = string.IsNullOrWhiteSpace(options.RomPath) ? string.Empty : EnsureTrailingSlash(options.RomPath);
    options.GameBaseFolder = !string.IsNullOrWhiteSpace(_existingDbPath)
      ? Path.GetDirectoryName(_existingDbPath) ?? string.Empty
      : Path.Combine(ToolboxPaths.GetGameBaseRoot(), options.GameBaseName);
    options.GameBaseFolderWithSlash = EnsureTrailingSlash(options.GameBaseFolder);
    options.DefaultTypes = "*.d64;*.t64;*.g64;*.tap";

    using var progress = new ProgressForm("Importing RomCenter 2.00 Data");
    using var cts = new CancellationTokenSource();
    progress.CancelRequested += () => cts.Cancel();
    progress.Show(this);
    progress.UpdateStatus("Initializing...");
    Application.DoEvents();

    var importer = new RomCenterImporter();
    RomCenterImportResult result;

    try
    {
      result = await importer.ImportAsync(options, progress, cts.Token);
    }
    catch (Exception ex)
    {
      result = RomCenterImportResult.Failed(ex.Message);
    }

    progress.Close();

    if (result.CancelledByUser)
    {
      MessageBox.Show("Import cancelled.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    if (!result.IsSuccess)
    {
      MessageBox.Show(result.Message, "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
      return;
    }

    MessageBox.Show("RomCenter import finished.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    DialogResult = DialogResult.OK;
    Close();
  }

  private static string EnsureTrailingSlash(string value)
  {
    if (string.IsNullOrWhiteSpace(value))
    {
      return string.Empty;
    }

    if (value.EndsWith(Path.DirectorySeparatorChar) || value.EndsWith(Path.AltDirectorySeparatorChar))
    {
      return value;
    }

    return value + Path.DirectorySeparatorChar;
  }

  private sealed record FieldOption(int Index, string Display)
  {
    public override string ToString() => Display;
  }
}
