using Gemus;

namespace GameBase
{
  public partial class OptionsForm : Form
  {
    public OptionsForm()
    {
      InitializeComponent();
    }

    private void OptionsForm_Load(object sender, EventArgs e)
    {
      string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string basePathRoot = string.IsNullOrWhiteSpace(appData)
          ? AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
          : appData;
      string defaultBase = Path.Combine(basePathRoot, "GameBase");
      txtBasePath.Text = defaultBase;
      txtGbName.Text = "GameBase";
      txtDbPath.Text = Path.Combine(defaultBase, "Data", "GameBase.db");
      cmbBackend.DataSource = Enum.GetValues(typeof(ArchiveBackend));

      TryLoadExisting();
    }

    private void TryLoadExisting()
    {
      try
      {
        string gemusIni = GetGemusIniPath();
        if (File.Exists(gemusIni))
        {
          string backendValue = GemusIniHelper.GetIniMessage("Settings", "ArchiveBackend", string.Empty, gemusIni);
          if (ArchiveUnpacker.TryParseBackend(backendValue, out var backend))
            cmbBackend.SelectedItem = backend;

          string dbPath = GemusIniHelper.GetIniMessage("Settings", "DatabasePath", string.Empty, gemusIni);
          if (!string.IsNullOrWhiteSpace(dbPath))
            txtDbPath.Text = dbPath;
        }
      }
      catch
      {
        // ignore load errors
      }
    }

    private string GetGemusIniPath()
    {
      string basePath = txtBasePath.Text;
      string candidate = Path.Combine(basePath, "Gemus.ini");
      if (File.Exists(candidate) || Directory.Exists(basePath))
        return candidate;
      return Path.Combine(basePath, txtGbName.Text, "Gemus.ini");
    }

    private void btnBrowse_Click(object sender, EventArgs e)
    {
      using var folder = new FolderBrowserDialog();
      folder.Description = "Select GameBase root folder";
      if (folder.ShowDialog(this) == DialogResult.OK)
      {
        txtBasePath.Text = folder.SelectedPath;
      }
    }

    private void btnSave_Click(object sender, EventArgs e)
    {
      try
      {
        string gemusIni = GetGemusIniPath();
        Directory.CreateDirectory(Path.GetDirectoryName(gemusIni)!);
        var backend = ArchiveUnpacker.GetDefaultBackend();
        if (cmbBackend.SelectedItem is ArchiveBackend selected)
          backend = selected;
        GemusIniHelper.WriteIniValue("Settings", "ArchiveBackend", backend.ToString(), gemusIni);

        var dbPath = txtDbPath.Text.Trim();
        if (!string.IsNullOrEmpty(dbPath))
        {
          Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
          GemusIniHelper.WriteIniValue("Settings", "DatabasePath", dbPath, gemusIni);
        }

        DialogResult = DialogResult.OK;
        MessageBox.Show(this, $"Saved settings to {gemusIni}", "Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
        Close();
      }
      catch (Exception ex)
      {
        MessageBox.Show(this, "Failed to save options: " + ex.Message, "Options", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void btnCancel_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.Cancel;
      Close();
    }

    private void btnBrowseDb_Click(object sender, EventArgs e)
    {
      using var dlg = new SaveFileDialog();
      dlg.Filter = "SQLite Database (*.db)|*.db|All files (*.*)|*.*";
      dlg.FileName = txtDbPath.Text;
      if (dlg.ShowDialog(this) == DialogResult.OK)
      {
        txtDbPath.Text = dlg.FileName;
      }
    }
  }
}
