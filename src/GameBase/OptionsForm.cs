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
      txtGamesPath.Text = Path.Combine(defaultBase, "Games");
      txtExtrasPath.Text = Path.Combine(defaultBase, "Extras");
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
          
          string gamesPath = GemusIniHelper.GetIniMessage("Settings", "GamesPath", string.Empty, gemusIni);
          if (!string.IsNullOrWhiteSpace(gamesPath))
            txtGamesPath.Text = gamesPath;
          
          string extrasPath = GemusIniHelper.GetIniMessage("Settings", "ExtrasPath", string.Empty, gemusIni);
          if (!string.IsNullOrWhiteSpace(extrasPath))
            txtExtrasPath.Text = extrasPath;
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
        
        var gamesPath = txtGamesPath.Text.Trim();
        if (!string.IsNullOrEmpty(gamesPath))
        {
          Directory.CreateDirectory(gamesPath);
          GemusIniHelper.WriteIniValue("Settings", "GamesPath", gamesPath, gemusIni);
        }
        
        var extrasPath = txtExtrasPath.Text.Trim();
        if (!string.IsNullOrEmpty(extrasPath))
        {
          Directory.CreateDirectory(extrasPath);
          GemusIniHelper.WriteIniValue("Settings", "ExtrasPath", extrasPath, gemusIni);
        }
        
        // Create default Emulators.ini if it doesn't exist
        CreateDefaultEmulatorsIni(gemusIni);

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
    
    private void btnBrowseGames_Click(object sender, EventArgs e)
    {
      using var folder = new FolderBrowserDialog();
      folder.Description = "Select Games folder";
      folder.SelectedPath = txtGamesPath.Text;
      if (folder.ShowDialog(this) == DialogResult.OK)
      {
        txtGamesPath.Text = folder.SelectedPath;
      }
    }
    
    private void btnBrowseExtras_Click(object sender, EventArgs e)
    {
      using var folder = new FolderBrowserDialog();
      folder.Description = "Select Extras folder";
      folder.SelectedPath = txtExtrasPath.Text;
      if (folder.ShowDialog(this) == DialogResult.OK)
      {
        txtExtrasPath.Text = folder.SelectedPath;
      }
    }
    
    private void CreateDefaultEmulatorsIni(string gemusIniPath)
    {
      try
      {
        string basePath = Path.GetDirectoryName(gemusIniPath)!;
        string emulatorsIni = Path.Combine(basePath, "GameBase", "Emulators.ini");
        
        if (File.Exists(emulatorsIni))
          return; // Already exists
          
        Directory.CreateDirectory(Path.GetDirectoryName(emulatorsIni)!);
        
        // Create a minimal default Emulators.ini
        var content = @"[Emulator0]
Name=Default Emulator
Path=
Params=
Directory=
FileTypes=d64;t64;prg;zip

[General]
NumberOfEmulators=1
";
        
        File.WriteAllText(emulatorsIni, content);
      }
      catch
      {
        // Silently ignore if we can't create it
      }
    }
  }
}
