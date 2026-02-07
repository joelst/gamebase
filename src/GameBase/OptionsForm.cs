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
      txtBasePath.Text = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
      txtGbName.Text = "GameBase";
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
        }
      }
      catch
      {
        // ignore load errors
      }
    }

    private string GetGemusIniPath()
    {
      return Path.Combine(txtBasePath.Text, txtGbName.Text, "Gemus.ini");
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
        MessageBox.Show(this, $"Saved backend {backend} to {gemusIni}", "Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
  }
}
