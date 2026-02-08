namespace GameBase;

public partial class BatchEditForm : Form
{
  private sealed class LookupOption
  {
    public int? Id { get; }
    public string Name { get; }

    public LookupOption(int? id, string name)
    {
      Id = id;
      Name = name;
    }

    public override string ToString() => Name;
  }

  private sealed class FlagOption
  {
    public int? Value { get; }
    public string Name { get; }

    public FlagOption(int? value, string name)
    {
      Value = value;
      Name = name;
    }

    public override string ToString() => Name;
  }

  public int? SelectedPublisherId => (cboPublisher.SelectedItem as LookupOption)?.Id;
  public int? SelectedGenreId => (cboGenre.SelectedItem as LookupOption)?.Id;
  public int? SelectedFavorite => (cboFavorite.SelectedItem as FlagOption)?.Value;
  public int? SelectedClassic => (cboClassic.SelectedItem as FlagOption)?.Value;
  public int? SelectedAdult => (cboAdult.SelectedItem as FlagOption)?.Value;

  public BatchEditForm(Dictionary<int, string> publishers, Dictionary<int, string> genres)
  {
    InitializeComponent();
    LoadLookupOptions(cboPublisher, publishers);
    LoadLookupOptions(cboGenre, genres);
    LoadFlagOptions(cboFavorite);
    LoadFlagOptions(cboClassic);
    LoadFlagOptions(cboAdult);
  }

  private static void LoadLookupOptions(ComboBox combo, Dictionary<int, string> source)
  {
    var options = new List<LookupOption> { new(null, "(No Change)") };
    foreach (var item in source.OrderBy(kvp => kvp.Value))
    {
      options.Add(new LookupOption(item.Key, item.Value));
    }

    combo.DataSource = options;
    combo.SelectedIndex = 0;
  }

  private static void LoadFlagOptions(ComboBox combo)
  {
    var options = new List<FlagOption>
        {
            new(null, "(No Change)"),
            new(1, "Yes"),
            new(0, "No")
        };

    combo.DataSource = options;
    combo.SelectedIndex = 0;
  }

  private void btnOk_Click(object? sender, EventArgs e)
  {
    DialogResult = DialogResult.OK;
    Close();
  }

  private void btnCancel_Click(object? sender, EventArgs e)
  {
    DialogResult = DialogResult.Cancel;
    Close();
  }
}
