namespace GameBase;

public partial class ColumnsForm : Form
{
  public Dictionary<string, bool> ColumnVisibility { get; } = new();

  public ColumnsForm(IEnumerable<(string Name, bool Visible)> columns)
  {
    InitializeComponent();
    foreach (var column in columns)
    {
      checkedListColumns.Items.Add(column.Name, column.Visible);
    }
  }

  private void btnOk_Click(object? sender, EventArgs e)
  {
    ColumnVisibility.Clear();
    for (int i = 0; i < checkedListColumns.Items.Count; i++)
    {
      string name = checkedListColumns.Items[i]?.ToString() ?? string.Empty;
      ColumnVisibility[name] = checkedListColumns.GetItemChecked(i);
    }

    DialogResult = DialogResult.OK;
    Close();
  }

  private void btnCancel_Click(object? sender, EventArgs e)
  {
    DialogResult = DialogResult.Cancel;
    Close();
  }
}
