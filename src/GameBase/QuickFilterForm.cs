namespace GameBase;

public partial class QuickFilterForm : Form
{
  public string FilterText => txtFilter.Text.Trim();
  public string SelectedField => cboField.SelectedItem?.ToString() ?? string.Empty;
  public bool IncludeSids => rbBothTables.Checked;

  private readonly string[] _fields = new[]
  {
        "Title","Publisher","Programmer","Musician","Language","Genre","Year","Comment","Favorite","Adult"
    };

  public QuickFilterForm()
  {
    InitializeComponent();
    cboField.Items.AddRange(_fields);
    if (_fields.Length > 0) cboField.SelectedIndex = 0;
  }

  private void txtFilter_TextChanged(object? sender, EventArgs e)
  {
    btnOK.Enabled = !string.IsNullOrWhiteSpace(txtFilter.Text);
  }

  private void btnOK_Click(object? sender, EventArgs e)
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
