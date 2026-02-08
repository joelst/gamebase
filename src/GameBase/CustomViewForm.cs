namespace GameBase;

public partial class CustomViewForm : Form
{
    public class FilterCondition
    {
        public string Field { get; set; } = string.Empty;
        public string Operator { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public List<FilterCondition> Filters { get; } = new();
    public bool MatchAny => rbMatchAny.Checked;

    public CustomViewForm(IEnumerable<FilterCondition>? initial = null, bool matchAny = false)
    {
        InitializeComponent();
        rbMatchAny.Checked = matchAny;
        rbMatchAll.Checked = !matchAny;
        ConfigureGrid();
        LoadFilters(initial);
    }

    private void ConfigureGrid()
    {
        dgvFilters.AutoGenerateColumns = false;
        dgvFilters.Columns.Clear();

        var colField = new DataGridViewComboBoxColumn
        {
            DataPropertyName = nameof(FilterCondition.Field),
            HeaderText = "Field",
            DataSource = new[] { "Title", "Publisher", "Programmer", "Musician", "Language", "Genre", "Year", "Comment", "Favorite", "Adult" }
        };
        var colOp = new DataGridViewComboBoxColumn
        {
            DataPropertyName = nameof(FilterCondition.Operator),
            HeaderText = "Operator",
            DataSource = new[] { "Contains", "Equals", "NotEquals", "StartsWith", "EndsWith", "GreaterThan", "LessThan", "IsEmpty", "IsNotEmpty" }
        };
        var colValue = new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(FilterCondition.Value),
            HeaderText = "Value"
        };

        dgvFilters.Columns.AddRange(colField, colOp, colValue);
        dgvFilters.DataSource = new BindingSource { DataSource = Filters };
    }

    private void LoadFilters(IEnumerable<FilterCondition>? initial)
    {
        if (initial == null) return;
        Filters.Clear();
        Filters.AddRange(initial);
        (dgvFilters.DataSource as BindingSource)?.ResetBindings(false);
    }

    private void btnAdd_Click(object? sender, EventArgs e)
    {
        Filters.Add(new FilterCondition { Field = "Title", Operator = "Contains", Value = string.Empty });
        (dgvFilters.DataSource as BindingSource)?.ResetBindings(false);
    }

    private void btnRemove_Click(object? sender, EventArgs e)
    {
        if (dgvFilters.SelectedRows.Count == 0) return;
        var row = dgvFilters.SelectedRows[0];
        if (row.DataBoundItem is FilterCondition fc)
        {
            Filters.Remove(fc);
            (dgvFilters.DataSource as BindingSource)?.ResetBindings(false);
        }
    }

    private void btnClear_Click(object? sender, EventArgs e)
    {
        Filters.Clear();
        (dgvFilters.DataSource as BindingSource)?.ResetBindings(false);
    }

    private void btnOK_Click(object? sender, EventArgs e)
    {
        dgvFilters.EndEdit();
        (dgvFilters.DataSource as BindingSource)?.EndEdit();
        DialogResult = DialogResult.OK;
        Close();
    }

    private void btnCancel_Click(object? sender, EventArgs e)
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }
}
