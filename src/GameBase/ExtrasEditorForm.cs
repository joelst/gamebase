using System.ComponentModel;
using GBSql.Data;
using GBSql.Models;

namespace GameBase;

public partial class ExtrasEditorForm : Form
{
  private readonly IGameBaseRepository _repository;
  private readonly int _gameId;
  private readonly string _gameName;
  private readonly BindingList<ExtraEntity> _extras = new();

  public ExtrasEditorForm(IGameBaseRepository repository, int gameId, string gameName)
  {
    _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    _gameId = gameId;
    _gameName = gameName;

    InitializeComponent();
    Load += ExtrasEditorForm_Load;
  }

  private async void ExtrasEditorForm_Load(object? sender, EventArgs e)
  {
    lblTitle.Text = $"Extras for {_gameName}";
    dgvExtras.AutoGenerateColumns = false;
    ConfigureGrid();
    await LoadExtrasAsync();
  }

  private void ConfigureGrid()
  {
    dgvExtras.Columns.Clear();

    var colId = new DataGridViewTextBoxColumn
    {
      DataPropertyName = nameof(ExtraEntity.EX_Id),
      HeaderText = "Id",
      Visible = false
    };
    var colName = new DataGridViewTextBoxColumn
    {
      DataPropertyName = nameof(ExtraEntity.Name),
      HeaderText = "Name"
    };
    var colFilename = new DataGridViewTextBoxColumn
    {
      DataPropertyName = nameof(ExtraEntity.Filename),
      HeaderText = "Filename"
    };
    var colType = new DataGridViewTextBoxColumn
    {
      DataPropertyName = nameof(ExtraEntity.EX_Type),
      HeaderText = "Type"
    };

    dgvExtras.Columns.AddRange(colId, colName, colFilename, colType);
    dgvExtras.DataSource = _extras;
  }

  private async Task LoadExtrasAsync()
  {
    _extras.Clear();
    var extras = await _repository.GetExtrasForGameAsync(_gameId);
    foreach (var extra in extras)
    {
      _extras.Add(extra);
    }
  }

  private void btnAdd_Click(object? sender, EventArgs e)
  {
    _extras.Add(new ExtraEntity { GA_Id = _gameId, Name = "New Extra", Filename = string.Empty, EX_Type = 0 });
    if (_extras.Count > 0)
    {
      dgvExtras.ClearSelection();
      dgvExtras.Rows[_extras.Count - 1].Selected = true;
    }
  }

  private void btnRemove_Click(object? sender, EventArgs e)
  {
    if (dgvExtras.SelectedRows.Count == 0) return;
    var row = dgvExtras.SelectedRows[0];
    if (row.DataBoundItem is ExtraEntity extra)
    {
      _extras.Remove(extra);
    }
  }

  private async void btnSave_Click(object? sender, EventArgs e)
  {
    try
    {
      // Commit any in-progress edits
      dgvExtras.EndEdit();
      var bc = BindingContext;
      if (dgvExtras.DataSource != null && bc != null && bc[dgvExtras.DataSource] is CurrencyManager cm)
      {
        cm.EndCurrentEdit();
      }

      await _repository.SaveExtrasForGameAsync(_gameId, _extras);
      await LoadExtrasAsync();
      MessageBox.Show(this, "Extras saved.", "Extras", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
      MessageBox.Show(this, "Failed to save extras: " + ex.Message, "Extras", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
  }

  private void btnClose_Click(object? sender, EventArgs e)
  {
    Close();
  }
}
