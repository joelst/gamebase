//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Select record dialog
//==============================================

using Microsoft.Data.Sqlite;

namespace GBToolbox;

internal sealed class SelectRecordForm : Form
{
  private readonly ListBox _listBox;
  private readonly Button _okButton;
  private readonly string _databasePath;
  private readonly string _table;
  private readonly string _idColumn;

  public int SelectedId { get; private set; }
  public string SelectedName { get; private set; } = string.Empty;

  public SelectRecordForm(string databasePath, string table, string idColumn, string title)
  {
    _databasePath = databasePath;
    _table = table;
    _idColumn = idColumn;

    Text = title;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(420, 360);

    _listBox = new ListBox
    {
      Left = 12,
      Top = 12,
      Width = 396,
      Height = 280
    };

    _okButton = new Button
    {
      Text = "OK",
      DialogResult = DialogResult.OK,
      Left = 228,
      Top = 310,
      Width = 85,
      Enabled = false
    };

    var cancelButton = new Button
    {
      Text = "Cancel",
      DialogResult = DialogResult.Cancel,
      Left = 323,
      Top = 310,
      Width = 85
    };

    _listBox.SelectedIndexChanged += (_, _) => _okButton.Enabled = _listBox.SelectedItem != null;

    Controls.Add(_listBox);
    Controls.Add(_okButton);
    Controls.Add(cancelButton);

    AcceptButton = _okButton;
    CancelButton = cancelButton;

    Load += SelectRecordForm_Load;
  }

  private void SelectRecordForm_Load(object? sender, EventArgs e)
  {
    using var conn = ToolboxDatabase.Open(_databasePath);
    if (!ToolboxDatabase.TableExists(conn, _table))
    {
      return;
    }

    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"SELECT {_idColumn}, Name FROM {_table} ORDER BY Name;";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
    {
      int id = reader.GetInt32(0);
      string name = reader.GetString(1);
      _listBox.Items.Add(new RecordOption(id, name));
    }
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    if (DialogResult == DialogResult.OK)
    {
      if (_listBox.SelectedItem is not RecordOption option)
      {
        e.Cancel = true;
        return;
      }

      SelectedId = option.Id;
      SelectedName = option.Name;
    }

    base.OnFormClosing(e);
  }

  private sealed record RecordOption(int Id, string Name)
  {
    public override string ToString() => Name;
  }
}
