//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Remove records dialog
//==============================================

using Microsoft.Data.Sqlite;

namespace GBToolbox;

internal sealed class RemoveRecordsForm : Form
{
  private readonly string _databasePath;
  private readonly ComboBox _comboTables;
  private readonly ListBox _listRecords;
  private readonly Button _btnDelete;
  private readonly Label _infoLabel;
  private readonly List<TableDescriptor> _tables;

  public RemoveRecordsForm(string databasePath)
  {
    _databasePath = databasePath;

    Text = "Delete Records from Table";
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(520, 420);

    var lblTable = new Label
    {
      Text = "Table:",
      AutoSize = false,
      Left = 12,
      Top = 16,
      Width = 60
    };

    _comboTables = new ComboBox
    {
      Left = 80,
      Top = 12,
      Width = 420,
      DropDownStyle = ComboBoxStyle.DropDownList
    };

    _listRecords = new ListBox
    {
      Left = 12,
      Top = 48,
      Width = 488,
      Height = 280
    };

    _infoLabel = new Label
    {
      Left = 12,
      Top = 336,
      Width = 488,
      Height = 20,
      Text = "Select a record to delete."
    };

    _btnDelete = new Button
    {
      Text = "Delete",
      Left = 320,
      Top = 360,
      Width = 85,
      Enabled = false
    };

    var btnClose = new Button
    {
      Text = "Close",
      Left = 415,
      Top = 360,
      Width = 85,
      DialogResult = DialogResult.Cancel
    };

    Controls.Add(lblTable);
    Controls.Add(_comboTables);
    Controls.Add(_listRecords);
    Controls.Add(_infoLabel);
    Controls.Add(_btnDelete);
    Controls.Add(btnClose);

    _tables = TableDescriptor.CreateDefaults();

    _comboTables.SelectedIndexChanged += (_, _) => LoadRecords();
    _listRecords.SelectedIndexChanged += (_, _) => _btnDelete.Enabled = _listRecords.SelectedItem != null;
    _btnDelete.Click += (_, _) => DeleteSelectedRecord();

    Load += RemoveRecordsForm_Load;
  }

  private void RemoveRecordsForm_Load(object? sender, EventArgs e)
  {
    using var conn = ToolboxDatabase.Open(_databasePath);
    foreach (var table in _tables)
    {
      if (ToolboxDatabase.TableExists(conn, table.TableName))
      {
        _comboTables.Items.Add(table);
      }
    }

    if (_comboTables.Items.Count > 0)
    {
      _comboTables.SelectedIndex = 0;
    }
  }

  private void LoadRecords()
  {
    _listRecords.Items.Clear();
    _btnDelete.Enabled = false;

    if (_comboTables.SelectedItem is not TableDescriptor table)
    {
      return;
    }

    using var conn = ToolboxDatabase.Open(_databasePath);
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"SELECT {table.IdColumn}, Name FROM {table.TableName} ORDER BY Name;";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
    {
      int id = reader.GetInt32(0);
      string name = reader.GetString(1);
      if (id <= table.MinReservedId)
      {
        continue;
      }
      _listRecords.Items.Add(new RecordOption(id, name));
    }
  }

  private void DeleteSelectedRecord()
  {
    if (_comboTables.SelectedItem is not TableDescriptor table)
    {
      return;
    }

    if (_listRecords.SelectedItem is not RecordOption record)
    {
      return;
    }

    if (MessageBox.Show($"Delete '{record.Name}'?", "GBToolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
    {
      return;
    }

    int replacementId = table.MinReservedId;
    string replacementName = "";

    using var conn = ToolboxDatabase.Open(_databasePath);

    if (!string.IsNullOrWhiteSpace(table.GamesColumn) && ToolboxDatabase.ColumnExists(conn, "Games", table.GamesColumn))
    {
      using var select = new SelectRecordForm(_databasePath, table.TableName, table.IdColumn, "Select replacement record");
      if (select.ShowDialog(this) != DialogResult.OK)
      {
        return;
      }

      replacementId = select.SelectedId;
      replacementName = select.SelectedName;

      if (replacementId == record.Id)
      {
        MessageBox.Show("You must select a different replacement record.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      ToolboxDatabase.ExecuteNonQuery(conn,
        $"UPDATE Games SET {table.GamesColumn}=@newId WHERE {table.GamesColumn}=@oldId;",
        new SqliteParameter("@newId", replacementId),
        new SqliteParameter("@oldId", record.Id));
    }

    ToolboxDatabase.ExecuteNonQuery(conn,
      $"DELETE FROM {table.TableName} WHERE {table.IdColumn}=@id;",
      new SqliteParameter("@id", record.Id));

    LoadRecords();

    if (!string.IsNullOrWhiteSpace(replacementName))
    {
      _infoLabel.Text = $"Reassigned to {replacementName}.";
    }
    else
    {
      _infoLabel.Text = "Record deleted.";
    }
  }

  private sealed record RecordOption(int Id, string Name)
  {
    public override string ToString() => Name;
  }

  private sealed class TableDescriptor
  {
    public string TableName { get; }
    public string IdColumn { get; }
    public string GamesColumn { get; }
    public int MinReservedId { get; }

    private TableDescriptor(string tableName, string idColumn, string gamesColumn, int minReservedId)
    {
      TableName = tableName;
      IdColumn = idColumn;
      GamesColumn = gamesColumn;
      MinReservedId = minReservedId;
    }

    public override string ToString() => TableName;

    public static List<TableDescriptor> CreateDefaults()
    {
      return new List<TableDescriptor>
      {
        new("Publishers", "PU_Id", "PU_Id", 1),
        new("Musicians", "MU_Id", "MU_Id", 1),
        new("Programmers", "PR_Id", "PR_Id", 1),
        new("Languages", "LA_Id", "LA_Id", 1),
        new("Crackers", "CR_Id", "CR_Id", 1),
        new("Artists", "AR_Id", "AR_Id", 1),
        new("Developers", "DE_Id", "DE_Id", 1),
        new("Licenses", "LI_Id", "LI_Id", 1),
        new("Rarities", "RA_Id", "RA_Id", 1)
      };
    }
  }
}
