//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Database settings dialog
//==============================================

using Microsoft.Data.Sqlite;

namespace GBToolbox;

internal sealed class DatabaseSettingsForm : Form
{
  private readonly string _databasePath;
  private readonly TextBox _txtDatabaseName;
  private readonly TextBox _txtWindowTitle;
  private readonly TextBox _txtMessage;
  private readonly CheckBox _chkGemusAsk;
  private readonly Button _okButton;
  private readonly Label _warningLabel;

  public DatabaseSettingsForm(string databasePath)
  {
    _databasePath = databasePath;

    Text = "Database Settings";
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(520, 330);

    var lblDatabaseName = new Label
    {
      Text = "Database name:",
      AutoSize = false,
      Left = 12,
      Top = 16,
      Width = 140,
      Height = 20
    };

    _txtDatabaseName = new TextBox
    {
      Left = 160,
      Top = 14,
      Width = 340
    };

    var lblWindowTitle = new Label
    {
      Text = "GameBase window title:",
      AutoSize = false,
      Left = 12,
      Top = 48,
      Width = 140,
      Height = 20
    };

    _txtWindowTitle = new TextBox
    {
      Left = 160,
      Top = 46,
      Width = 340
    };

    var lblMessage = new Label
    {
      Text = "First load message:",
      AutoSize = false,
      Left = 12,
      Top = 80,
      Width = 140,
      Height = 20
    };

    _txtMessage = new TextBox
    {
      Left = 160,
      Top = 78,
      Width = 340,
      Height = 140,
      Multiline = true,
      ScrollBars = ScrollBars.Vertical
    };

    _chkGemusAsk = new CheckBox
    {
      Text = "Prompt for GEMUS Setup Wizard",
      Left = 160,
      Top = 230,
      Width = 320
    };

    _warningLabel = new Label
    {
      Text = "The Config table (DatabaseName/GameBaseWindowTitle) was not found.",
      AutoSize = false,
      Left = 12,
      Top = 255,
      Width = 488,
      Height = 20,
      ForeColor = Color.Firebrick,
      Visible = false
    };

    _okButton = new Button
    {
      Text = "OK",
      DialogResult = DialogResult.OK,
      Left = 316,
      Top = 280,
      Width = 90
    };

    var cancelButton = new Button
    {
      Text = "Cancel",
      DialogResult = DialogResult.Cancel,
      Left = 410,
      Top = 280,
      Width = 90
    };

    Controls.Add(lblDatabaseName);
    Controls.Add(_txtDatabaseName);
    Controls.Add(lblWindowTitle);
    Controls.Add(_txtWindowTitle);
    Controls.Add(lblMessage);
    Controls.Add(_txtMessage);
    Controls.Add(_chkGemusAsk);
    Controls.Add(_warningLabel);
    Controls.Add(_okButton);
    Controls.Add(cancelButton);

    AcceptButton = _okButton;
    CancelButton = cancelButton;

    Load += DatabaseSettingsForm_Load;
  }

  private void DatabaseSettingsForm_Load(object? sender, EventArgs e)
  {
    using var conn = ToolboxDatabase.Open(_databasePath);
    if (!ToolboxDatabase.TableExists(conn, "Config")
        || !ToolboxDatabase.ColumnExists(conn, "Config", "DatabaseName")
        || !ToolboxDatabase.ColumnExists(conn, "Config", "GameBaseWindowTitle")
        || !ToolboxDatabase.ColumnExists(conn, "Config", "FirstLoadMessage")
        || !ToolboxDatabase.ColumnExists(conn, "Config", "FirstLoadGemusAsk"))
    {
      _warningLabel.Visible = true;
      _okButton.Enabled = false;
      return;
    }

    LoadConfigRow(conn);
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    if (DialogResult == DialogResult.OK)
    {
      using var conn = ToolboxDatabase.Open(_databasePath);
      if (!ToolboxDatabase.TableExists(conn, "Config")
          || !ToolboxDatabase.ColumnExists(conn, "Config", "DatabaseName")
          || !ToolboxDatabase.ColumnExists(conn, "Config", "GameBaseWindowTitle")
          || !ToolboxDatabase.ColumnExists(conn, "Config", "FirstLoadMessage")
          || !ToolboxDatabase.ColumnExists(conn, "Config", "FirstLoadGemusAsk"))
      {
        e.Cancel = true;
        return;
      }

      SaveConfigRow(conn);
    }

    base.OnFormClosing(e);
  }

  private void LoadConfigRow(SqliteConnection conn)
  {
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "SELECT DatabaseName, GameBaseWindowTitle, FirstLoadMessage, FirstLoadGemusAsk FROM Config LIMIT 1;";
    using var reader = cmd.ExecuteReader();
    if (reader.Read())
    {
      _txtDatabaseName.Text = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
      _txtWindowTitle.Text = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
      _txtMessage.Text = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
      int ask = reader.IsDBNull(3) ? 0 : reader.GetInt32(3);
      _chkGemusAsk.Checked = ask != 0;
    }
  }

  private void SaveConfigRow(SqliteConnection conn)
  {
    string message = _txtMessage.Text.Trim();
    if (message.Length > 65526)
    {
      message = message.Substring(0, 65526);
    }

    using var update = conn.CreateCommand();
    update.CommandText = @"UPDATE Config
SET DatabaseName=@DatabaseName,
    GameBaseWindowTitle=@WindowTitle,
    FirstLoadMessage=@Message,
    FirstLoadGemusAsk=@Ask;";
    update.Parameters.AddWithValue("@DatabaseName", _txtDatabaseName.Text.Trim());
    update.Parameters.AddWithValue("@WindowTitle", _txtWindowTitle.Text.Trim());
    update.Parameters.AddWithValue("@Message", message);
    update.Parameters.AddWithValue("@Ask", _chkGemusAsk.Checked ? 1 : 0);
    if (update.ExecuteNonQuery() > 0)
    {
      return;
    }

    using var insert = conn.CreateCommand();
    insert.CommandText = @"INSERT INTO Config (DatabaseName, GameBaseWindowTitle, FirstLoadMessage, FirstLoadGemusAsk)
VALUES (@DatabaseName, @WindowTitle, @Message, @Ask);";
    insert.Parameters.AddWithValue("@DatabaseName", _txtDatabaseName.Text.Trim());
    insert.Parameters.AddWithValue("@WindowTitle", _txtWindowTitle.Text.Trim());
    insert.Parameters.AddWithValue("@Message", message);
    insert.Parameters.AddWithValue("@Ask", _chkGemusAsk.Checked ? 1 : 0);
    insert.ExecuteNonQuery();
  }
}
