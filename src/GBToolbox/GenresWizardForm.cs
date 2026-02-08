//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Genres wizard (partial)
//==============================================

using Microsoft.Data.Sqlite;

namespace GBToolbox;

internal sealed class GenresWizardForm : Form
{
  private readonly string _databasePath;
  private readonly RadioButton _optClear;
  private readonly RadioButton _optRemoveGenres;
  private readonly RadioButton _optRemovePGenres;
  private readonly Button _btnRun;

  public GenresWizardForm(string databasePath)
  {
    _databasePath = databasePath;

    Text = "Genres Wizard";
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(520, 320);

    var lblInfo = new Label
    {
      Text = "Select an action to perform on Genres/Parent Genres:",
      AutoSize = false,
      Left = 12,
      Top = 16,
      Width = 496,
      Height = 20
    };

    _optClear = new RadioButton
    {
      Text = "Clear all Genres and Parent Genres",
      Left = 12,
      Top = 50,
      Width = 496,
      Checked = true
    };

    var lblClear = new Label
    {
      Text = "  Removes all genres and recategorizes games to [Uncategorized].",
      AutoSize = false,
      Left = 12,
      Top = 75,
      Width = 496,
      Height = 20,
      ForeColor = Color.Gray
    };

    _optRemovePGenres = new RadioButton
    {
      Text = "Remove Parent Genres",
      Left = 12,
      Top = 105,
      Width = 496
    };

    var lblPGenres = new Label
    {
      Text = "  Removes selected parent genres and reassigns affected games.",
      AutoSize = false,
      Left = 12,
      Top = 130,
      Width = 496,
      Height = 20,
      ForeColor = Color.Gray
    };

    _optRemoveGenres = new RadioButton
    {
      Text = "Remove Genres",
      Left = 12,
      Top = 160,
      Width = 496
    };

    var lblGenres = new Label
    {
      Text = "  Removes selected genres and reassigns affected games.",
      AutoSize = false,
      Left = 12,
      Top = 185,
      Width = 496,
      Height = 20,
      ForeColor = Color.Gray
    };

    _btnRun = new Button
    {
      Text = "Run",
      Left = 340,
      Top = 260,
      Width = 80
    };

    var btnClose = new Button
    {
      Text = "Close",
      Left = 428,
      Top = 260,
      Width = 80,
      DialogResult = DialogResult.Cancel
    };

    Controls.Add(lblInfo);
    Controls.Add(_optClear);
    Controls.Add(lblClear);
    Controls.Add(_optRemovePGenres);
    Controls.Add(lblPGenres);
    Controls.Add(_optRemoveGenres);
    Controls.Add(lblGenres);
    Controls.Add(_btnRun);
    Controls.Add(btnClose);

    _btnRun.Click += (_, _) => RunAction();
  }

  private void RunAction()
  {
    if (_optClear.Checked)
    {
      ClearAllGenres();
    }
    else if (_optRemovePGenres.Checked)
    {
      RemoveParentGenres();
    }
    else if (_optRemoveGenres.Checked)
    {
      RemoveGenres();
    }
  }

  private void ClearAllGenres()
  {
    if (MessageBox.Show("Clear all Genres and Parent Genres?\n\nThis will recategorize all games to [Uncategorized].", "GBToolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
    {
      return;
    }

    try
    {
      using var conn = ToolboxDatabase.Open(_databasePath);
      EnsureUncategorized(conn, "Genres", "GE_Id");
      EnsureUncategorized(conn, "PGenres", "PG_Id");

      if (ToolboxDatabase.TableExists(conn, "Games"))
      {
        if (ToolboxDatabase.ColumnExists(conn, "Games", "GE_Id"))
        {
          ToolboxDatabase.ExecuteNonQuery(conn, "UPDATE Games SET GE_Id=1 WHERE GE_Id IS NULL OR GE_Id <> 1;");
        }
        if (ToolboxDatabase.ColumnExists(conn, "Games", "PG_Id"))
        {
          ToolboxDatabase.ExecuteNonQuery(conn, "UPDATE Games SET PG_Id=1 WHERE PG_Id IS NULL OR PG_Id <> 1;");
        }
      }

      if (ToolboxDatabase.TableExists(conn, "Genres"))
      {
        ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM Genres WHERE GE_Id <> 1;");
      }
      if (ToolboxDatabase.TableExists(conn, "PGenres"))
      {
        ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM PGenres WHERE PG_Id <> 1;");
      }

      MessageBox.Show("Genres cleared.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Unable to clear genres: {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
  }

  private void RemoveParentGenres()
  {
    using var conn = ToolboxDatabase.Open(_databasePath);
    if (!ToolboxDatabase.TableExists(conn, "PGenres"))
    {
      MessageBox.Show("The PGenres table does not exist.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    var genres = LoadGenreList(conn, "PGenres", "PG_Id");
    if (genres.Count == 0)
    {
      MessageBox.Show("There are no Parent Genres to remove.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    using var selectForm = new SelectRecordForm(_databasePath, "PGenres", "PG_Id", "Select Parent Genre to Remove");
    if (selectForm.ShowDialog(this) != DialogResult.OK)
    {
      return;
    }

    int removeId = selectForm.SelectedId;
    string removeName = selectForm.SelectedName;

    if (removeId <= 1)
    {
      MessageBox.Show("Cannot remove the default Parent Genre.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    using var replaceForm = new SelectRecordForm(_databasePath, "PGenres", "PG_Id", "Select Replacement Parent Genre");
    if (replaceForm.ShowDialog(this) != DialogResult.OK)
    {
      return;
    }

    int replaceId = replaceForm.SelectedId;
    if (replaceId == removeId)
    {
      MessageBox.Show("Replacement must be different from the genre being removed.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      return;
    }

    try
    {
      if (ToolboxDatabase.ColumnExists(conn, "Games", "PG_Id"))
      {
        ToolboxDatabase.ExecuteNonQuery(conn, "UPDATE Games SET PG_Id=@newId WHERE PG_Id=@oldId;",
          new SqliteParameter("@newId", replaceId),
          new SqliteParameter("@oldId", removeId));
      }

      ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM PGenres WHERE PG_Id=@id;",
        new SqliteParameter("@id", removeId));

      MessageBox.Show($"Parent Genre '{removeName}' removed.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Unable to remove parent genre: {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
  }

  private void RemoveGenres()
  {
    using var conn = ToolboxDatabase.Open(_databasePath);
    if (!ToolboxDatabase.TableExists(conn, "Genres"))
    {
      MessageBox.Show("The Genres table does not exist.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    var genres = LoadGenreList(conn, "Genres", "GE_Id");
    if (genres.Count == 0)
    {
      MessageBox.Show("There are no Genres to remove.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    using var selectForm = new SelectRecordForm(_databasePath, "Genres", "GE_Id", "Select Genre to Remove");
    if (selectForm.ShowDialog(this) != DialogResult.OK)
    {
      return;
    }

    int removeId = selectForm.SelectedId;
    string removeName = selectForm.SelectedName;

    if (removeId <= 1)
    {
      MessageBox.Show("Cannot remove the default Genre.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
      return;
    }

    using var replaceForm = new SelectRecordForm(_databasePath, "Genres", "GE_Id", "Select Replacement Genre");
    if (replaceForm.ShowDialog(this) != DialogResult.OK)
    {
      return;
    }

    int replaceId = replaceForm.SelectedId;
    if (replaceId == removeId)
    {
      MessageBox.Show("Replacement must be different from the genre being removed.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      return;
    }

    try
    {
      if (ToolboxDatabase.ColumnExists(conn, "Games", "GE_Id"))
      {
        ToolboxDatabase.ExecuteNonQuery(conn, "UPDATE Games SET GE_Id=@newId WHERE GE_Id=@oldId;",
          new SqliteParameter("@newId", replaceId),
          new SqliteParameter("@oldId", removeId));
      }

      ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM Genres WHERE GE_Id=@id;",
        new SqliteParameter("@id", removeId));

      MessageBox.Show($"Genre '{removeName}' removed.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Unable to remove genre: {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
  }

  private static List<GenreOption> LoadGenreList(SqliteConnection conn, string table, string idColumn)
  {
    var list = new List<GenreOption>();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"SELECT {idColumn}, Name FROM {table} WHERE {idColumn} > 1 ORDER BY Name;";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
    {
      int id = reader.GetInt32(0);
      string name = reader.GetString(1);
      list.Add(new GenreOption(id, name));
    }
    return list;
  }

  private sealed record GenreOption(int Id, string Name);

  private static void EnsureUncategorized(SqliteConnection conn, string table, string idColumn)
  {
    if (!ToolboxDatabase.TableExists(conn, table))
    {
      return;
    }

    using var cmd = conn.CreateCommand();
    cmd.CommandText = $"INSERT OR IGNORE INTO {table} ({idColumn}, Name) VALUES (1, 'Uncategorized');";
    cmd.ExecuteNonQuery();
  }
}
