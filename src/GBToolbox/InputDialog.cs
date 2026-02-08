//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Input dialog
//==============================================

namespace GBToolbox;

internal sealed class InputDialog : Form
{
  private readonly TextBox _textInput;
  private readonly Button _okButton;
  private readonly string _invalidChars;
  private readonly bool _allowEmpty;

  public string ResultText { get; private set; } = string.Empty;

  public InputDialog(string caption, string prompt, string startingText, string invalidChars, int maxLength, bool allowEmpty)
  {
    _invalidChars = invalidChars;
    _allowEmpty = allowEmpty;

    Text = caption;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ClientSize = new Size(460, 140);

    var label = new Label
    {
      Text = prompt,
      AutoSize = false,
      Left = 12,
      Top = 12,
      Width = 436,
      Height = 20
    };

    _textInput = new TextBox
    {
      Left = 12,
      Top = 36,
      Width = 436,
      Text = startingText ?? string.Empty,
      MaxLength = maxLength
    };

    _okButton = new Button
    {
      Text = "OK",
      DialogResult = DialogResult.OK,
      Left = 260,
      Top = 80,
      Width = 90,
      Enabled = allowEmpty || !string.IsNullOrWhiteSpace(_textInput.Text)
    };

    var cancelButton = new Button
    {
      Text = "Cancel",
      DialogResult = DialogResult.Cancel,
      Left = 358,
      Top = 80,
      Width = 90
    };

    _textInput.TextChanged += (_, _) =>
    {
      if (!_allowEmpty)
      {
        _okButton.Enabled = !string.IsNullOrWhiteSpace(_textInput.Text);
      }
    };

    _textInput.KeyPress += (_, args) =>
    {
      if (!string.IsNullOrEmpty(_invalidChars) && _invalidChars.Contains(args.KeyChar))
      {
        args.Handled = true;
      }
    };

    Controls.Add(label);
    Controls.Add(_textInput);
    Controls.Add(_okButton);
    Controls.Add(cancelButton);

    AcceptButton = _okButton;
    CancelButton = cancelButton;
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    if (DialogResult == DialogResult.OK)
    {
      if (!_allowEmpty && string.IsNullOrWhiteSpace(_textInput.Text))
      {
        e.Cancel = true;
        return;
      }

      if (!string.IsNullOrEmpty(_invalidChars) && _textInput.Text.IndexOfAny(_invalidChars.ToCharArray()) >= 0)
      {
        MessageBox.Show("You have entered one or more invalid characters!", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        e.Cancel = true;
        return;
      }

      ResultText = _textInput.Text.Trim();
    }

    base.OnFormClosing(e);
  }
}
