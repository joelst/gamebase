//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Wait dialog
//==============================================

namespace GBToolbox;

internal sealed class WaitForm : Form
{
  private readonly Label _label;

  public WaitForm(string message)
  {
    Text = "Please wait";
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ControlBox = false;
    ClientSize = new Size(320, 70);

    _label = new Label
    {
      Text = message,
      AutoSize = false,
      TextAlign = ContentAlignment.MiddleCenter,
      Dock = DockStyle.Fill
    };

    Controls.Add(_label);
  }

  public void UpdateMessage(string message)
  {
    _label.Text = message;
  }
}
