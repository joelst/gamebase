//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBToolbox: Progress dialog
//==============================================

namespace GBToolbox;

internal sealed class ProgressForm : Form
{
  private readonly Label _statusLabel;
  private readonly Label _countLabel;
  private readonly ProgressBar _progressBar;
  private readonly Button _cancelButton;

  public event Action? CancelRequested;

  public ProgressForm(string title)
  {
    Text = title;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    StartPosition = FormStartPosition.CenterParent;
    MaximizeBox = false;
    MinimizeBox = false;
    ControlBox = false;
    ClientSize = new Size(520, 120);

    _statusLabel = new Label
    {
      AutoSize = false,
      Left = 12,
      Top = 12,
      Width = 380,
      Height = 20,
      Text = "Initializing..."
    };

    _countLabel = new Label
    {
      AutoSize = false,
      Left = 404,
      Top = 12,
      Width = 92,
      Height = 20,
      TextAlign = ContentAlignment.MiddleRight,
      Text = "0"
    };

    _progressBar = new ProgressBar
    {
      Left = 12,
      Top = 36,
      Width = 484,
      Height = 20,
      Minimum = 0,
      Maximum = 100
    };

    _cancelButton = new Button
    {
      Text = "Cancel",
      Left = 206,
      Top = 72,
      Width = 100
    };
    _cancelButton.Click += (_, _) => CancelRequested?.Invoke();

    Controls.Add(_statusLabel);
    Controls.Add(_countLabel);
    Controls.Add(_progressBar);
    Controls.Add(_cancelButton);
  }

  public void UpdateStatus(string status, int? count = null)
  {
    _statusLabel.Text = status;
    if (count.HasValue)
    {
      _countLabel.Text = count.Value.ToString();
    }
  }

  public void UpdateProgress(int value, int max)
  {
    if (max <= 0)
    {
      _progressBar.Value = 0;
      return;
    }

    if (_progressBar.Maximum != max)
    {
      _progressBar.Maximum = max;
    }

    _progressBar.Value = Math.Clamp(value, _progressBar.Minimum, _progressBar.Maximum);
  }

  public void SetCancelEnabled(bool enabled)
  {
    _cancelButton.Enabled = enabled;
  }
}
