namespace GameBase;

public partial class BigScreenshotForm : Form
{
  private Image? _image;
  private float _zoom = 1f;

  public BigScreenshotForm(Image? image)
  {
    _image = image;
    InitializeComponent();
    Load += BigScreenshotForm_Load;
    KeyPreview = true;
    KeyDown += BigScreenshotForm_KeyDown;
    MouseWheel += BigScreenshotForm_MouseWheel;
  }

  private void BigScreenshotForm_Load(object? sender, EventArgs e)
  {
    if (_image == null)
    {
      MessageBox.Show(this, "No screenshot available.", "Screenshot", MessageBoxButtons.OK, MessageBoxIcon.Information);
      Close();
      return;
    }

    pictureBox.Image = _image;
    UpdateZoom();
  }

  private void trackZoom_Scroll(object? sender, EventArgs e)
  {
    _zoom = trackZoom.Value / 100f;
    UpdateZoom();
  }

  private void UpdateZoom()
  {
    if (_image == null) return;
    lblZoom.Text = $"{trackZoom.Value}% Zoom";
    pictureBox.Width = (int)(_image.Width * _zoom);
    pictureBox.Height = (int)(_image.Height * _zoom);
    pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
  }

  private void BigScreenshotForm_MouseWheel(object? sender, MouseEventArgs e)
  {
    int delta = e.Delta > 0 ? 5 : -5;
    int newVal = Math.Clamp(trackZoom.Value + delta, trackZoom.Minimum, trackZoom.Maximum);
    if (newVal != trackZoom.Value)
    {
      trackZoom.Value = newVal;
      _zoom = trackZoom.Value / 100f;
      UpdateZoom();
    }
  }

  private void BigScreenshotForm_KeyDown(object? sender, KeyEventArgs e)
  {
    if (e.KeyCode == Keys.Escape)
    {
      Close();
    }
  }
}
