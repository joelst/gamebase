namespace GameBase;

partial class BigScreenshotForm
{
    private System.ComponentModel.IContainer components = null;
    private Panel panelScroll;
    private PictureBox pictureBox;
    private TrackBar trackZoom;
    private Label lblZoom;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        panelScroll = new Panel();
        pictureBox = new PictureBox();
        trackZoom = new TrackBar();
        lblZoom = new Label();
        panelScroll.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)pictureBox).BeginInit();
        ((System.ComponentModel.ISupportInitialize)trackZoom).BeginInit();
        SuspendLayout();
        // 
        // panelScroll
        // 
        panelScroll.AutoScroll = true;
        panelScroll.Controls.Add(pictureBox);
        panelScroll.Dock = DockStyle.Fill;
        panelScroll.Location = new Point(0, 45);
        panelScroll.Name = "panelScroll";
        panelScroll.Size = new Size(984, 616);
        panelScroll.TabIndex = 0;
        // 
        // pictureBox
        // 
        pictureBox.Location = new Point(0, 0);
        pictureBox.Name = "pictureBox";
        pictureBox.Size = new Size(100, 50);
        pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
        pictureBox.TabIndex = 0;
        pictureBox.TabStop = false;
        // 
        // trackZoom
        // 
        trackZoom.Dock = DockStyle.Top;
        trackZoom.LargeChange = 25;
        trackZoom.Location = new Point(0, 0);
        trackZoom.Maximum = 400;
        trackZoom.Minimum = 10;
        trackZoom.Name = "trackZoom";
        trackZoom.Size = new Size(984, 45);
        trackZoom.TabIndex = 1;
        trackZoom.TickFrequency = 10;
        trackZoom.Value = 100;
        trackZoom.Scroll += trackZoom_Scroll;
        // 
        // lblZoom
        // 
        lblZoom.AutoSize = true;
        lblZoom.BackColor = SystemColors.Control;
        lblZoom.Location = new Point(12, 12);
        lblZoom.Name = "lblZoom";
        lblZoom.Size = new Size(62, 15);
        lblZoom.TabIndex = 2;
        lblZoom.Text = "100% Zoom";
        // 
        // BigScreenshotForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(984, 661);
        Controls.Add(lblZoom);
        Controls.Add(panelScroll);
        Controls.Add(trackZoom);
        Name = "BigScreenshotForm";
        StartPosition = FormStartPosition.CenterParent;
        Text = "Screenshot";
        panelScroll.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)pictureBox).EndInit();
        ((System.ComponentModel.ISupportInitialize)trackZoom).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
}
