namespace GameBase;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void btnOptions_Click(object sender, EventArgs e)
    {
        using var dlg = new OptionsForm();
        dlg.ShowDialog(this);
    }
}
