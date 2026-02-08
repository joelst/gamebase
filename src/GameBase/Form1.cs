using GBSql.Data;
using GBSql.Models;

namespace GameBase;

public partial class Form1 : Form
{
    private readonly IGameBaseRepository _repo;

    public Form1(IGameBaseRepository repo)
    {
        _repo = repo ?? throw new ArgumentNullException(nameof(repo));
        InitializeComponent();
        Load += Form1_Load;
        lstGames.SelectedIndexChanged += LstGames_SelectedIndexChanged;
    }

    private async void Form1_Load(object? sender, EventArgs e)
    {
        try
        {
            var games = (await _repo.GetGamesAsync(limit: 200)).ToList();
            lstGames.DataSource = games;
            lstGames.DisplayMember = nameof(GameEntity.Name);
            lblStatus.Text = $"Loaded {games.Count} games.";
        }
        catch (Exception ex)
        {
            lblStatus.Text = $"Failed to load games: {ex.Message}";
        }
    }

    private async void LstGames_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (lstGames.SelectedItem is GameEntity game)
        {
            try
            {
                var extras = (await _repo.GetExtrasForGameAsync(game.GA_Id)).ToList();
                lstExtras.DataSource = extras;
                lstExtras.DisplayMember = nameof(ExtraEntity.Name);
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Failed to load extras: {ex.Message}";
            }
        }
    }

    private void btnOptions_Click(object sender, EventArgs e)
    {
        using var dlg = new OptionsForm();
        dlg.ShowDialog(this);
    }
}
