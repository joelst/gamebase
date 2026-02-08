using GBSql.Data;
using GBSql.Models;

namespace GameBase;

public partial class Form1 : Form
{
    private readonly IGameBaseRepository _repository;
    private List<GameEntity> _allGames = new();
    private List<GameEntity> _filteredGames = new();

    private readonly Dictionary<int, string> _publishers = new();
    private readonly Dictionary<int, string> _years = new();
    private readonly Dictionary<int, string> _programmers = new();
    private readonly Dictionary<int, string> _genres = new();
    private readonly Dictionary<int, string> _musicians = new();
    private readonly Dictionary<int, string> _languages = new();
    private readonly Dictionary<int, string> _artists = new();
    private readonly Dictionary<int, string> _developers = new();
    private readonly Dictionary<int, string> _licenses = new();
    private readonly Dictionary<int, string> _rarities = new();
    private readonly Dictionary<int, string> _pgenres = new();
    private readonly Dictionary<int, string> _difficulty = new();
    private readonly Dictionary<int, string> _crackers = new();

    public Form1(IGameBaseRepository repository)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        InitializeComponent();
        Load += Form1_Load;
        editGameToolStripMenuItem.Click += editGameToolStripMenuItem_Click;
        manageExtrasToolStripMenuItem.Click += manageExtrasToolStripMenuItem_Click;
        picScreenshot.DoubleClick += picScreenshot_DoubleClick;
    }

    private async void Form1_Load(object? sender, EventArgs e)
    {
        await InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            lblStatus.Text = "Initializing...";

            await LoadLookupsAsync();
            // Load games
            await LoadGamesAsync();

            lblStatus.Text = "Ready";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Initialization failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"Error: {ex.Message}";
        }
    }

    private async Task LoadLookupsAsync()
    {
        async Task LoadAsync(string table, Dictionary<int, string> target)
        {
            target.Clear();
            var items = await _repository.GetLookupTableAsync(table);
            foreach (var item in items)
            {
                if (!target.ContainsKey(item.Id))
                    target[item.Id] = item.Name;
            }
        }

        await LoadAsync("Publishers", _publishers);
        await LoadAsync("Years", _years);
        await LoadAsync("Programmers", _programmers);
        await LoadAsync("Genres", _genres);
        await LoadAsync("Musicians", _musicians);
        await LoadAsync("Languages", _languages);
        await LoadAsync("Artists", _artists);
        await LoadAsync("Developers", _developers);
        await LoadAsync("Licenses", _licenses);
        await LoadAsync("Rarities", _rarities);
        await LoadAsync("PGenres", _pgenres);
        await LoadAsync("Difficulty", _difficulty);
        await LoadAsync("Crackers", _crackers);
    }

    private async Task LoadGamesAsync()
    {
        try
        {
            lblStatus.Text = "Loading games...";

            _allGames = (await _repository.GetGamesAsync()).ToList();
            _filteredGames = new List<GameEntity>(_allGames);

            UpdateGameGrid();

            lblStatus.Text = $"Loaded {_allGames.Count} games";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load games: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"Error loading games: {ex.Message}";
        }
    }

    private void UpdateGameGrid()
    {
        dgvGames.DataSource = null;

        if (_filteredGames.Count == 0)
        {
            lblSearch.Text = "Games (0 of 0)";
            return;
        }

        // Create display data with selected columns
        var displayData = _filteredGames.Select(g => new
        {
            g.GA_Id,
            Game = g.Name,
            Year = Resolve(_years, g.YE_Id),
            Publisher = Resolve(_publishers, g.PU_Id),
            Programmer = Resolve(_programmers, g.PR_Id),
            Genre = Resolve(_genres, g.GE_Id)
        }).ToList();

        dgvGames.DataSource = displayData;

        // Hide ID column
        if (dgvGames.Columns["GA_Id"] != null)
            dgvGames.Columns["GA_Id"].Visible = false;

        // Auto-size columns
        dgvGames.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        lblSearch.Text = $"Games ({_filteredGames.Count} of {_allGames.Count})";
    }

    private void dgvGames_SelectionChanged(object? sender, EventArgs e)
    {
        if (dgvGames.SelectedRows.Count == 0)
        {
            ClearGameDetails();
            return;
        }

        var selectedId = (int)dgvGames.SelectedRows[0].Cells["GA_Id"].Value;
        var game = _filteredGames.FirstOrDefault(g => g.GA_Id == selectedId);

        if (game != null)
        {
            DisplayGameDetails(game);
        }
    }

    private async void DisplayGameDetails(GameEntity game)
    {
        // Display game info
        rtbGameInfo.Clear();
        rtbGameInfo.AppendText($"Title: {game.Name}\n");
        rtbGameInfo.AppendText($"Year: {Resolve(_years, game.YE_Id)}\n");
        rtbGameInfo.AppendText($"Publisher: {Resolve(_publishers, game.PU_Id)}\n");
        rtbGameInfo.AppendText($"Programmer: {Resolve(_programmers, game.PR_Id)}\n");
        rtbGameInfo.AppendText($"Genre: {Resolve(_genres, game.GE_Id)}\n");
        rtbGameInfo.AppendText($"Filename: {game.Filename}\n\n");

        if (!string.IsNullOrWhiteSpace(game.Comment))
        {
            rtbGameInfo.AppendText($"Comments:\n{game.Comment}\n");
        }

        // Load screenshot if available (not yet implemented; schema has no screenshot field)
        picScreenshot.Image = null;

        // Load extras
        try
        {
            var extras = (await _repository.GetExtrasForGameAsync(game.GA_Id)).ToList();
            lstExtras.Items.Clear();

            foreach (var extra in extras)
            {
                lstExtras.Items.Add(new ExtraListItem(extra));
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = $"Error loading extras: {ex.Message}";
        }
    }

    private void LoadScreenshot(GameEntity game)
    {
        // Placeholder for future screenshot support
        picScreenshot.Image = null;
    }

    private void ClearGameDetails()
    {
        rtbGameInfo.Clear();
        picScreenshot.Image = null;
        lstExtras.Items.Clear();
    }

    private void txtSearch_TextChanged(object? sender, EventArgs e)
    {
        var searchText = txtSearch.Text.Trim();

        if (string.IsNullOrWhiteSpace(searchText))
        {
            _filteredGames = new List<GameEntity>(_allGames);
        }
        else
        {
            _filteredGames = _allGames
                .Where(g =>
                    (g.Name?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (_publishers.TryGetValue(g.PU_Id ?? -1, out var pub) && pub.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (_programmers.TryGetValue(g.PR_Id ?? -1, out var prog) && prog.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                    (_genres.TryGetValue(g.GE_Id ?? -1, out var gen) && gen.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                )
                .ToList();
        }

        UpdateGameGrid();
    }

    private async void playGameToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        await PlaySelectedGameAsync();
    }

    private void editGameToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        if (dgvGames.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a game to view.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedId = (int)dgvGames.SelectedRows[0].Cells["GA_Id"].Value;
        var game = _filteredGames.FirstOrDefault(g => g.GA_Id == selectedId);
        if (game == null) return;

        using var dlg = new GameDetailsForm(game, _years, _publishers, _programmers, _genres,
            _musicians, _languages, _artists, _developers, _licenses, _rarities, _pgenres, _difficulty, _crackers);
        dlg.ShowDialog(this);
    }

    private void manageExtrasToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        if (dgvGames.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a game to manage extras.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedId = (int)dgvGames.SelectedRows[0].Cells["GA_Id"].Value;
        var game = _filteredGames.FirstOrDefault(g => g.GA_Id == selectedId);
        if (game == null) return;

        using var dlg = new ExtrasEditorForm(_repository, game.GA_Id, game.Name);
        dlg.ShowDialog(this);
    }

    private void picScreenshot_DoubleClick(object? sender, EventArgs e)
    {
        if (picScreenshot.Image == null)
        {
            MessageBox.Show(this, "No screenshot loaded.", "Screenshot", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new BigScreenshotForm(picScreenshot.Image);
        dlg.ShowDialog(this);
    }

    private async void dgvGames_DoubleClick(object? sender, EventArgs e)
    {
        await PlaySelectedGameAsync();
    }

    private async Task PlaySelectedGameAsync()
    {
        if (dgvGames.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a game to play.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedId = (int)dgvGames.SelectedRows[0].Cells["GA_Id"].Value;
        var game = _filteredGames.FirstOrDefault(g => g.GA_Id == selectedId);

        if (game == null) return;

        // TODO: Implement GEMUS engine integration for launching games
        MessageBox.Show($"Game launching not yet implemented.\n\nGame: {game.Name}\nFile: {game.Filename}",
            "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information);

        await Task.CompletedTask;
    }

    private async void lstExtras_DoubleClick(object? sender, EventArgs e)
    {
        if (lstExtras.SelectedItem is not ExtraListItem extraItem)
            return;

        // TODO: Implement GEMUS engine integration for running extras
        MessageBox.Show($"Extra launching not yet implemented.\n\nExtra: {extraItem.Extra.Name}\nFile: {extraItem.Extra.Filename}",
            "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information);

        await Task.CompletedTask;
    }

    private void exitToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        Application.Exit();
    }

    private void btnOptions_Click(object? sender, EventArgs e)
    {
        using var optionsForm = new OptionsForm();
        if (optionsForm.ShowDialog() == DialogResult.OK)
        {
            // Reload data with new settings
            _ = InitializeAsync();
        }
    }

    private static string Resolve(Dictionary<int, string> map, int? id)
        => id.HasValue && map.TryGetValue(id.Value, out var name) ? name : string.Empty;

    // Helper class for displaying extras in ListBox
    private class ExtraListItem
    {
        public ExtraEntity Extra { get; }

        public ExtraListItem(ExtraEntity extra)
        {
            Extra = extra;
        }

        public override string ToString()
        {
            return Extra.Name ?? "Unnamed Extra";
        }
    }
}
