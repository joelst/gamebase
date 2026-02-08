using GBSql.Data;
using GBSql.Models;
using Gemus;
using GBFuncs;

namespace GameBase;

public partial class Form1 : Form
{
    private readonly IGameBaseRepository _repository;
    private readonly GemusEngine? _gemusEngine;
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

    private readonly Dictionary<string, bool> _columnVisibility = new();

    private List<CustomViewForm.FilterCondition> _customFilters = new();
    private bool _customMatchAny;

    public event EventHandler? DataLoaded;

    public Form1(IGameBaseRepository repository)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        _gemusEngine = InitializeGemusEngine();
        InitializeComponent();
        Load += Form1_Load;
        editGameToolStripMenuItem.Click += editGameToolStripMenuItem_Click;
        manageExtrasToolStripMenuItem.Click += manageExtrasToolStripMenuItem_Click;
        picScreenshot.DoubleClick += picScreenshot_DoubleClick;
    }

    private GemusEngine? InitializeGemusEngine()
    {
        try
        {
            var funcs = new FunctionsHelper();
            var registry = new RegistryService();
            var logger = new GemusLogger();
            var unpacker = new ArchiveUnpacker();

            var engine = new GemusEngine(funcs, registry, logger, unpacker);

            // Set app context
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string basePathRoot = string.IsNullOrWhiteSpace(appData)
                ? AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                : appData;
            string basePath = Path.Combine(basePathRoot, "GameBase");

            // Ensure directory exists
            Directory.CreateDirectory(basePath);

            engine.SetAppContext(basePath + Path.DirectorySeparatorChar, string.Empty, "GameBase");
            engine.Init("GameBase");
            engine.SetSilentMode(false);

            return engine;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to initialize GEMUS engine: {ex.Message}\n\nGame launching will be disabled.",
                "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
    }

    private async void Form1_Load(object? sender, EventArgs e)
    {
        try
        {
            await InitializeAsync();
        }
        finally
        {
            DataLoaded?.Invoke(this, EventArgs.Empty);
        }
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

            UpdateStatus($"Loaded {_allGames.Count} games");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load games: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"Error loading games: {ex.Message}";
        }
    }

    private void UpdateGameGrid(int? selectedId = null)
    {
        dgvGames.DataSource = null;

        if (_filteredGames.Count == 0)
        {
            lblSearch.Text = "Games (0 of 0)";
            lblStatus.Text = "Filters: none";
            return;
        }

        // Create display data with selected columns
        var displayData = _filteredGames.Select(g => new
        {
            g.GA_Id,
            Fav = g.Fav == 1,
            Classic = g.Classic == 1,
            Adult = g.Adult == 1,
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

        ApplyColumnVisibility();

        if (selectedId.HasValue)
        {
            foreach (DataGridViewRow row in dgvGames.Rows)
            {
                if (row.Cells["GA_Id"].Value is int id && id == selectedId.Value)
                {
                    row.Selected = true;
                    dgvGames.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                    break;
                }
            }
        }

        lblSearch.Text = $"Games ({_filteredGames.Count} of {_allGames.Count})";
        UpdateStatus();
    }

    private void ApplyColumnVisibility()
    {
        foreach (DataGridViewColumn column in dgvGames.Columns)
        {
            if (string.Equals(column.Name, "GA_Id", StringComparison.OrdinalIgnoreCase))
            {
                column.Visible = false;
                continue;
            }

            if (_columnVisibility.TryGetValue(column.Name, out bool visible))
            {
                column.Visible = visible;
            }
        }
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

        // Load screenshot if available
        LoadScreenshot(game);

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

        UpdateStatus($"Selected: {game.Name}");
    }

    private void LoadScreenshot(GameEntity game)
    {
        if (picScreenshot.Image != null)
        {
            picScreenshot.Image.Dispose();
            picScreenshot.Image = null;
        }
        if (string.IsNullOrWhiteSpace(game.ScreenshotFilename)) return;
        var path = Path.Combine(GetScreenshotsPath(), game.ScreenshotFilename);
        if (File.Exists(path))
        {
            try
            {
                using var img = Image.FromFile(path);
                picScreenshot.Image = (Image)img.Clone();
            }
            catch { /* ignore */ }
        }
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
                    Contains(g.Name, searchText) ||
                    Contains(Resolve(_publishers, g.PU_Id), searchText) ||
                    Contains(Resolve(_programmers, g.PR_Id), searchText) ||
                    Contains(Resolve(_genres, g.GE_Id), searchText) ||
                    Contains(g.Comment, searchText)
                )
                .ToList();
        }

        UpdateGameGrid();
    }

    private void customViewsToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        using var dlg = new CustomViewForm(_customFilters, _customMatchAny);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _customFilters = dlg.Filters.ToList();
            _customMatchAny = dlg.MatchAny;
            ApplyCustomFilters();
        }
    }

    private void quickFilterToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        using var dlg = new QuickFilterForm();
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _customFilters = new List<CustomViewForm.FilterCondition>
            {
                new CustomViewForm.FilterCondition
                {
                    Field = dlg.SelectedField,
                    Operator = "Contains",
                    Value = dlg.FilterText
                }
            };
            _customMatchAny = false;
            ApplyCustomFilters();
            UpdateStatus($"Quick filter: {dlg.SelectedField} contains '{dlg.FilterText}'");
        }
    }

    private void clearFiltersToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        _customFilters.Clear();
        _customMatchAny = false;
        txtSearch.Text = string.Empty;
        _filteredGames = new List<GameEntity>(_allGames);
        UpdateGameGrid();
        UpdateStatus("Filters cleared");
    }

    private string GetScreenshotsPath()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var basePath = string.IsNullOrWhiteSpace(appData)
            ? AppContext.BaseDirectory
            : Path.Combine(appData, "GameBase", "Screenshots");
        Directory.CreateDirectory(basePath);
        return basePath;
    }

    private void ApplyCustomFilters()
    {
        if (_customFilters.Count == 0)
        {
            _filteredGames = new List<GameEntity>(_allGames);
            UpdateGameGrid();
            return;
        }

        Func<GameEntity, CustomViewForm.FilterCondition, bool> match = (g, f) =>
        {
            string val = f.Field switch
            {
                "Title" => g.Name ?? string.Empty,
                "Publisher" => Resolve(_publishers, g.PU_Id),
                "Programmer" => Resolve(_programmers, g.PR_Id),
                "Musician" => Resolve(_musicians, g.MU_Id),
                "Language" => Resolve(_languages, g.LA_Id),
                "Genre" => Resolve(_genres, g.GE_Id),
                "Year" => Resolve(_years, g.YE_Id),
                "Comment" => g.Comment ?? string.Empty,
                "Favorite" => g.Fav.ToString(),
                "Classic" => g.Classic.ToString(),
                "Adult" => g.Adult.ToString(),
                _ => string.Empty
            };

            string value = f.Value ?? string.Empty;
            return f.Operator switch
            {
                "Contains" => val.Contains(value, StringComparison.OrdinalIgnoreCase),
                "Equals" => string.Equals(val, value, StringComparison.OrdinalIgnoreCase),
                "NotEquals" => !string.Equals(val, value, StringComparison.OrdinalIgnoreCase),
                "StartsWith" => val.StartsWith(value, StringComparison.OrdinalIgnoreCase),
                "EndsWith" => val.EndsWith(value, StringComparison.OrdinalIgnoreCase),
                "GreaterThan" => int.TryParse(val, out var vi) && int.TryParse(value, out var v2) && vi > v2,
                "LessThan" => int.TryParse(val, out var vi2) && int.TryParse(value, out var v22) && vi2 < v22,
                "IsEmpty" => string.IsNullOrWhiteSpace(val),
                "IsNotEmpty" => !string.IsNullOrWhiteSpace(val),
                _ => false
            };
        };

        _filteredGames = _allGames.Where(g =>
        {
            var results = _customFilters.Select(f => match(g, f));
            return _customMatchAny ? results.Any(r => r) : results.All(r => r);
        }).ToList();

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

        using var dlg = new GameDetailsForm(_repository, game, _years, _publishers, _programmers, _genres,
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

        if (_gemusEngine == null)
        {
            MessageBox.Show("Game launching requires GEMUS emulator configuration.\n\n" +
                "To enable game launching:\n" +
                "1. Configure emulators in Options\n" +
                "2. Set game paths and scripts\n" +
                "3. Ensure Emulators.ini is configured\n\n" +
                $"Game: {game.Name}\nFile: {game.Filename}",
                "Configuration Required",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (string.IsNullOrWhiteSpace(game.Filename))
        {
            MessageBox.Show($"No filename specified for {game.Name}.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            lblStatus.Text = $"Launching {game.Name}...";

            string gamesPath = GetConfiguredPath("GamesPath", "Games");
            string extractPath = Path.Combine(GetAppDataBasePath(), "Temp");

            Directory.CreateDirectory(gamesPath);
            Directory.CreateDirectory(extractPath);

            // Build full game path
            string gameFile = Path.Combine(gamesPath, game.Filename ?? string.Empty);

            if (!File.Exists(gameFile))
            {
                MessageBox.Show($"Game file not found:\n{gameFile}\n\n" +
                    $"Please ensure game files are in:\n{gamesPath}",
                    "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lblStatus.Text = $"Game file not found: {game.Name}";
                return;
            }

            string defaultFile = string.Empty;

            // Run emulator (index 0 = first emulator in Emulators.ini)
            await Task.Run(() =>
            {
                string result = _gemusEngine.RunEmulator(
                    emuIdx: 0,
                    gamePathFile: gameFile,
                    imageIndex: 0,
                    defaultFile: ref defaultFile,
                    extractDestFolder: extractPath,
                    fromDb: true
                );

                if (!string.IsNullOrWhiteSpace(result))
                {
                    Invoke(() =>
                    {
                        MessageBox.Show($"Error launching game:\n{result}", "Launch Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            });

            lblStatus.Text = $"Launched {game.Name}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to launch game:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"Failed to launch {game.Name}";
        }
    }

    private async void lstExtras_DoubleClick(object? sender, EventArgs e)
    {
        if (lstExtras.SelectedItem is not ExtraListItem extraItem)
            return;

        if (_gemusEngine == null)
        {
            MessageBox.Show("Extra launching requires GEMUS emulator configuration.\n\n" +
                "To enable extra launching:\n" +
                "1. Configure emulators and viewers in Options\n" +
                "2. Set paths and scripts\n" +
                "3. Ensure Emulators.ini is configured\n\n" +
                $"Extra: {extraItem.Extra.Name}\nFile: {extraItem.Extra.Filename}",
                "Configuration Required",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (string.IsNullOrWhiteSpace(extraItem.Extra.Filename))
        {
            MessageBox.Show($"No filename specified for {extraItem.Extra.Name}.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            lblStatus.Text = $"Launching {extraItem.Extra.Name}...";

            string extrasPath = GetConfiguredPath("ExtrasPath", "Extras");
            string extractPath = Path.Combine(GetAppDataBasePath(), "Temp");

            Directory.CreateDirectory(extrasPath);
            Directory.CreateDirectory(extractPath);

            string extraFile = Path.Combine(extrasPath, extraItem.Extra.Filename ?? string.Empty);

            if (!File.Exists(extraFile))
            {
                MessageBox.Show($"Extra file not found:\n{extraFile}\n\n" +
                    $"Please ensure extra files are in:\n{extrasPath}",
                    "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lblStatus.Text = $"Extra file not found: {extraItem.Extra.Name}";
                return;
            }

            string defaultFile = string.Empty;

            await Task.Run(() =>
            {
                string result = _gemusEngine.RunEmulator(
                    emuIdx: 0,
                    gamePathFile: extraFile,
                    imageIndex: 0,
                    defaultFile: ref defaultFile,
                    extractDestFolder: extractPath,
                    fromDb: true
                );

                if (!string.IsNullOrWhiteSpace(result))
                {
                    Invoke(() =>
                    {
                        MessageBox.Show($"Error launching extra:\n{result}", "Launch Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            });

            lblStatus.Text = $"Launched {extraItem.Extra.Name}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to launch extra:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"Failed to launch {extraItem.Extra.Name}";
        }
    }

    private void exitToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        Application.Exit();
    }

    private void importAccessDatabaseToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        using var importForm = new ImportDatabaseForm();
        if (importForm.ShowDialog() == DialogResult.OK)
        {
            // Import successful - ask if user wants to switch to the new database
            var result = MessageBox.Show(
                "Database imported successfully!\n\n" +
                "Would you like to restart GameBase and use the new SQLite database?\n\n" +
                "(You will need to configure the database path in Options)",
                "Import Complete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                MessageBox.Show(
                    "Please restart GameBase and configure the database path in Options.\n\n" +
                    $"New database location:\n{importForm.SqliteDatabasePath}",
                    "Restart Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                Application.Exit();
            }
        }
    }

    private void exportToAccessDatabaseToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        // Get current database path from repository context
        string currentDbPath = string.Empty;
        try
        {
            var contextType = _repository.GetType();
            var contextField = contextType.GetField("_context", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (contextField != null)
            {
                var context = contextField.GetValue(_repository);
                if (context != null)
                {
                    var dbPathProp = context.GetType().GetProperty("DatabasePath");
                    if (dbPathProp != null)
                    {
                        currentDbPath = dbPathProp.GetValue(context)?.ToString() ?? string.Empty;
                    }
                }
            }
        }
        catch
        {
            // If we can't get the path via reflection, user will need to browse for it
        }

        using var exportForm = new ExportDatabaseForm(currentDbPath);
        if (exportForm.ShowDialog() == DialogResult.OK)
        {
            MessageBox.Show(
                "Database exported successfully!\n\n" +
                "The Access database is compatible with VB6 GameBase versions.",
                "Export Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }

    private void columnsToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        if (dgvGames.Columns.Count == 0)
            return;

        var columns = dgvGames.Columns.Cast<DataGridViewColumn>()
            .Where(c => !string.Equals(c.Name, "GA_Id", StringComparison.OrdinalIgnoreCase))
            .Select(c => (c.Name, c.Visible))
            .ToList();

        using var dlg = new ColumnsForm(columns);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _columnVisibility.Clear();
            foreach (var entry in dlg.ColumnVisibility)
            {
                _columnVisibility[entry.Key] = entry.Value;
            }

            ApplyColumnVisibility();
        }
    }

    private async void batchEditToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        if (dgvGames.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select games to edit.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selectedIds = dgvGames.SelectedRows
            .Cast<DataGridViewRow>()
            .Select(r => (int)r.Cells["GA_Id"].Value)
            .ToHashSet();

        var selectedGames = _filteredGames
            .Where(g => selectedIds.Contains(g.GA_Id))
            .ToList();

        if (selectedGames.Count == 0)
            return;

        using var dlg = new BatchEditForm(_publishers, _genres);
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;

        foreach (var game in selectedGames)
        {
            if (dlg.SelectedPublisherId.HasValue)
                game.PU_Id = dlg.SelectedPublisherId;
            if (dlg.SelectedGenreId.HasValue)
                game.GE_Id = dlg.SelectedGenreId;
            if (dlg.SelectedFavorite.HasValue)
                game.Fav = dlg.SelectedFavorite.Value;
            if (dlg.SelectedClassic.HasValue)
                game.Classic = dlg.SelectedClassic.Value;
            if (dlg.SelectedAdult.HasValue)
                game.Adult = dlg.SelectedAdult.Value;

            await _repository.UpsertGameAsync(game);
        }

        UpdateGameGrid();
        UpdateStatus($"Batch updated {selectedGames.Count} games");
    }

    private async void verifyFilesToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        UseWaitCursor = true;
        try
        {
            var issues = await Task.Run(async () => await CollectMissingFilesAsync());
            using var dlg = new VerifyFilesForm(issues);
            dlg.ShowDialog(this);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private async Task<List<VerifyFilesForm.FileIssue>> CollectMissingFilesAsync()
    {
        var issues = new List<VerifyFilesForm.FileIssue>();
        string gamesPath = GetConfiguredPath("GamesPath", "Games");
        string extrasPath = GetConfiguredPath("ExtrasPath", "Extras");

        foreach (var game in _allGames)
        {
            if (string.IsNullOrWhiteSpace(game.Filename))
            {
                issues.Add(new VerifyFilesForm.FileIssue("Game", game.Name, string.Empty, "(missing filename)"));
            }
            else
            {
                string gameFile = Path.Combine(gamesPath, game.Filename);
                if (!File.Exists(gameFile))
                {
                    issues.Add(new VerifyFilesForm.FileIssue("Game", game.Name, string.Empty, gameFile));
                }
            }

            var extras = await _repository.GetExtrasForGameAsync(game.GA_Id);
            foreach (var extra in extras)
            {
                if (string.IsNullOrWhiteSpace(extra.Filename))
                {
                    issues.Add(new VerifyFilesForm.FileIssue("Extra", game.Name, extra.Name, "(missing filename)"));
                    continue;
                }

                string extraFile = Path.Combine(extrasPath, extra.Filename);
                if (!File.Exists(extraFile))
                {
                    issues.Add(new VerifyFilesForm.FileIssue("Extra", game.Name, extra.Name, extraFile));
                }
            }
        }

        return issues;
    }

    private string GetAppDataBasePath()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string basePathRoot = string.IsNullOrWhiteSpace(appData)
            ? AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            : appData;
        return Path.Combine(basePathRoot, "GameBase");
    }

    private string GetDocumentsBasePath()
    {
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return string.IsNullOrWhiteSpace(documentsPath)
            ? GetAppDataBasePath()
            : Path.Combine(documentsPath, "GameBase");
    }

    private string GetConfiguredPath(string settingKey, string defaultFolder)
    {
        string appBase = GetAppDataBasePath();
        string gemusIni = Path.Combine(appBase, "Gemus.ini");
        string docsBase = GetDocumentsBasePath();
        string defaultPath = Path.Combine(docsBase, defaultFolder);
        return GemusIniHelper.GetIniMessage("Settings", settingKey, defaultPath, gemusIni);
    }

    private void aboutToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        using var aboutForm = new AboutForm();
        aboutForm.ShowDialog(this);
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

    private GameEntity? GetSelectedGame()
    {
        if (dgvGames.SelectedRows.Count == 0) return null;
        var selectedId = (int)dgvGames.SelectedRows[0].Cells["GA_Id"].Value;
        return _filteredGames.FirstOrDefault(g => g.GA_Id == selectedId);
    }

    private async void toggleFavoriteToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        var game = GetSelectedGame();
        if (game == null) return;
        game.Fav = game.Fav == 1 ? 0 : 1;
        await _repository.UpsertGameAsync(game);
        UpdateGameGrid(game.GA_Id);
        DisplayGameDetails(game);
    }

    private async void toggleClassicToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        var game = GetSelectedGame();
        if (game == null) return;
        game.Classic = game.Classic == 1 ? 0 : 1;
        await _repository.UpsertGameAsync(game);
        UpdateGameGrid(game.GA_Id);
        DisplayGameDetails(game);
    }

    private async void toggleAdultToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        var game = GetSelectedGame();
        if (game == null) return;
        game.Adult = game.Adult == 1 ? 0 : 1;
        await _repository.UpsertGameAsync(game);
        UpdateGameGrid(game.GA_Id);
        DisplayGameDetails(game);
    }

    private async void selectScreenshotToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        var game = GetSelectedGame();
        if (game == null) return;

        using var ofd = new OpenFileDialog
        {
            Title = "Select Screenshot",
            Filter = "Image Files|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.tif;*.tiff|All Files|*.*",
            Multiselect = false
        };

        if (ofd.ShowDialog(this) == DialogResult.OK)
        {
            try
            {
                var ext = Path.GetExtension(ofd.FileName);
                var fileName = $"GA_{game.GA_Id}{ext}";
                var destPath = Path.Combine(GetScreenshotsPath(), fileName);
                File.Copy(ofd.FileName, destPath, overwrite: true);
                game.ScreenshotFilename = fileName;
                await _repository.UpsertGameAsync(game);
                LoadScreenshot(game);
                UpdateGameGrid(game.GA_Id);
                UpdateStatus($"Screenshot set for {game.Name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to set screenshot: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"Failed to set screenshot: {ex.Message}");
            }
        }
    }

    private async void clearScreenshotToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        var game = GetSelectedGame();
        if (game == null) return;
        game.ScreenshotFilename = null;
        await _repository.UpsertGameAsync(game);
        LoadScreenshot(game);
        UpdateGameGrid(game.GA_Id);
        UpdateStatus($"Screenshot cleared for {game.Name}");
    }

    private void dgvGames_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
        {
            dgvGames.ClearSelection();
            dgvGames.Rows[e.RowIndex].Selected = true;
            dgvGames.CurrentCell = dgvGames.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }
    }

    private void UpdateStatus(string? message = null)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(message)) parts.Add(message);

        if (!string.IsNullOrWhiteSpace(txtSearch.Text))
            parts.Add($"Search: \"{txtSearch.Text.Trim()}\"");

        if (_customFilters.Count > 0)
        {
            var filtersDesc = string.Join("; ", _customFilters.Select(f => $"{f.Field} {f.Operator} \"{f.Value}\""));
            parts.Add(_customMatchAny ? $"Filters (ANY): {filtersDesc}" : $"Filters (ALL): {filtersDesc}");
        }
        else
        {
            parts.Add("Filters: none");
        }

        lblStatus.Text = string.Join(" | ", parts);
    }

    private static bool Contains(string? target, string text)
        => target?.Contains(text, StringComparison.OrdinalIgnoreCase) ?? false;

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
