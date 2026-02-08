using GBSql.Models;

namespace GameBase;

public partial class GameDetailsForm : Form
{
    private readonly GameEntity _game;
    private readonly Dictionary<int, string> _years;
    private readonly Dictionary<int, string> _publishers;
    private readonly Dictionary<int, string> _programmers;
    private readonly Dictionary<int, string> _genres;

    public GameDetailsForm(GameEntity game,
        Dictionary<int, string> years,
        Dictionary<int, string> publishers,
        Dictionary<int, string> programmers,
        Dictionary<int, string> genres)
    {
        _game = game ?? throw new ArgumentNullException(nameof(game));
        _years = years ?? throw new ArgumentNullException(nameof(years));
        _publishers = publishers ?? throw new ArgumentNullException(nameof(publishers));
        _programmers = programmers ?? throw new ArgumentNullException(nameof(programmers));
        _genres = genres ?? throw new ArgumentNullException(nameof(genres));

        InitializeComponent();
        Load += GameDetailsForm_Load;
    }

    private void GameDetailsForm_Load(object? sender, EventArgs e)
    {
        txtName.Text = _game.Name;
        FillCombo(cboYear, _years, _game.YE_Id);
        FillCombo(cboPublisher, _publishers, _game.PU_Id);
        FillCombo(cboProgrammer, _programmers, _game.PR_Id);
        FillCombo(cboGenre, _genres, _game.GE_Id);
        rtbComment.Text = _game.Comment ?? string.Empty;
    }

    private static void FillCombo(ComboBox combo, Dictionary<int, string> map, int? id)
    {
        combo.BeginUpdate();
        combo.Items.Clear();

        foreach (var kv in map.OrderBy(k => k.Value))
        {
            combo.Items.Add(new ComboItem(kv.Key, kv.Value));
        }

        combo.EndUpdate();

        if (id.HasValue && map.TryGetValue(id.Value, out var name))
        {
            var idx = combo.Items.Cast<ComboItem>().ToList().FindIndex(i => i.Id == id.Value);
            if (idx >= 0)
                combo.SelectedIndex = idx;
        }
    }

    private record ComboItem(int Id, string Name)
    {
        public override string ToString() => Name;
    }

    private void btnClose_Click(object? sender, EventArgs e)
    {
        DialogResult = DialogResult.OK;
        Close();
    }
}
