using GBSql.Data;
using GBSql.Models;

namespace GameBase;

public partial class GameDetailsForm : Form
{
  private readonly GameEntity _game;
  private readonly Dictionary<int, string> _years;
  private readonly Dictionary<int, string> _publishers;
  private readonly Dictionary<int, string> _programmers;
  private readonly Dictionary<int, string> _genres;
  private readonly Dictionary<int, string> _musicians;
  private readonly Dictionary<int, string> _languages;
  private readonly Dictionary<int, string> _artists;
  private readonly Dictionary<int, string> _developers;
  private readonly Dictionary<int, string> _licenses;
  private readonly Dictionary<int, string> _rarities;
  private readonly Dictionary<int, string> _pgenres;
  private readonly Dictionary<int, string> _difficulty;
  private readonly Dictionary<int, string> _crackers;

  public GameDetailsForm(IGameBaseRepository repository, GameEntity game,
      Dictionary<int, string> years,
      Dictionary<int, string> publishers,
      Dictionary<int, string> programmers,
      Dictionary<int, string> genres,
      Dictionary<int, string> musicians,
      Dictionary<int, string> languages,
      Dictionary<int, string> artists,
      Dictionary<int, string> developers,
      Dictionary<int, string> licenses,
      Dictionary<int, string> rarities,
      Dictionary<int, string> pgenres,
      Dictionary<int, string> difficulty,
      Dictionary<int, string> crackers)
  {
    _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    _game = game ?? throw new ArgumentNullException(nameof(game));
    _years = years ?? throw new ArgumentNullException(nameof(years));
    _publishers = publishers ?? throw new ArgumentNullException(nameof(publishers));
    _programmers = programmers ?? throw new ArgumentNullException(nameof(programmers));
    _genres = genres ?? throw new ArgumentNullException(nameof(genres));
    _musicians = musicians ?? throw new ArgumentNullException(nameof(musicians));
    _languages = languages ?? throw new ArgumentNullException(nameof(languages));
    _artists = artists ?? throw new ArgumentNullException(nameof(artists));
    _developers = developers ?? throw new ArgumentNullException(nameof(developers));
    _licenses = licenses ?? throw new ArgumentNullException(nameof(licenses));
    _rarities = rarities ?? throw new ArgumentNullException(nameof(rarities));
    _pgenres = pgenres ?? throw new ArgumentNullException(nameof(pgenres));
    _difficulty = difficulty ?? throw new ArgumentNullException(nameof(difficulty));
    _crackers = crackers ?? throw new ArgumentNullException(nameof(crackers));

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

  private readonly IGameBaseRepository _repository;

  private record ComboItem(int Id, string Name)
  {
    public override string ToString() => Name;
  }

  private void btnAdvanced_Click(object? sender, EventArgs e)
  {
    using var dlg = new GameInfoPropsForm(_repository, _game, _years, _publishers, _programmers, _musicians, _languages,
        _genres, _pgenres, _difficulty, _crackers, _artists, _developers, _licenses, _rarities);
    dlg.ShowDialog(this);
  }

  private void btnClose_Click(object? sender, EventArgs e)
  {
    DialogResult = DialogResult.OK;
    Close();
  }
}
