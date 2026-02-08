using GBSql.Data;
using GBSql.Models;

namespace GameBase;

public partial class GameInfoPropsForm : Form
{
  private readonly IGameBaseRepository _repository;
  private readonly GameEntity _game;
  private readonly Dictionary<int, string> _years;
  private readonly Dictionary<int, string> _publishers;
  private readonly Dictionary<int, string> _programmers;
  private readonly Dictionary<int, string> _musicians;
  private readonly Dictionary<int, string> _languages;
  private readonly Dictionary<int, string> _genres;
  private readonly Dictionary<int, string> _pgenres;
  private readonly Dictionary<int, string> _difficulty;
  private readonly Dictionary<int, string> _crackers;
  private readonly Dictionary<int, string> _artists;
  private readonly Dictionary<int, string> _developers;
  private readonly Dictionary<int, string> _licenses;
  private readonly Dictionary<int, string> _rarities;

  public GameInfoPropsForm(IGameBaseRepository repository, GameEntity game,
      Dictionary<int, string> years,
      Dictionary<int, string> publishers,
      Dictionary<int, string> programmers,
      Dictionary<int, string> musicians,
      Dictionary<int, string> languages,
      Dictionary<int, string> genres,
      Dictionary<int, string> pgenres,
      Dictionary<int, string> difficulty,
      Dictionary<int, string> crackers,
      Dictionary<int, string> artists,
      Dictionary<int, string> developers,
      Dictionary<int, string> licenses,
      Dictionary<int, string> rarities)
  {
    _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    _game = game;
    _years = years;
    _publishers = publishers;
    _programmers = programmers;
    _musicians = musicians;
    _languages = languages;
    _genres = genres;
    _pgenres = pgenres;
    _difficulty = difficulty;
    _crackers = crackers;
    _artists = artists;
    _developers = developers;
    _licenses = licenses;
    _rarities = rarities;

    InitializeComponent();
    Load += GameInfoPropsForm_Load;
  }

  private void GameInfoPropsForm_Load(object? sender, EventArgs e)
  {
    txtName.Text = _game.Name;
    txtCCode.Text = _game.CCode;
    txtYear.Text = Resolve(_years, _game.YE_Id);
    txtPublisher.Text = Resolve(_publishers, _game.PU_Id);
    txtProgrammer.Text = Resolve(_programmers, _game.PR_Id);
    txtMusician.Text = Resolve(_musicians, _game.MU_Id);
    txtLanguage.Text = Resolve(_languages, _game.LA_Id);
    txtGenre.Text = Resolve(_genres, _game.GE_Id);
    txtPGenre.Text = Resolve(_pgenres, _game.PG_Id);
    txtDifficulty.Text = Resolve(_difficulty, _game.DI_Id);
    txtCracker.Text = Resolve(_crackers, _game.CR_Id);
    txtArtist.Text = Resolve(_artists, _game.AR_Id);
    txtDeveloper.Text = Resolve(_developers, _game.DE_Id);
    txtLicense.Text = Resolve(_licenses, _game.LI_Id);
    txtRarity.Text = Resolve(_rarities, _game.RA_Id);
    txtPlayers.Text = (_game.PlayersFrom == _game.PlayersTo)
        ? _game.PlayersFrom.ToString()
        : $"{_game.PlayersFrom}-{_game.PlayersTo}";
    txtReviewRating.Text = _game.ReviewRating?.ToString() ?? string.Empty;
    txtFilename.Text = _game.Filename;
    rtbComment.Text = _game.Comment ?? string.Empty;
  }

  private static string Resolve(Dictionary<int, string> map, int? id)
      => id.HasValue && map.TryGetValue(id.Value, out var name) ? name : string.Empty;

  private async void btnSave_Click(object? sender, EventArgs e)
  {
    _game.Comment = rtbComment.Text;
    _game.ReviewRating = int.TryParse(txtReviewRating.Text, out var rr) ? rr : _game.ReviewRating;
    await _repository.UpsertGameAsync(_game);
    MessageBox.Show(this, "Game saved.", "Game", MessageBoxButtons.OK, MessageBoxIcon.Information);
  }

  private void btnClose_Click(object? sender, EventArgs e)
  {
    Close();
  }
}
