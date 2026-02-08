namespace GameBase;

public partial class VerifyFilesForm : Form
{
  public sealed class FileIssue
  {
    public string Type { get; }
    public string GameName { get; }
    public string ExtraName { get; }
    public string FilePath { get; }

    public FileIssue(string type, string gameName, string extraName, string filePath)
    {
      Type = type;
      GameName = gameName;
      ExtraName = extraName;
      FilePath = filePath;
    }
  }

  public VerifyFilesForm(IEnumerable<FileIssue> issues)
  {
    InitializeComponent();
    PopulateIssues(issues);
  }

  private void PopulateIssues(IEnumerable<FileIssue> issues)
  {
    listIssues.Items.Clear();
    int count = 0;
    foreach (var issue in issues)
    {
      var item = new ListViewItem(issue.Type);
      item.SubItems.Add(issue.GameName);
      item.SubItems.Add(issue.ExtraName);
      item.SubItems.Add(issue.FilePath);
      listIssues.Items.Add(item);
      count++;
    }

    lblSummary.Text = count == 0
        ? "No missing files found."
        : $"Missing files: {count}";
  }

  private void btnClose_Click(object? sender, EventArgs e)
  {
    Close();
  }
}
