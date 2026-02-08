namespace GameBase;

partial class GameInfoPropsForm
{
  private System.ComponentModel.IContainer components = null;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  #region Windows Form Designer generated code

  private TableLayoutPanel tableLayoutPanel1;
  private Button btnClose;
  private Button btnSave;
  private Label lblName;
  private TextBox txtName;
  private Label lblCCode;
  private TextBox txtCCode;
  private Label lblYear;
  private TextBox txtYear;
  private Label lblPublisher;
  private TextBox txtPublisher;
  private Label lblProgrammer;
  private TextBox txtProgrammer;
  private Label lblMusician;
  private TextBox txtMusician;
  private Label lblLanguage;
  private TextBox txtLanguage;
  private Label lblGenre;
  private TextBox txtGenre;
  private Label lblPGenre;
  private TextBox txtPGenre;
  private Label lblDifficulty;
  private TextBox txtDifficulty;
  private Label lblCracker;
  private TextBox txtCracker;
  private Label lblArtist;
  private TextBox txtArtist;
  private Label lblDeveloper;
  private TextBox txtDeveloper;
  private Label lblLicense;
  private TextBox txtLicense;
  private Label lblRarity;
  private TextBox txtRarity;
  private Label lblPlayers;
  private TextBox txtPlayers;
  private Label lblReviewRating;
  private TextBox txtReviewRating;
  private Label lblFilename;
  private TextBox txtFilename;
  private Label lblComment;
  private RichTextBox rtbComment;

  private void InitializeComponent()
  {
    tableLayoutPanel1 = new TableLayoutPanel();
    lblName = new Label();
    txtName = new TextBox();
    lblCCode = new Label();
    txtCCode = new TextBox();
    lblYear = new Label();
    txtYear = new TextBox();
    lblPublisher = new Label();
    txtPublisher = new TextBox();
    lblProgrammer = new Label();
    txtProgrammer = new TextBox();
    lblMusician = new Label();
    txtMusician = new TextBox();
    lblLanguage = new Label();
    txtLanguage = new TextBox();
    lblGenre = new Label();
    txtGenre = new TextBox();
    lblPGenre = new Label();
    txtPGenre = new TextBox();
    lblDifficulty = new Label();
    txtDifficulty = new TextBox();
    lblCracker = new Label();
    txtCracker = new TextBox();
    lblArtist = new Label();
    txtArtist = new TextBox();
    lblDeveloper = new Label();
    txtDeveloper = new TextBox();
    lblLicense = new Label();
    txtLicense = new TextBox();
    lblRarity = new Label();
    txtRarity = new TextBox();
    lblPlayers = new Label();
    txtPlayers = new TextBox();
    lblReviewRating = new Label();
    txtReviewRating = new TextBox();
    lblFilename = new Label();
    txtFilename = new TextBox();
    lblComment = new Label();
    rtbComment = new RichTextBox();
    btnClose = new Button();
    tableLayoutPanel1.SuspendLayout();
    SuspendLayout();
    //
    // tableLayoutPanel1
    //
    tableLayoutPanel1.ColumnCount = 2;
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.Controls.Add(lblName, 0, 0);
    tableLayoutPanel1.Controls.Add(txtName, 1, 0);
    tableLayoutPanel1.Controls.Add(lblCCode, 0, 1);
    tableLayoutPanel1.Controls.Add(txtCCode, 1, 1);
    tableLayoutPanel1.Controls.Add(lblYear, 0, 2);
    tableLayoutPanel1.Controls.Add(txtYear, 1, 2);
    tableLayoutPanel1.Controls.Add(lblPublisher, 0, 3);
    tableLayoutPanel1.Controls.Add(txtPublisher, 1, 3);
    tableLayoutPanel1.Controls.Add(lblProgrammer, 0, 4);
    tableLayoutPanel1.Controls.Add(txtProgrammer, 1, 4);
    tableLayoutPanel1.Controls.Add(lblMusician, 0, 5);
    tableLayoutPanel1.Controls.Add(txtMusician, 1, 5);
    tableLayoutPanel1.Controls.Add(lblLanguage, 0, 6);
    tableLayoutPanel1.Controls.Add(txtLanguage, 1, 6);
    tableLayoutPanel1.Controls.Add(lblGenre, 0, 7);
    tableLayoutPanel1.Controls.Add(txtGenre, 1, 7);
    tableLayoutPanel1.Controls.Add(lblPGenre, 0, 8);
    tableLayoutPanel1.Controls.Add(txtPGenre, 1, 8);
    tableLayoutPanel1.Controls.Add(lblDifficulty, 0, 9);
    tableLayoutPanel1.Controls.Add(txtDifficulty, 1, 9);
    tableLayoutPanel1.Controls.Add(lblCracker, 0, 10);
    tableLayoutPanel1.Controls.Add(txtCracker, 1, 10);
    tableLayoutPanel1.Controls.Add(lblArtist, 0, 11);
    tableLayoutPanel1.Controls.Add(txtArtist, 1, 11);
    tableLayoutPanel1.Controls.Add(lblDeveloper, 0, 12);
    tableLayoutPanel1.Controls.Add(txtDeveloper, 1, 12);
    tableLayoutPanel1.Controls.Add(lblLicense, 0, 13);
    tableLayoutPanel1.Controls.Add(txtLicense, 1, 13);
    tableLayoutPanel1.Controls.Add(lblRarity, 0, 14);
    tableLayoutPanel1.Controls.Add(txtRarity, 1, 14);
    tableLayoutPanel1.Controls.Add(lblPlayers, 0, 15);
    tableLayoutPanel1.Controls.Add(txtPlayers, 1, 15);
    tableLayoutPanel1.Controls.Add(lblReviewRating, 0, 16);
    tableLayoutPanel1.Controls.Add(txtReviewRating, 1, 16);
    tableLayoutPanel1.Controls.Add(lblFilename, 0, 17);
    tableLayoutPanel1.Controls.Add(txtFilename, 1, 17);
    tableLayoutPanel1.Controls.Add(lblComment, 0, 18);
    tableLayoutPanel1.Controls.Add(rtbComment, 1, 18);
    tableLayoutPanel1.Controls.Add(btnClose, 1, 19);
    tableLayoutPanel1.Dock = DockStyle.Fill;
    tableLayoutPanel1.Location = new Point(0, 0);
    tableLayoutPanel1.Name = "tableLayoutPanel1";
    tableLayoutPanel1.RowCount = 20;
    for (int i = 0; i < 18; i++)
      tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.Size = new Size(600, 600);
    tableLayoutPanel1.TabIndex = 0;
    //
    // labels and textboxes
    //
    ConfigureLabel(lblName, "Name:");
    ConfigureTextBox(txtName);
    ConfigureLabel(lblCCode, "CCode:");
    ConfigureTextBox(txtCCode);
    ConfigureLabel(lblYear, "Year:");
    ConfigureTextBox(txtYear);
    ConfigureLabel(lblPublisher, "Publisher:");
    ConfigureTextBox(txtPublisher);
    ConfigureLabel(lblProgrammer, "Programmer:");
    ConfigureTextBox(txtProgrammer);
    ConfigureLabel(lblMusician, "Musician:");
    ConfigureTextBox(txtMusician);
    ConfigureLabel(lblLanguage, "Language:");
    ConfigureTextBox(txtLanguage);
    ConfigureLabel(lblGenre, "Genre:");
    ConfigureTextBox(txtGenre);
    ConfigureLabel(lblPGenre, "PGenre:");
    ConfigureTextBox(txtPGenre);
    ConfigureLabel(lblDifficulty, "Difficulty:");
    ConfigureTextBox(txtDifficulty);
    ConfigureLabel(lblCracker, "Cracker:");
    ConfigureTextBox(txtCracker);
    ConfigureLabel(lblArtist, "Artist:");
    ConfigureTextBox(txtArtist);
    ConfigureLabel(lblDeveloper, "Developer:");
    ConfigureTextBox(txtDeveloper);
    ConfigureLabel(lblLicense, "License:");
    ConfigureTextBox(txtLicense);
    ConfigureLabel(lblRarity, "Rarity:");
    ConfigureTextBox(txtRarity);
    ConfigureLabel(lblPlayers, "Players:");
    ConfigureTextBox(txtPlayers);
    ConfigureLabel(lblReviewRating, "Review Rating:");
    ConfigureTextBox(txtReviewRating);
    ConfigureLabel(lblFilename, "Filename:");
    ConfigureTextBox(txtFilename);
    ConfigureLabel(lblComment, "Comment:");
    rtbComment.Dock = DockStyle.Fill;
    rtbComment.ReadOnly = false;
    //
    // btnClose
    //
    btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
    btnClose.Location = new Point(497, 562);
    btnClose.Name = "btnClose";
    btnClose.Size = new Size(100, 23);
    btnClose.TabIndex = 19;
    btnClose.Text = "Close";
    btnClose.UseVisualStyleBackColor = true;
    btnClose.Click += btnClose_Click;
    //
    // GameInfoPropsForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(600, 600);
    Controls.Add(tableLayoutPanel1);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "GameInfoPropsForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Game Info";
    tableLayoutPanel1.ResumeLayout(false);
    tableLayoutPanel1.PerformLayout();
    ResumeLayout(false);

    void ConfigureLabel(Label lbl, string text)
    {
      lbl.AutoSize = true;
      lbl.Dock = DockStyle.Fill;
      lbl.Text = text;
      lbl.TextAlign = ContentAlignment.MiddleLeft;
    }

    void ConfigureTextBox(TextBox txt)
    {
      txt.Dock = DockStyle.Fill;
      txt.ReadOnly = true;
    }
  }

  #endregion
}
