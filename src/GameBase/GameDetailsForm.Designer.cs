namespace GameBase;

partial class GameDetailsForm
{
  /// <summary>
  /// Required designer variable.
  /// </summary>
  private System.ComponentModel.IContainer components = null;

  /// <summary>
  /// Clean up any resources being used.
  /// </summary>
  /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  #region Windows Form Designer generated code

  private Label lblName;
  private TextBox txtName;
  private Label lblYear;
  private ComboBox cboYear;
  private Label lblPublisher;
  private ComboBox cboPublisher;
  private Label lblProgrammer;
  private ComboBox cboProgrammer;
  private Label lblGenre;
  private ComboBox cboGenre;
  private Label lblComment;
  private RichTextBox rtbComment;
  private Button btnClose;
  private TableLayoutPanel tableLayoutPanel1;

  /// <summary>
  /// Required method for Designer support - do not modify
  /// the contents of this method with the code editor.
  /// </summary>
  private void InitializeComponent()
  {
    lblName = new Label();
    txtName = new TextBox();
    lblYear = new Label();
    cboYear = new ComboBox();
    lblPublisher = new Label();
    cboPublisher = new ComboBox();
    lblProgrammer = new Label();
    cboProgrammer = new ComboBox();
    lblGenre = new Label();
    cboGenre = new ComboBox();
    lblComment = new Label();
    rtbComment = new RichTextBox();
    btnClose = new Button();
    tableLayoutPanel1 = new TableLayoutPanel();
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
    tableLayoutPanel1.Controls.Add(lblYear, 0, 1);
    tableLayoutPanel1.Controls.Add(cboYear, 1, 1);
    tableLayoutPanel1.Controls.Add(lblPublisher, 0, 2);
    tableLayoutPanel1.Controls.Add(cboPublisher, 1, 2);
    tableLayoutPanel1.Controls.Add(lblProgrammer, 0, 3);
    tableLayoutPanel1.Controls.Add(cboProgrammer, 1, 3);
    tableLayoutPanel1.Controls.Add(lblGenre, 0, 4);
    tableLayoutPanel1.Controls.Add(cboGenre, 1, 4);
    tableLayoutPanel1.Controls.Add(lblComment, 0, 5);
    tableLayoutPanel1.Controls.Add(rtbComment, 1, 5);
    tableLayoutPanel1.Controls.Add(btnClose, 1, 6);
    tableLayoutPanel1.Dock = DockStyle.Fill;
    tableLayoutPanel1.Location = new Point(0, 0);
    tableLayoutPanel1.Name = "tableLayoutPanel1";
    tableLayoutPanel1.RowCount = 7;
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.Size = new Size(500, 400);
    tableLayoutPanel1.TabIndex = 0;
    //
    // lblName
    //
    lblName.AutoSize = true;
    lblName.Dock = DockStyle.Fill;
    lblName.Location = new Point(3, 0);
    lblName.Name = "lblName";
    lblName.Size = new Size(114, 30);
    lblName.TabIndex = 0;
    lblName.Text = "Name:";
    lblName.TextAlign = ContentAlignment.MiddleLeft;
    //
    // txtName
    //
    txtName.Dock = DockStyle.Fill;
    txtName.Location = new Point(123, 3);
    txtName.Name = "txtName";
    txtName.ReadOnly = true;
    txtName.Size = new Size(374, 23);
    txtName.TabIndex = 1;
    //
    // lblYear
    //
    lblYear.AutoSize = true;
    lblYear.Dock = DockStyle.Fill;
    lblYear.Location = new Point(3, 30);
    lblYear.Name = "lblYear";
    lblYear.Size = new Size(114, 30);
    lblYear.TabIndex = 2;
    lblYear.Text = "Year:";
    lblYear.TextAlign = ContentAlignment.MiddleLeft;
    //
    // cboYear
    //
    cboYear.Dock = DockStyle.Fill;
    cboYear.DropDownStyle = ComboBoxStyle.DropDownList;
    cboYear.Enabled = false;
    cboYear.FormattingEnabled = true;
    cboYear.Location = new Point(123, 33);
    cboYear.Name = "cboYear";
    cboYear.Size = new Size(374, 23);
    cboYear.TabIndex = 3;
    //
    // lblPublisher
    //
    lblPublisher.AutoSize = true;
    lblPublisher.Dock = DockStyle.Fill;
    lblPublisher.Location = new Point(3, 60);
    lblPublisher.Name = "lblPublisher";
    lblPublisher.Size = new Size(114, 30);
    lblPublisher.TabIndex = 4;
    lblPublisher.Text = "Publisher:";
    lblPublisher.TextAlign = ContentAlignment.MiddleLeft;
    //
    // cboPublisher
    //
    cboPublisher.Dock = DockStyle.Fill;
    cboPublisher.DropDownStyle = ComboBoxStyle.DropDownList;
    cboPublisher.Enabled = false;
    cboPublisher.FormattingEnabled = true;
    cboPublisher.Location = new Point(123, 63);
    cboPublisher.Name = "cboPublisher";
    cboPublisher.Size = new Size(374, 23);
    cboPublisher.TabIndex = 5;
    //
    // lblProgrammer
    //
    lblProgrammer.AutoSize = true;
    lblProgrammer.Dock = DockStyle.Fill;
    lblProgrammer.Location = new Point(3, 90);
    lblProgrammer.Name = "lblProgrammer";
    lblProgrammer.Size = new Size(114, 30);
    lblProgrammer.TabIndex = 6;
    lblProgrammer.Text = "Programmer:";
    lblProgrammer.TextAlign = ContentAlignment.MiddleLeft;
    //
    // cboProgrammer
    //
    cboProgrammer.Dock = DockStyle.Fill;
    cboProgrammer.DropDownStyle = ComboBoxStyle.DropDownList;
    cboProgrammer.Enabled = false;
    cboProgrammer.FormattingEnabled = true;
    cboProgrammer.Location = new Point(123, 93);
    cboProgrammer.Name = "cboProgrammer";
    cboProgrammer.Size = new Size(374, 23);
    cboProgrammer.TabIndex = 7;
    //
    // lblGenre
    //
    lblGenre.AutoSize = true;
    lblGenre.Dock = DockStyle.Fill;
    lblGenre.Location = new Point(3, 120);
    lblGenre.Name = "lblGenre";
    lblGenre.Size = new Size(114, 30);
    lblGenre.TabIndex = 8;
    lblGenre.Text = "Genre:";
    lblGenre.TextAlign = ContentAlignment.MiddleLeft;
    //
    // cboGenre
    //
    cboGenre.Dock = DockStyle.Fill;
    cboGenre.DropDownStyle = ComboBoxStyle.DropDownList;
    cboGenre.Enabled = false;
    cboGenre.FormattingEnabled = true;
    cboGenre.Location = new Point(123, 123);
    cboGenre.Name = "cboGenre";
    cboGenre.Size = new Size(374, 23);
    cboGenre.TabIndex = 9;
    //
    // lblComment
    //
    lblComment.AutoSize = true;
    lblComment.Dock = DockStyle.Fill;
    lblComment.Location = new Point(3, 150);
    lblComment.Name = "lblComment";
    lblComment.Size = new Size(114, 30);
    lblComment.TabIndex = 10;
    lblComment.Text = "Comment:";
    lblComment.TextAlign = ContentAlignment.MiddleLeft;
    //
    // rtbComment
    //
    rtbComment.Dock = DockStyle.Fill;
    rtbComment.Location = new Point(123, 153);
    rtbComment.Name = "rtbComment";
    rtbComment.ReadOnly = true;
    rtbComment.Size = new Size(374, 204);
    rtbComment.TabIndex = 11;
    rtbComment.Text = "";
    //
    // btnClose
    //
    btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
    btnClose.Location = new Point(422, 363);
    btnClose.Name = "btnClose";
    btnClose.Size = new Size(75, 23);
    btnClose.TabIndex = 12;
    btnClose.Text = "Close";
    btnClose.UseVisualStyleBackColor = true;
    btnClose.Click += btnClose_Click;
    //
    // GameDetailsForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(500, 400);
    Controls.Add(tableLayoutPanel1);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "GameDetailsForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Game Details";
    tableLayoutPanel1.ResumeLayout(false);
    tableLayoutPanel1.PerformLayout();
    ResumeLayout(false);
  }

  #endregion
}
