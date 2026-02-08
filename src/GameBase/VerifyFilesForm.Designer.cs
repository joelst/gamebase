namespace GameBase;

partial class VerifyFilesForm
{
  private System.ComponentModel.IContainer components = null;
  private ListView listIssues;
  private ColumnHeader colType;
  private ColumnHeader colGame;
  private ColumnHeader colExtra;
  private ColumnHeader colPath;
  private Label lblSummary;
  private Button btnClose;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  private void InitializeComponent()
  {
    listIssues = new ListView();
    colType = new ColumnHeader();
    colGame = new ColumnHeader();
    colExtra = new ColumnHeader();
    colPath = new ColumnHeader();
    lblSummary = new Label();
    btnClose = new Button();
    SuspendLayout();
    //
    // listIssues
    //
    listIssues.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
    listIssues.Columns.AddRange(new ColumnHeader[] { colType, colGame, colExtra, colPath });
    listIssues.FullRowSelect = true;
    listIssues.GridLines = true;
    listIssues.Location = new Point(12, 12);
    listIssues.Name = "listIssues";
    listIssues.Size = new Size(760, 360);
    listIssues.TabIndex = 0;
    listIssues.UseCompatibleStateImageBehavior = false;
    listIssues.View = View.Details;
    //
    // colType
    //
    colType.Text = "Type";
    colType.Width = 80;
    //
    // colGame
    //
    colGame.Text = "Game";
    colGame.Width = 180;
    //
    // colExtra
    //
    colExtra.Text = "Extra";
    colExtra.Width = 180;
    //
    // colPath
    //
    colPath.Text = "Path";
    colPath.Width = 300;
    //
    // lblSummary
    //
    lblSummary.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
    lblSummary.AutoSize = true;
    lblSummary.Location = new Point(12, 385);
    lblSummary.Name = "lblSummary";
    lblSummary.Size = new Size(102, 15);
    lblSummary.TabIndex = 1;
    lblSummary.Text = "Missing files: 0";
    //
    // btnClose
    //
    btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnClose.Location = new Point(697, 380);
    btnClose.Name = "btnClose";
    btnClose.Size = new Size(75, 23);
    btnClose.TabIndex = 2;
    btnClose.Text = "Close";
    btnClose.UseVisualStyleBackColor = true;
    btnClose.Click += btnClose_Click;
    //
    // VerifyFilesForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(784, 411);
    Controls.Add(btnClose);
    Controls.Add(lblSummary);
    Controls.Add(listIssues);
    Name = "VerifyFilesForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Verify Files";
    ResumeLayout(false);
    PerformLayout();
  }
}
