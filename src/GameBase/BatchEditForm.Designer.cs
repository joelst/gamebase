namespace GameBase;

partial class BatchEditForm
{
  private System.ComponentModel.IContainer components = null;
  private Label lblPublisher;
  private ComboBox cboPublisher;
  private Label lblGenre;
  private ComboBox cboGenre;
  private Label lblFavorite;
  private ComboBox cboFavorite;
  private Label lblClassic;
  private ComboBox cboClassic;
  private Label lblAdult;
  private ComboBox cboAdult;
  private Button btnOk;
  private Button btnCancel;

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
    lblPublisher = new Label();
    cboPublisher = new ComboBox();
    lblGenre = new Label();
    cboGenre = new ComboBox();
    lblFavorite = new Label();
    cboFavorite = new ComboBox();
    lblClassic = new Label();
    cboClassic = new ComboBox();
    lblAdult = new Label();
    cboAdult = new ComboBox();
    btnOk = new Button();
    btnCancel = new Button();
    SuspendLayout();
    //
    // lblPublisher
    //
    lblPublisher.AutoSize = true;
    lblPublisher.Location = new Point(12, 15);
    lblPublisher.Name = "lblPublisher";
    lblPublisher.Size = new Size(60, 15);
    lblPublisher.TabIndex = 0;
    lblPublisher.Text = "Publisher:";
    //
    // cboPublisher
    //
    cboPublisher.DropDownStyle = ComboBoxStyle.DropDownList;
    cboPublisher.Location = new Point(120, 12);
    cboPublisher.Name = "cboPublisher";
    cboPublisher.Size = new Size(300, 23);
    cboPublisher.TabIndex = 1;
    //
    // lblGenre
    //
    lblGenre.AutoSize = true;
    lblGenre.Location = new Point(12, 50);
    lblGenre.Name = "lblGenre";
    lblGenre.Size = new Size(41, 15);
    lblGenre.TabIndex = 2;
    lblGenre.Text = "Genre:";
    //
    // cboGenre
    //
    cboGenre.DropDownStyle = ComboBoxStyle.DropDownList;
    cboGenre.Location = new Point(120, 47);
    cboGenre.Name = "cboGenre";
    cboGenre.Size = new Size(300, 23);
    cboGenre.TabIndex = 3;
    //
    // lblFavorite
    //
    lblFavorite.AutoSize = true;
    lblFavorite.Location = new Point(12, 85);
    lblFavorite.Name = "lblFavorite";
    lblFavorite.Size = new Size(54, 15);
    lblFavorite.TabIndex = 4;
    lblFavorite.Text = "Favorite:";
    //
    // cboFavorite
    //
    cboFavorite.DropDownStyle = ComboBoxStyle.DropDownList;
    cboFavorite.Location = new Point(120, 82);
    cboFavorite.Name = "cboFavorite";
    cboFavorite.Size = new Size(300, 23);
    cboFavorite.TabIndex = 5;
    //
    // lblClassic
    //
    lblClassic.AutoSize = true;
    lblClassic.Location = new Point(12, 120);
    lblClassic.Name = "lblClassic";
    lblClassic.Size = new Size(46, 15);
    lblClassic.TabIndex = 6;
    lblClassic.Text = "Classic:";
    //
    // cboClassic
    //
    cboClassic.DropDownStyle = ComboBoxStyle.DropDownList;
    cboClassic.Location = new Point(120, 117);
    cboClassic.Name = "cboClassic";
    cboClassic.Size = new Size(300, 23);
    cboClassic.TabIndex = 7;
    //
    // lblAdult
    //
    lblAdult.AutoSize = true;
    lblAdult.Location = new Point(12, 155);
    lblAdult.Name = "lblAdult";
    lblAdult.Size = new Size(38, 15);
    lblAdult.TabIndex = 8;
    lblAdult.Text = "Adult:";
    //
    // cboAdult
    //
    cboAdult.DropDownStyle = ComboBoxStyle.DropDownList;
    cboAdult.Location = new Point(120, 152);
    cboAdult.Name = "cboAdult";
    cboAdult.Size = new Size(300, 23);
    cboAdult.TabIndex = 9;
    //
    // btnOk
    //
    btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnOk.Location = new Point(264, 195);
    btnOk.Name = "btnOk";
    btnOk.Size = new Size(75, 23);
    btnOk.TabIndex = 10;
    btnOk.Text = "OK";
    btnOk.UseVisualStyleBackColor = true;
    btnOk.Click += btnOk_Click;
    //
    // btnCancel
    //
    btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnCancel.Location = new Point(345, 195);
    btnCancel.Name = "btnCancel";
    btnCancel.Size = new Size(75, 23);
    btnCancel.TabIndex = 11;
    btnCancel.Text = "Cancel";
    btnCancel.UseVisualStyleBackColor = true;
    btnCancel.Click += btnCancel_Click;
    //
    // BatchEditForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(434, 230);
    Controls.Add(btnCancel);
    Controls.Add(btnOk);
    Controls.Add(cboAdult);
    Controls.Add(lblAdult);
    Controls.Add(cboClassic);
    Controls.Add(lblClassic);
    Controls.Add(cboFavorite);
    Controls.Add(lblFavorite);
    Controls.Add(cboGenre);
    Controls.Add(lblGenre);
    Controls.Add(cboPublisher);
    Controls.Add(lblPublisher);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "BatchEditForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Batch Edit";
    ResumeLayout(false);
    PerformLayout();
  }
}
