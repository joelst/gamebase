namespace GameBase;

partial class ColumnsForm
{
  private System.ComponentModel.IContainer components = null;
  private CheckedListBox checkedListColumns;
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
    checkedListColumns = new CheckedListBox();
    btnOk = new Button();
    btnCancel = new Button();
    SuspendLayout();
    //
    // checkedListColumns
    //
    checkedListColumns.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
    checkedListColumns.CheckOnClick = true;
    checkedListColumns.FormattingEnabled = true;
    checkedListColumns.Location = new Point(12, 12);
    checkedListColumns.Name = "checkedListColumns";
    checkedListColumns.Size = new Size(260, 202);
    checkedListColumns.TabIndex = 0;
    //
    // btnOk
    //
    btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnOk.Location = new Point(116, 230);
    btnOk.Name = "btnOk";
    btnOk.Size = new Size(75, 23);
    btnOk.TabIndex = 1;
    btnOk.Text = "OK";
    btnOk.UseVisualStyleBackColor = true;
    btnOk.Click += btnOk_Click;
    //
    // btnCancel
    //
    btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    btnCancel.Location = new Point(197, 230);
    btnCancel.Name = "btnCancel";
    btnCancel.Size = new Size(75, 23);
    btnCancel.TabIndex = 2;
    btnCancel.Text = "Cancel";
    btnCancel.UseVisualStyleBackColor = true;
    btnCancel.Click += btnCancel_Click;
    //
    // ColumnsForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(284, 265);
    Controls.Add(btnCancel);
    Controls.Add(btnOk);
    Controls.Add(checkedListColumns);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "ColumnsForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Columns";
    ResumeLayout(false);
  }
}
