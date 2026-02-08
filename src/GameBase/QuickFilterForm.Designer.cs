namespace GameBase;

partial class QuickFilterForm
{
  private System.ComponentModel.IContainer components = null;
  private TextBox txtFilter;
  private ComboBox cboField;
  private RadioButton rbGamesOnly;
  private RadioButton rbBothTables;
  private Button btnOK;
  private Button btnCancel;
  private Label lblContains;
  private Label lblField;
  private TableLayoutPanel tableLayoutPanel1;
  private FlowLayoutPanel flowButtons;
  private FlowLayoutPanel flowTables;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  #region Windows Form Designer generated code

  private void InitializeComponent()
  {
    txtFilter = new TextBox();
    cboField = new ComboBox();
    rbGamesOnly = new RadioButton();
    rbBothTables = new RadioButton();
    btnOK = new Button();
    btnCancel = new Button();
    lblContains = new Label();
    lblField = new Label();
    tableLayoutPanel1 = new TableLayoutPanel();
    flowButtons = new FlowLayoutPanel();
    flowTables = new FlowLayoutPanel();
    tableLayoutPanel1.SuspendLayout();
    flowButtons.SuspendLayout();
    flowTables.SuspendLayout();
    SuspendLayout();
    //
    // txtFilter
    //
    txtFilter.Dock = DockStyle.Fill;
    txtFilter.PlaceholderText = "Contains text...";
    txtFilter.TextChanged += txtFilter_TextChanged;
    //
    // cboField
    //
    cboField.Dock = DockStyle.Fill;
    cboField.DropDownStyle = ComboBoxStyle.DropDownList;
    //
    // rbGamesOnly
    //
    rbGamesOnly.AutoSize = true;
    rbGamesOnly.Checked = true;
    rbGamesOnly.Text = "Games only";
    //
    // rbBothTables
    //
    rbBothTables.AutoSize = true;
    rbBothTables.Text = "Games + SIDs";
    //
    // btnOK
    //
    btnOK.AutoSize = true;
    btnOK.Text = "OK";
    btnOK.Enabled = false;
    btnOK.Click += btnOK_Click;
    //
    // btnCancel
    //
    btnCancel.AutoSize = true;
    btnCancel.Text = "Cancel";
    btnCancel.Click += btnCancel_Click;
    //
    // lblContains
    //
    lblContains.AutoSize = true;
    lblContains.Text = "Contains text:";
    //
    // lblField
    //
    lblField.AutoSize = true;
    lblField.Text = "Field:";
    //
    // tableLayoutPanel1
    //
    tableLayoutPanel1.ColumnCount = 2;
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.Controls.Add(lblContains, 0, 0);
    tableLayoutPanel1.Controls.Add(txtFilter, 1, 0);
    tableLayoutPanel1.Controls.Add(lblField, 0, 1);
    tableLayoutPanel1.Controls.Add(cboField, 1, 1);
    tableLayoutPanel1.Controls.Add(flowTables, 1, 2);
    tableLayoutPanel1.Controls.Add(flowButtons, 1, 3);
    tableLayoutPanel1.Dock = DockStyle.Fill;
    tableLayoutPanel1.RowCount = 4;
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
    tableLayoutPanel1.Padding = new Padding(10);
    //
    // flowButtons
    //
    flowButtons.AutoSize = true;
    flowButtons.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    flowButtons.Controls.Add(btnOK);
    flowButtons.Controls.Add(btnCancel);
    flowButtons.FlowDirection = FlowDirection.LeftToRight;
    flowButtons.Dock = DockStyle.Fill;
    flowButtons.Padding = new Padding(0, 10, 0, 0);
    //
    // flowTables
    //
    flowTables.AutoSize = true;
    flowTables.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    flowTables.Controls.Add(rbGamesOnly);
    flowTables.Controls.Add(rbBothTables);
    flowTables.FlowDirection = FlowDirection.LeftToRight;
    flowTables.Dock = DockStyle.Fill;
    //
    // QuickFilterForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(320, 200);
    Controls.Add(tableLayoutPanel1);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "QuickFilterForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Quick Filter";
    tableLayoutPanel1.ResumeLayout(false);
    tableLayoutPanel1.PerformLayout();
    flowButtons.ResumeLayout(false);
    flowButtons.PerformLayout();
    flowTables.ResumeLayout(false);
    flowTables.PerformLayout();
    ResumeLayout(false);
  }

  #endregion
}
