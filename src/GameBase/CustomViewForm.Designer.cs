namespace GameBase;

partial class CustomViewForm
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

  private DataGridView dgvFilters;
  private Button btnAdd;
  private Button btnRemove;
  private Button btnClear;
  private Button btnOK;
  private Button btnCancel;
  private RadioButton rbMatchAll;
  private RadioButton rbMatchAny;
  private FlowLayoutPanel flowButtons;
  private FlowLayoutPanel flowMode;
  private TableLayoutPanel tableLayoutPanel1;

  private void InitializeComponent()
  {
    dgvFilters = new DataGridView();
    btnAdd = new Button();
    btnRemove = new Button();
    btnClear = new Button();
    btnOK = new Button();
    btnCancel = new Button();
    rbMatchAll = new RadioButton();
    rbMatchAny = new RadioButton();
    flowButtons = new FlowLayoutPanel();
    flowMode = new FlowLayoutPanel();
    tableLayoutPanel1 = new TableLayoutPanel();
    ((System.ComponentModel.ISupportInitialize)dgvFilters).BeginInit();
    flowButtons.SuspendLayout();
    flowMode.SuspendLayout();
    tableLayoutPanel1.SuspendLayout();
    SuspendLayout();
    //
    // dgvFilters
    //
    dgvFilters.AllowUserToAddRows = false;
    dgvFilters.AllowUserToDeleteRows = false;
    dgvFilters.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
    dgvFilters.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
    dgvFilters.Dock = DockStyle.Fill;
    dgvFilters.Location = new Point(3, 53);
    dgvFilters.MultiSelect = false;
    dgvFilters.Name = "dgvFilters";
    dgvFilters.RowHeadersVisible = false;
    dgvFilters.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
    dgvFilters.Size = new Size(694, 311);
    dgvFilters.TabIndex = 0;
    //
    // btnAdd
    //
    btnAdd.AutoSize = true;
    btnAdd.Text = "Add";
    btnAdd.Click += btnAdd_Click;
    //
    // btnRemove
    //
    btnRemove.AutoSize = true;
    btnRemove.Text = "Remove";
    btnRemove.Click += btnRemove_Click;
    //
    // btnClear
    //
    btnClear.AutoSize = true;
    btnClear.Text = "Clear";
    btnClear.Click += btnClear_Click;
    //
    // btnOK
    //
    btnOK.AutoSize = true;
    btnOK.Text = "OK";
    btnOK.Click += btnOK_Click;
    //
    // btnCancel
    //
    btnCancel.AutoSize = true;
    btnCancel.Text = "Cancel";
    btnCancel.Click += btnCancel_Click;
    //
    // rbMatchAll
    //
    rbMatchAll.AutoSize = true;
    rbMatchAll.Text = "All filters must match";
    rbMatchAll.Checked = true;
    //
    // rbMatchAny
    //
    rbMatchAny.AutoSize = true;
    rbMatchAny.Text = "Any filter can match";
    //
    // flowButtons
    //
    flowButtons.AutoSize = true;
    flowButtons.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    flowButtons.Controls.Add(btnAdd);
    flowButtons.Controls.Add(btnRemove);
    flowButtons.Controls.Add(btnClear);
    flowButtons.Controls.Add(btnOK);
    flowButtons.Controls.Add(btnCancel);
    flowButtons.Dock = DockStyle.Fill;
    flowButtons.FlowDirection = FlowDirection.LeftToRight;
    flowButtons.Location = new Point(3, 370);
    flowButtons.Name = "flowButtons";
    flowButtons.Size = new Size(694, 30);
    flowButtons.TabIndex = 1;
    //
    // flowMode
    //
    flowMode.AutoSize = true;
    flowMode.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    flowMode.Controls.Add(rbMatchAll);
    flowMode.Controls.Add(rbMatchAny);
    flowMode.Dock = DockStyle.Fill;
    flowMode.FlowDirection = FlowDirection.LeftToRight;
    flowMode.Location = new Point(3, 3);
    flowMode.Name = "flowMode";
    flowMode.Size = new Size(248, 25);
    flowMode.TabIndex = 2;
    //
    // tableLayoutPanel1
    //
    tableLayoutPanel1.ColumnCount = 1;
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.Controls.Add(flowMode, 0, 0);
    tableLayoutPanel1.Controls.Add(dgvFilters, 0, 1);
    tableLayoutPanel1.Controls.Add(flowButtons, 0, 2);
    tableLayoutPanel1.Dock = DockStyle.Fill;
    tableLayoutPanel1.Location = new Point(0, 0);
    tableLayoutPanel1.Name = "tableLayoutPanel1";
    tableLayoutPanel1.RowCount = 3;
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
    tableLayoutPanel1.Size = new Size(700, 450);
    tableLayoutPanel1.TabIndex = 3;
    //
    // CustomViewForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(700, 450);
    Controls.Add(tableLayoutPanel1);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "CustomViewForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Custom Views";
    ((System.ComponentModel.ISupportInitialize)dgvFilters).EndInit();
    flowButtons.ResumeLayout(false);
    flowButtons.PerformLayout();
    flowMode.ResumeLayout(false);
    flowMode.PerformLayout();
    tableLayoutPanel1.ResumeLayout(false);
    tableLayoutPanel1.PerformLayout();
    ResumeLayout(false);
  }

  #endregion
}
