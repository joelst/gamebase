namespace GameBase;

partial class ExtrasEditorForm
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

  private Label lblTitle;
  private DataGridView dgvExtras;
  private Button btnAdd;
  private Button btnRemove;
  private Button btnSave;
  private Button btnClose;
  private TableLayoutPanel tableLayoutPanel1;
  private FlowLayoutPanel flowLayoutPanel1;

  private void InitializeComponent()
  {
    lblTitle = new Label();
    dgvExtras = new DataGridView();
    btnAdd = new Button();
    btnRemove = new Button();
    btnSave = new Button();
    btnClose = new Button();
    tableLayoutPanel1 = new TableLayoutPanel();
    flowLayoutPanel1 = new FlowLayoutPanel();
    ((System.ComponentModel.ISupportInitialize)dgvExtras).BeginInit();
    tableLayoutPanel1.SuspendLayout();
    flowLayoutPanel1.SuspendLayout();
    SuspendLayout();
    //
    // lblTitle
    //
    lblTitle.AutoSize = true;
    lblTitle.Dock = DockStyle.Fill;
    lblTitle.Font = new Font("Segoe UI", 11F, FontStyle.Bold, GraphicsUnit.Point);
    lblTitle.Location = new Point(3, 0);
    lblTitle.Name = "lblTitle";
    lblTitle.Size = new Size(794, 30);
    lblTitle.TabIndex = 0;
    lblTitle.Text = "Extras";
    lblTitle.TextAlign = ContentAlignment.MiddleLeft;
    //
    // dgvExtras
    //
    dgvExtras.AllowUserToAddRows = false;
    dgvExtras.AllowUserToDeleteRows = false;
    dgvExtras.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
    dgvExtras.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
    dgvExtras.Dock = DockStyle.Fill;
    dgvExtras.Location = new Point(3, 33);
    dgvExtras.MultiSelect = false;
    dgvExtras.Name = "dgvExtras";
    dgvExtras.RowHeadersVisible = false;
    dgvExtras.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
    dgvExtras.Size = new Size(794, 355);
    dgvExtras.TabIndex = 1;
    //
    // btnAdd
    //
    btnAdd.AutoSize = true;
    btnAdd.Location = new Point(3, 3);
    btnAdd.Name = "btnAdd";
    btnAdd.Size = new Size(75, 25);
    btnAdd.TabIndex = 0;
    btnAdd.Text = "Add";
    btnAdd.UseVisualStyleBackColor = true;
    btnAdd.Click += btnAdd_Click;
    //
    // btnRemove
    //
    btnRemove.AutoSize = true;
    btnRemove.Location = new Point(84, 3);
    btnRemove.Name = "btnRemove";
    btnRemove.Size = new Size(75, 25);
    btnRemove.TabIndex = 1;
    btnRemove.Text = "Remove";
    btnRemove.UseVisualStyleBackColor = true;
    btnRemove.Click += btnRemove_Click;
    //
    // btnSave
    //
    btnSave.AutoSize = true;
    btnSave.Location = new Point(165, 3);
    btnSave.Name = "btnSave";
    btnSave.Size = new Size(75, 25);
    btnSave.TabIndex = 2;
    btnSave.Text = "Save";
    btnSave.UseVisualStyleBackColor = true;
    btnSave.Click += btnSave_Click;
    //
    // btnClose
    //
    btnClose.AutoSize = true;
    btnClose.Location = new Point(246, 3);
    btnClose.Name = "btnClose";
    btnClose.Size = new Size(75, 25);
    btnClose.TabIndex = 3;
    btnClose.Text = "Close";
    btnClose.UseVisualStyleBackColor = true;
    btnClose.Click += btnClose_Click;
    //
    // tableLayoutPanel1
    //
    tableLayoutPanel1.ColumnCount = 1;
    tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.Controls.Add(lblTitle, 0, 0);
    tableLayoutPanel1.Controls.Add(dgvExtras, 0, 1);
    tableLayoutPanel1.Controls.Add(flowLayoutPanel1, 0, 2);
    tableLayoutPanel1.Dock = DockStyle.Fill;
    tableLayoutPanel1.Location = new Point(0, 0);
    tableLayoutPanel1.Name = "tableLayoutPanel1";
    tableLayoutPanel1.RowCount = 3;
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
    tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
    tableLayoutPanel1.Size = new Size(800, 428);
    tableLayoutPanel1.TabIndex = 0;
    //
    // flowLayoutPanel1
    //
    flowLayoutPanel1.AutoSize = true;
    flowLayoutPanel1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    flowLayoutPanel1.Controls.Add(btnAdd);
    flowLayoutPanel1.Controls.Add(btnRemove);
    flowLayoutPanel1.Controls.Add(btnSave);
    flowLayoutPanel1.Controls.Add(btnClose);
    flowLayoutPanel1.Dock = DockStyle.Fill;
    flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight;
    flowLayoutPanel1.Location = new Point(3, 391);
    flowLayoutPanel1.Name = "flowLayoutPanel1";
    flowLayoutPanel1.Size = new Size(291, 34);
    flowLayoutPanel1.TabIndex = 2;
    //
    // ExtrasEditorForm
    //
    AutoScaleDimensions = new SizeF(7F, 15F);
    AutoScaleMode = AutoScaleMode.Font;
    ClientSize = new Size(800, 428);
    Controls.Add(tableLayoutPanel1);
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    Name = "ExtrasEditorForm";
    StartPosition = FormStartPosition.CenterParent;
    Text = "Extras Editor";
    ((System.ComponentModel.ISupportInitialize)dgvExtras).EndInit();
    tableLayoutPanel1.ResumeLayout(false);
    tableLayoutPanel1.PerformLayout();
    flowLayoutPanel1.ResumeLayout(false);
    flowLayoutPanel1.PerformLayout();
    ResumeLayout(false);
  }

  #endregion
}
