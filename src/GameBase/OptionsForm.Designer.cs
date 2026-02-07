namespace GameBase
{
  partial class OptionsForm
  {
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    private System.Windows.Forms.Label lblBasePath;
    private System.Windows.Forms.TextBox txtBasePath;
    private System.Windows.Forms.Button btnBrowse;
    private System.Windows.Forms.Label lblGbName;
    private System.Windows.Forms.TextBox txtGbName;
    private System.Windows.Forms.Label lblBackend;
    private System.Windows.Forms.ComboBox cmbBackend;
    private System.Windows.Forms.Button btnSave;
    private System.Windows.Forms.Button btnCancel;

    /// <summary>
    ///  Clean up any resources being used.
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

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      lblBasePath = new System.Windows.Forms.Label();
      txtBasePath = new System.Windows.Forms.TextBox();
      btnBrowse = new System.Windows.Forms.Button();
      lblGbName = new System.Windows.Forms.Label();
      txtGbName = new System.Windows.Forms.TextBox();
      lblBackend = new System.Windows.Forms.Label();
      cmbBackend = new System.Windows.Forms.ComboBox();
      btnSave = new System.Windows.Forms.Button();
      btnCancel = new System.Windows.Forms.Button();
      SuspendLayout();
      //
      // lblBasePath
      //
      lblBasePath.AutoSize = true;
      lblBasePath.Location = new System.Drawing.Point(12, 15);
      lblBasePath.Name = "lblBasePath";
      lblBasePath.Size = new System.Drawing.Size(97, 15);
      lblBasePath.TabIndex = 0;
      lblBasePath.Text = "GameBase Root:";
      //
      // txtBasePath
      //
      txtBasePath.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      txtBasePath.Location = new System.Drawing.Point(115, 12);
      txtBasePath.Name = "txtBasePath";
      txtBasePath.Size = new System.Drawing.Size(328, 23);
      txtBasePath.TabIndex = 1;
      //
      // btnBrowse
      //
      btnBrowse.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
      btnBrowse.Location = new System.Drawing.Point(449, 11);
      btnBrowse.Name = "btnBrowse";
      btnBrowse.Size = new System.Drawing.Size(75, 23);
      btnBrowse.TabIndex = 2;
      btnBrowse.Text = "Browse...";
      btnBrowse.UseVisualStyleBackColor = true;
      btnBrowse.Click += btnBrowse_Click;
      //
      // lblGbName
      //
      lblGbName.AutoSize = true;
      lblGbName.Location = new System.Drawing.Point(12, 52);
      lblGbName.Name = "lblGbName";
      lblGbName.Size = new System.Drawing.Size(61, 15);
      lblGbName.TabIndex = 3;
      lblGbName.Text = "GB Name:";
      //
      // txtGbName
      //
      txtGbName.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      txtGbName.Location = new System.Drawing.Point(115, 49);
      txtGbName.Name = "txtGbName";
      txtGbName.Size = new System.Drawing.Size(328, 23);
      txtGbName.TabIndex = 4;
      //
      // lblBackend
      //
      lblBackend.AutoSize = true;
      lblBackend.Location = new System.Drawing.Point(12, 90);
      lblBackend.Name = "lblBackend";
      lblBackend.Size = new System.Drawing.Size(90, 15);
      lblBackend.TabIndex = 5;
      lblBackend.Text = "Archive backend";
      //
      // cmbBackend
      //
      cmbBackend.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      cmbBackend.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
      cmbBackend.FormattingEnabled = true;
      cmbBackend.Location = new System.Drawing.Point(115, 87);
      cmbBackend.Name = "cmbBackend";
      cmbBackend.Size = new System.Drawing.Size(328, 23);
      cmbBackend.TabIndex = 6;
      //
      // btnSave
      //
      btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
      btnSave.Location = new System.Drawing.Point(368, 139);
      btnSave.Name = "btnSave";
      btnSave.Size = new System.Drawing.Size(75, 23);
      btnSave.TabIndex = 7;
      btnSave.Text = "Save";
      btnSave.UseVisualStyleBackColor = true;
      btnSave.Click += btnSave_Click;
      //
      // btnCancel
      //
      btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
      btnCancel.Location = new System.Drawing.Point(449, 139);
      btnCancel.Name = "btnCancel";
      btnCancel.Size = new System.Drawing.Size(75, 23);
      btnCancel.TabIndex = 8;
      btnCancel.Text = "Cancel";
      btnCancel.UseVisualStyleBackColor = true;
      btnCancel.Click += btnCancel_Click;
      //
      // OptionsForm
      //
      AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
      AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      ClientSize = new System.Drawing.Size(536, 174);
      Controls.Add(btnCancel);
      Controls.Add(btnSave);
      Controls.Add(cmbBackend);
      Controls.Add(lblBackend);
      Controls.Add(txtGbName);
      Controls.Add(lblGbName);
      Controls.Add(btnBrowse);
      Controls.Add(txtBasePath);
      Controls.Add(lblBasePath);
      Name = "OptionsForm";
      Text = "Options";
      Load += OptionsForm_Load;
      ResumeLayout(false);
      PerformLayout();
    }

    #endregion
  }
}
