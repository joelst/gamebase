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
    private System.Windows.Forms.Label lblDbPath;
    private System.Windows.Forms.TextBox txtDbPath;
    private System.Windows.Forms.Button btnBrowseDb;
    private System.Windows.Forms.Label lblGamesPath;
    private System.Windows.Forms.TextBox txtGamesPath;
    private System.Windows.Forms.Button btnBrowseGames;
    private System.Windows.Forms.Label lblExtrasPath;
    private System.Windows.Forms.TextBox txtExtrasPath;
    private System.Windows.Forms.Button btnBrowseExtras;
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
      lblDbPath = new System.Windows.Forms.Label();
      txtDbPath = new System.Windows.Forms.TextBox();
      btnBrowseDb = new System.Windows.Forms.Button();
      lblGamesPath = new System.Windows.Forms.Label();
      txtGamesPath = new System.Windows.Forms.TextBox();
      btnBrowseGames = new System.Windows.Forms.Button();
      lblExtrasPath = new System.Windows.Forms.Label();
      txtExtrasPath = new System.Windows.Forms.TextBox();
      btnBrowseExtras = new System.Windows.Forms.Button();
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
      // lblDbPath
      //
      lblDbPath.AutoSize = true;
      lblDbPath.Location = new System.Drawing.Point(12, 125);
      lblDbPath.Name = "lblDbPath";
      lblDbPath.Size = new System.Drawing.Size(86, 15);
      lblDbPath.TabIndex = 7;
      lblDbPath.Text = "Database Path:";
      //
      // txtDbPath
      //
      txtDbPath.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      txtDbPath.Location = new System.Drawing.Point(115, 122);
      txtDbPath.Name = "txtDbPath";
      txtDbPath.Size = new System.Drawing.Size(328, 23);
      txtDbPath.TabIndex = 8;
      //
      // btnBrowseDb
      //
      btnBrowseDb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
      btnBrowseDb.Location = new System.Drawing.Point(449, 121);
      btnBrowseDb.Name = "btnBrowseDb";
      btnBrowseDb.Size = new System.Drawing.Size(75, 23);
      btnBrowseDb.TabIndex = 9;
      btnBrowseDb.Text = "Browse...";
      btnBrowseDb.UseVisualStyleBackColor = true;
      btnBrowseDb.Click += btnBrowseDb_Click;
      //
      // lblGamesPath
      //
      lblGamesPath.AutoSize = true;
      lblGamesPath.Location = new System.Drawing.Point(12, 160);
      lblGamesPath.Name = "lblGamesPath";
      lblGamesPath.Size = new System.Drawing.Size(74, 15);
      lblGamesPath.TabIndex = 10;
      lblGamesPath.Text = "Games Path:";
      //
      // txtGamesPath
      //
      txtGamesPath.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      txtGamesPath.Location = new System.Drawing.Point(115, 157);
      txtGamesPath.Name = "txtGamesPath";
      txtGamesPath.Size = new System.Drawing.Size(328, 23);
      txtGamesPath.TabIndex = 11;
      //
      // btnBrowseGames
      //
      btnBrowseGames.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
      btnBrowseGames.Location = new System.Drawing.Point(449, 156);
      btnBrowseGames.Name = "btnBrowseGames";
      btnBrowseGames.Size = new System.Drawing.Size(75, 23);
      btnBrowseGames.TabIndex = 12;
      btnBrowseGames.Text = "Browse...";
      btnBrowseGames.UseVisualStyleBackColor = true;
      btnBrowseGames.Click += btnBrowseGames_Click;
      //
      // lblExtrasPath
      //
      lblExtrasPath.AutoSize = true;
      lblExtrasPath.Location = new System.Drawing.Point(12, 195);
      lblExtrasPath.Name = "lblExtrasPath";
      lblExtrasPath.Size = new System.Drawing.Size(70, 15);
      lblExtrasPath.TabIndex = 13;
      lblExtrasPath.Text = "Extras Path:";
      //
      // txtExtrasPath
      //
      txtExtrasPath.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
      txtExtrasPath.Location = new System.Drawing.Point(115, 192);
      txtExtrasPath.Name = "txtExtrasPath";
      txtExtrasPath.Size = new System.Drawing.Size(328, 23);
      txtExtrasPath.TabIndex = 14;
      //
      // btnBrowseExtras
      //
      btnBrowseExtras.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
      btnBrowseExtras.Location = new System.Drawing.Point(449, 191);
      btnBrowseExtras.Name = "btnBrowseExtras";
      btnBrowseExtras.Size = new System.Drawing.Size(75, 23);
      btnBrowseExtras.TabIndex = 15;
      btnBrowseExtras.Text = "Browse...";
      btnBrowseExtras.UseVisualStyleBackColor = true;
      btnBrowseExtras.Click += btnBrowseExtras_Click;
      //
      // btnSave
      //
      btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
      btnSave.Location = new System.Drawing.Point(368, 240);
      btnSave.Name = "btnSave";
      btnSave.Size = new System.Drawing.Size(75, 23);
      btnSave.TabIndex = 16;
      btnSave.Text = "Save";
      btnSave.UseVisualStyleBackColor = true;
      btnSave.Click += btnSave_Click;
      //
      // btnCancel
      //
      btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
      btnCancel.Location = new System.Drawing.Point(449, 240);
      btnCancel.Name = "btnCancel";
      btnCancel.Size = new System.Drawing.Size(75, 23);
      btnCancel.TabIndex = 17;
      btnCancel.Text = "Cancel";
      btnCancel.UseVisualStyleBackColor = true;
      btnCancel.Click += btnCancel_Click;
      //
      // OptionsForm
      //
      AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
      AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      ClientSize = new System.Drawing.Size(536, 280);
      Controls.Add(btnCancel);
      Controls.Add(btnSave);
      Controls.Add(btnBrowseExtras);
      Controls.Add(txtExtrasPath);
      Controls.Add(lblExtrasPath);
      Controls.Add(btnBrowseGames);
      Controls.Add(txtGamesPath);
      Controls.Add(lblGamesPath);
      Controls.Add(btnBrowseDb);
      Controls.Add(txtDbPath);
      Controls.Add(lblDbPath);
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
