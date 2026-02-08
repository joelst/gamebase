namespace GameBase;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

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
    private Button btnOptions;
    private TableLayoutPanel tableLayoutPanel1;
    private SplitContainer splitContainer1;
    private DataGridView dgvGames;
    private TabControl tabDetails;
    private TabPage tabInfo;
    private TabPage tabScreenshots;
    private TabPage tabExtras;
    private PictureBox picScreenshot;
    private ListBox lstExtras;
    private RichTextBox rtbGameInfo;
    private Label lblStatus;
    private MenuStrip menuStrip1;
    private ToolStripMenuItem fileToolStripMenuItem;
    private ToolStripMenuItem viewToolStripMenuItem;
    private ToolStripMenuItem gameToolStripMenuItem;
    private ToolStripMenuItem helpToolStripMenuItem;
    private ToolStripMenuItem exitToolStripMenuItem;
    private ToolStripMenuItem playGameToolStripMenuItem;
    private ToolStripMenuItem editGameToolStripMenuItem;
    private ToolStripMenuItem manageExtrasToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator1;
    private ToolStripMenuItem batchEditToolStripMenuItem;
    private ToolStripMenuItem verifyFilesToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator2;
    private ToolStripMenuItem customViewsToolStripMenuItem;
    private ToolStripMenuItem quickFilterToolStripMenuItem;
    private ToolStripMenuItem clearFiltersToolStripMenuItem;
    private ToolStripMenuItem columnsToolStripMenuItem;
    private ToolStripMenuItem aboutToolStripMenuItem;
    private TextBox txtSearch;
    private Button btnClearFilters;
    private FlowLayoutPanel flowSearch;
    private ContextMenuStrip dgvContextMenu;
    private ToolStripMenuItem toggleFavoriteToolStripMenuItem;
    private ToolStripMenuItem toggleClassicToolStripMenuItem;
    private ToolStripMenuItem toggleAdultToolStripMenuItem;
    private ContextMenuStrip screenshotContextMenu;
    private ToolStripMenuItem selectScreenshotToolStripMenuItem;
    private ToolStripMenuItem clearScreenshotToolStripMenuItem;
    private Label lblSearch;

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        menuStrip1 = new MenuStrip();
        fileToolStripMenuItem = new ToolStripMenuItem();
        exitToolStripMenuItem = new ToolStripMenuItem();
        viewToolStripMenuItem = new ToolStripMenuItem();
        customViewsToolStripMenuItem = new ToolStripMenuItem();
        quickFilterToolStripMenuItem = new ToolStripMenuItem();
        clearFiltersToolStripMenuItem = new ToolStripMenuItem();
        columnsToolStripMenuItem = new ToolStripMenuItem();
        gameToolStripMenuItem = new ToolStripMenuItem();
        playGameToolStripMenuItem = new ToolStripMenuItem();
        editGameToolStripMenuItem = new ToolStripMenuItem();
        manageExtrasToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator1 = new ToolStripSeparator();
        batchEditToolStripMenuItem = new ToolStripMenuItem();
        verifyFilesToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator2 = new ToolStripSeparator();
        helpToolStripMenuItem = new ToolStripMenuItem();
        aboutToolStripMenuItem = new ToolStripMenuItem();
        btnOptions = new Button();
        tableLayoutPanel1 = new TableLayoutPanel();
        splitContainer1 = new SplitContainer();
        dgvGames = new DataGridView();
        lblSearch = new Label();
        tabDetails = new TabControl();
        tabInfo = new TabPage();
        rtbGameInfo = new RichTextBox();
        tabScreenshots = new TabPage();
        picScreenshot = new PictureBox();
        screenshotContextMenu = new ContextMenuStrip(components);
        selectScreenshotToolStripMenuItem = new ToolStripMenuItem();
        clearScreenshotToolStripMenuItem = new ToolStripMenuItem();
        tabExtras = new TabPage();
        lstExtras = new ListBox();
        lblStatus = new Label();
        txtSearch = new TextBox();
        dgvContextMenu = new ContextMenuStrip(components);
        toggleFavoriteToolStripMenuItem = new ToolStripMenuItem();
        toggleClassicToolStripMenuItem = new ToolStripMenuItem();
        toggleAdultToolStripMenuItem = new ToolStripMenuItem();
        flowSearch = new FlowLayoutPanel();
        btnClearFilters = new Button();
        menuStrip1.SuspendLayout();
        tableLayoutPanel1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
        splitContainer1.Panel1.SuspendLayout();
        splitContainer1.Panel2.SuspendLayout();
        splitContainer1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvGames).BeginInit();
        tabDetails.SuspendLayout();
        tabInfo.SuspendLayout();
        tabScreenshots.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)picScreenshot).BeginInit();
        dgvContextMenu.SuspendLayout();
        screenshotContextMenu.SuspendLayout();
        tabExtras.SuspendLayout();
        SuspendLayout();
        //
        // menuStrip1
        //
        menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, viewToolStripMenuItem, gameToolStripMenuItem, helpToolStripMenuItem });
        menuStrip1.Location = new Point(0, 0);
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(1000, 24);
        menuStrip1.TabIndex = 3;
        menuStrip1.Text = "menuStrip1";
        //
        // fileToolStripMenuItem
        //
        fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { exitToolStripMenuItem });
        fileToolStripMenuItem.Name = "fileToolStripMenuItem";
        fileToolStripMenuItem.Size = new Size(37, 20);
        fileToolStripMenuItem.Text = "&File";
        //
        // exitToolStripMenuItem
        //
        exitToolStripMenuItem.Name = "exitToolStripMenuItem";
        exitToolStripMenuItem.Size = new Size(93, 22);
        exitToolStripMenuItem.Text = "E&xit";
        exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
        //
        // viewToolStripMenuItem
        //
        viewToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { customViewsToolStripMenuItem, quickFilterToolStripMenuItem, clearFiltersToolStripMenuItem, columnsToolStripMenuItem });
        viewToolStripMenuItem.Name = "viewToolStripMenuItem";
        viewToolStripMenuItem.Size = new Size(44, 20);
        viewToolStripMenuItem.Text = "&View";
        //
        // customViewsToolStripMenuItem
        //
        customViewsToolStripMenuItem.Name = "customViewsToolStripMenuItem";
        customViewsToolStripMenuItem.Size = new Size(180, 22);
        customViewsToolStripMenuItem.Text = "Custom &Views...";
        //
        // quickFilterToolStripMenuItem
        //
        quickFilterToolStripMenuItem.Name = "quickFilterToolStripMenuItem";
        quickFilterToolStripMenuItem.Size = new Size(180, 22);
        quickFilterToolStripMenuItem.Text = "Quick &Filter...";
        quickFilterToolStripMenuItem.Click += quickFilterToolStripMenuItem_Click;
        //
        // clearFiltersToolStripMenuItem
        //
        clearFiltersToolStripMenuItem.Name = "clearFiltersToolStripMenuItem";
        clearFiltersToolStripMenuItem.Size = new Size(180, 22);
        clearFiltersToolStripMenuItem.Text = "Clear Filters";
        clearFiltersToolStripMenuItem.Click += clearFiltersToolStripMenuItem_Click;
        //
        // columnsToolStripMenuItem
        //
        columnsToolStripMenuItem.Name = "columnsToolStripMenuItem";
        columnsToolStripMenuItem.Size = new Size(180, 22);
        columnsToolStripMenuItem.Text = "Columns...";
        columnsToolStripMenuItem.Click += columnsToolStripMenuItem_Click;
        //
        // gameToolStripMenuItem
        //
        gameToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { playGameToolStripMenuItem, editGameToolStripMenuItem, manageExtrasToolStripMenuItem, toolStripSeparator1, batchEditToolStripMenuItem, verifyFilesToolStripMenuItem, toolStripSeparator2 });
        gameToolStripMenuItem.Name = "gameToolStripMenuItem";
        gameToolStripMenuItem.Size = new Size(50, 20);
        gameToolStripMenuItem.Text = "&Game";
        //
        // playGameToolStripMenuItem
        //
        playGameToolStripMenuItem.Name = "playGameToolStripMenuItem";
        playGameToolStripMenuItem.ShortcutKeys = Keys.F5;
        playGameToolStripMenuItem.Size = new Size(144, 22);
        playGameToolStripMenuItem.Text = "&Play Game";
        playGameToolStripMenuItem.Click += playGameToolStripMenuItem_Click;
        //
        // editGameToolStripMenuItem
        //
        editGameToolStripMenuItem.Name = "editGameToolStripMenuItem";
        editGameToolStripMenuItem.Size = new Size(144, 22);
        editGameToolStripMenuItem.Text = "&Edit Game...";
        //
        // manageExtrasToolStripMenuItem
        //
        manageExtrasToolStripMenuItem.Name = "manageExtrasToolStripMenuItem";
        manageExtrasToolStripMenuItem.Size = new Size(144, 22);
        manageExtrasToolStripMenuItem.Text = "Manage &Extras...";
        manageExtrasToolStripMenuItem.Click += manageExtrasToolStripMenuItem_Click;
        //
        // toolStripSeparator1
        //
        toolStripSeparator1.Name = "toolStripSeparator1";
        toolStripSeparator1.Size = new Size(141, 6);
        //
        // batchEditToolStripMenuItem
        //
        batchEditToolStripMenuItem.Name = "batchEditToolStripMenuItem";
        batchEditToolStripMenuItem.Size = new Size(144, 22);
        batchEditToolStripMenuItem.Text = "&Batch Edit...";
        batchEditToolStripMenuItem.Click += batchEditToolStripMenuItem_Click;
        //
        // verifyFilesToolStripMenuItem
        //
        verifyFilesToolStripMenuItem.Name = "verifyFilesToolStripMenuItem";
        verifyFilesToolStripMenuItem.Size = new Size(144, 22);
        verifyFilesToolStripMenuItem.Text = "&Verify Files...";
        verifyFilesToolStripMenuItem.Click += verifyFilesToolStripMenuItem_Click;
        //
        // toolStripSeparator2
        //
        toolStripSeparator2.Name = "toolStripSeparator2";
        toolStripSeparator2.Size = new Size(141, 6);
        //
        // helpToolStripMenuItem
        //
        helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { aboutToolStripMenuItem });
        helpToolStripMenuItem.Name = "helpToolStripMenuItem";
        helpToolStripMenuItem.Size = new Size(44, 20);
        helpToolStripMenuItem.Text = "&Help";
        //
        // aboutToolStripMenuItem
        //
        aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
        aboutToolStripMenuItem.Size = new Size(116, 22);
        aboutToolStripMenuItem.Text = "&About...";
        aboutToolStripMenuItem.Click += aboutToolStripMenuItem_Click;
        //
        // btnOptions
        //
        btnOptions.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnOptions.Location = new Point(912, 568);
        btnOptions.Name = "btnOptions";
        btnOptions.Size = new Size(75, 23);
        btnOptions.TabIndex = 1;
        btnOptions.Text = "Options";
        btnOptions.UseVisualStyleBackColor = true;
        btnOptions.Click += btnOptions_Click;
        //
        // tableLayoutPanel1
        //
        tableLayoutPanel1.ColumnCount = 1;
        tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel1.Controls.Add(splitContainer1, 0, 1);
        tableLayoutPanel1.Controls.Add(lblStatus, 0, 3);
        tableLayoutPanel1.Controls.Add(btnOptions, 0, 2);
        tableLayoutPanel1.Controls.Add(flowSearch, 0, 0);
        tableLayoutPanel1.Dock = DockStyle.Fill;
        tableLayoutPanel1.Location = new Point(0, 24);
        tableLayoutPanel1.Name = "tableLayoutPanel1";
        tableLayoutPanel1.RowCount = 4;
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
        tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableLayoutPanel1.Size = new Size(1000, 626);
        tableLayoutPanel1.TabIndex = 2;
        //
        // splitContainer1
        //
        splitContainer1.Dock = DockStyle.Fill;
        splitContainer1.Location = new Point(3, 38);
        splitContainer1.Name = "splitContainer1";
        //
        // splitContainer1.Panel1
        //
        splitContainer1.Panel1.Controls.Add(dgvGames);
        splitContainer1.Panel1.Controls.Add(lblSearch);
        //
        // splitContainer1.Panel2
        //
        splitContainer1.Panel2.Controls.Add(tabDetails);
        splitContainer1.Size = new Size(994, 520);
        splitContainer1.SplitterDistance = 550;
        splitContainer1.TabIndex = 0;
        //
        // dgvGames
        //
        dgvGames.AllowUserToAddRows = false;
        dgvGames.AllowUserToDeleteRows = false;
        dgvGames.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        dgvGames.ContextMenuStrip = dgvContextMenu;
        dgvGames.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvGames.Location = new Point(3, 23);
        dgvGames.MultiSelect = true;
        dgvGames.Name = "dgvGames";
        dgvGames.ReadOnly = true;
        dgvGames.RowHeadersVisible = false;
        dgvGames.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvGames.Size = new Size(544, 494);
        dgvGames.TabIndex = 1;
        dgvGames.SelectionChanged += dgvGames_SelectionChanged;
        dgvGames.DoubleClick += dgvGames_DoubleClick;
        dgvGames.CellMouseDown += dgvGames_CellMouseDown;
        //
        // lblSearch
        //
        lblSearch.AutoSize = true;
        lblSearch.Location = new Point(3, 3);
        lblSearch.Name = "lblSearch";
        lblSearch.Size = new Size(85, 15);
        lblSearch.TabIndex = 2;
        lblSearch.Text = "Games (0 of 0)";
        //
        // tabDetails
        //
        tabDetails.Controls.Add(tabInfo);
        tabDetails.Controls.Add(tabScreenshots);
        tabDetails.Controls.Add(tabExtras);
        tabDetails.Dock = DockStyle.Fill;
        tabDetails.Location = new Point(0, 0);
        tabDetails.Name = "tabDetails";
        tabDetails.SelectedIndex = 0;
        tabDetails.Size = new Size(440, 520);
        tabDetails.TabIndex = 0;
        //
        // tabInfo
        //
        tabInfo.Controls.Add(rtbGameInfo);
        tabInfo.Location = new Point(4, 24);
        tabInfo.Name = "tabInfo";
        tabInfo.Padding = new Padding(3);
        tabInfo.Size = new Size(432, 492);
        tabInfo.TabIndex = 0;
        tabInfo.Text = "Game Info";
        tabInfo.UseVisualStyleBackColor = true;
        //
        // rtbGameInfo
        //
        rtbGameInfo.Dock = DockStyle.Fill;
        rtbGameInfo.Location = new Point(3, 3);
        rtbGameInfo.Name = "rtbGameInfo";
        rtbGameInfo.ReadOnly = true;
        rtbGameInfo.Size = new Size(426, 486);
        rtbGameInfo.TabIndex = 0;
        rtbGameInfo.Text = "";
        //
        // tabScreenshots
        //
        tabScreenshots.Controls.Add(picScreenshot);
        tabScreenshots.Location = new Point(4, 24);
        tabScreenshots.Name = "tabScreenshots";
        tabScreenshots.Padding = new Padding(3);
        tabScreenshots.Size = new Size(432, 492);
        tabScreenshots.TabIndex = 1;
        tabScreenshots.Text = "Screenshots";
        tabScreenshots.UseVisualStyleBackColor = true;
        //
        // picScreenshot
        //
        picScreenshot.ContextMenuStrip = screenshotContextMenu;
        picScreenshot.Dock = DockStyle.Fill;
        picScreenshot.Location = new Point(3, 3);
        picScreenshot.Name = "picScreenshot";
        picScreenshot.Size = new Size(426, 486);
        picScreenshot.SizeMode = PictureBoxSizeMode.Zoom;
        picScreenshot.TabIndex = 0;
        picScreenshot.TabStop = false;
        //
        // dgvContextMenu
        //
        dgvContextMenu.Items.AddRange(new ToolStripItem[] { toggleFavoriteToolStripMenuItem, toggleClassicToolStripMenuItem, toggleAdultToolStripMenuItem });
        //
        // toggleFavoriteToolStripMenuItem
        //
        toggleFavoriteToolStripMenuItem.Name = "toggleFavoriteToolStripMenuItem";
        toggleFavoriteToolStripMenuItem.Size = new Size(159, 22);
        toggleFavoriteToolStripMenuItem.Text = "Toggle Favorite";
        toggleFavoriteToolStripMenuItem.Click += toggleFavoriteToolStripMenuItem_Click;
        //
        // toggleClassicToolStripMenuItem
        //
        toggleClassicToolStripMenuItem.Name = "toggleClassicToolStripMenuItem";
        toggleClassicToolStripMenuItem.Size = new Size(159, 22);
        toggleClassicToolStripMenuItem.Text = "Toggle Classic";
        toggleClassicToolStripMenuItem.Click += toggleClassicToolStripMenuItem_Click;
        //
        // toggleAdultToolStripMenuItem
        //
        toggleAdultToolStripMenuItem.Name = "toggleAdultToolStripMenuItem";
        toggleAdultToolStripMenuItem.Size = new Size(159, 22);
        toggleAdultToolStripMenuItem.Text = "Toggle Adult";
        toggleAdultToolStripMenuItem.Click += toggleAdultToolStripMenuItem_Click;

        // screenshotContextMenu
        screenshotContextMenu.Items.AddRange(new ToolStripItem[] { selectScreenshotToolStripMenuItem, clearScreenshotToolStripMenuItem });
        screenshotContextMenu.Name = "screenshotContextMenu";
        screenshotContextMenu.Size = new Size(171, 48);

        // selectScreenshotToolStripMenuItem
        selectScreenshotToolStripMenuItem.Name = "selectScreenshotToolStripMenuItem";
        selectScreenshotToolStripMenuItem.Size = new Size(170, 22);
        selectScreenshotToolStripMenuItem.Text = "Select Screenshot";
        selectScreenshotToolStripMenuItem.Click += selectScreenshotToolStripMenuItem_Click;

        // clearScreenshotToolStripMenuItem
        clearScreenshotToolStripMenuItem.Name = "clearScreenshotToolStripMenuItem";
        clearScreenshotToolStripMenuItem.Size = new Size(170, 22);
        clearScreenshotToolStripMenuItem.Text = "Clear Screenshot";
        clearScreenshotToolStripMenuItem.Click += clearScreenshotToolStripMenuItem_Click;

        // tabExtras
        //
        tabExtras.Controls.Add(lstExtras);
        tabExtras.Location = new Point(4, 24);
        tabExtras.Name = "tabExtras";
        tabExtras.Size = new Size(432, 492);
        tabExtras.TabIndex = 2;
        tabExtras.Text = "Extras";
        tabExtras.UseVisualStyleBackColor = true;
        //
        // lstExtras
        //
        lstExtras.Dock = DockStyle.Fill;
        lstExtras.FormattingEnabled = true;
        lstExtras.ItemHeight = 15;
        lstExtras.Location = new Point(0, 0);
        lstExtras.Name = "lstExtras";
        lstExtras.Size = new Size(432, 492);
        lstExtras.TabIndex = 0;
        lstExtras.DoubleClick += lstExtras_DoubleClick;
        //
        // lblStatus
        //
        lblStatus.AutoSize = true;
        lblStatus.Dock = DockStyle.Fill;
        lblStatus.Location = new Point(3, 596);
        lblStatus.Name = "lblStatus";
        lblStatus.Size = new Size(994, 30);
        lblStatus.TabIndex = 1;
        lblStatus.Text = "Ready";
        lblStatus.TextAlign = ContentAlignment.MiddleLeft;
        //
        // txtSearch
        //
        flowSearch.AutoSize = true;
        flowSearch.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        flowSearch.Dock = DockStyle.Fill;
        flowSearch.FlowDirection = FlowDirection.LeftToRight;
        flowSearch.Controls.Add(txtSearch);
        flowSearch.Controls.Add(btnClearFilters);

        txtSearch.Width = 400;
        txtSearch.PlaceholderText = "Search games...";
        txtSearch.TextChanged += txtSearch_TextChanged;

        btnClearFilters.AutoSize = true;
        btnClearFilters.Text = "Clear";
        btnClearFilters.Click += clearFiltersToolStripMenuItem_Click;
        //
        // Form1
        //
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1000, 650);
        Controls.Add(tableLayoutPanel1);
        Controls.Add(menuStrip1);
        MainMenuStrip = menuStrip1;
        Name = "Form1";
        Text = "GameBase";
        menuStrip1.ResumeLayout(false);
        menuStrip1.PerformLayout();
        splitContainer1.Panel1.ResumeLayout(false);
        splitContainer1.Panel1.PerformLayout();
        splitContainer1.Panel2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
        splitContainer1.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvGames).EndInit();
        tabDetails.ResumeLayout(false);
        tabInfo.ResumeLayout(false);
        tabScreenshots.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)picScreenshot).EndInit();
        dgvContextMenu.ResumeLayout(false);
        screenshotContextMenu.ResumeLayout(false);
        tabExtras.ResumeLayout(false);
        tableLayoutPanel1.ResumeLayout(false);
        tableLayoutPanel1.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
}
