namespace GBToolbox;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;
    private MenuStrip mainMenu;
    private ToolStripMenuItem mnuFile;
    private ToolStripMenuItem mnuFileNew;
    private ToolStripMenuItem mnuFileImportRc2;
    private ToolStripSeparator mnuFileSep;
    private ToolStripMenuItem mnuFileExit;
    private TreeView tvGameBase;
    private StatusStrip statusStrip;
    private ToolStripStatusLabel statusLabel;
    private ContextMenuStrip ctxRoot;
    private ToolStripMenuItem ctxRootNew;
    private ToolStripMenuItem ctxRootImportRc2;
    private ContextMenuStrip ctxGameBase;
    private ToolStripMenuItem ctxGameBaseDelete;
    private ContextMenuStrip ctxDatabase;
    private ToolStripMenuItem ctxDbCompact;
    private ToolStripMenuItem ctxDbFix;
    private ToolStripSeparator ctxDbSep1;
    private ToolStripMenuItem ctxDbImportRc2;
    private ToolStripSeparator ctxDbSep2;
    private ToolStripMenuItem ctxDbGenres;
    private ToolStripMenuItem ctxDbSettings;
    private ToolStripSeparator ctxDbSep3;
    private ToolStripMenuItem ctxDbClearExtras;
    private ToolStripMenuItem ctxDbRemoveViews;
    private ToolStripMenuItem ctxDbDeleteRecords;
    private ContextMenuStrip ctxScriptRoot;
    private ToolStripMenuItem ctxScriptNew;

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
        components = new System.ComponentModel.Container();
        mainMenu = new MenuStrip();
        mnuFile = new ToolStripMenuItem();
        mnuFileNew = new ToolStripMenuItem();
        mnuFileImportRc2 = new ToolStripMenuItem();
        mnuFileSep = new ToolStripSeparator();
        mnuFileExit = new ToolStripMenuItem();
        tvGameBase = new TreeView();
        statusStrip = new StatusStrip();
        statusLabel = new ToolStripStatusLabel();
        ctxRoot = new ContextMenuStrip(components);
        ctxRootNew = new ToolStripMenuItem();
        ctxRootImportRc2 = new ToolStripMenuItem();
        ctxGameBase = new ContextMenuStrip(components);
        ctxGameBaseDelete = new ToolStripMenuItem();
        ctxDatabase = new ContextMenuStrip(components);
        ctxDbCompact = new ToolStripMenuItem();
        ctxDbFix = new ToolStripMenuItem();
        ctxDbSep1 = new ToolStripSeparator();
        ctxDbImportRc2 = new ToolStripMenuItem();
        ctxDbSep2 = new ToolStripSeparator();
        ctxDbGenres = new ToolStripMenuItem();
        ctxDbSettings = new ToolStripMenuItem();
        ctxDbSep3 = new ToolStripSeparator();
        ctxDbClearExtras = new ToolStripMenuItem();
        ctxDbRemoveViews = new ToolStripMenuItem();
        ctxDbDeleteRecords = new ToolStripMenuItem();
        ctxScriptRoot = new ContextMenuStrip(components);
        ctxScriptNew = new ToolStripMenuItem();
        mainMenu.SuspendLayout();
        statusStrip.SuspendLayout();
        SuspendLayout();
        //
        // mainMenu
        //
        mainMenu.Items.AddRange(new ToolStripItem[] { mnuFile });
        mainMenu.Location = new Point(0, 0);
        mainMenu.Name = "mainMenu";
        mainMenu.Size = new Size(900, 24);
        mainMenu.TabIndex = 0;
        mainMenu.Text = "menuStrip";
        //
        // mnuFile
        //
        mnuFile.DropDownItems.AddRange(new ToolStripItem[] { mnuFileNew, mnuFileImportRc2, mnuFileSep, mnuFileExit });
        mnuFile.Name = "mnuFile";
        mnuFile.Size = new Size(37, 20);
        mnuFile.Text = "&File";
        //
        // mnuFileNew
        //
        mnuFileNew.Name = "mnuFileNew";
        mnuFileNew.ShortcutKeys = Keys.Control | Keys.N;
        mnuFileNew.Size = new Size(257, 22);
        mnuFileNew.Text = "&New GameBase...";
        mnuFileNew.Click += mnuFileNew_Click;
        //
        // mnuFileImportRc2
        //
        mnuFileImportRc2.Name = "mnuFileImportRc2";
        mnuFileImportRc2.ShortcutKeys = Keys.Control | Keys.R;
        mnuFileImportRc2.Size = new Size(257, 22);
        mnuFileImportRc2.Text = "Import &RomCenter 2.00...";
        mnuFileImportRc2.Click += mnuFileImportRc2_Click;
        //
        // mnuFileExit
        //
        mnuFileExit.Name = "mnuFileExit";
        mnuFileExit.Size = new Size(257, 22);
        mnuFileExit.Text = "E&xit";
        mnuFileExit.Click += mnuFileExit_Click;
        //
        // tvGameBase
        //
        tvGameBase.Dock = DockStyle.Fill;
        tvGameBase.HideSelection = false;
        tvGameBase.Location = new Point(0, 24);
        tvGameBase.Name = "tvGameBase";
        tvGameBase.Size = new Size(900, 526);
        tvGameBase.TabIndex = 1;
        tvGameBase.NodeMouseClick += tvGameBase_NodeMouseClick;
        tvGameBase.NodeMouseDoubleClick += tvGameBase_NodeMouseDoubleClick;
        //
        // statusStrip
        //
        statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel });
        statusStrip.Location = new Point(0, 550);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(900, 22);
        statusStrip.TabIndex = 2;
        //
        // statusLabel
        //
        statusLabel.Name = "statusLabel";
        statusLabel.Size = new Size(187, 17);
        statusLabel.Text = "Right-click each node for options";
        //
        // ctxRoot
        //
        ctxRoot.Items.AddRange(new ToolStripItem[] { ctxRootNew, ctxRootImportRc2 });
        ctxRoot.Name = "ctxRoot";
        //
        // ctxRootNew
        //
        ctxRootNew.Name = "ctxRootNew";
        ctxRootNew.Size = new Size(257, 22);
        ctxRootNew.Text = "&New GameBase...";
        ctxRootNew.Click += mnuFileNew_Click;
        //
        // ctxRootImportRc2
        //
        ctxRootImportRc2.Name = "ctxRootImportRc2";
        ctxRootImportRc2.Size = new Size(257, 22);
        ctxRootImportRc2.Text = "Import &RomCenter 2.00...";
        ctxRootImportRc2.Click += mnuFileImportRc2_Click;
        //
        // ctxGameBase
        //
        ctxGameBase.Items.AddRange(new ToolStripItem[] { ctxGameBaseDelete });
        ctxGameBase.Name = "ctxGameBase";
        //
        // ctxGameBaseDelete
        //
        ctxGameBaseDelete.Name = "ctxGameBaseDelete";
        ctxGameBaseDelete.Size = new Size(128, 22);
        ctxGameBaseDelete.Text = "&Delete...";
        ctxGameBaseDelete.Click += ctxGameBaseDelete_Click;
        //
        // ctxDatabase
        //
        ctxDatabase.Items.AddRange(new ToolStripItem[]
        {
            ctxDbCompact,
            ctxDbFix,
            ctxDbSep1,
            ctxDbImportRc2,
            ctxDbSep2,
            ctxDbGenres,
            ctxDbSettings,
            ctxDbSep3,
            ctxDbClearExtras,
            ctxDbRemoveViews,
            ctxDbDeleteRecords
        });
        ctxDatabase.Name = "ctxDatabase";
        //
        // ctxDbCompact
        //
        ctxDbCompact.Name = "ctxDbCompact";
        ctxDbCompact.Size = new Size(224, 22);
        ctxDbCompact.Text = "&Compact";
        ctxDbCompact.Click += ctxDbCompact_Click;
        //
        // ctxDbFix
        //
        ctxDbFix.Name = "ctxDbFix";
        ctxDbFix.Size = new Size(224, 22);
        ctxDbFix.Text = "&Scan and Fix Errors";
        ctxDbFix.Click += ctxDbFix_Click;
        //
        // ctxDbImportRc2
        //
        ctxDbImportRc2.Name = "ctxDbImportRc2";
        ctxDbImportRc2.Size = new Size(224, 22);
        ctxDbImportRc2.Text = "Import &RomCenter 2.00...";
        ctxDbImportRc2.Click += ctxDbImportRc2_Click;
        //
        // ctxDbGenres
        //
        ctxDbGenres.Name = "ctxDbGenres";
        ctxDbGenres.Size = new Size(224, 22);
        ctxDbGenres.Text = "&Genres Wizard...";
        ctxDbGenres.Click += ctxDbGenres_Click;
        //
        // ctxDbSettings
        //
        ctxDbSettings.Name = "ctxDbSettings";
        ctxDbSettings.Size = new Size(224, 22);
        ctxDbSettings.Text = "&Database Settings...";
        ctxDbSettings.Click += ctxDbSettings_Click;
        //
        // ctxDbClearExtras
        //
        ctxDbClearExtras.Name = "ctxDbClearExtras";
        ctxDbClearExtras.Size = new Size(224, 22);
        ctxDbClearExtras.Text = "Clear all Game &Extras...";
        ctxDbClearExtras.Click += ctxDbClearExtras_Click;
        //
        // ctxDbRemoveViews
        //
        ctxDbRemoveViews.Name = "ctxDbRemoveViews";
        ctxDbRemoveViews.Size = new Size(224, 22);
        ctxDbRemoveViews.Text = "&Remove All Custom Views...";
        ctxDbRemoveViews.Click += ctxDbRemoveViews_Click;
        //
        // ctxDbDeleteRecords
        //
        ctxDbDeleteRecords.Name = "ctxDbDeleteRecords";
        ctxDbDeleteRecords.Size = new Size(224, 22);
        ctxDbDeleteRecords.Text = "Delete Rec&ords from Table...";
        ctxDbDeleteRecords.Click += ctxDbDeleteRecords_Click;
        //
        // ctxScriptRoot
        //
        ctxScriptRoot.Items.AddRange(new ToolStripItem[] { ctxScriptNew });
        ctxScriptRoot.Name = "ctxScriptRoot";
        //
        // ctxScriptNew
        //
        ctxScriptNew.Name = "ctxScriptNew";
        ctxScriptNew.Size = new Size(144, 22);
        ctxScriptNew.Text = "&New Script...";
        ctxScriptNew.Enabled = false;
        //
        // Form1
        //
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(900, 572);
        Controls.Add(tvGameBase);
        Controls.Add(statusStrip);
        Controls.Add(mainMenu);
        MainMenuStrip = mainMenu;
        Name = "Form1";
        Text = "GameBase Toolbox";
        Load += Form1_Load;
        FormClosing += Form1_FormClosing;
        mainMenu.ResumeLayout(false);
        mainMenu.PerformLayout();
        statusStrip.ResumeLayout(false);
        statusStrip.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
}
