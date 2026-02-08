using GBFuncs;
using GBSql.Data;
using Microsoft.Data.Sqlite;
using System.Diagnostics;

namespace GBToolbox;

public partial class Form1 : Form
{
    private readonly RegistryService _registry = new();
    private readonly FunctionsHelper _functions = new();
    private bool _expertMode;

    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object? sender, EventArgs e)
    {
        ApplyWindowSettings();
        ParseCommandLine();
        RefreshTree();
        if (!_expertMode)
        {
            MessageBox.Show("Please make sure the GameBase frontend is NOT running when using this tool.", "GameBase", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
    {
        SaveWindowSettings();
    }

    private void ParseCommandLine()
    {
        var args = Environment.GetCommandLineArgs().Select(a => a.ToUpperInvariant()).ToArray();
        _expertMode = args.Contains("-EXPERT");
        if (args.Contains("-NOMSGBOX"))
        {
            _expertMode = true;
        }
    }

    private void ApplyWindowSettings()
    {
        _registry.GetSetting_HKCU_Fallback_HKLM("Software\\GB64\\GBToolbox\\Window", "MainWidth", out var widthValue, 4770);
        _registry.GetSetting_HKCU_Fallback_HKLM("Software\\GB64\\GBToolbox\\Window", "MainHeight", out var heightValue, 4230);
        _registry.GetSetting_HKCU_Fallback_HKLM("Software\\GB64\\GBToolbox\\Window", "MainTop", out var topValue, 120);
        _registry.GetSetting_HKCU_Fallback_HKLM("Software\\GB64\\GBToolbox\\Window", "MainLeft", out var leftValue, 120);

        int width = Convert.ToInt32(widthValue ?? 900);
        int height = Convert.ToInt32(heightValue ?? 600);
        int top = Convert.ToInt32(topValue ?? 120);
        int left = Convert.ToInt32(leftValue ?? 120);

        var primaryScreen = Screen.PrimaryScreen ?? Screen.AllScreens.FirstOrDefault();
        int maxWidth = primaryScreen?.Bounds.Width ?? 1920;
        int maxHeight = primaryScreen?.Bounds.Height ?? 1080;

        width = Math.Clamp(width, 600, maxWidth);
        height = Math.Clamp(height, 400, maxHeight);
        if (top < 0 || left < 0)
        {
            top = 120;
            left = 120;
        }

        StartPosition = FormStartPosition.Manual;
        Width = width;
        Height = height;
        Top = top;
        Left = left;
    }

    private void SaveWindowSettings()
    {
        _registry.SaveSetting("HKEY_CURRENT_USER", "Software\\GB64\\GBToolbox\\Window", "MainWidth", Width, false);
        _registry.SaveSetting("HKEY_CURRENT_USER", "Software\\GB64\\GBToolbox\\Window", "MainHeight", Height, false);
        _registry.SaveSetting("HKEY_CURRENT_USER", "Software\\GB64\\GBToolbox\\Window", "MainTop", Top, false);
        _registry.SaveSetting("HKEY_CURRENT_USER", "Software\\GB64\\GBToolbox\\Window", "MainLeft", Left, false);
    }

    private void RefreshTree()
    {
        tvGameBase.Nodes.Clear();

        var root = new TreeNode("Gamebases") { Tag = new NodeInfo(NodeKind.Root, ToolboxPaths.GetGameBaseRoot(), ToolboxPaths.GetGameBaseRoot()) };
        tvGameBase.Nodes.Add(root);

        string basePath = ToolboxPaths.GetGameBaseRoot();
        if (Directory.Exists(basePath))
        {
            foreach (string dir in Directory.GetDirectories(basePath))
            {
                var folderName = new DirectoryInfo(dir).Name;
                if (folderName.Equals("DOCS", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string dbFile = ToolboxPaths.GetDatabaseFile(dir);
                if (!File.Exists(dbFile))
                {
                    continue;
                }

                var gbNode = new TreeNode(folderName) { Tag = new NodeInfo(NodeKind.GameBase, dir, dir) };
                root.Nodes.Add(gbNode);

                gbNode.Nodes.Add(new TreeNode(Path.GetFileName(dbFile))
                {
                    Tag = new NodeInfo(NodeKind.Database, dbFile, dir)
                });

                gbNode.Nodes.Add(new TreeNode("INI Files") { Tag = new NodeInfo(NodeKind.IniFolder, dir, dir) });
                gbNode.Nodes.Add(new TreeNode("GEMUS Scripts") { Tag = new NodeInfo(NodeKind.ScriptFolder, ToolboxPaths.GetScriptsFolder(dir), dir) });
                gbNode.Nodes.Add(new TreeNode("Text Files") { Tag = new NodeInfo(NodeKind.TextFolder, dir, dir) });

                PopulateFiles(gbNode, dir);
            }
        }

        root.Expand();
        statusLabel.Text = "Right-click each node for options";
    }

    private void PopulateFiles(TreeNode gbNode, string baseFolder)
    {
        var iniFolder = gbNode.Nodes.Cast<TreeNode>().FirstOrDefault(n => ((NodeInfo)n.Tag!).Kind == NodeKind.IniFolder);
        var scriptFolder = gbNode.Nodes.Cast<TreeNode>().FirstOrDefault(n => ((NodeInfo)n.Tag!).Kind == NodeKind.ScriptFolder);
        var textFolder = gbNode.Nodes.Cast<TreeNode>().FirstOrDefault(n => ((NodeInfo)n.Tag!).Kind == NodeKind.TextFolder);

        if (iniFolder != null)
        {
            foreach (string ini in Directory.GetFiles(baseFolder, "*.ini"))
            {
                iniFolder.Nodes.Add(new TreeNode(Path.GetFileName(ini))
                {
                    Tag = new NodeInfo(NodeKind.IniFile, ini, baseFolder)
                });
            }
        }

        if (textFolder != null)
        {
            foreach (string txt in Directory.GetFiles(baseFolder, "*.txt"))
            {
                textFolder.Nodes.Add(new TreeNode(Path.GetFileName(txt))
                {
                    Tag = new NodeInfo(NodeKind.TextFile, txt, baseFolder)
                });
            }
        }

        string scriptsPath = ToolboxPaths.GetScriptsFolder(baseFolder);
        if (!Directory.Exists(scriptsPath))
        {
            Directory.CreateDirectory(scriptsPath);
        }

        if (scriptFolder != null)
        {
            foreach (string txt in Directory.GetFiles(scriptsPath, "*.txt"))
            {
                scriptFolder.Nodes.Add(new TreeNode(Path.GetFileName(txt))
                {
                    Tag = new NodeInfo(NodeKind.ScriptFile, txt, baseFolder)
                });
            }
        }
    }

    private void tvGameBase_NodeMouseClick(object? sender, TreeNodeMouseClickEventArgs e)
    {
        tvGameBase.SelectedNode = e.Node;
        var info = e.Node.Tag as NodeInfo;
        if (info == null)
        {
            return;
        }

        statusLabel.Text = info.Path;
        if (e.Button == MouseButtons.Right)
        {
            switch (info.Kind)
            {
                case NodeKind.Root:
                    ctxRoot.Show(tvGameBase, e.Location);
                    break;
                case NodeKind.GameBase:
                    ctxGameBase.Show(tvGameBase, e.Location);
                    break;
                case NodeKind.Database:
                    ctxDatabase.Show(tvGameBase, e.Location);
                    break;
                case NodeKind.ScriptFolder:
                    ctxScriptRoot.Show(tvGameBase, e.Location);
                    break;
            }
        }
    }

    private void tvGameBase_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
    {
        var info = e.Node.Tag as NodeInfo;
        if (info == null)
        {
            return;
        }

        if (info.Kind == NodeKind.Database)
        {
            if (!_expertMode)
            {
                MessageBox.Show("WARNING: Only EXPERTS should change the GameBase database directly!", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            ShellOpen(info.Path);
        }
        else if (info.Kind == NodeKind.IniFile || info.Kind == NodeKind.TextFile || info.Kind == NodeKind.ScriptFile)
        {
            ShellOpen(info.Path);
        }
    }

    private void ShellOpen(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            MessageBox.Show("Unable to open file.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private NodeInfo? GetSelectedNodeInfo()
    {
        return tvGameBase.SelectedNode?.Tag as NodeInfo;
    }

    private async void mnuFileNew_Click(object? sender, EventArgs e)
    {
        using var input = new InputDialog("Create New GameBase", "Enter the new GameBase name:", string.Empty, "\\/:*?\"<>|", 64, false);
        if (input.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        string name = input.ResultText;
        string baseFolder = Path.Combine(ToolboxPaths.GetGameBaseRoot(), name);
        if (Directory.Exists(baseFolder))
        {
            MessageBox.Show("There is already a GameBase folder with that name!", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        Directory.CreateDirectory(baseFolder);
        Directory.CreateDirectory(Path.Combine(baseFolder, "Gfx"));
        Directory.CreateDirectory(Path.Combine(baseFolder, "Scripts"));
        Directory.CreateDirectory(Path.Combine(baseFolder, "Sounds"));
        Directory.CreateDirectory(Path.Combine(baseFolder, "Splash"));

        ToolboxResourceLoader.EnsureDefaultFiles(baseFolder);
        ToolboxResourceLoader.EnsureDefaultFiles(Path.Combine(baseFolder, "Gfx"));

        File.WriteAllText(Path.Combine(baseFolder, "Credits.txt"), $"Created by GBToolbox on {DateTime.Now:dddd dd MMMM yyyy}.\n");
        File.WriteAllText(Path.Combine(baseFolder, "Sounds", "sounds.txt"), "Place your Classic Game WAV files here.");
        File.WriteAllText(Path.Combine(baseFolder, "Splash", "Splash.txt"), "Place your splash screens here.");

        string dbPath = Path.Combine(baseFolder, name + ".db");
        var context = new GameBaseSqliteContext(dbPath);
        await context.InitializeDatabaseAsync();
        context.Dispose();

        RefreshTree();
    }

    private void mnuFileImportRc2_Click(object? sender, EventArgs e)
    {
        StartRomCenterImport(existingDbPath: null);
    }

    private void mnuFileExit_Click(object? sender, EventArgs e)
    {
        Close();
    }

    private void ctxGameBaseDelete_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.GameBase)
        {
            return;
        }

        if (MessageBox.Show($"Are you sure you want to delete '{info.Path}'?", "GBToolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
        {
            return;
        }

        try
        {
            Directory.Delete(info.Path, recursive: true);
            RefreshTree();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to delete folder: {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ctxDbCompact_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        using var wait = new WaitForm("Compacting...");
        wait.Show(this);
        Application.DoEvents();

        try
        {
            long oldSize = new FileInfo(info.Path).Length;
            using var conn = ToolboxDatabase.Open(info.Path);
            ToolboxDatabase.ExecuteNonQuery(conn, "VACUUM;");
            long newSize = new FileInfo(info.Path).Length;
            MessageBox.Show($"Compacted database from {oldSize / 1024}Kb to {newSize / 1024}Kb.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Unable to compact database. {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            wait.Close();
        }
    }

    private void ctxDbFix_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        using var wait = new WaitForm("Removing NULL data...");
        wait.Show(this);
        Application.DoEvents();

        try
        {
            using var conn = ToolboxDatabase.Open(info.Path);
            ReplaceNullData(conn);
            MessageBox.Show("Fixed: All NULL data removed.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Unable to fix data: {ex.Message}", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            wait.Close();
        }
    }

    private void ReplaceNullData(SqliteConnection conn)
    {
        var updates = new List<(string Table, string Column, string Value)>
        {
            ("Games", "Name", ""),
            ("Games", "Filename", ""),
            ("Games", "ScreenshotFilename", ""),
            ("Games", "Comment", ""),
            ("Games", "CCode", ""),
            ("Extras", "Name", ""),
            ("Extras", "Filename", "")
        };

        foreach (var update in updates)
        {
            if (!ToolboxDatabase.TableExists(conn, update.Table) || !ToolboxDatabase.ColumnExists(conn, update.Table, update.Column))
            {
                continue;
            }

            ToolboxDatabase.ExecuteNonQuery(conn, $"UPDATE {update.Table} SET {update.Column}=@value WHERE {update.Column} IS NULL",
                new SqliteParameter("@value", update.Value));
        }
    }

    private void ctxDbImportRc2_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        StartRomCenterImport(info.Path);
    }

    private void ctxDbSettings_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        using var dialog = new DatabaseSettingsForm(info.Path);
        dialog.ShowDialog(this);
    }

    private void ctxDbGenres_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        using var dialog = new GenresWizardForm(info.Path);
        dialog.ShowDialog(this);
    }

    private void ctxDbClearExtras_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        if (MessageBox.Show("Do you really want to clear all Game Extras?", "GBToolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
        {
            return;
        }

        using var conn = ToolboxDatabase.Open(info.Path);
        if (ToolboxDatabase.TableExists(conn, "Extras"))
        {
            ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM Extras");
        }
        if (ToolboxDatabase.TableExists(conn, "Games") && ToolboxDatabase.ColumnExists(conn, "Games", "Extras"))
        {
            ToolboxDatabase.ExecuteNonQuery(conn, "UPDATE Games SET Extras=0");
        }
        MessageBox.Show("Cleared all Game Extras.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ctxDbRemoveViews_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        if (MessageBox.Show("Do you really want to delete all custom views from this database?", "GBToolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
        {
            return;
        }

        using var conn = ToolboxDatabase.Open(info.Path);
        if (ToolboxDatabase.TableExists(conn, "ViewData"))
        {
            ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM ViewData");
        }
        if (ToolboxDatabase.TableExists(conn, "ViewFilters"))
        {
            ToolboxDatabase.ExecuteNonQuery(conn, "DELETE FROM ViewFilters");
        }

        string dbName = Path.GetFileNameWithoutExtension(info.Path);
        _registry.SaveSetting("HKEY_CURRENT_USER", $"Software\\GB64\\GameBase\\GameBase\\{dbName}\\QuickFilter", "NumFilters", 0, false);
        MessageBox.Show("Removed all custom views.", "GBToolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ctxDbDeleteRecords_Click(object? sender, EventArgs e)
    {
        var info = GetSelectedNodeInfo();
        if (info == null || info.Kind != NodeKind.Database)
        {
            return;
        }

        using var dialog = new RemoveRecordsForm(info.Path);
        dialog.ShowDialog(this);
    }

    private void StartRomCenterImport(string? existingDbPath)
    {
        using var openFile = new OpenFileDialog
        {
            Filter = "RomCenter 2.0 Datafiles (*.dat)|*.dat|All Files (*.*)|*.*",
            Title = "Select the RC-2.0 Datafile"
        };

        if (openFile.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        using var dialog = new RomCenterImportForm(openFile.FileName, existingDbPath);
        dialog.ShowDialog(this);
        RefreshTree();
    }

    private sealed record NodeInfo(NodeKind Kind, string Path, string BaseFolder);

    private enum NodeKind
    {
        Root,
        GameBase,
        Database,
        IniFolder,
        ScriptFolder,
        TextFolder,
        IniFile,
        TextFile,
        ScriptFile
    }
}
