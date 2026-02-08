namespace GameBase;

using GBSql.Data;
using Gemus;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();

        var repo = CreateRepository();
        Application.Run(new Form1(repo));
    }

    private static GameBaseSqliteRepository CreateRepository()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string basePathRoot = string.IsNullOrWhiteSpace(appData)
            ? AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            : appData;
        string basePath = Path.Combine(basePathRoot, "GameBase");
        Directory.CreateDirectory(basePath);
        string gemusIni = Path.Combine(basePath, "Gemus.ini");

        // Env override
        string? dbPath = Environment.GetEnvironmentVariable("GAMEBASE_DB_PATH");
        if (string.IsNullOrWhiteSpace(dbPath))
        {
            dbPath = GemusIniHelper.GetIniMessage("Settings", "DatabasePath", string.Empty, gemusIni);
        }

        if (string.IsNullOrWhiteSpace(dbPath))
        {
            dbPath = Path.Combine(basePath, "Data", "GameBase.db");
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
            // Persist default path for future runs
            GemusIniHelper.WriteIniValue("Settings", "DatabasePath", dbPath, gemusIni);
        }

        return GameBaseSqliteRepository.CreateAsync(dbPath, initializeSchema: true).GetAwaiter().GetResult();
    }
}