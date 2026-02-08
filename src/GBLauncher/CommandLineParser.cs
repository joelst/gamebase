using System.Text.RegularExpressions;

namespace GBLauncher;

/// <summary>
/// Parses command-line arguments for GBLauncher.
/// </summary>
internal static class CommandLineParser
{
    public static LauncherOptions Parse(string[] args)
    {
        var options = new LauncherOptions();
        var commandLine = string.Join(" ", args);

        // db= or database=
        var dbMatch = Regex.Match(commandLine, @"(db|database)=(""([^""]+)""|([^"" ]+))", RegexOptions.IgnoreCase);
        if (dbMatch.Success)
        {
            options.DatabaseName = dbMatch.Groups[3].Success ? dbMatch.Groups[3].Value : dbMatch.Groups[4].Value;
        }

        // id= or game= or gameid=
        var idMatch = Regex.Match(commandLine, @"(id|game|gameid)=(""([0-9]+)""|([0-9]+))", RegexOptions.IgnoreCase);
        if (idMatch.Success)
        {
            options.GameId = idMatch.Groups[3].Success ? idMatch.Groups[3].Value : idMatch.Groups[4].Value;
        }

        // fn= or filename=
        var fnMatch = Regex.Match(commandLine, @"(fn|filename)=(""([^""]+)""|([^"" ]+))", RegexOptions.IgnoreCase);
        if (fnMatch.Success)
        {
            options.GameFilename = fnMatch.Groups[3].Success ? fnMatch.Groups[3].Value : fnMatch.Groups[4].Value;
        }

        // emu= or emulator=
        var emuMatch = Regex.Match(commandLine, @"(emu|emulator)=(""([0-9]+)""|([0-9]+))", RegexOptions.IgnoreCase);
        if (emuMatch.Success)
        {
            var emuValue = emuMatch.Groups[3].Success ? emuMatch.Groups[3].Value : emuMatch.Groups[4].Value;
            options.EmulatorIndex = int.Parse(emuValue);
        }

        // wait=
        var waitMatch = Regex.Match(commandLine, @"(wait)=(""([01])""|([01]))", RegexOptions.IgnoreCase);
        if (waitMatch.Success)
        {
            var waitValue = waitMatch.Groups[3].Success ? waitMatch.Groups[3].Value : waitMatch.Groups[4].Value;
            options.WaitForExit = waitValue == "1";
        }

        return options;
    }

    public static void ShowUsage()
    {
        Console.WriteLine();
        Console.WriteLine("GameBase Launcher v1.0");
        Console.WriteLine("Copyright (C) 1997-2026 James Burrows");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  GBLauncher.exe db=<db name> [id=<game id>|fn=<game filename>] [emu=<emu index>] [wait=1]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  GBLauncher.exe db=\"GameBase Amiga\" fn=Cannon_Fodder.adf emu=2 wait=1");
        Console.WriteLine();
        Console.WriteLine("Parameters:");
        Console.WriteLine("  db=<db name>         Required. Name of GameBase database containing game to launch.");
        Console.WriteLine("  id=<game id>         Required if 'fn' not specified. GameBase Game ID.");
        Console.WriteLine("  fn=<game filename>   Required if 'id' not specified. Game filename (can be");
        Console.WriteLine("                       path-less or include relative path with/without extension).");
        Console.WriteLine("  emu=<emu index>      Optional. Emulator index from Emulators.ini. Defaults");
        Console.WriteLine("                       to the default emulator.");
        Console.WriteLine("  wait=1               Optional. Wait until emulator is closed before quitting.");
        Console.WriteLine();
    }
}
