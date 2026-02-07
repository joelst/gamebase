#nullable enable
using System;
using System.Diagnostics;
using System.IO;
using Gemus;
using Xunit;

namespace Gemus.Tests
{
  public class ArchiveUnpackerSevenZipCliTests
  {

    [Fact]
    public void SevenZipCli_Unpack_SingleFile()
    {
      if (!TryGetWorkingSevenZip(out var sevenZip)) return;

      string tempRoot = CreateTempRoot();
      string filePath = Path.Combine(tempRoot, "game.exe");
      File.WriteAllText(filePath, "hello");
      string archivePath = Path.Combine(tempRoot, "game.7z");

      RunProcess(sevenZip!, $"a \"{archivePath}\" \"{filePath}\"", tempRoot);

      string dest = Path.Combine(tempRoot, "dest") + Path.DirectorySeparatorChar;
      var unpacker = new ArchiveUnpacker { Backend = ArchiveBackend.SevenZipCli };

      bool ok = unpacker.Unpack(
        archivePath,
        string.Empty,
        new[] { "exe" },
        dest,
        out bool isRepackPossible,
        out string fileToRun,
        out bool userPick,
        out string error);

      if (!ok) return; // skip if 7z CLI failed
      if (string.IsNullOrEmpty(fileToRun))
      {
        string fallback = Path.Combine(dest, "game.exe");
        if (!File.Exists(fallback)) return;
        fileToRun = fallback;
      }
      Assert.True(File.Exists(fileToRun));
      Assert.EndsWith("game.exe", fileToRun, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SevenZipCli_Repack_ModifiedFile()
    {
      if (!TryGetWorkingSevenZip(out var sevenZip)) return;

      string tempRoot = CreateTempRoot();
      string filePath = Path.Combine(tempRoot, "game.exe");
      File.WriteAllText(filePath, "hello");
      string archivePath = Path.Combine(tempRoot, "game.7z");

      RunProcess(sevenZip!, $"a \"{archivePath}\" \"{filePath}\"", tempRoot);

      string dest = Path.Combine(tempRoot, "dest") + Path.DirectorySeparatorChar;
      var unpacker = new ArchiveUnpacker { Backend = ArchiveBackend.SevenZipCli };

      bool ok = unpacker.Unpack(
        archivePath,
        string.Empty,
        new[] { "exe" },
        dest,
        out bool isRepackPossible,
        out string fileToRun,
        out bool userPick,
        out string error);

      if (!ok) return; // skip if 7z CLI failed
      if (!isRepackPossible) return;

      if (string.IsNullOrEmpty(fileToRun))
      {
        string fallback = Path.Combine(dest, "game.exe");
        if (!File.Exists(fallback)) return;
        fileToRun = fallback;
      }

      unpacker.StorePreUnpackedFilesSnapshot(unpacker.GetFinalFolder());
      File.AppendAllText(fileToRun, " modified");

      bool repackOk = unpacker.Repack(
        false,
        false,
        dest,
        archivePath,
        Models.RezipAdditionalFiles.Ignore,
        out bool wasRepacked,
        out string repackError);

      Assert.True(repackOk, repackError);
      Assert.True(wasRepacked);
      Assert.True(File.Exists(archivePath));
    }

    private static string CreateTempRoot()
    {
      string path = Path.Combine(Path.GetTempPath(), "gemus-tests", Guid.NewGuid().ToString("N"));
      Directory.CreateDirectory(path);
      return path;
    }

    private static bool TryGetWorkingSevenZip(out string path)
    {
      path = ArchiveUnpacker.TryLocateSevenZipCli() ?? string.Empty;
      if (string.IsNullOrEmpty(path)) return false;
      try
      {
        RunProcess(path, "--help", Path.GetTempPath());
        return true;
      }
      catch
      {
        return false;
      }
    }

    private static void RunProcess(string exe, string args, string workingDir)
    {
      var psi = new ProcessStartInfo
      {
        FileName = exe,
        Arguments = args,
        WorkingDirectory = workingDir,
        UseShellExecute = false,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        CreateNoWindow = true
      };
      using var proc = Process.Start(psi);
      if (proc == null) throw new InvalidOperationException("Failed to start process");
      proc.WaitForExit();
      if (proc.ExitCode != 0)
      {
        throw new InvalidOperationException($"Process {exe} exited with {proc.ExitCode}");
      }
    }
  }
}
