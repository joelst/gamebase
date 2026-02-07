using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Gemus;
using Gemus.Models;
using Xunit;

namespace Gemus.Tests
{
  public class ArchiveUnpackerTests
  {
    [Fact]
    public void ManagedZip_Unpack_SingleFile_ReturnsFileToRun()
    {
      string tempRoot = CreateTempRoot();
      string archivePath = Path.Combine(tempRoot, "game.zip");
      CreateZipWithEntry(archivePath, ("game.exe", "hello"));

      string dest = Path.Combine(tempRoot, "dest") + Path.DirectorySeparatorChar;
      var unpacker = new ArchiveUnpacker { Backend = ArchiveBackend.ManagedZip };

      bool ok = unpacker.Unpack(
        archivePath,
        string.Empty,
        new[] { "exe" },
        dest,
        out bool isRepackPossible,
        out string fileToRun,
        out bool userPick,
        out string error);

      Assert.True(ok, error);
      Assert.False(userPick);
      Assert.True(isRepackPossible);
      Assert.False(string.IsNullOrEmpty(fileToRun));
      Assert.True(File.Exists(fileToRun));
      Assert.EndsWith("game.exe", fileToRun, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Managed7Zip_RarExtension_RepackCreatesZip()
    {
      string tempRoot = CreateTempRoot();
      string archivePathRar = Path.Combine(tempRoot, "game.rar");
      // Create a ZIP archive but name it .rar; SharpCompress opens by content
      CreateZipWithEntry(archivePathRar, ("game.exe", "hello"));

      string dest = Path.Combine(tempRoot, "dest") + Path.DirectorySeparatorChar;
      var unpacker = new ArchiveUnpacker { Backend = ArchiveBackend.Managed7Zip };

      bool ok = unpacker.Unpack(
        archivePathRar,
        string.Empty,
        new[] { "exe" },
        dest,
        out bool isRepackPossible,
        out string fileToRun,
        out bool userPick,
        out string error);

      Assert.True(ok, error);
      Assert.False(userPick);
      Assert.False(isRepackPossible);
      Assert.False(string.IsNullOrEmpty(fileToRun));
      Assert.True(File.Exists(fileToRun));

      // Snapshot pre-state, modify file, then repack
      string finalFolder = unpacker.GetFinalFolder();
      unpacker.StorePreUnpackedFilesSnapshot(finalFolder);
      File.AppendAllText(fileToRun, " modified");

      bool repackOk = unpacker.Repack(
        false,
        false,
        dest,
        archivePathRar,
        RezipAdditionalFiles.Ignore,
        out bool wasRepacked,
        out string repackError);

      Assert.True(repackOk, repackError);
      Assert.True(wasRepacked);

      string zipPath = Path.ChangeExtension(archivePathRar, ".zip");
      Assert.True(File.Exists(zipPath));
    }

    private static string CreateTempRoot()
    {
      string path = Path.Combine(Path.GetTempPath(), "gemus-tests", Guid.NewGuid().ToString("N"));
      Directory.CreateDirectory(path);
      return path;
    }

    private static void CreateZipWithEntry(string archivePath, params (string entry, string content)[] entries)
    {
      using var zip = ZipFile.Open(archivePath, ZipArchiveMode.Create);
      foreach (var (entry, content) in entries)
      {
        var e = zip.CreateEntry(entry);
        using var writer = new StreamWriter(e.Open());
        writer.Write(content);
      }
    }
  }
}
