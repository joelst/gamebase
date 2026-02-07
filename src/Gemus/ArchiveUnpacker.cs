// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Archive unpacking and repacking helper
// Converted from VB6 clsUnpack usage with 7-Zip CLI.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using SharpCompress.Archives;
using SharpCompress.Common;
using SharpCompress.Writers;

namespace Gemus
{
  public enum ArchiveBackend
  {
    SevenZipCli,
    ManagedZip,
    Managed7Zip
  }

  public class ArchiveUnpacker
  {
    private const string GemusMarkerFile = "__gemus";
    private const string EnvArchiveBackend = "GEMUS_ARCHIVE_BACKEND";
    private readonly CrcCalculator _crc = new();
    private readonly List<SnapshotEntry> _preSnapshot = new();
    private readonly List<SnapshotEntry> _postSnapshot = new();
    private readonly List<string> _recursiveSnapshot = new();
    private readonly string[] _unpackableFileTypes =
    {
      "zip", "7z", "gz", "gzip", "tgz", "bz2", "bzip2", "tbz2", "tbz", "tar", "lzma", "rar", "cab", "arj", "z", "taz", "iso", "lzh", "lha"
    };
    private readonly string[] _repackableFileTypes =
    {
      "7z", "zip", "gz", "gzip", "tgz", "bz2", "bzip2", "tbz2", "tbz", "tar"
    };

    private int _finalFolderIndex = -1;
    private string _finalFolder = string.Empty;
    private string _runnableRelativePathFile = string.Empty;
    private readonly string _appPathWbs;
    private ArchiveBackend _backendUsed = ArchiveBackend.SevenZipCli;

    public Func<string, IReadOnlyList<string>, string?>? RunnableFileSelector { get; set; }
    public Func<ArchiveBackend>? BackendSelector { get; set; }
    public ArchiveBackend Backend { get; set; } = ArchiveBackend.SevenZipCli;

    public ArchiveUnpacker()
    {
      _appPathWbs = AppContext.BaseDirectory;
      if (!_appPathWbs.EndsWith("\\", StringComparison.Ordinal))
        _appPathWbs += "\\";

      Backend = GetDefaultBackend();
    }

    public bool Unpack(string gamePathFile, string defaultFile, IReadOnlyList<string> runnableFileTypes,
      string destFolderWithBackSlash, out bool isRepackPossible, out string fileToRun,
      out bool userPick, out string error)
    {
      isRepackPossible = false;
      userPick = false;
      error = string.Empty;
      fileToRun = string.Empty;
      _finalFolderIndex = -1;
      _finalFolder = string.Empty;
      _runnableRelativePathFile = string.Empty;

      if (runnableFileTypes == null || runnableFileTypes.Count == 0)
      {
        error = "No configured filetypes";
        return false;
      }

      if (!File.Exists(gamePathFile))
      {
        error = "Game file not found.";
        return false;
      }

      if (!IsFileMatchingFileTypes(gamePathFile, runnableFileTypes) && !IsFileMatchingFileTypes(gamePathFile, _unpackableFileTypes))
      {
        error = "Specified file type is not supported";
        return false;
      }

      string[] runnableChain = Array.Empty<string>();
      bool hasRunnableChain = false;
      if (!string.IsNullOrWhiteSpace(defaultFile))
      {
        if (!IsFileMatchingFileTypes(defaultFile, runnableFileTypes))
        {
          error = "Specified runnable file does not match the configured filetypes";
          return false;
        }

        runnableChain = defaultFile.Trim().Split(':', StringSplitOptions.RemoveEmptyEntries);
        hasRunnableChain = runnableChain.Length > 0;
      }

      string archivePathFile = gamePathFile;
      if (!hasRunnableChain && IsFileMatchingFileTypes(archivePathFile, runnableFileTypes))
      {
        fileToRun = archivePathFile;
        return true;
      }

      if (!IsFileMatchingFileTypes(archivePathFile, _unpackableFileTypes))
      {
        error = "Specified archive type is not supported";
        return false;
      }

      ArchiveBackend backend = BackendSelector?.Invoke() ?? Backend;
      _backendUsed = backend;

      isRepackPossible = IsFileMatchingFileTypes(archivePathFile, _repackableFileTypes);
      if (backend == ArchiveBackend.Managed7Zip)
      {
        if (!IsSharpCompressRepackSupported(archivePathFile))
          isRepackPossible = false;
      }
      if (hasRunnableChain && isRepackPossible)
      {
        for (int i = 0; i < runnableChain.Length - 1; i++)
        {
          if (!IsFileMatchingFileTypes(runnableChain[i], _repackableFileTypes))
          {
            isRepackPossible = false;
            break;
          }
          if (backend == ArchiveBackend.Managed7Zip && !IsSharpCompressRepackSupported(runnableChain[i]))
          {
            isRepackPossible = false;
            break;
          }
        }
      }

      string? sevenZip = null;
      if (backend == ArchiveBackend.SevenZipCli)
      {
        sevenZip = GetExtractorPath("7z.exe");
        if (string.IsNullOrEmpty(sevenZip))
        {
          error = "Unable to locate 7z.exe for unpacking archives";
          return false;
        }
      }

      string[] generatedRunnable = Array.Empty<string>();
      _finalFolderIndex = 0;
      while (true)
      {
        (string archivePath, string archiveFile) = SplitPath(archivePathFile);
        _finalFolder = Path.Combine(destFolderWithBackSlash, _finalFolderIndex.ToString());
        SafeDeleteDirectory(_finalFolder);
        Directory.CreateDirectory(_finalFolder);

        if (!UnpackArchive(backend, sevenZip, archivePath, archiveFile, _finalFolder, out error))
        {
          return false;
        }

        if (hasRunnableChain)
        {
          string runnable = runnableChain[_finalFolderIndex];
          string runnablePathFile = Path.Combine(_finalFolder, runnable);
          if (!File.Exists(runnablePathFile))
          {
            error = "Expected file not found in archive: " + runnable;
            return false;
          }

          if (_finalFolderIndex == runnableChain.Length - 1)
          {
            fileToRun = runnablePathFile;
            _runnableRelativePathFile = runnable;
            return true;
          }

          WriteGemusMarker(_finalFolder, runnable);
          _finalFolderIndex++;
          archivePathFile = runnablePathFile;
          continue;
        }

        var files = ListArchive(backend, sevenZip, archivePathFile, out error);
        if (!string.IsNullOrEmpty(error))
        {
          return false;
        }
        if (files.Count == 0)
        {
          error = "No runnable files found in archive matching configured file types.";
          return false;
        }

        var runnableFiles = files.Where(f => IsFileMatchingFileTypes(f, runnableFileTypes)).ToList();
        var unpackableFiles = files.Where(f => IsFileMatchingFileTypes(f, _unpackableFileTypes)).ToList();

        if (runnableFiles.Count == 1 && unpackableFiles.Count == 0)
        {
          fileToRun = Path.Combine(_finalFolder, runnableFiles[0]);
          _runnableRelativePathFile = runnableFiles[0];
          return true;
        }

        if (runnableFiles.Count == 0 && unpackableFiles.Count == 1)
        {
          archivePathFile = Path.Combine(_finalFolder, unpackableFiles[0]);
          Array.Resize(ref generatedRunnable, _finalFolderIndex + 1);
          generatedRunnable[_finalFolderIndex] = unpackableFiles[0];
          _finalFolderIndex++;
          continue;
        }

        if ((runnableFiles.Count > 1 || unpackableFiles.Count > 1) || (runnableFiles.Count == 1 && unpackableFiles.Count == 1))
        {
          userPick = true;
          return true;
        }

        error = "No runnable files found in archive matching configured file types.";
        return false;
      }
    }

    public bool GetRunnableFileInArchiveFromUser(string gamePathFile, string destFolderWithBackSlash,
      IReadOnlyList<string> runnableFileTypes, out string selectedFile)
    {
      selectedFile = string.Empty;
      if (RunnableFileSelector == null)
        return false;

      string? result = RunnableFileSelector(gamePathFile, runnableFileTypes);
      if (string.IsNullOrEmpty(result))
        return false;

      selectedFile = result;
      return true;
    }

    public void StorePreUnpackedFilesSnapshot(string folder)
    {
      _preSnapshot.Clear();
      if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        return;

      foreach (string pathFile in Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories))
      {
        if (IsGemusMarker(pathFile))
          continue;

        string relative = GetRelativePath(folder, pathFile);
        _preSnapshot.Add(new SnapshotEntry
        {
          RelativePathFile = relative,
          UpperRelativePathFile = relative.ToUpperInvariant(),
          Crc = _crc.CalculateFile(pathFile).ToString()
        });
      }
    }

    public string GetFinalFolder() => _finalFolder;

    public static ArchiveBackend GetDefaultBackend()
    {
      string? env = Environment.GetEnvironmentVariable(EnvArchiveBackend);
      if (TryParseBackend(env, out var backend))
        return backend;

      return OperatingSystem.IsWindows() ? ArchiveBackend.SevenZipCli : ArchiveBackend.Managed7Zip;
    }

    public static bool TryParseBackend(string? value, out ArchiveBackend backend)
    {
      backend = ArchiveBackend.SevenZipCli;
      if (string.IsNullOrWhiteSpace(value))
        return false;

      string v = value.Trim().ToLowerInvariant();
      switch (v)
      {
        case "sevenzipcli":
        case "sevenzip":
        case "7zip":
        case "7z":
        case "cli":
          backend = ArchiveBackend.SevenZipCli;
          return true;
        case "managedzip":
        case "zip":
          backend = ArchiveBackend.ManagedZip;
          return true;
        case "managed7zip":
        case "sharpcompress":
        case "managed":
          backend = ArchiveBackend.Managed7Zip;
          return true;
        default:
          return Enum.TryParse(value, true, out backend);
      }
    }

    public bool Repack(bool askUser, bool notifyUser, string destFolderWithBackSlash, string originalArchivePath,
      Models.RezipAdditionalFiles additionalFilesRule, out bool wasRepacked, out string error)
    {
      wasRepacked = false;
      error = string.Empty;

      if (_finalFolderIndex < 0 || string.IsNullOrEmpty(_finalFolder))
      {
        error = "Final unpacked archive folder not known.";
        return false;
      }

      if (string.IsNullOrEmpty(_runnableRelativePathFile))
      {
        error = "Runnable file not known for repack.";
        return false;
      }

      if (_preSnapshot.Count == 0)
      {
        error = "Snapshot of unpacked folder before playing does not exist.";
        return false;
      }

      StorePostUnpackedFilesSnapshot(_finalFolder);
      if (_postSnapshot.Count == 0)
      {
        error = "No files exist in the unpacked folder after playing.";
        return false;
      }

      if (!DidSnapshotsChangeDuringPlay())
      {
        return true;
      }

      if (askUser && !ConfirmRepack())
      {
        return true;
      }

      string? sevenZip = null;
      if (_backendUsed == ArchiveBackend.SevenZipCli)
      {
        sevenZip = GetExtractorPath("7z.exe");
        if (string.IsNullOrEmpty(sevenZip))
        {
          error = "Unable to locate 7z.exe for repacking archives";
          return false;
        }
      }

      if (!GenerateRepackListFile(_finalFolder, out string listFile, out error))
        return false;

      try
      {
        int currentIdx = _finalFolderIndex;
        string currentFolder = Path.Combine(destFolderWithBackSlash, currentIdx.ToString());
        string destArchive = GetParentArchivePath(destFolderWithBackSlash, currentIdx - 1, originalArchivePath, out error);
        if (string.IsNullOrEmpty(destArchive))
          return false;

        if (_backendUsed == ArchiveBackend.ManagedZip)
        {
          if (!ManagedZipRepackWithListFile(currentFolder, listFile, destArchive, out error))
            return false;
        }
        else if (_backendUsed == ArchiveBackend.Managed7Zip)
        {
          if (!SharpCompressRepackWithListFile(currentFolder, listFile, destArchive, out error))
            return false;
        }
        else
        {
          string sevenZipPath = sevenZip ?? string.Empty;
          if (!SevenZipRepackWithListFile(sevenZipPath, currentFolder, listFile, destArchive, out error))
            return false;
        }

        for (currentIdx = _finalFolderIndex - 1; currentIdx >= 0; currentIdx--)
        {
          currentFolder = Path.Combine(destFolderWithBackSlash, currentIdx.ToString());
          SafeDeleteFile(Path.Combine(currentFolder, GemusMarkerFile));
          destArchive = GetParentArchivePath(destFolderWithBackSlash, currentIdx - 1, originalArchivePath, out error);
          if (string.IsNullOrEmpty(destArchive))
            return false;

          if (_backendUsed == ArchiveBackend.ManagedZip)
          {
            if (!ManagedZipRepackFolder(currentFolder, destArchive, out error))
              return false;
          }
          else if (_backendUsed == ArchiveBackend.Managed7Zip)
          {
            if (!SharpCompressRepackFolder(currentFolder, destArchive, out error))
              return false;
          }
          else
          {
            string sevenZipPath = sevenZip ?? string.Empty;
            if (!SevenZipRepackFolder(sevenZipPath, currentFolder, destArchive, out error))
              return false;
          }
        }

        wasRepacked = true;
        if (notifyUser)
          ShowRepackSuccess();
        return true;
      }
      finally
      {
        SafeDeleteFile(Path.Combine(_finalFolder, "__gemus_repack.lst"));
        SafeDeleteFile(Path.Combine(_finalFolder, "__gemus_repack_check.txt"));
      }
    }

    private void StorePostUnpackedFilesSnapshot(string folder)
    {
      _postSnapshot.Clear();
      if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        return;

      foreach (string pathFile in Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories))
      {
        if (IsGemusMarker(pathFile))
          continue;

        string relative = GetRelativePath(folder, pathFile);
        _postSnapshot.Add(new SnapshotEntry
        {
          RelativePathFile = relative,
          UpperRelativePathFile = relative.ToUpperInvariant(),
          Crc = _crc.CalculateFile(pathFile).ToString()
        });
      }
    }

    private void StoreRecursiveUnpackedFilesSnapshot(string folder)
    {
      _recursiveSnapshot.Clear();
      if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        return;

      foreach (string pathFile in Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories))
      {
        if (IsGemusMarker(pathFile))
          continue;

        _recursiveSnapshot.Add(GetRelativePath(folder, pathFile));
      }
    }

    private bool DidSnapshotsChangeDuringPlay()
    {
      if (_preSnapshot.Count != _postSnapshot.Count)
        return true;

      var postLookup = _postSnapshot.ToDictionary(s => s.UpperRelativePathFile, s => s.Crc, StringComparer.OrdinalIgnoreCase);
      foreach (var pre in _preSnapshot)
      {
        if (!postLookup.TryGetValue(pre.UpperRelativePathFile, out string? crc))
          return true;
        if (!string.Equals(crc, pre.Crc, StringComparison.OrdinalIgnoreCase))
          return true;
      }

      return false;
    }

    private bool IsFileMatchingFileTypes(string fileOrName, IReadOnlyList<string> types)
    {
      if (types == null || types.Count == 0)
        return false;

      string ext = Path.GetExtension(fileOrName).TrimStart('.');
      if (string.IsNullOrEmpty(ext))
        return false;

      return types.Any(t => string.Equals(t, ext, StringComparison.OrdinalIgnoreCase));
    }

    private bool IsFileMatchingFileTypes(string fileOrName, string[] types)
    {
      if (types.Length == 0)
        return false;

      string ext = Path.GetExtension(fileOrName).TrimStart('.');
      if (string.IsNullOrEmpty(ext))
        return false;

      return types.Any(t => string.Equals(t, ext, StringComparison.OrdinalIgnoreCase));
    }

    private string? GetExtractorPath(string filename)
    {
      string appPath = Path.Combine(_appPathWbs, filename);
      if (File.Exists(appPath))
        return appPath;
      if (File.Exists(filename))
        return filename;
      return null;
    }

    private static (string directory, string file) SplitPath(string pathFile)
    {
      string dir = Path.GetDirectoryName(pathFile) ?? string.Empty;
      string file = Path.GetFileName(pathFile);
      return (dir, file);
    }

    private static string GetRelativePath(string root, string pathFile)
    {
      string relative = Path.GetRelativePath(root, pathFile);
      return relative.Replace('/', '\\');
    }

    private static bool IsGemusMarker(string pathFile)
    {
      string name = Path.GetFileName(pathFile);
      return string.Equals(name, GemusMarkerFile, StringComparison.OrdinalIgnoreCase);
    }

    private static void SafeDeleteDirectory(string folder)
    {
      try
      {
        if (Directory.Exists(folder))
          Directory.Delete(folder, true);
      }
      catch
      {
        // Ignore cleanup failures
      }
    }

    private static void SafeDeleteFile(string pathFile)
    {
      try
      {
        if (File.Exists(pathFile))
          File.Delete(pathFile);
      }
      catch
      {
        // Ignore cleanup failures
      }
    }

    private static bool RunProcess(string fileName, string arguments, string workingDirectory, out string output)
    {
      output = string.Empty;
      try
      {
        var startInfo = new ProcessStartInfo
        {
          FileName = fileName,
          Arguments = arguments,
          WorkingDirectory = workingDirectory,
          UseShellExecute = false,
          RedirectStandardOutput = true,
          RedirectStandardError = true,
          CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process == null)
          return false;

        output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (!string.IsNullOrWhiteSpace(error))
          output += error;

        return process.ExitCode == 0;
      }
      catch
      {
        return false;
      }
    }

    private bool UnpackArchive(ArchiveBackend backend, string? sevenZip, string archivePath, string archiveFile,
      string destFolder, out string error)
    {
      error = string.Empty;
      if (backend == ArchiveBackend.ManagedZip)
      {
        string pathFile = Path.Combine(archivePath, archiveFile);
        if (!string.Equals(Path.GetExtension(pathFile), ".zip", StringComparison.OrdinalIgnoreCase))
        {
          error = "Managed ZIP backend only supports .zip archives";
          return false;
        }

        try
        {
          ZipFile.ExtractToDirectory(pathFile, destFolder, true);
          return true;
        }
        catch (Exception ex)
        {
          error = "ZIP extract failed: " + ex.Message;
          return false;
        }
      }

      if (backend == ArchiveBackend.Managed7Zip)
      {
        string pathFile = Path.Combine(archivePath, archiveFile);
        return SharpCompressUnpack(pathFile, destFolder, out error);
      }

      string args = "x \"" + archiveFile + "\" -y -r -o\"" + destFolder + "\"";
      if (!RunProcess(sevenZip ?? string.Empty, args, archivePath, out _))
      {
        error = "7-Zip failed to unpack archive " + archiveFile + " to " + destFolder;
        return false;
      }

      return true;
    }

    private List<string> ListArchive(ArchiveBackend backend, string? sevenZip, string archivePathFile, out string error)
    {
      error = string.Empty;
      var results = new List<string>();
      if (backend == ArchiveBackend.ManagedZip)
      {
        if (!string.Equals(Path.GetExtension(archivePathFile), ".zip", StringComparison.OrdinalIgnoreCase))
        {
          error = "Managed ZIP backend only supports .zip archives";
          return results;
        }

        try
        {
          using var archive = ZipFile.OpenRead(archivePathFile);
          foreach (var entry in archive.Entries)
          {
            if (string.IsNullOrEmpty(entry.Name))
              continue;
            results.Add(entry.FullName.Replace('/', '\\'));
          }
          return results;
        }
        catch (Exception ex)
        {
          error = "ZIP list failed: " + ex.Message;
          return results;
        }
      }

      if (backend == ArchiveBackend.Managed7Zip)
      {
        return SharpCompressList(archivePathFile, out error);
      }

      (string archivePath, string archiveFile) = SplitPath(archivePathFile);
      string args = "l -slt \"" + archiveFile + "\"";

      if (!RunProcess(sevenZip ?? string.Empty, args, archivePath, out string output))
      {
        error = "Unable to list contents of archive";
        return results;
      }

      using var reader = new StringReader(output);
      string? line;
      string? currentPath = null;
      bool isDirectory = false;
      while ((line = reader.ReadLine()) != null)
      {
        if (line.StartsWith("Path = ", StringComparison.OrdinalIgnoreCase))
        {
          currentPath = line.Substring(7).Trim();
          isDirectory = false;
        }
        else if (line.StartsWith("Attributes = ", StringComparison.OrdinalIgnoreCase))
        {
          string attr = line.Substring(13).Trim();
          isDirectory = attr.Contains('D');
        }
        else if (string.IsNullOrWhiteSpace(line))
        {
          if (!string.IsNullOrEmpty(currentPath) && !isDirectory)
          {
            results.Add(currentPath.Replace('/', '\\'));
          }
          currentPath = null;
          isDirectory = false;
        }
      }

      if (!string.IsNullOrEmpty(currentPath) && !isDirectory)
      {
        results.Add(currentPath.Replace('/', '\\'));
      }

      return results;
    }

    private bool GenerateRepackListFile(string repackFolder, out string listFile, out string error)
    {
      error = string.Empty;
      listFile = Path.Combine(repackFolder, "__gemus_repack.lst");
      try
      {
        using var writer = new StreamWriter(listFile, false, Encoding.UTF8);
        foreach (var entry in _postSnapshot)
        {
          writer.WriteLine(entry.RelativePathFile);
        }
        return true;
      }
      catch (Exception ex)
      {
        error = "Unable to create repack list file: " + ex.Message;
        return false;
      }
    }

    private bool SevenZipRepackWithListFile(string sevenZip, string repackFolder, string listFile, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      SafeDeleteFile(destArchivePathFile);

      string args = "a \"" + destArchivePathFile + "\" -y @\"" + listFile + "\"";
      if (!RunProcess(sevenZip, args, repackFolder, out _))
      {
        error = "Unable to create archive " + destArchivePathFile;
        return false;
      }

      if (!File.Exists(destArchivePathFile))
      {
        error = "Unable to create archive " + destArchivePathFile;
        return false;
      }

      return true;
    }

    private bool SevenZipRepackFolder(string sevenZip, string repackFolder, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      SafeDeleteFile(destArchivePathFile);

      string args = "a \"" + destArchivePathFile + "\" * -ry";
      if (!RunProcess(sevenZip, args, repackFolder, out _))
      {
        error = "Unable to create archive " + destArchivePathFile;
        return false;
      }

      if (!File.Exists(destArchivePathFile))
      {
        error = "Unable to create archive " + destArchivePathFile;
        return false;
      }

      return true;
    }

    private bool ManagedZipRepackWithListFile(string repackFolder, string listFile, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      if (!string.Equals(Path.GetExtension(destArchivePathFile), ".zip", StringComparison.OrdinalIgnoreCase))
      {
        error = "Managed ZIP backend only supports .zip archives";
        return false;
      }

      SafeDeleteFile(destArchivePathFile);

      try
      {
        using var archive = ZipFile.Open(destArchivePathFile, ZipArchiveMode.Create);
        foreach (string line in File.ReadLines(listFile))
        {
          string relative = line.Trim();
          if (string.IsNullOrEmpty(relative))
            continue;

          string pathFile = Path.Combine(repackFolder, relative);
          if (!File.Exists(pathFile))
          {
            error = "Repack aborted. Missing file: " + relative;
            return false;
          }

          string entryName = relative.Replace('\\', '/');
          archive.CreateEntryFromFile(pathFile, entryName, CompressionLevel.Optimal);
        }

        return true;
      }
      catch (Exception ex)
      {
        error = "ZIP repack failed: " + ex.Message;
        return false;
      }
    }

    private bool ManagedZipRepackFolder(string repackFolder, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      if (!string.Equals(Path.GetExtension(destArchivePathFile), ".zip", StringComparison.OrdinalIgnoreCase))
      {
        error = "Managed ZIP backend only supports .zip archives";
        return false;
      }

      SafeDeleteFile(destArchivePathFile);

      try
      {
        ZipFile.CreateFromDirectory(repackFolder, destArchivePathFile, CompressionLevel.Optimal, false);
        return true;
      }
      catch (Exception ex)
      {
        error = "ZIP repack failed: " + ex.Message;
        return false;
      }
    }

    private bool SharpCompressUnpack(string archivePathFile, string destFolder, out string error)
    {
      error = string.Empty;
      try
      {
        using var archive = ArchiveFactory.Open(archivePathFile);
        foreach (var entry in archive.Entries.Where(e => !e.IsDirectory))
        {
          entry.WriteToDirectory(destFolder, new ExtractionOptions
          {
            ExtractFullPath = true,
            Overwrite = true
          });
        }
        return true;
      }
      catch (Exception ex)
      {
        error = "Archive extract failed: " + ex.Message;
        return false;
      }
    }

    private List<string> SharpCompressList(string archivePathFile, out string error)
    {
      error = string.Empty;
      var results = new List<string>();
      try
      {
        using var archive = ArchiveFactory.Open(archivePathFile);
        foreach (var entry in archive.Entries)
        {
          if (entry.IsDirectory)
            continue;
          string key = entry.Key ?? string.Empty;
          if (string.IsNullOrEmpty(key))
            continue;
          results.Add(key.Replace('/', '\\'));
        }
      }
      catch (Exception ex)
      {
        error = "Archive list failed: " + ex.Message;
      }
      return results;
    }

    private bool SharpCompressRepackWithListFile(string repackFolder, string listFile, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      if (!TryGetSharpCompressRepackTarget(destArchivePathFile, out string targetArchivePathFile,
        out ArchiveType archiveType, out CompressionType compressionType, out _, out error))
      {
        return false;
      }

      SafeDeleteFile(targetArchivePathFile);

      try
      {
        using var stream = File.Create(targetArchivePathFile);
        var options = new WriterOptions(compressionType);
        using var writer = WriterFactory.Open(stream, archiveType, options);

        foreach (string line in File.ReadLines(listFile))
        {
          string relative = line.Trim();
          if (string.IsNullOrEmpty(relative))
            continue;

          string pathFile = Path.Combine(repackFolder, relative);
          if (!File.Exists(pathFile))
          {
            error = "Repack aborted. Missing file: " + relative;
            return false;
          }

          string entryName = relative.Replace('\\', '/');
          writer.Write(entryName, pathFile);
        }

        return true;
      }
      catch (Exception ex)
      {
        error = "Archive repack failed: " + ex.Message;
        return false;
      }
    }

    private bool SharpCompressRepackFolder(string repackFolder, string destArchivePathFile, out string error)
    {
      error = string.Empty;
      if (!TryGetSharpCompressRepackTarget(destArchivePathFile, out string targetArchivePathFile,
        out ArchiveType archiveType, out CompressionType compressionType, out _, out error))
      {
        return false;
      }

      SafeDeleteFile(targetArchivePathFile);

      try
      {
        using var stream = File.Create(targetArchivePathFile);
        var options = new WriterOptions(compressionType);
        using var writer = WriterFactory.Open(stream, archiveType, options);

        foreach (string pathFile in Directory.GetFiles(repackFolder, "*.*", SearchOption.AllDirectories))
        {
          if (IsGemusMarker(pathFile))
            continue;

          string name = Path.GetFileName(pathFile);
          if (string.Equals(name, "__gemus_repack.lst", StringComparison.OrdinalIgnoreCase))
            continue;
          if (string.Equals(name, "__gemus_repack_check.txt", StringComparison.OrdinalIgnoreCase))
            continue;

          string relative = GetRelativePath(repackFolder, pathFile).Replace('\\', '/');
          writer.Write(relative, pathFile);
        }

        return true;
      }
      catch (Exception ex)
      {
        error = "Archive repack failed: " + ex.Message;
        return false;
      }
    }

    private bool IsSharpCompressRepackSupported(string pathFile)
    {
      return TryGetSharpCompressRepackTarget(pathFile, out _, out _, out _, out _, out _);
    }

    private bool TryGetSharpCompressRepackTarget(string pathFile, out string targetArchivePathFile,
      out ArchiveType archiveType, out CompressionType compressionType, out bool usedFallbackZip, out string error)
    {
      error = string.Empty;
      usedFallbackZip = false;
      targetArchivePathFile = pathFile;
      archiveType = ArchiveType.Zip;
      compressionType = CompressionType.Deflate;

      string lower = pathFile.ToLowerInvariant();
      if (lower.EndsWith(".tar.gz", StringComparison.Ordinal) || lower.EndsWith(".tgz", StringComparison.Ordinal))
      {
        archiveType = ArchiveType.Tar;
        compressionType = CompressionType.GZip;
        return true;
      }
      if (lower.EndsWith(".tar.bz2", StringComparison.Ordinal) || lower.EndsWith(".tbz2", StringComparison.Ordinal) || lower.EndsWith(".tbz", StringComparison.Ordinal))
      {
        archiveType = ArchiveType.Tar;
        compressionType = CompressionType.BZip2;
        return true;
      }

      string ext = Path.GetExtension(pathFile).ToLowerInvariant();
      switch (ext)
      {
        case ".zip":
          archiveType = ArchiveType.Zip;
          compressionType = CompressionType.Deflate;
          return true;
        case ".7z":
          archiveType = ArchiveType.SevenZip;
          compressionType = CompressionType.LZMA;
          return true;
        case ".tar":
          archiveType = ArchiveType.Tar;
          compressionType = CompressionType.None;
          return true;
        case ".rar":
          targetArchivePathFile = Path.ChangeExtension(pathFile, ".zip");
          archiveType = ArchiveType.Zip;
          compressionType = CompressionType.Deflate;
          usedFallbackZip = true;
          return true;
        default:
          error = "Managed 7-Zip backend does not support repacking this archive type";
          return false;
      }
    }

    private void WriteGemusMarker(string folder, string relativePathFile)
    {
      string pathFile = Path.Combine(folder, GemusMarkerFile);
      File.WriteAllText(pathFile, relativePathFile, Encoding.UTF8);
    }

    private string GetParentArchivePath(string baseFolderWbs, int parentIdx, string originalArchivePath, out string error)
    {
      error = string.Empty;
      if (parentIdx < 0)
        return originalArchivePath;

      string parentFolder = Path.Combine(baseFolderWbs, parentIdx.ToString());
      string marker = Path.Combine(parentFolder, GemusMarkerFile);
      if (!File.Exists(marker))
      {
        error = "Unable to repack. Could not find required __gemus file in unpacked folder";
        return string.Empty;
      }

      string? line = File.ReadLines(marker).FirstOrDefault();
      if (string.IsNullOrWhiteSpace(line))
      {
        error = "Unable to repack. Could not read parent archive from __gemus file";
        return string.Empty;
      }

      string parentArchivePath = Path.Combine(parentFolder, line.Trim());
      if (!File.Exists(parentArchivePath))
      {
        error = "Unable to repack. Could not find parent archive from __gemus file";
        return string.Empty;
      }

      return parentArchivePath;
    }

    private bool ConfirmRepack()
    {
      try
      {
        return System.Windows.Forms.MessageBox.Show(
          "There have been changes made to the game file(s). Do you want to repack them?",
          "GEMUS",
          System.Windows.Forms.MessageBoxButtons.YesNo,
          System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes;
      }
      catch
      {
        return false;
      }
    }

    private void ShowRepackSuccess()
    {
      try
      {
        System.Windows.Forms.MessageBox.Show(
          "The game was repacked successfully!",
          "GEMUS",
          System.Windows.Forms.MessageBoxButtons.OK,
          System.Windows.Forms.MessageBoxIcon.Information);
      }
      catch
      {
        // Ignore UI failures
      }
    }

    private sealed class SnapshotEntry
    {
      public string Crc { get; set; } = string.Empty;
      public string RelativePathFile { get; set; } = string.Empty;
      public string UpperRelativePathFile { get; set; } = string.Empty;
    }
  }
}
