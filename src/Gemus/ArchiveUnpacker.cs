// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Archive unpacking stub
// Converted from VB6 clsUnpack usage; extraction logic to be completed.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Gemus
{
  public class ArchiveUnpacker
  {
    private string _finalFolder = string.Empty;

    public Func<string, IReadOnlyList<string>, string?>? RunnableFileSelector { get; set; }

    public bool Unpack(string gamePathFile, string defaultFile, IReadOnlyList<string> runnableFileTypes,
      string destFolderWithBackSlash, out bool isRepackPossible, out string fileToRun,
      out bool userPick, out string error)
    {
      isRepackPossible = false;
      userPick = false;
      error = string.Empty;
      fileToRun = string.Empty;
      _finalFolder = destFolderWithBackSlash;

      if (!File.Exists(gamePathFile))
      {
        error = "Game file not found.";
        return false;
      }

      string ext = Path.GetExtension(gamePathFile).TrimStart('.');
      if (runnableFileTypes.Any(t => string.Equals(t, ext, StringComparison.OrdinalIgnoreCase)))
      {
        fileToRun = gamePathFile;
        return true;
      }

      if (!string.IsNullOrEmpty(defaultFile))
      {
        fileToRun = defaultFile;
        return true;
      }

      userPick = true;
      return true;
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
      _finalFolder = folder;
    }

    public string GetFinalFolder() => _finalFolder;

    public bool Repack(bool askUser, bool notifyUser, string destFolderWithBackSlash, string originalArchivePath,
      Models.RezipAdditionalFiles additionalFilesRule, out bool wasRepacked, out string error)
    {
      wasRepacked = false;
      error = string.Empty;
      return true;
    }
  }
}
