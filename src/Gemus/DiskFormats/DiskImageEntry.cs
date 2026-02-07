// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Disk Image Entry base class

namespace Gemus.DiskFormats
{
  /// <summary>
  /// Represents a single entry (file) in a disk/tape image.
  /// </summary>
  public class DiskImageEntry
  {
    public string Filename { get; set; } = string.Empty;
    public string FileSizeInBlocks { get; set; } = string.Empty;
    public string FileType { get; set; } = string.Empty;
  }
}
