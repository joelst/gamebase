// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1997-2004 James Burrows - Migrated to .NET 2026
// HVSC2GB64: Data models

namespace HVSC2GB64;

/// <summary>
/// Musician information from HVSC collection.
/// </summary>
internal class MusicianData
{
  public long MusicianId { get; set; }
  public string Name { get; set; } = string.Empty;
  public string Group { get; set; } = string.Empty;
  public string Nickname { get; set; } = string.Empty;
  public string PhotoPath { get; set; } = string.Empty;
}

/// <summary>
/// SID file metadata from HVSC.
/// </summary>
internal class SidData
{
  public string RelativePath { get; set; } = string.Empty;
  public string Filename { get; set; } = string.Empty;
  public string Name { get; set; } = string.Empty;
  public string Author { get; set; } = string.Empty;
  public string Copyright { get; set; } = string.Empty;
  public int Songs { get; set; }
  public int DefaultSong { get; set; }
  public bool IsGame { get; set; }
}
