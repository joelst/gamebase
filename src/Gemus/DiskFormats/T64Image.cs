// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: T64 tape image reader (Commodore 64 tape archive format)
// Converted from VB6 clsT64, clsT64Entries, clsT64Entry

namespace Gemus.DiskFormats
{
  /// <summary>
  /// Represents an entry in a T64 tape image.
  /// </summary>
  public class T64Entry
  {
    public string C64Filename { get; set; } = string.Empty;
  }

  /// <summary>
  /// Reads and parses T64 tape image files.
  /// </summary>
  public class T64Image
  {
    public string Signature { get; private set; } = string.Empty;
    public string Version { get; private set; } = string.Empty;
    public long EntriesCount { get; private set; }
    public long EntriesUsed { get; private set; }
    public string TapeImageName { get; private set; } = string.Empty;
    public List<T64Entry> Entries { get; } = new();

    /// <summary>
    /// Loads and parses a T64 tape image file.
    /// Returns empty string on success, error message on failure.
    /// </summary>
    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        // Read header
        byte[] signatureBytes = br.ReadBytes(32);
        Signature = System.Text.Encoding.ASCII.GetString(signatureBytes).Trim();

        byte[] versionBytes = br.ReadBytes(2);
        Version = $"${Get2ByteAddress(versionBytes[1], versionBytes[0]):X4}";

        byte[] entriesCountBytes = br.ReadBytes(2);
        EntriesCount = Get2ByteAddress(entriesCountBytes[1], entriesCountBytes[0]);

        byte[] entriesUsedBytes = br.ReadBytes(2);
        EntriesUsed = Get2ByteAddress(entriesUsedBytes[1], entriesUsedBytes[0]);

        byte[] notUsed = br.ReadBytes(2);

        byte[] tapeNameBytes = br.ReadBytes(24);
        TapeImageName = System.Text.Encoding.ASCII.GetString(tapeNameBytes).Trim();

        // Read entries
        for (int i = 0; i < EntriesUsed; i++)
        {
          byte[] entryData = br.ReadBytes(16);
          byte[] c64FilenameBytes = br.ReadBytes(16);
          string c64Filename = System.Text.Encoding.ASCII.GetString(c64FilenameBytes).Trim();
          Entries.Add(new T64Entry { C64Filename = c64Filename });
        }

        return string.Empty;
      }
      catch
      {
        return "There was an error reading from the file or the file does not conform to the standard.";
      }
    }

    private static long Get2ByteAddress(byte hiByte, byte loByte)
    {
      return (long)loByte + (long)hiByte * 256L;
    }
  }
}
