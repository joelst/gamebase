// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: D64 disk image reader (Commodore 1541 disk format)
// Converted from VB6 clsD64, clsD64Entries, clsD64Entry

namespace Gemus.DiskFormats
{
  /// <summary>
  /// Reads and parses D64 disk images (Commodore 1541 format).
  /// </summary>
  public class D64Image
  {
    // File type constants
    private const byte Scratched = 0;
    private const byte SeqClosed = 1;
    private const byte PrgClosed = 2;
    private const byte UsrClosed = 3;
    private const byte RelClosed = 4;
    private const byte Del = 128;
    private const byte Seq = 129;
    private const byte Prg = 130;
    private const byte Usr = 131;
    private const byte Rel = 132;
    private const byte DelSaved = 160;
    private const byte SeqSaved = 161;
    private const byte PrgSaved = 162;
    private const byte UsrSaved = 163;
    private const byte RelSaved = 164;
    private const byte DelLocked = 192;
    private const byte SeqLocked = 193;
    private const byte PrgLocked = 194;
    private const byte UsrLocked = 195;
    private const byte RelLocked = 196;

    public string DiskName { get; private set; } = string.Empty;
    public string DiskID { get; private set; } = string.Empty;
    public List<DiskImageEntry> Entries { get; } = new();

    /// <summary>
    /// Loads and parses a D64 disk image file.
    /// Returns empty string on success, error message on failure.
    /// </summary>
    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        // Read BAM (Block Availability Map) at track 18, sector 0
        // Offset 91392 (0x16500) in a standard D64
        fs.Seek(91392, SeekOrigin.Begin);
        byte trackOfFirstDirBlock = br.ReadByte();
        byte sectorOfFirstDirBlock = br.ReadByte();
        byte diskDosVersionType = br.ReadByte();
        byte notUsed = br.ReadByte();
        byte[] bamTrackEntries = br.ReadBytes(140);
        byte[] diskNameChars = br.ReadBytes(16);
        byte[] paddedWithA0 = br.ReadBytes(2);
        byte[] diskIdBytes = br.ReadBytes(2);
        byte usuallyA0 = br.ReadByte();
        byte[] dosType = br.ReadBytes(2);

        DiskName = GetName(diskNameChars);

        // Build disk ID string
        string rawId = $"{(char)diskIdBytes[0]}{(char)diskIdBytes[1]}{(char)usuallyA0}{(char)dosType[0]}{(char)dosType[1]}";
        var cleanId = new char[5];
        for (int i = 0; i < 5; i++)
        {
          char c = rawId[i];
          cleanId[i] = (c == '\0' || c == '\xA0') ? ' ' : c;
        }
        DiskID = new string(cleanId);

        // Read directory entries starting at offset 91648 (0x16600)
        long startAddress = 91648;
        for (int block = 1; block <= 18; block++)
        {
          long nextTrack = 0;
          long nextSector = 0;

          for (int entry = 0; entry <= 7; entry++)
          {
            fs.Seek(startAddress + (entry * 32), SeekOrigin.Begin);

            byte dirNextTrack = br.ReadByte();
            byte dirNextSector = br.ReadByte();
            byte fileType = br.ReadByte();
            byte trackOfFile = br.ReadByte();
            byte sectorOfFile = br.ReadByte();
            byte[] fileNameChars = br.ReadBytes(16);
            byte relFileTrack = br.ReadByte();
            byte relFileSector = br.ReadByte();
            byte relFileRecordLen = br.ReadByte();
            byte[] geosUnused = br.ReadBytes(6);
            byte fileSizeLo = br.ReadByte();
            byte fileSizeHi = br.ReadByte();

            // Capture next block pointer from first entry
            if (entry == 0)
            {
              nextTrack = dirNextTrack;
              nextSector = dirNextSector;
            }

            // Add entry if not scratched
            AddDirEntry(fileType, fileNameChars, fileSizeLo, fileSizeHi);

            // At end of block, follow chain
            if (entry == 7)
            {
              startAddress = GetNextBlock((byte)nextTrack, (byte)nextSector);
            }
          }

          if (nextTrack == 0) break; // No more blocks
        }

        return string.Empty;
      }
      catch
      {
        return GetErrorMessage();
      }
    }

    private void AddDirEntry(byte fileType, byte[] fileNameChars, byte sizeLo, byte sizeHi)
    {
      if (fileType == Scratched) return;

      string entryType = fileType switch
      {
        Del => "DEL",
        Seq => "SEQ",
        Prg => "PRG",
        Usr => "USR",
        Rel => "REL",
        DelLocked => "DEL<",
        SeqLocked => "SEQ<",
        PrgLocked => "PRG<",
        UsrLocked => "USR<",
        RelLocked => "REL<",
        SeqClosed => "*SEQ",
        PrgClosed => "*PRG",
        UsrClosed => "*USR",
        RelClosed => "*REL",
        DelSaved => "@0:DEL",
        SeqSaved => "@0:SEQ",
        PrgSaved => "@0:PRG",
        UsrSaved => "@0:USR",
        RelSaved => "@0:REL",
        _ => "?"
      };

      string entryName = GetName(fileNameChars);
      string entrySize = (sizeHi * 256 + sizeLo).ToString().PadLeft(5);

      Entries.Add(new DiskImageEntry
      {
        Filename = entryName,
        FileSizeInBlocks = entrySize,
        FileType = entryType
      });
    }

    /// <summary>
    /// Calculates the byte offset for a given track/sector in a D64 image.
    /// </summary>
    private static long GetNextBlock(byte track, byte block)
    {
      long offset = track switch
      {
        0 => 0,
        < 19 => ((track - 1) * 21L) * 256L + block * 256L,
        < 26 => 357L * 256L + ((track - 19) * 21L) * 256L + block * 256L,
        < 32 => 490L * 256L + ((track - 26) * 21L) * 256L + block * 256L,
        < 41 => 598L * 256L + ((track - 32) * 21L) * 256L + block * 256L,
        _ => 0
      };
      return offset; // Note: VB6 used 1-based file positioning; .NET uses 0-based Seek
    }

    /// <summary>
    /// Converts a 16-byte C64 filename to a .NET string, trimming 0xA0 padding.
    /// </summary>
    private static string GetName(byte[] chars)
    {
      // Build raw string from bytes
      var raw = new char[16];
      for (int i = 0; i < 16 && i < chars.Length; i++)
        raw[i] = (char)chars[i];

      // Trim trailing 0xA0 values
      int length = 16;
      while (length > 0 && raw[length - 1] == '\xA0')
        length--;

      // Convert remaining 0xA0 to spaces
      var result = new char[length];
      for (int i = 0; i < length; i++)
        result[i] = raw[i] == '\xA0' ? ' ' : raw[i];

      return new string(result);
    }

    private static string GetErrorMessage()
    {
      return "There was an error reading from the file or the file does not conform to the standard.";
    }
  }
}
