// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Simple binary file format readers (P00, PRG, G64, TAP, CRT)
// Converted from VB6 clsP00, clsPRG, clsG64, clsTAP, clsCRT

using System.Text;

namespace Gemus.DiskFormats
{
  /// <summary>
  /// Trims null bytes from binary-read strings (shared helper).
  /// </summary>
  internal static class BinaryStringHelper
  {
    /// <summary>
    /// Converts a byte array to a display-ready string, trimming trailing nulls
    /// and replacing embedded nulls with spaces.
    /// </summary>
    public static string TrimForDisplay(byte[] bytes)
    {
      // Build raw string
      var raw = Encoding.ASCII.GetString(bytes);

      // Strip trailing nulls
      int length = raw.Length;
      while (length > 0 && raw[length - 1] == '\0')
        length--;

      // Replace embedded nulls with spaces
      var sb = new StringBuilder(length);
      for (int i = 0; i < length; i++)
        sb.Append(raw[i] == '\0' ? ' ' : raw[i]);

      return sb.ToString();
    }

    private const string DefaultError =
        "There was an error reading from the file or the file does not conform to the standard.";

    public static string ErrorMessage => DefaultError;
  }

  /// <summary>
  /// Reads P00 files (PC64 archive format for C64 files).
  /// </summary>
  public class P00Image
  {
    public string Signature { get; private set; } = string.Empty;
    public string Filename { get; private set; } = string.Empty;
    public string FileSize { get; private set; } = string.Empty;

    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        FileSize = ((fs.Length - 26) / 1024.0).ToString("#####0.00");

        byte[] signatureBytes = br.ReadBytes(7);
        byte notUsed = br.ReadByte();
        byte[] filenameBytes = br.ReadBytes(16);

        Signature = BinaryStringHelper.TrimForDisplay(signatureBytes);
        Filename = BinaryStringHelper.TrimForDisplay(filenameBytes);

        return string.Empty;
      }
      catch
      {
        return BinaryStringHelper.ErrorMessage;
      }
    }
  }

  /// <summary>
  /// Reads PRG files (Commodore 64 program files).
  /// </summary>
  public class PRGImage
  {
    public string FileSize { get; private set; } = string.Empty;
    public string LoadAddress { get; private set; } = string.Empty;
    public string SYSAddress { get; private set; } = string.Empty;

    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        FileSize = (fs.Length / 1024.0).ToString("#####0.00");

        byte loByte = br.ReadByte();
        byte hiByte = br.ReadByte();

        LoadAddress = $"${hiByte:X2}{loByte:X2}";
        int sysAddr = (hiByte * 256) + loByte;
        SYSAddress = $"(SYS {sysAddr})";

        return string.Empty;
      }
      catch
      {
        return BinaryStringHelper.ErrorMessage;
      }
    }
  }

  /// <summary>
  /// Reads G64 files (GCR-encoded 1541 disk image format).
  /// </summary>
  public class G64Image
  {
    public string Signature { get; private set; } = string.Empty;
    public string Version { get; private set; } = string.Empty;
    public string NumTracks { get; private set; } = string.Empty;
    public string TrackSize { get; private set; } = string.Empty;

    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        byte[] sigBytes = br.ReadBytes(8);
        Signature = BinaryStringHelper.TrimForDisplay(sigBytes);

        byte version = br.ReadByte();
        Version = $"${version:X2}";

        byte numTracks = br.ReadByte();
        NumTracks = numTracks.ToString();

        byte trackSizeLo = br.ReadByte();
        byte trackSizeHi = br.ReadByte();
        TrackSize = ((long)(trackSizeHi * 256) + trackSizeLo).ToString();

        return string.Empty;
      }
      catch
      {
        return BinaryStringHelper.ErrorMessage;
      }
    }
  }

  /// <summary>
  /// Reads TAP files (Commodore datasette tape image format).
  /// </summary>
  public class TAPImage
  {
    public string Signature { get; private set; } = string.Empty;
    public string Version { get; private set; } = string.Empty;
    public string FileSize { get; private set; } = string.Empty;

    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        FileSize = (fs.Length / 1024.0).ToString("0.00");

        byte[] sigBytes = br.ReadBytes(12);
        Signature = BinaryStringHelper.TrimForDisplay(sigBytes);

        byte version = br.ReadByte();
        Version = $"${version:X2}";

        return string.Empty;
      }
      catch
      {
        return BinaryStringHelper.ErrorMessage;
      }
    }
  }

  /// <summary>
  /// Reads CRT files (C64 cartridge image format).
  /// </summary>
  public class CRTImage
  {
    public string Signature { get; private set; } = string.Empty;
    public string Version { get; private set; } = string.Empty;
    public string FileSize { get; private set; } = string.Empty;
    public long HardwareType { get; private set; }
    public bool ExromLine { get; private set; }
    public bool GameLine { get; private set; }
    public string CartName { get; private set; } = string.Empty;

    public string Load(string filename)
    {
      try
      {
        using var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs);

        FileSize = (fs.Length / 1024.0).ToString("0.00");

        byte[] sigBytes = br.ReadBytes(16);
        Signature = BinaryStringHelper.TrimForDisplay(sigBytes).Trim();

        byte[] headerLen = br.ReadBytes(4);

        byte[] versionBytes = br.ReadBytes(2);
        Version = $"${versionBytes[0]:X2}.{versionBytes[1]:X2}";

        byte[] hwType = br.ReadBytes(2);
        HardwareType = (long)(hwType[0] * 256) + hwType[1];

        ExromLine = br.ReadByte() != 0;
        GameLine = br.ReadByte() != 0;

        byte[] reserved = br.ReadBytes(6);

        byte[] cartNameBytes = br.ReadBytes(32);
        CartName = BinaryStringHelper.TrimForDisplay(cartNameBytes);

        return string.Empty;
      }
      catch
      {
        return BinaryStringHelper.ErrorMessage;
      }
    }
  }
}
