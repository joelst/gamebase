// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// CRC checksum calculator
// Converted from VB6 clsCRC (which used inline x86 assembly)
// Replaced with pure managed .NET CRC32 implementation

namespace Gemus
{
  /// <summary>
  /// CRC checksum calculation algorithms.
  /// </summary>
  public enum CrcAlgorithm
  {
    Crc16,
    Crc32
  }

  /// <summary>
  /// CRC checksum calculator supporting CRC16 and CRC32.
  /// Converted from VB6 clsCRC (replaced P/Invoke inline assembly with managed .NET).
  /// </summary>
  public class CrcCalculator
  {
    private CrcAlgorithm _algorithm = CrcAlgorithm.Crc32;
    private uint _crc32;
    private ushort _crc16;

    // CRC32 lookup table (standard polynomial 0xEDB88320)
    private static readonly uint[] Crc32Table = InitCrc32Table();

    // CRC16 lookup table (CCITT polynomial 0x1021)
    private static readonly ushort[] Crc16Table = InitCrc16Table();

    public CrcCalculator()
    {
      Clear();
    }

    public CrcCalculator(CrcAlgorithm algorithm)
    {
      Algorithm = algorithm;
    }

    /// <summary>
    /// Gets or sets the CRC algorithm to use.
    /// </summary>
    public CrcAlgorithm Algorithm
    {
      get => _algorithm;
      set
      {
        _algorithm = value;
        Clear();
      }
    }

    /// <summary>
    /// Gets the current CRC value.
    /// </summary>
    public long Value
    {
      get => _algorithm switch
      {
        CrcAlgorithm.Crc16 => _crc16 & 0xFFFF,
        CrcAlgorithm.Crc32 => (long)(~_crc32),
        _ => 0
      };
    }

    /// <summary>
    /// Resets the CRC state.
    /// </summary>
    public void Clear()
    {
      _crc16 = 0;
      _crc32 = 0xFFFFFFFF;
    }

    /// <summary>
    /// Adds bytes to the running CRC calculation.
    /// </summary>
    public long AddBytes(byte[] data)
    {
      if (data == null || data.Length == 0) return Value;

      switch (_algorithm)
      {
        case CrcAlgorithm.Crc16:
          foreach (byte b in data)
          {
            _crc16 = (ushort)((_crc16 << 8) ^ Crc16Table[((_crc16 >> 8) ^ b) & 0xFF]);
          }
          break;

        case CrcAlgorithm.Crc32:
          foreach (byte b in data)
          {
            _crc32 = (_crc32 >> 8) ^ Crc32Table[(_crc32 ^ b) & 0xFF];
          }
          break;
      }

      return Value;
    }

    /// <summary>
    /// Adds a string to the running CRC calculation.
    /// </summary>
    public long AddString(string text)
    {
      return AddBytes(System.Text.Encoding.ASCII.GetBytes(text));
    }

    /// <summary>
    /// Calculates CRC from a byte array (resets first).
    /// </summary>
    public long CalculateBytes(byte[] data)
    {
      Clear();
      return AddBytes(data);
    }

    /// <summary>
    /// Calculates CRC from a file.
    /// </summary>
    public long CalculateFile(string filename)
    {
      if (!File.Exists(filename) || new FileInfo(filename).Length == 0)
        return 0;

      byte[] data = File.ReadAllBytes(filename);
      return CalculateBytes(data);
    }

    /// <summary>
    /// Calculates CRC from a string (resets first).
    /// </summary>
    public long CalculateString(string text)
    {
      return CalculateBytes(System.Text.Encoding.ASCII.GetBytes(text));
    }

    private static uint[] InitCrc32Table()
    {
      // Standard CRC32 table (polynomial 0xEDB88320, reflected)
      uint[] table = new uint[256];
      for (uint i = 0; i < 256; i++)
      {
        uint crc = i;
        for (int j = 0; j < 8; j++)
        {
          if ((crc & 1) != 0)
            crc = (crc >> 1) ^ 0xEDB88320;
          else
            crc >>= 1;
        }
        table[i] = crc;
      }
      return table;
    }

    private static ushort[] InitCrc16Table()
    {
      // CRC16-CCITT table (polynomial 0x1021)
      ushort[] table = new ushort[256];
      for (int i = 0; i < 256; i++)
      {
        int k = i << 8;
        ushort crc = 0;
        for (int j = 0; j < 8; j++)
        {
          if (((crc ^ k) & 0x8000) != 0)
            crc = (ushort)((crc << 1) ^ 0x1021);
          else
            crc = (ushort)(crc << 1);
          k <<= 1;
        }
        table[i] = crc;
      }
      return table;
    }
  }
}
