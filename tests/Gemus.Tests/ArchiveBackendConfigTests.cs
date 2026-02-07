using System;
using Gemus;
using Xunit;

namespace Gemus.Tests
{
  public class ArchiveBackendConfigTests
  {
    [Theory]
    [InlineData("SevenZipCli", ArchiveBackend.SevenZipCli)]
    [InlineData("7zip", ArchiveBackend.SevenZipCli)]
    [InlineData("ManagedZip", ArchiveBackend.ManagedZip)]
    [InlineData("zip", ArchiveBackend.ManagedZip)]
    [InlineData("Managed7Zip", ArchiveBackend.Managed7Zip)]
    [InlineData("sharpcompress", ArchiveBackend.Managed7Zip)]
    public void TryParseBackend_ParsesAliases(string input, ArchiveBackend expected)
    {
      Assert.True(ArchiveUnpacker.TryParseBackend(input, out var backend));
      Assert.Equal(expected, backend);
    }

    [Fact]
    public void TryParseBackend_Invalid_ReturnsFalse()
    {
      Assert.False(ArchiveUnpacker.TryParseBackend("bogus", out _));
      Assert.False(ArchiveUnpacker.TryParseBackend(null, out _));
    }

    [Fact]
    public void DefaultBackend_Windows_SevenZipCli_WhenNoEnv()
    {
      var prev = Environment.GetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND");
      try
      {
        Environment.SetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND", null);
        if (OperatingSystem.IsWindows())
          Assert.Equal(ArchiveBackend.SevenZipCli, ArchiveUnpacker.GetDefaultBackend());
      }
      finally
      {
        Environment.SetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND", prev);
      }
    }

    [Fact]
    public void EnvOverrideBackend_Managed7Zip()
    {
      var prev = Environment.GetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND");
      try
      {
        Environment.SetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND", "Managed7Zip");
        var unpacker = new ArchiveUnpacker();
        Assert.Equal(ArchiveBackend.Managed7Zip, unpacker.Backend);
      }
      finally
      {
        Environment.SetEnvironmentVariable("GEMUS_ARCHIVE_BACKEND", prev);
      }
    }
  }
}
