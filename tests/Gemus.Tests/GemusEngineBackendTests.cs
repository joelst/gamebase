using Gemus;
using GBFuncs;
using Xunit;

namespace Gemus.Tests
{
  public class GemusEngineBackendTests
  {
    [Fact]
    public void TrySetArchiveBackend_SetsUnpackerBackend()
    {
      var engine = new GemusEngine(new FunctionsHelper(), new RegistryService(), new GemusLogger(), new ArchiveUnpacker());
      bool ok = engine.TrySetArchiveBackend("Managed7Zip", out var parsed);
      Assert.True(ok);
      Assert.Equal(ArchiveBackend.Managed7Zip, parsed);
    }

    [Fact]
    public void SetArchiveBackend_PersistsToGemusIni()
    {
      string root = CreateTempRoot();
      string gbName = "GB";
      string gbPath = Path.Combine(root, gbName);
      Directory.CreateDirectory(gbPath);
      File.WriteAllText(Path.Combine(gbPath, "Emulators.ini"), string.Empty);
      File.WriteAllText(Path.Combine(gbPath, "Paths.ini"), string.Empty);
      File.WriteAllText(Path.Combine(root, "lang.ini"), string.Empty);

      var engine = new GemusEngine(new FunctionsHelper(), new RegistryService(), new GemusLogger(), new ArchiveUnpacker());
      engine.SetAppContext(root + Path.DirectorySeparatorChar, Path.Combine(root, "lang.ini"), "GameBase");
      engine.Init(gbName);

      engine.SetArchiveBackend(ArchiveBackend.Managed7Zip, persistToIni: true);

      string gemusIni = Path.Combine(gbPath, "Gemus.ini");
      string val = GemusIniHelper.GetIniMessage("Settings", "ArchiveBackend", string.Empty, gemusIni);
      Assert.Equal("Managed7Zip", val);
    }

    private static string CreateTempRoot()
    {
      string path = Path.Combine(Path.GetTempPath(), "gemus-tests", Guid.NewGuid().ToString("N"));
      Directory.CreateDirectory(path);
      return path;
    }
  }
}
