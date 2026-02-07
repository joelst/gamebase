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
  }
}
