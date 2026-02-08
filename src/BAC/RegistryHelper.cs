using GBFuncs;

namespace BAC;

/// <summary>
/// Helper class for registry operations used by BAC.
/// Wraps GBFuncs.RegistryService with simplified methods.
/// </summary>
public class RegistryHelper
{
  private readonly RegistryService _registry = new();

  public int GetIntValue(string keyName, string valueName, int defaultValue)
  {
    if (_registry.GetSetting("HKEY_LOCAL_MACHINE", keyName, valueName, out var value, false, defaultValue))
    {
      return Convert.ToInt32(value);
    }
    return defaultValue;
  }

  public string GetStringValue(string keyName, string valueName, string defaultValue)
  {
    if (_registry.GetSetting("HKEY_LOCAL_MACHINE", keyName, valueName, out var value, true, defaultValue))
    {
      return value?.ToString() ?? defaultValue;
    }
    return defaultValue;
  }

  public bool GetBoolValue(string keyName, string valueName, bool defaultValue)
  {
    if (_registry.GetSetting("HKEY_LOCAL_MACHINE", keyName, valueName, out var value, false, defaultValue ? 1 : 0))
    {
      return Convert.ToInt32(value) != 0;
    }
    return defaultValue;
  }

  public void SetIntValue(string keyName, string valueName, int value)
  {
    _registry.SaveSetting("HKEY_LOCAL_MACHINE", keyName, valueName, value, false);
  }

  public void SetStringValue(string keyName, string valueName, string value)
  {
    _registry.SaveSetting("HKEY_LOCAL_MACHINE", keyName, valueName, value, true);
  }

  public void SetBoolValue(string keyName, string valueName, bool value)
  {
    _registry.SaveSetting("HKEY_LOCAL_MACHINE", keyName, valueName, value ? 1 : 0, false);
  }
}
