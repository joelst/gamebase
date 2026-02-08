namespace BAC;

/// <summary>
/// Provides path and registry key utilities for BAC.
/// </summary>
internal static class BacPaths
{
  /// <summary>
  /// Gets the registry base path for BAC settings.
  /// </summary>
  public static string RegistryBase => @"Software\GB64\BAC";

  /// <summary>
  /// Gets the registry path for window settings.
  /// </summary>
  public static string RegistryWindow => $@"{RegistryBase}\Window";

  /// <summary>
  /// Gets the registry path for last directory settings.
  /// </summary>
  public static string RegistryLastDirs => $@"{RegistryBase}\LastDirs";
}
