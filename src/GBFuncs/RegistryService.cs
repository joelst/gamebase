//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBFuncs:  General Functions used by GB
//==============================================

using Microsoft.Win32;

namespace GBFuncs
{
  /// <summary>
  /// Provides Windows Registry access functionality.
  /// Converted from VB6 clsRegistry class.
  /// </summary>
  public class RegistryService
  {
    /// <summary>
    /// Gets a registry value, trying HKEY_CURRENT_USER first, then falling back to HKEY_LOCAL_MACHINE.
    /// </summary>
    public bool GetSetting_HKCU_With_HKLM_Fallback(
        string keyName,
        string valueName,
        out object? value,
        object? defaultValue = null)
    {
      // Try HKEY_CURRENT_USER first
      if (GetSettingFromRoot(Registry.CurrentUser, keyName, valueName, out value))
        return true;

      // Fallback to HKEY_LOCAL_MACHINE
      if (GetSettingFromRoot(Registry.LocalMachine, keyName, valueName, out value))
        return true;

      // Use default if both failed
      value = defaultValue;
      return false;
    }

    /// <summary>
    /// Gets a registry value, from HKEY_CURRENT_USER, falling back to HKEY_LOCAL_MACHINE if not found.
    /// </summary>
    public bool GetSetting_HKCU_Fallback_HKLM(
        string keyName,
        string valueName,
        out object? value,
        object? defaultValue = null)
    {
      return GetSetting_HKCU_With_HKLM_Fallback(keyName, valueName, out value, defaultValue);
    }

    /// <summary>
    /// Saves a setting to the Windows Registry.
    /// </summary>
    public bool SaveSetting(string rootKey, string keyName, string valueName, object value, bool isString)
    {
      try
      {
        RegistryKey? root = GetRootKey(rootKey);
        if (root == null)
          return false;

        using (RegistryKey? key = root.CreateSubKey(keyName))
        {
          if (key == null)
            return false;

          if (isString)
          {
            key.SetValue(valueName, value?.ToString() ?? string.Empty, RegistryValueKind.String);
          }
          else
          {
            // Assume DWORD for non-string values
            int intValue = Convert.ToInt32(value);
            key.SetValue(valueName, intValue, RegistryValueKind.DWord);
          }

          return true;
        }
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Gets a setting from the Windows Registry.
    /// </summary>
    public bool GetSetting(
        string rootKey,
        string keyName,
        string valueName,
        out object? value,
        bool isString,
        object? defaultValue = null)
    {
      value = defaultValue;

      try
      {
        RegistryKey? root = GetRootKey(rootKey);
        if (root == null)
          return false;

        return GetSettingFromRoot(root, keyName, valueName, out value);
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Helper method to get a setting from a specific registry root.
    /// </summary>
    private bool GetSettingFromRoot(
        RegistryKey root,
        string keyName,
        string valueName,
        out object? value)
    {
      value = null;

      try
      {
        using (RegistryKey? key = root.OpenSubKey(keyName))
        {
          if (key == null)
            return false;

          value = key.GetValue(valueName);
          return value != null;
        }
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Converts a root key string to a RegistryKey object.
    /// </summary>
    private RegistryKey? GetRootKey(string rootKeyName)
    {
      return rootKeyName.ToUpperInvariant() switch
      {
        "HKEY_CLASSES_ROOT" => Registry.ClassesRoot,
        "HKEY_CURRENT_USER" => Registry.CurrentUser,
        "HKEY_LOCAL_MACHINE" => Registry.LocalMachine,
        "HKEY_USERS" => Registry.Users,
        _ => null
      };
    }

    /// <summary>
    /// Deletes a registry key.
    /// </summary>
    public bool DeleteKey(string rootKey, string keyName, string valueName)
    {
      try
      {
        RegistryKey? root = GetRootKey(rootKey);
        if (root == null)
          return false;

        using (RegistryKey? key = root.OpenSubKey(keyName, writable: true))
        {
          if (key == null)
            return false;

          key.DeleteValue(valueName, throwOnMissingValue: false);
          return true;
        }
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Deletes an entire registry key and all its subkeys.
    /// </summary>
    public bool DeleteKeyTree(string rootKey, string keyName)
    {
      try
      {
        RegistryKey? root = GetRootKey(rootKey);
        if (root == null)
          return false;

        root.DeleteSubKeyTree(keyName, throwOnMissingSubKey: false);
        return true;
      }
      catch
      {
        return false;
      }
    }
  }
}
