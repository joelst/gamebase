// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: INI/KV file operations used by the script engine
// Converted from clsGemus KV methods and INI reading

using System.Runtime.InteropServices;
using System.Text;
using Gemus.Models;

namespace Gemus
{
  /// <summary>
  /// Provides INI file and Key=Value set operations for GEMUS.
  /// Converted from VB6 clsGemus KV methods and modFuncs INI helper.
  /// </summary>
  public class GemusIniHelper
  {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern int GetPrivateProfileString(
        string lpAppName, string lpKeyName, string lpDefault,
        StringBuilder lpReturnedString, int nSize, string lpFileName);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern int WritePrivateProfileString(
        string lpAppName, string? lpKeyName, string? lpString, string lpFileName);

    /// <summary>
    /// Reads a message string from an INI file.
    /// </summary>
    public static string GetIniMessage(string section, string key, string defaultValue, string iniFile)
    {
      if (string.IsNullOrEmpty(iniFile))
        return defaultValue;

      var sb = new StringBuilder(2096);
      int len = GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, iniFile);
      return sb.ToString(0, len);
    }

    /// <summary>
    /// Writes a value to an INI file.
    /// </summary>
    public static bool WriteIniValue(string section, string key, string value, string iniFile)
    {
      return WritePrivateProfileString(section, key, value, iniFile) != 0;
    }

    /// <summary>
    /// Reads all key=value pairs from a named set (section) in a Gemus.ini file.
    /// </summary>
    public static List<NotesKeyValue> GetKvSet(string setName, string gemusIniFile)
    {
      var result = new List<NotesKeyValue>();

      if (!File.Exists(gemusIniFile))
        return result;

      try
      {
        bool inDesiredSet = false;
        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string line = rawLine.Trim();

          if (line.StartsWith('[') && line.EndsWith(']') && line.Length > 2)
          {
            string sectionName = line[1..^1];
            if (string.Equals(sectionName, setName, StringComparison.OrdinalIgnoreCase))
            {
              inDesiredSet = true;
            }
            else if (inDesiredSet)
            {
              break; // New section after our set = done
            }
          }
          else if (inDesiredSet)
          {
            int eqPos = line.IndexOf('=');
            if (eqPos > 0 && !line.StartsWith(';'))
            {
              string key = line[..eqPos].Trim();
              if (!string.IsNullOrEmpty(key))
              {
                string value = line[(eqPos + 1)..].Trim();
                result.Add(new NotesKeyValue { Key = key, Value = value });
              }
            }
          }
        }
      }
      catch
      {
        return new List<NotesKeyValue>();
      }

      return result;
    }

    /// <summary>
    /// Gets the text content of a named KV set.
    /// </summary>
    public static (bool success, string text, string error) GetKvSetAsText(string setName, string gemusIniFile)
    {
      try
      {
        if (!File.Exists(gemusIniFile))
          return (true, string.Empty, string.Empty);

        var sb = new StringBuilder();
        bool inSet = false;
        bool doneSet = false;

        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string trimmedLine = rawLine.Trim();

          if (string.Equals(trimmedLine, $"[{setName}]", StringComparison.OrdinalIgnoreCase))
          {
            inSet = true;
            doneSet = true;
          }
          else
          {
            if (trimmedLine.StartsWith('[') && trimmedLine.EndsWith(']'))
            {
              if (doneSet) break;
            }

            if (inSet)
            {
              sb.AppendLine(rawLine);
            }
          }
        }

        string text = sb.ToString().TrimEnd();
        return (true, text, string.Empty);
      }
      catch (Exception ex)
      {
        return (false, string.Empty, $"{ex.Message} ({ex.HResult})");
      }
    }

    /// <summary>
    /// Gets all section names from a Gemus.ini file (excluding "global").
    /// </summary>
    public static List<string> GetKvSetNames(string gemusIniFile)
    {
      var result = new List<string>();

      if (!File.Exists(gemusIniFile))
        return result;

      try
      {
        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string line = rawLine.Trim();
          if (line.StartsWith('[') && line.EndsWith(']') && line.Length > 2)
          {
            string sectionName = line[1..^1];
            if (!string.Equals(sectionName, "global", StringComparison.OrdinalIgnoreCase))
            {
              result.Add(sectionName);
            }
          }
        }
      }
      catch
      {
        return new List<string>();
      }

      return result;
    }

    /// <summary>
    /// Parses a multi-line string containing key=value pairs into an array.
    /// </summary>
    public static List<NotesKeyValue> GetKvsFromString(string kvText)
    {
      var result = new List<NotesKeyValue>();

      if (string.IsNullOrWhiteSpace(kvText))
        return result;

      foreach (string rawLine in kvText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
      {
        string line = rawLine.Trim();
        if (line.Length == 0 || line == "=") continue;

        int eqPos = line.IndexOf('=');
        string key;
        string value;

        if (eqPos > 0)
        {
          key = line[..eqPos].Trim();
          value = line[(eqPos + 1)..].Trim();
        }
        else
        {
          key = line;
          value = string.Empty;
        }

        if (!string.IsNullOrEmpty(key))
        {
          result.Add(new NotesKeyValue { Key = key, Value = value });
        }
      }

      return result;
    }

    /// <summary>
    /// Parses a multi-line string into structured lines, some of which may be KV pairs.
    /// </summary>
    public static List<NotesKeyValueLine> GetKvLinesFromString(string kvText)
    {
      var result = new List<NotesKeyValueLine>();

      if (string.IsNullOrWhiteSpace(kvText))
        return result;

      foreach (string rawLine in kvText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
      {
        var lineObj = new NotesKeyValueLine
        {
          Line = rawLine,
          IsKeyValue = false,
          Key = string.Empty,
          Value = string.Empty,
          Flagged = false
        };

        string trimmed = rawLine.Trim();
        int eqPos = trimmed.IndexOf('=');
        if (eqPos > 0)
        {
          lineObj.Key = trimmed[..eqPos].Trim();
          lineObj.Value = trimmed[(eqPos + 1)..].Trim();
          lineObj.IsKeyValue = true;
        }

        result.Add(lineObj);
      }

      return result;
    }

    /// <summary>
    /// Saves a KV set to the Gemus.ini file.
    /// </summary>
    public static (bool success, string error) SaveKvSet(string setName, string kvText, string gemusIniFile)
    {
      try
      {
        string textToWrite = $"[{setName.ToLower()}]\r\n{kvText}";

        if (!File.Exists(gemusIniFile))
        {
          File.WriteAllText(gemusIniFile, textToWrite + "\r\n");
          return (true, string.Empty);
        }

        var sb = new StringBuilder();
        bool inSet = false;
        bool doneSet = false;

        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string trimmedLine = rawLine.Trim();

          if (string.Equals(trimmedLine, $"[{setName}]", StringComparison.OrdinalIgnoreCase))
          {
            inSet = true;
            doneSet = true;
            sb.AppendLine(textToWrite);
          }
          else
          {
            if (trimmedLine.StartsWith('[') && trimmedLine.EndsWith(']'))
            {
              inSet = false;
            }

            if (!inSet)
            {
              sb.AppendLine(rawLine);
            }
          }
        }

        if (!doneSet)
        {
          sb.AppendLine(textToWrite);
        }

        File.WriteAllText(gemusIniFile, sb.ToString());
        return (true, string.Empty);
      }
      catch (Exception ex)
      {
        return (false, $"{ex.Message} ({ex.HResult})");
      }
    }

    /// <summary>
    /// Deletes a KV set from the Gemus.ini file.
    /// </summary>
    public static (bool success, string error) DeleteKvSet(string setName, string gemusIniFile)
    {
      try
      {
        if (!File.Exists(gemusIniFile))
          return (true, string.Empty);

        var sb = new StringBuilder();
        bool inSet = false;

        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string trimmedLine = rawLine.Trim().ToLower();

          if (trimmedLine == $"[{setName.ToLower()}]")
          {
            inSet = true;
          }
          else
          {
            if (trimmedLine.StartsWith('[') && trimmedLine.EndsWith(']'))
            {
              inSet = false;
            }

            if (!inSet)
            {
              sb.AppendLine(rawLine);
            }
          }
        }

        File.WriteAllText(gemusIniFile, sb.ToString());
        return (true, string.Empty);
      }
      catch (Exception ex)
      {
        return (false, $"{ex.Message} ({ex.HResult})");
      }
    }

    /// <summary>
    /// Renames a KV set in the Gemus.ini file.
    /// </summary>
    public static (bool success, string error) RenameKvSet(string currentName, string newName, string gemusIniFile)
    {
      if (string.Equals(currentName, "global", StringComparison.OrdinalIgnoreCase) ||
          string.Equals(newName, "global", StringComparison.OrdinalIgnoreCase))
      {
        return (false, "The KVSet cannot be renamed from/to \"global\".");
      }

      try
      {
        if (!File.Exists(gemusIniFile))
          return (false, "The KVSet could not be found.");

        var sb = new StringBuilder();

        foreach (string rawLine in File.ReadLines(gemusIniFile))
        {
          string trimmedLine = rawLine.Trim();

          if (string.Equals(trimmedLine, $"[{currentName}]", StringComparison.OrdinalIgnoreCase))
          {
            sb.AppendLine($"[{newName}]");
          }
          else
          {
            sb.AppendLine(rawLine);
          }
        }

        File.WriteAllText(gemusIniFile, sb.ToString());
        return (true, string.Empty);
      }
      catch (Exception ex)
      {
        return (false, $"{ex.Message} ({ex.HResult})");
      }
    }

    /// <summary>
    /// Writes a value to a config file (key=value format without sections).
    /// </summary>
    public static bool WriteCfgValue(string filePath, string key, string value)
    {
      try
      {
        var lines = File.Exists(filePath) ? new List<string>(File.ReadAllLines(filePath)) : new List<string>();
        bool found = false;

        for (int i = 0; i < lines.Count; i++)
        {
          string line = lines[i].TrimStart();
          int eqPos = line.IndexOf('=');
          if (eqPos > 0)
          {
            string lineKey = line[..eqPos].Trim();
            if (string.Equals(lineKey, key, StringComparison.OrdinalIgnoreCase))
            {
              lines[i] = $"{key}={value}";
              found = true;
              break;
            }
          }
        }

        if (!found)
        {
          lines.Add($"{key}={value}");
        }

        File.WriteAllLines(filePath, lines);
        return true;
      }
      catch
      {
        return false;
      }
    }
  }
}
