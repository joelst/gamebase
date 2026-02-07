// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Core model types shared across the Gemus library
// Converted from VB6 modFuncs.bas and clsGemus.cls types

namespace Gemus.Models
{
  /// <summary>
  /// Represents a key=value pair from the game notes.
  /// </summary>
  public class NotesKeyValue
  {
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
  }

  /// <summary>
  /// Represents a line from the key=value data, which may or may not be a KV pair.
  /// </summary>
  public class NotesKeyValueLine
  {
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public bool IsKeyValue { get; set; }
    public string Line { get; set; } = string.Empty;
    public bool Flagged { get; set; }
  }

  /// <summary>
  /// Fields that can be manually specified for running games.
  /// </summary>
  public class ManualGemusFields
  {
    public int Control { get; set; }
    public int Players { get; set; }
    public int PalNtsc { get; set; }
    public bool TrueDriveEmu { get; set; }
    public string Comment { get; set; } = string.Empty;
    public string VersionComment { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
    public bool OkClicked { get; set; }
    public bool Defaults { get; set; }
  }

  /// <summary>
  /// Emulator configuration data.
  /// </summary>
  public class EmulatorConfig
  {
    public string Name { get; set; } = string.Empty;
    public string FullPathAndFilename { get; set; } = string.Empty;
    public string FullPathOnly { get; set; } = string.Empty;
    public string FullPathOnlyWithBackSlash { get; set; } = string.Empty;
    public string FullFilenameOnly { get; set; } = string.Empty;
    public string ShortPathAndFilename { get; set; } = string.Empty;
    public string ShortPathOnly { get; set; } = string.Empty;
    public string ShortPathOnlyWithBackSlash { get; set; } = string.Empty;
    public string ShortFilenameOnly { get; set; } = string.Empty;
    public string Script { get; set; } = string.Empty;
    public bool UseShortFilenames { get; set; }
    public PathLocationType PathType { get; set; }
    public List<string> Files { get; set; } = new();
    public string Types { get; set; } = string.Empty;
    public bool UseDefaultTypes { get; set; }
    public long Wait { get; set; }
  }

  /// <summary>
  /// Emulator properties for the editor dialog.
  /// </summary>
  public class EmulatorProperties
  {
    public string Name { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;
    public string Script { get; set; } = string.Empty;
    public bool UseShort { get; set; }
    public List<string> Files { get; set; } = new();
    public string Types { get; set; } = string.Empty;
    public bool UseDefaultTypes { get; set; }
    public string GBName { get; set; } = string.Empty;
    public string EmuFile { get; set; } = string.Empty;
    public long Wait { get; set; }
  }

  /// <summary>
  /// Path location type (drive type).
  /// </summary>
  public enum PathLocationType
  {
    Removable,
    FixedHD,
    CD,
    Invalid
  }

  /// <summary>
  /// Path configuration with resolved path and type info.
  /// </summary>
  public class PathInfo
  {
    public string WithBackSlash { get; set; } = string.Empty;
    public string RealPath { get; set; } = string.Empty;
    public PathLocationType PathType { get; set; } = PathLocationType.Invalid;
    public bool IsSet { get; set; }
  }

  /// <summary>
  /// Script comparison field definition.
  /// </summary>
  public class ScriptField
  {
    public string Name { get; set; } = string.Empty;
    public List<string> Operators { get; set; } = new();
    public ScriptValueType ValueType { get; set; }
    public List<string> ListValues { get; set; } = new();
  }

  /// <summary>
  /// Script value types for comparison fields.
  /// </summary>
  public enum ScriptValueType
  {
    Number,
    List,
    Text
  }

  /// <summary>
  /// Script function parameter type.
  /// </summary>
  public enum ScriptFunctionParamType
  {
    FileNumber,
    IniSection,
    Text,
    List
  }

  /// <summary>
  /// Script function definition.
  /// </summary>
  public class ScriptFunction
  {
    public string Name { get; set; } = string.Empty;
    public List<ScriptFunctionParamType> ParamTypes { get; set; } = new();
    public List<string> ListValues { get; set; } = new();
  }

  /// <summary>
  /// Rezip additional files handling mode.
  /// </summary>
  public enum RezipAdditionalFiles
  {
    Ignore = 0,
    Zip = 1,
    Ask = 2
  }

  /// <summary>
  /// Archive browsing state.
  /// </summary>
  public class ArchiveBrowseInfo
  {
    public string ArchivePathFile { get; set; } = string.Empty;
    public string TempPathWithBackSlash { get; set; } = string.Empty;
    public List<string> RunnableFileTypes { get; set; } = new();
    public bool OkPressed { get; set; }
    public string RunnableFile { get; set; } = string.Empty;
  }

  /// <summary>
  /// Input box dialog configuration.
  /// </summary>
  public class InputBoxConfig
  {
    public bool AllowEmptyReturn { get; set; }
    public bool OkClicked { get; set; }
    public int MaxLength { get; set; }
    public string CancelButtonText { get; set; } = "&Cancel";
    public string Caption { get; set; } = string.Empty;
    public string InvalidChars { get; set; } = string.Empty;
    public string OkButtonText { get; set; } = "&OK";
    public string Prompt { get; set; } = string.Empty;
    public string ReturnedText { get; set; } = string.Empty;
    public string StartingText { get; set; } = string.Empty;
  }

  /// <summary>
  /// Script selector dialog state.
  /// </summary>
  public class ScriptSelectorConfig
  {
    public bool OkClicked { get; set; }
    public string SelectedFile { get; set; } = string.Empty;
    public string InitPath { get; set; } = string.Empty;
  }

  /// <summary>
  /// Wizard DAT file structures for emulator setup.
  /// </summary>
  public class WizardDatFile
  {
    public string MessagesFile { get; set; } = string.Empty;
    public bool Aborted { get; set; }
    public bool GlobalDat { get; set; }
    public string DatFile { get; set; } = string.Empty;
    public string EmuFile { get; set; } = string.Empty;
    public string GBFolder { get; set; } = string.Empty;
    public string ScriptFolder { get; set; } = string.Empty;
    public float DatVersion { get; set; }
    public string LastUpdated { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
    public string HomePage { get; set; } = string.Empty;
    public List<WizardDatSection> Sections { get; set; } = new();
  }

  public class WizardDatSection
  {
    public string Name { get; set; } = string.Empty;
    public List<WizardDatEmulator> Emulators { get; set; } = new();
  }

  public class WizardDatEmulator
  {
    public string EmulatorName { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string Info { get; set; } = string.Empty;
    public string Filename { get; set; } = string.Empty;
    public List<string> Crc32Values { get; set; } = new();
    public List<string> FileTypes { get; set; } = new();
    public bool UseShort { get; set; }
    public string ScriptFile { get; set; } = string.Empty;
    public List<string> AssociatedFiles { get; set; } = new();
    public string Found { get; set; } = string.Empty;
  }
}
