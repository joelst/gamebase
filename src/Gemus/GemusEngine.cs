// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Core script engine
// Converted from VB6 clsGemus.cls

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using GBFuncs;
using Gemus.DiskFormats;
using Gemus.Models;

namespace Gemus
{
  public interface IGemusUi
  {
    bool Confirm(string message, string caption);
    void ShowInfo(string message, string caption);
    void ShowWarning(string message, string caption);
    void ShowError(string message, string caption);
    string? PromptInput(InputBoxConfig config);
    void ShowKv(string kvText);
  }

  public class GemusEngine
  {
    private readonly FunctionsHelper _funcs;
    private readonly RegistryService _registry;
    private readonly GemusLogger _logger;
    private readonly ArchiveUnpacker _unpacker;

    public IGemusUi? Ui { get; set; }

    private bool _silentMode;
    private bool _simulateRun;
    private string _simulateKv = string.Empty;

    private string _pathsIniFile = string.Empty;
    private PathInfo _sidPath = new();
    private List<PathInfo> _gamePaths = new();
    private List<PathInfo> _screenshotPaths = new();
    private List<PathInfo> _extrasPaths = new();
    private PathInfo _photosPath = new();
    private bool _validGamePathExists;

    private const string NumbersAllowed = "0123456789";
    private bool _useShort;
    private int _emuIdx;
    private const int PlayersUnknown = -1;
    private bool _scriptFileErrorDone;

    private string _appPathWithBackSlash = string.Empty;
    private string _langIniPathFile = string.Empty;
    private string _appName = "GEMUS";
    private string _commandCom = "";
    private string _emuFile = string.Empty;
    private string _gbName = string.Empty;
    private string _gbGamePath = string.Empty;
    private bool _waitIfNotRezipping;

    private bool _userWantsRepack;
    private bool _askBeforeRepack;
    private bool _notifyRepackOk;
    private bool _attemptRepack;
    private bool _rezipLegacy;
    private RezipAdditionalFiles _repackAdditional = RezipAdditionalFiles.Ignore;

    private readonly List<EmulatorConfig> _emulators = new();

    private readonly List<string> _emuFileTypes = new();
    private const string DefaultEmuFileTypes = "";

    private readonly List<ScriptField> _scriptFields = new();
    private readonly List<ScriptFunction> _scriptFunctions = new();
    private readonly List<string> _scriptEnvVars = new();

    private List<string> _scriptLines = new();
    private List<string> _scriptLinesPre = new();
    private readonly List<string> _scriptLinesToProcess = new();

    private string _langEndOfLineExpected = string.Empty;
    private string _langScriptError = string.Empty;
    private string _langLine = string.Empty;
    private string _langUnknownCommand = string.Empty;
    private string _langMissingParams = string.Empty;
    private string _langTooManyParams = string.Empty;
    private string _langExpected = string.Empty;
    private string _langGameNotSupported = string.Empty;
    private string _langInvalidFunctionParam = string.Empty;
    private string _langNumber = string.Empty;
    private string _langExpressionReqd = string.Empty;
    private string _langUnknownField = string.Empty;
    private string _langInvalidOp = string.Empty;
    private string _langInvalidValue = string.Empty;
    private string _langWithout = string.Empty;
    private string _langTooMany = string.Empty;
    private string _langStatements = string.Empty;

    private int _fldNumPlayersFrom;
    private int _fldNumPlayersTo;
    private int _fldControl;
    private string _fldGameComment = string.Empty;
    private string _fldVersionComment = string.Empty;
    private int _fldPalNtsc;
    private bool _fldTrueDriveEmu;

    private readonly List<NotesKeyValue> _notesKeys = new();

    private string _fldGameType = string.Empty;
    private int _fldNumGameFiles;

    private int _imageIndex;
    private string _imageName = string.Empty;
    private string _c64ImageName = string.Empty;
    private string _envGamePathFile = string.Empty;
    private string _envGameFile = string.Empty;
    private string _envGamePath = string.Empty;
    private string _envGameFileNoExt = string.Empty;
    private readonly string[] _extraGamePathFile = new string[21];
    private readonly string[] _extraGamePath = new string[21];
    private readonly string[] _extraGameFile = new string[21];
    private readonly string[] _extraGameFileNoExt = new string[21];

    private string _commandLine = string.Empty;

    public GemusEngine(FunctionsHelper funcs, RegistryService registry, GemusLogger logger, ArchiveUnpacker unpacker)
    {
      _funcs = funcs;
      _registry = registry;
      _logger = logger;
      _unpacker = unpacker;
      _commandCom = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
      _commandCom = _commandCom + " /c ";
    }

    public void SetAppContext(string appPathWithBackSlash, string langIniPathFile, string appName)
    {
      _appPathWithBackSlash = appPathWithBackSlash.EndsWith("\\") ? appPathWithBackSlash : appPathWithBackSlash + "\\";
      _langIniPathFile = langIniPathFile;
      _appName = appName;
    }

    public void SetSilentMode(bool silentMode) => _silentMode = silentMode;

    public void SetWaitIfNotRezipping(bool wait) => _waitIfNotRezipping = wait;

    public void SetRezipLegacy(bool useLegacy) => _rezipLegacy = useLegacy;

    public void SetRezipNotify(bool notify) => _notifyRepackOk = notify;

    public void SetRezip(bool doRezip, bool userConfirm, RezipAdditionalFiles additionalFilesRule)
    {
      _userWantsRepack = doRezip;
      _askBeforeRepack = userConfirm;
      _repackAdditional = additionalFilesRule;
    }

    public void SetArchiveBackend(ArchiveBackend backend) => SetArchiveBackend(backend, false);

    public void SetArchiveBackend(ArchiveBackend backend, bool persistToIni)
    {
      _unpacker.Backend = backend;
      if (persistToIni && !string.IsNullOrEmpty(_gbName) && !string.IsNullOrEmpty(_appPathWithBackSlash))
      {
        string gemusIni = Path.Combine(_appPathWithBackSlash, _gbName, "Gemus.ini");
        GemusIniHelper.WriteIniValue("Settings", "ArchiveBackend", backend.ToString(), gemusIni);
      }
    }

    public bool TrySetArchiveBackend(string? backend, out ArchiveBackend parsed) => TrySetArchiveBackend(backend, false, out parsed);

    public bool TrySetArchiveBackend(string? backend, bool persistToIni, out ArchiveBackend parsed)
    {
      if (ArchiveUnpacker.TryParseBackend(backend, out parsed))
      {
        SetArchiveBackend(parsed, persistToIni);
        return true;
      }
      return false;
    }

    public void Init(string gbName)
    {
      _emuFile = Path.Combine(_appPathWithBackSlash, gbName, "Emulators.ini");
      _gbName = gbName;

      // Prefer Gemus.ini [Settings] ArchiveBackend if present
      string gemusIni = Path.Combine(_appPathWithBackSlash, _gbName, "Gemus.ini");
      string backendValue = GemusIniHelper.GetIniMessage("Settings", "ArchiveBackend", string.Empty, gemusIni);
      if (!string.IsNullOrWhiteSpace(backendValue))
      {
        if (TrySetArchiveBackend(backendValue, out var parsedBackend))
          _logger.LogStatus($"Archive backend set from Gemus.ini: {parsedBackend}");
        else
          _logger.LogStatus($"Archive backend in Gemus.ini not recognized: '{backendValue}'");
      }

      _pathsIniFile = Path.Combine(_appPathWithBackSlash, _gbName, "Paths.ini");
      LoadPaths();

      LoadScriptData();

      string fileTypes = _funcs.INI_GetSetting("File Types", "SupportTypes", "", _emuFile);
      SetEmuFileTypes(fileTypes);

      ReadInEmulators();

      _langEndOfLineExpected = GetIniMessage("Emulator Script", "Error_EndOfLineExpected", "End of line expected!");
      _langScriptError = GetIniMessage("Emulator Script", "GEMUSError", "GEMUS Error");
      _langLine = GetIniMessage("Emulator Script", "Line", "Line");
      _langUnknownCommand = GetIniMessage("Emulator Script", "Error_UnknownCommand", "Unknown function or statement!");
      _langMissingParams = GetIniMessage("Emulator Script", "Error_MissingParams", "Too few function parameters!");
      _langTooManyParams = GetIniMessage("Emulator Script", "Error_TooManyParams", "Too many function parameters!");
      _langExpected = GetIniMessage("Emulator Script", "Error_Expected", "Expected");
      _langGameNotSupported = GetIniMessage("Emulator Script", "Error_GameNotSupported", "This game type is not supported by this emulator!");
      _langInvalidFunctionParam = GetIniMessage("Emulator Script", "Error_InvalidFunctionParam", "Invalid function parameter");
      _langNumber = GetIniMessage("Emulator Script", "Error_Number", "Number");
      _langExpressionReqd = GetIniMessage("Emulator Script", "Error_ExpressionRequired", "Expression Required!");
      _langUnknownField = GetIniMessage("Emulator Script", "Error_UnknownField", "Unknown comparison field!");
      _langInvalidOp = GetIniMessage("Emulator Script", "Error_InvalidOp", "Missing or invalid comparison operator!");
      _langInvalidValue = GetIniMessage("Emulator Script", "Error_InvalidValue", "Missing or invalid comparison value!");
      _langWithout = GetIniMessage("Emulator Script", "Error_Without", "without");
      _langTooMany = GetIniMessage("Emulator Script", "Error_TooMany", "Too many");
      _langStatements = GetIniMessage("Emulator Script", "Error_Statements", "statements");
    }

    public string SimulateRunEmulator(out string kvText, int emuIdx, string gamePathFile, int imageIndex,
      ref string defaultFile, string extractDestFolder, bool fromDb,
      int? numPlayersFrom = null, int? numPlayersTo = null, int? control = null, string? gameComment = null,
      string? versionComment = null, int? palNtsc = null, bool? trueDriveEmu = null, string? notes = null)
    {
      _simulateRun = true;
      _simulateKv = string.Empty;
      string result = RunEmulator(emuIdx, gamePathFile, imageIndex, ref defaultFile, extractDestFolder, fromDb,
        numPlayersFrom, numPlayersTo, control, gameComment, versionComment, palNtsc, trueDriveEmu, notes);
      _simulateRun = false;
      kvText = _simulateKv;
      return result;
    }

    public string RunEmulator(int emuIdx, string gamePathFile, int imageIndex,
      ref string defaultFile, string extractDestFolder, bool fromDb,
      int? numPlayersFrom = null, int? numPlayersTo = null, int? control = null, string? gameComment = null,
      string? versionComment = null, int? palNtsc = null, bool? trueDriveEmu = null, string? notes = null)
    {
      _logger.CreateLogFile(Path.Combine(_appPathWithBackSlash, "Gemus.log"));
      _logger.PushLogFileTrace("GemusEngine.RunEmulator");
      _logger.LogStatus($"Arg Silent Mode: [{_silentMode}]");
      _logger.LogStatus($"Arg gamePathFile: [{gamePathFile}]");
      _logger.LogStatus($"Arg imageIndex: [{imageIndex}]");
      _logger.LogStatus($"Arg defaultFile: [{defaultFile}]");
      _logger.LogStatus($"Arg extractDestFolder: [{extractDestFolder}]");

      _gbGamePath = extractDestFolder;
      _scriptFileErrorDone = false;

      string destFolderWbs = extractDestFolder.EndsWith("\\") ? extractDestFolder : extractDestFolder + "\\";

      List<string> extractedTypes = GetAllValidFileTypes();
      _logger.LogStatus("GetAllValidFileTypesArray: " + string.Join(";", extractedTypes));

      _logger.LogStatus("Starting Unpacker...");
      bool unpackOk = _unpacker.Unpack(gamePathFile, defaultFile, extractedTypes, destFolderWbs,
        out bool isRepackPossible, out string fileToRun, out bool userPick, out string unpackError);

      if (!unpackOk)
      {
        _logger.LogStatus($"Unpacker failed with error [{unpackError}]");
        _logger.CloseLogFile();
        return unpackError;
      }

      if (userPick)
      {
        _logger.LogStatus("User needs to pick runnable file within archive");
        if (!_silentMode)
        {
          if (_unpacker.GetRunnableFileInArchiveFromUser(gamePathFile, destFolderWbs, extractedTypes, out string selected))
          {
            defaultFile = selected;
            _logger.LogStatus($"User picked runnable file [{defaultFile}]");

            _logger.LogStatus("Starting Unpacker (again)...");
            unpackOk = _unpacker.Unpack(gamePathFile, defaultFile, extractedTypes, destFolderWbs,
              out isRepackPossible, out fileToRun, out userPick, out unpackError);
            if (!unpackOk)
            {
              _logger.LogStatus($"Unpacker failed with error [{unpackError}]");
              _logger.CloseLogFile();
              return unpackError;
            }
            if (userPick)
            {
              string msg = GetIniMessage("Running Games", "RunnableFileNotDetermined",
                "Runnable file within archive could not be determined");
              _logger.LogStatus("Unpacker unable to determine runnable file within archive.");
              _logger.CloseLogFile();
              return msg;
            }
          }
          else
          {
            string msg = GetIniMessage("Running Games", "AbortedRunningGamePleaseSelect",
              "Aborted running game. Please select a runnable file for this archive");
            _logger.LogStatus("Unpacker unable to determine runnable file within archive. User aborted.");
            _logger.CloseLogFile();
            return msg;
          }
        }
        else
        {
          string msg = GetIniMessage("Running Games", "RunnableFileNotAutoDetermined",
            "Runnable file within archive could not be automatically determined");
          _logger.LogStatus("Unpacker unable to auto-determine runnable file within archive. Silent mode.");
          _logger.CloseLogFile();
          return msg;
        }
      }
      else
      {
        _logger.LogStatus($"Runnable file determined [{fileToRun}]");
      }

      _attemptRepack = false;
      if (_userWantsRepack && isRepackPossible)
      {
        if (!ProcessHelper.IsCdRomDrive(gamePathFile.Length >= 2 ? gamePathFile[..2] : gamePathFile))
        {
          _attemptRepack = true;
          _logger.LogStatus("Will attempt to repack after emulator closes");
          _unpacker.StorePreUnpackedFilesSnapshot(_unpacker.GetFinalFolder());
          _logger.LogStatus($"Stored unpacked files snapshot from [{_unpacker.GetFinalFolder()}]");
        }
      }

      string notesText = notes ?? string.Empty;
      if (!fromDb)
      {
        if (_registry.GetSetting_HKCU_Fallback_HKLM(
          $"Software\\GB64\\GameBase\\GameBase\\{_gbName}\\GEMUSFields", "Notes",
          out object? regNotes, string.Empty))
        {
          notesText = regNotes?.ToString() ?? string.Empty;
        }
      }

      _notesKeys.Clear();
      _notesKeys.AddRange(GemusIniHelper.GetKvSet("global", Path.Combine(_appPathWithBackSlash, _gbName, "Gemus.ini")));
      _logger.LogStatus($"Loaded {_notesKeys.Count} global k=v pairs");

      if (!string.IsNullOrWhiteSpace(notesText))
      {
        _logger.LogStatus("Looking for kvsets to load...");
        foreach (string rawLine in notesText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
        {
          string line = rawLine.Trim();
          int eqPos = line.IndexOf('=');
          if (eqPos > 1)
          {
            string key = line[..eqPos].Trim();
            if (string.Equals(key, "kvset", StringComparison.OrdinalIgnoreCase))
            {
              string val = line[(eqPos + 1)..].Trim();
              if (!string.IsNullOrEmpty(val) && !string.Equals(val, "global", StringComparison.OrdinalIgnoreCase))
              {
                List<NotesKeyValue> setItems = GemusIniHelper.GetKvSet(val, Path.Combine(_appPathWithBackSlash, _gbName, "Gemus.ini"));
                _logger.LogStatus($"Loaded {setItems.Count} k=v pairs in kvset \"{val}\"");
                MergeKeyValues(setItems);
              }
            }
          }
        }

        foreach (string rawLine in notesText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
        {
          string line = rawLine.Trim();
          int eqPos = line.IndexOf('=');
          if (eqPos > 1)
          {
            string key = line[..eqPos].Trim();
            if (!string.IsNullOrEmpty(key) && !string.Equals(key, "kvset", StringComparison.OrdinalIgnoreCase) && !key.StartsWith(";"))
            {
              string val = line[(eqPos + 1)..].Trim();
              UpsertKeyValue(key, val);
            }
          }
        }

        _logger.LogStatus($"Finished loading all k=v pairs. Total loaded: {_notesKeys.Count}");
      }
      else
      {
        _logger.LogStatus("No game-level k=v pairs to process");
      }

      if (emuIdx == -1)
      {
        string testName = string.Empty;
        string testNameDisplay = string.Empty;
        foreach (var kv in _notesKeys)
        {
          if (string.Equals(kv.Key, "EMU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(kv.Value))
          {
            testName = kv.Value.ToUpperInvariant();
            testNameDisplay = kv.Value;
            _logger.LogStatus($"Found [emu={testNameDisplay}] k=v pair");
            break;
          }
        }

        if (!string.IsNullOrEmpty(testName))
        {
          bool found = false;
          for (int i = 0; i < _emulators.Count; i++)
          {
            if (string.Equals(_emulators[i].Name, testName, StringComparison.OrdinalIgnoreCase))
            {
              emuIdx = i;
              found = true;
              break;
            }
          }

          if (!found)
          {
            string msg = GetIniMessage("Running Games", "NoMatchingEmulator",
              "No emulator name was found to match your 'emu=' key=value pair.");
            _logger.CloseLogFile();
            return msg + "\r\n\r\n[emu=" + testNameDisplay + "]";
          }
        }
      }

      string ext = Path.GetExtension(fileToRun).TrimStart('.');

      if (emuIdx >= 0 && emuIdx < _emulators.Count)
      {
        var emu = _emulators[emuIdx];
        if (!emu.UseDefaultTypes)
        {
          if (FileTypeSupported(emu.Types, ext))
          {
            _emuIdx = emuIdx;
          }
          else
          {
            string msg = GetIniMessage("Running Games", "Status_NotSupportedByEmulator",
              "This game type is not supported by this emulator.");
            _logger.CloseLogFile();
            return msg + " [" + ext + "]";
          }
        }
        else
        {
          _emuIdx = emuIdx;
        }

        _logger.LogStatus($"Using emulator with index {_emuIdx} [{_emulators[_emuIdx].Name}]");
      }
      else
      {
        _logger.LogStatus($"Finding first emulator that supports gamefile type [{ext}]");
        _emuIdx = -1;
        for (int i = 0; i < _emulators.Count; i++)
        {
          var emu = _emulators[i];
          if (!emu.UseDefaultTypes)
          {
            if (FileTypeSupported(emu.Types, ext))
            {
              _emuIdx = i;
              break;
            }
          }
          else
          {
            _emuIdx = i;
            break;
          }
        }

        if (_emuIdx == -1)
        {
          string msg = GetIniMessage("Running Games", "Status_NotSupportedByAnyEmulator",
            "An emulator that supports this game type was not found.");
          _logger.CloseLogFile();
          return msg + " [" + ext + "]";
        }

        _logger.LogStatus($"Using emulator with index {_emuIdx} [{_emulators[_emuIdx].Name}]");
      }

      _useShort = _emulators[_emuIdx].UseShortFilenames;
      _imageIndex = imageIndex;
      _imageName = GetImageNameFromGameFile(fileToRun, _imageIndex);
      _c64ImageName = ConvertToC64Lower(_imageName);

      if (fromDb)
      {
        _fldNumPlayersFrom = numPlayersFrom ?? 0;
        _fldNumPlayersTo = numPlayersTo ?? 0;
        _fldControl = control ?? 0;
        _fldGameComment = gameComment ?? string.Empty;
        _fldVersionComment = versionComment ?? string.Empty;
        _fldPalNtsc = palNtsc ?? 0;
        _fldTrueDriveEmu = trueDriveEmu ?? false;
      }
      else
      {
        _fldNumPlayersFrom = 1;
        _fldNumPlayersTo = GetRegistryInt("Players", 1);
        _fldControl = GetRegistryInt("Control", 0);
        _fldPalNtsc = GetRegistryInt("PalNTSC", 0);
        _fldTrueDriveEmu = GetRegistryBool("TDE", false);
        _fldGameComment = GetRegistryString("Comment", string.Empty);
        _fldVersionComment = GetRegistryString("VComment", string.Empty);
      }

      _logger.LogStatus("Will be using these game field values:");
      _logger.LogStatus($" m_lngFldNumPlayersFrom [{_fldNumPlayersFrom}]");
      _logger.LogStatus($" m_lngFldNumPlayersTo [{_fldNumPlayersTo}]");
      _logger.LogStatus($" m_lngFldControl [{_fldControl}]");
      _logger.LogStatus($" m_strFldGameComment [{_fldGameComment}]");
      _logger.LogStatus($" m_strFldVersionComment [{_fldVersionComment}]");
      _logger.LogStatus($" m_lngFldPALNTSC [{_fldPalNtsc}]");
      _logger.LogStatus($" m_blnFldTrueDriveEmu [{_fldTrueDriveEmu}]");

      _envGamePathFile = fileToRun.ToLowerInvariant();
      _fldGameType = Path.GetExtension(fileToRun).TrimStart('.').ToUpperInvariant();
      _fldNumGameFiles = 1;
      _funcs.GetFileAndPath(_envGamePathFile, out _envGamePath, out _envGameFile);
      _envGamePath = _envGamePath.ToLowerInvariant();
      _envGameFile = _envGameFile.ToLowerInvariant();

      string pathWbs = _envGamePath;
      if (!pathWbs.EndsWith("\\")) pathWbs += "\\";
      _envGameFileNoExt = Path.GetFileNameWithoutExtension(_envGameFile);

      string lastChar = _envGameFileNoExt.Length > 0 ? _envGameFileNoExt[^1].ToString().ToUpperInvariant() : "";
      if (lastChar.Length > 0 && "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(lastChar))
      {
        int extraIdx = lastChar[0] + 1;
        for (int i = 1; i <= 20; i++)
        {
          _extraGamePathFile[i] = pathWbs + _envGameFileNoExt[..^1] + char.ToLower((char)extraIdx) + "." + _fldGameType.ToLowerInvariant();
          if (File.Exists(_extraGamePathFile[i]))
          {
            _fldNumGameFiles++;
            extraIdx++;
            _funcs.GetFileAndPath(_extraGamePathFile[i], out _extraGamePath[i], out _extraGameFile[i]);
            _extraGameFileNoExt[i] = Path.GetFileNameWithoutExtension(_extraGameFile[i]);
          }
          else
          {
            break;
          }
        }
      }
      else if (lastChar.Length > 0 && "0123456789".Contains(lastChar))
      {
        int extraIdx = int.Parse(lastChar) + 1;
        for (int i = 1; i <= 20; i++)
        {
          _extraGamePathFile[i] = pathWbs + _envGameFileNoExt[..^1] + extraIdx + "." + _fldGameType.ToLowerInvariant();
          if (File.Exists(_extraGamePathFile[i]))
          {
            _fldNumGameFiles++;
            extraIdx++;
            _funcs.GetFileAndPath(_extraGamePathFile[i], out _extraGamePath[i], out _extraGameFile[i]);
            _extraGameFileNoExt[i] = Path.GetFileNameWithoutExtension(_extraGameFile[i]);
          }
          else
          {
            break;
          }
        }
      }
      else
      {
        for (int i = 1; i <= 20; i++)
        {
          _extraGamePathFile[i] = _envGamePathFile;
          _extraGamePath[i] = _envGamePath;
          _extraGameFile[i] = _envGameFile;
          _extraGameFileNoExt[i] = _envGameFileNoExt;
        }
      }

      _logger.LogStatus("All script environment vars created");

      _logger.LogStatus("Final k=v pairs are:");
      foreach (var kv in _notesKeys)
      {
        string value = kv.Value;
        ReplaceEnvironmentVars(ref value, false, false);
        kv.Value = value;
        _logger.LogStatus($" {kv.Key}={kv.Value}");
      }

      if (_simulateRun)
      {
        var sb = new StringBuilder();
        foreach (var kv in _notesKeys)
        {
          sb.AppendLine(kv.Key + "=" + kv.Value);
        }
        _simulateKv = sb.ToString().TrimEnd();
        _simulateRun = false;
        _logger.LogStatus("Simulating run. Aborting.");
        _logger.CloseLogFile();
        return string.Empty;
      }

      if (_emuIdx >= 0 && _emuIdx < _emulators.Count)
      {
        if (File.Exists(_emulators[_emuIdx].FullPathAndFilename))
        {
          _logger.LogStatus($"Emulator executable found: [{_emulators[_emuIdx].FullPathAndFilename}]");

          string scriptFile = _emulators[_emuIdx].Script;
          if (!string.IsNullOrWhiteSpace(scriptFile))
          {
            string scriptPath = Path.Combine(_appPathWithBackSlash, _gbName, "Scripts", scriptFile);
            if (File.Exists(scriptPath))
            {
              string scriptText = File.ReadAllText(scriptPath);
              _logger.LogStatus($"Loaded GEMUS Script [{scriptPath}]");

              string validationError = ValidateAndGetScriptLines(scriptText);
              if (string.IsNullOrEmpty(validationError))
              {
                _logger.LogStatus("Validated Script lines");

                string structureError = ValidateScriptStructure(scriptText, out int errorLine);
                if (string.IsNullOrEmpty(structureError))
                {
                  _logger.LogStatus("Validated Script structure");

                  _scriptLinesToProcess.Clear();
                  GetScriptFunctionsToProcess(0, out int nextLine, 0, true);

                  if (_scriptLinesToProcess.Count > 0)
                  {
                    _logger.LogStatus($"Processing {_scriptLinesToProcess.Count} script functions...");
                    _commandLine = string.Empty;

                    foreach (string line in _scriptLinesToProcess)
                    {
                      _logger.LogStatus(line);
                      string err = ProcessScriptLine(line);
                      if (!string.IsNullOrEmpty(err))
                      {
                        _logger.LogStatus($"Error returned: [{err}]");
                        _logger.CloseLogFile();
                        return err;
                      }
                    }

                    if (_attemptRepack)
                    {
                      bool repackAsk = !_silentMode && _askBeforeRepack;
                      bool repackNotify = !_silentMode && _notifyRepackOk;
                      if (_unpacker.Repack(repackAsk, repackNotify, destFolderWbs, gamePathFile,
                        _repackAdditional, out bool repacked, out string repackError))
                      {
                        _logger.CloseLogFile();
                        return string.Empty;
                      }

                      _logger.LogStatus($"Error returned: [{repackError}]");
                      _logger.CloseLogFile();
                      return repackError;
                    }

                    _logger.CloseLogFile();
                    return string.Empty;
                  }

                  string noFunctions = GetIniMessage("Running Games", "NoFunctions",
                    "The script for this game has no functions to process!");
                  _logger.LogStatus($"Error returned: [{noFunctions}]");
                  _logger.CloseLogFile();
                  return noFunctions;
                }

                string errorMsg = errorLine > -1
                  ? _langScriptError + "\r\n\r\n" + structureError + "\r\n\r\n" + _langLine + " " + errorLine
                  : _langScriptError + "\r\n\r\n" + structureError;

                _logger.LogStatus($"Error returned: [{errorMsg}]");
                _logger.CloseLogFile();
                return errorMsg;
              }

              if (_scriptLines.Count == 0)
              {
                string noScript = GetIniMessage("Running Games", "NoScript", "There is no script set for this emulator!");
                _logger.LogStatus($"Error returned: [{noScript}]");
                _logger.CloseLogFile();
                return noScript;
              }

              string badLine = _langScriptError + "\r\n\r\n" + validationError;
              _logger.LogStatus($"Error returned: [{badLine}]");
              _logger.CloseLogFile();
              return badLine;
            }

            string missingScript = GetIniMessage("Running Games", "NoScript", "There is no script set for this emulator!");
            _logger.LogStatus($"Error returned: [{missingScript}]");
            _logger.CloseLogFile();
            return missingScript;
          }

          string noScriptSet = GetIniMessage("Running Games", "NoScript", "There is no script set for this emulator!");
          _logger.LogStatus($"Error returned: [{noScriptSet}]");
          _logger.CloseLogFile();
          return noScriptSet;
        }

        string emuNotFound = GetIniMessage("Running Games", "EmulatorNotFound", "The emulator file could not be found.")
          + "\r\n\r\n'" + _emulators[_emuIdx].FullPathAndFilename + "'";
        _logger.LogStatus($"Error returned: [{emuNotFound}]");
        _logger.CloseLogFile();
        return emuNotFound;
      }

      string invalidEmu = GetIniMessage("Running Games", "InvalidEmulator", "Invalid Emulator!");
      _logger.LogStatus($"Error returned: [{invalidEmu}]");
      _logger.CloseLogFile();
      return invalidEmu;
    }

    private void MergeKeyValues(List<NotesKeyValue> items)
    {
      foreach (var kv in items)
      {
        UpsertKeyValue(kv.Key, kv.Value);
      }
    }

    private void UpsertKeyValue(string key, string value)
    {
      for (int i = 0; i < _notesKeys.Count; i++)
      {
        if (string.Equals(_notesKeys[i].Key, key, StringComparison.OrdinalIgnoreCase))
        {
          _notesKeys[i].Value = value;
          return;
        }
      }
      _notesKeys.Add(new NotesKeyValue { Key = key, Value = value });
    }

    private int GetRegistryInt(string valueName, int defaultValue)
    {
      if (_registry.GetSetting_HKCU_Fallback_HKLM(
        $"Software\\GB64\\GameBase\\GameBase\\{_gbName}\\GEMUSFields", valueName,
        out object? value, defaultValue))
      {
        return Convert.ToInt32(value ?? defaultValue);
      }
      return defaultValue;
    }

    private bool GetRegistryBool(string valueName, bool defaultValue)
    {
      if (_registry.GetSetting_HKCU_Fallback_HKLM(
        $"Software\\GB64\\GameBase\\GameBase\\{_gbName}\\GEMUSFields", valueName,
        out object? value, defaultValue))
      {
        return Convert.ToBoolean(value ?? defaultValue);
      }
      return defaultValue;
    }

    private string GetRegistryString(string valueName, string defaultValue)
    {
      if (_registry.GetSetting_HKCU_Fallback_HKLM(
        $"Software\\GB64\\GameBase\\GameBase\\{_gbName}\\GEMUSFields", valueName,
        out object? value, defaultValue))
      {
        return value?.ToString() ?? defaultValue;
      }
      return defaultValue;
    }

    private string GetIniMessage(string section, string key, string defaultValue)
    {
      return GemusIniHelper.GetIniMessage(section, key, defaultValue, _langIniPathFile);
    }

    private void LoadPaths()
    {
      string sidPath = _funcs.INI_GetSetting("Misc", "Music", "", _pathsIniFile);
      _sidPath = GetPathInfo(sidPath);

      string photosPath = _funcs.INI_GetSetting("Misc", "Photos", "", _pathsIniFile);
      _photosPath = GetPathInfo(photosPath);

      _gamePaths = LoadPathList("Games");
      _screenshotPaths = LoadPathList("Pictures");
      _extrasPaths = LoadPathList("Extras");

      _validGamePathExists = _gamePaths.Any(p => p.PathType != PathLocationType.Invalid);
    }

    private List<PathInfo> LoadPathList(string section)
    {
      var results = new List<PathInfo>();
      int idx = 1;
      while (true)
      {
        string value = _funcs.INI_GetSetting(section, idx.ToString(), "", _pathsIniFile);
        if (string.IsNullOrEmpty(value))
          break;

        results.Add(GetPathInfo(value));
        idx++;
      }
      return results;
    }

    private PathInfo GetPathInfo(string path)
    {
      var info = new PathInfo { RealPath = path };
      if (string.IsNullOrEmpty(info.RealPath) || info.RealPath.Length <= 2)
      {
        info.PathType = PathLocationType.Invalid;
        return info;
      }

      if (info.RealPath.StartsWith("\\"))
      {
        info.RealPath = _appPathWithBackSlash[..3] + info.RealPath[1..];
      }

      if (info.RealPath.Length >= 2 && info.RealPath[1] != ':')
      {
        info.RealPath = Path.Combine(_appPathWithBackSlash, info.RealPath);
        if (Directory.Exists(info.RealPath))
        {
          info.RealPath = Directory.GetCurrentDirectory();
        }
      }

      if (!Directory.Exists(info.RealPath))
      {
        if (info.RealPath.Length >= 3)
        {
          string validateLocal = _appPathWithBackSlash[..3] + info.RealPath[3..];
          if (Directory.Exists(validateLocal))
          {
            info.RealPath = validateLocal;
          }
        }
      }

      info.WithBackSlash = info.RealPath.EndsWith("\\") ? info.RealPath : info.RealPath + "\\";
      info.PathType = ProcessHelper.GetDriveLocationType(info.RealPath.Length >= 3 ? info.RealPath[..3] : info.RealPath);

      if (info.PathType == PathLocationType.FixedHD && !Directory.Exists(info.RealPath))
      {
        info.PathType = PathLocationType.Invalid;
      }

      return info;
    }

    private List<string> GetAllValidFileTypes()
    {
      var types = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

      string fileTypes = _funcs.INI_GetSetting("File Types", "SupportTypes", "", _emuFile);
      SetEmuFileTypes(fileTypes);

      bool defaultsDone = false;
      foreach (var emu in _emulators)
      {
        if (emu.UseDefaultTypes)
        {
          if (!defaultsDone)
          {
            foreach (string t in _emuFileTypes)
            {
              types.Add(t.ToUpperInvariant());
            }
            defaultsDone = true;
          }
        }
        else if (!string.IsNullOrWhiteSpace(emu.Types))
        {
          foreach (string type in emu.Types.Split(';'))
          {
            string value = type.Trim();
            if (ValidateEmuFileType(ref value))
            {
              types.Add(value.ToUpperInvariant());
            }
          }
        }
      }

      return types.ToList();
    }

    private bool FileTypeSupported(string fileTypes, string ext)
    {
      if (string.IsNullOrEmpty(fileTypes))
        return false;

      foreach (string type in fileTypes.Split(';'))
      {
        string value = type.Trim();
        if (ValidateEmuFileType(ref value) && string.Equals(value, ext, StringComparison.OrdinalIgnoreCase))
          return true;
      }

      return false;
    }

    private string GetImageNameFromGameFile(string filePath, int imageIndex)
    {
      string ext = Path.GetExtension(filePath).ToUpperInvariant();
      if (ext == ".D64")
      {
        var d64 = new D64Image();
        string error = d64.Load(filePath);
        if (string.IsNullOrEmpty(error) && imageIndex < d64.Entries.Count)
        {
          return d64.Entries[imageIndex].Filename;
        }
      }
      else if (ext == ".T64")
      {
        var t64 = new T64Image();
        string error = t64.Load(filePath);
        if (string.IsNullOrEmpty(error) && imageIndex < t64.Entries.Count)
        {
          return t64.Entries[imageIndex].C64Filename;
        }
      }

      return string.Empty;
    }

    private string ConvertToC64Lower(string text)
    {
      if (string.IsNullOrEmpty(text))
        return string.Empty;

      var sb = new StringBuilder(text.Length);
      foreach (char c in text)
      {
        int code = c;
        if (code > 64)
        {
          if (code < 94)
          {
            sb.Append((char)(code + 32));
          }
          else if (code == 95)
          {
            sb.Append('?');
          }
          else
          {
            sb.Append(c);
          }
        }
        else
        {
          sb.Append(c);
        }
      }
      return sb.ToString();
    }

    private void ReadInEmulators()
    {
      _emulators.Clear();

      int idx = 1;
      var sections = new List<string>();
      while (true)
      {
        string section = _funcs.INI_GetSetting("Emulators", idx.ToString(), "", _emuFile);
        if (string.IsNullOrEmpty(section))
          break;
        sections.Add(section);
        idx++;
      }

      foreach (string section in sections)
      {
        var emu = new EmulatorConfig
        {
          Name = section,
          UseShortFilenames = _funcs.INI_GetSetting(section, "UseShortFilenames", "0", _emuFile) == "1",
          UseDefaultTypes = _funcs.INI_GetSetting(section, "UseDefaultTypes", "0", _emuFile) == "1",
          FullPathAndFilename = _funcs.INI_GetSetting(section, "EmulatorPathFile", "", _emuFile),
          Script = _funcs.INI_GetSetting(section, "ScriptFile", "", _emuFile),
          Types = _funcs.INI_GetSetting(section, "SupportedTypes", "", _emuFile),
          Wait = long.TryParse(_funcs.INI_GetSetting(section, "WaitFinish", "1", _emuFile), out long wait) ? wait : 1
        };

        var files = new List<string>();
        int fileIdx = 1;
        while (true)
        {
          string file = _funcs.INI_GetSetting(section, fileIdx.ToString(), "", _emuFile);
          if (string.IsNullOrEmpty(file))
            break;
          files.Add(file);
          fileIdx++;
        }

        ApplyEmulatorPathInfo(emu, files);
        _emulators.Add(emu);
      }
    }

    private void ApplyEmulatorPathInfo(EmulatorConfig emu, List<string> files)
    {
      emu.Files = new List<string>();
      if (string.IsNullOrEmpty(emu.FullPathAndFilename))
      {
        emu.PathType = PathLocationType.Invalid;
        return;
      }

      _funcs.GetFileAndPath(emu.FullPathAndFilename, out string path, out string file);
      string realPath = path;
      if (!string.IsNullOrEmpty(realPath))
      {
        if (realPath.StartsWith("\\"))
        {
          realPath = _appPathWithBackSlash[..3] + realPath[1..];
        }

        if (realPath.Length >= 2 && realPath[1] != ':')
        {
          realPath = Path.Combine(_appPathWithBackSlash, realPath);
          if (Directory.Exists(realPath))
          {
            realPath = Directory.GetCurrentDirectory();
          }
        }

        if (!Directory.Exists(realPath))
        {
          if (realPath.Length >= 3)
          {
            string validateLocal = _appPathWithBackSlash[..3] + realPath[3..];
            if (Directory.Exists(validateLocal))
            {
              realPath = validateLocal;
            }
          }
        }

        emu.FullPathOnly = realPath.TrimEnd('\\');
        emu.FullFilenameOnly = file;
        emu.FullPathAndFilename = Path.Combine(emu.FullPathOnly, emu.FullFilenameOnly);
      }

      emu.PathType = ProcessHelper.GetDriveLocationType(emu.FullPathAndFilename.Length >= 3 ? emu.FullPathAndFilename[..3] : emu.FullPathAndFilename);

      if (emu.PathType != PathLocationType.Invalid && File.Exists(emu.FullPathAndFilename))
      {
        emu.ShortPathAndFilename = emu.FullPathAndFilename;
        _funcs.GetFileAndPath(emu.FullPathAndFilename, out string fullPathOnly, out string fullFile);
        emu.FullPathOnly = fullPathOnly.TrimEnd('\\');
        emu.FullFilenameOnly = fullFile;
        emu.ShortPathOnly = emu.FullPathOnly;
        emu.ShortFilenameOnly = emu.FullFilenameOnly;
        emu.FullPathOnlyWithBackSlash = emu.FullPathOnly + "\\";
        emu.ShortPathOnlyWithBackSlash = emu.ShortPathOnly + "\\";
      }
      else if (emu.PathType == PathLocationType.FixedHD)
      {
        emu.PathType = PathLocationType.Invalid;
      }

      foreach (string f in files)
      {
        _funcs.GetFileAndPath(f, out string fPath, out string fName);
        if (!string.IsNullOrEmpty(fPath))
        {
          if (fPath.StartsWith("\\"))
          {
            fPath = _appPathWithBackSlash[..3] + fPath[1..];
          }
          if (fPath.Length >= 2 && fPath[1] != ':')
          {
            fPath = Path.Combine(_appPathWithBackSlash, fPath);
            if (Directory.Exists(fPath))
            {
              fPath = Directory.GetCurrentDirectory();
            }
          }
          if (!Directory.Exists(fPath) && fPath.Length >= 3)
          {
            string validateLocal = _appPathWithBackSlash[..3] + fPath[3..];
            if (Directory.Exists(validateLocal))
            {
              fPath = validateLocal;
            }
          }
        }

        emu.Files.Add(Path.Combine(fPath, fName));
      }
    }

    private void SetEmuFileTypes(string fileTypes)
    {
      _emuFileTypes.Clear();
      if (!string.IsNullOrWhiteSpace(fileTypes))
      {
        foreach (string part in fileTypes.Split(';'))
        {
          string type = part.Trim();
          if (ValidateEmuFileType(ref type))
          {
            _emuFileTypes.Add(type.ToUpperInvariant());
          }
        }
      }
    }

    private bool ValidateEmuFileType(ref string fileType)
    {
      if (string.IsNullOrWhiteSpace(fileType))
        return false;

      if (fileType.StartsWith("."))
      {
        if (fileType.Length == 1)
          return false;
        fileType = fileType[1..];
      }
      return true;
    }

    private string ProcessScriptLine(string line)
    {
      string upper = line.ToUpperInvariant();
      if (upper.StartsWith("ADD_CLP2"))
      {
        string err = Validate_ADD_CLP2(ref line, out string clp);
        return err == "" ? Process_ADD_CLP2(clp) : err;
      }
      if (upper.StartsWith("ADD_CLP"))
      {
        string err = Validate_ADD_CLP(ref line, out string clp);
        return err == "" ? Process_ADD_CLP(clp) : err;
      }
      if (upper.StartsWith("EDIT_CLP"))
      {
        string err = Validate_EDIT_CLP(ref line, out string prompt);
        return err == "" ? Process_EDIT_CLP(prompt) : err;
      }
      if (upper.StartsWith("CLEAR_CLP"))
      {
        string err = Validate_CLEAR_CLP(ref line);
        return err == "" ? Process_CLEAR_CLP() : err;
      }
      if (upper.StartsWith("SHOW_KV"))
      {
        string err = Validate_SHOW_KV(ref line);
        return err == "" ? Process_SHOW_KV() : err;
      }
      if (upper.StartsWith("WAIT"))
      {
        string err = Validate_WAIT(ref line, out string ms);
        return err == "" ? Process_WAIT(ms) : err;
      }
      if (upper.StartsWith("KILL_PROCESS"))
      {
        string err = Validate_KILL_PROCESS(ref line, out string filename);
        return err == "" ? Process_KILL_PROCESS(filename) : err;
      }
      if (upper.StartsWith("SET_INI_VALUE"))
      {
        string err = Validate_SET_INI_VALUE(ref line, out int fileNum, out string section, out string key, out string value, out string filename);
        return err == "" ? Process_SET_INI_VALUE(fileNum, section, key, value, filename) : err;
      }
      if (upper.StartsWith("SET_CFG_ITEM"))
      {
        string err = Validate_SET_CFG_ITEM(ref line, out int fileNum, out string key, out string delim, out string value, out string filename);
        return err == "" ? Process_SET_CFG_ITEM(fileNum, key, delim, value, filename) : err;
      }
      if (upper.StartsWith("SET_CFG_VALUE"))
      {
        string err = Validate_SET_CFG_VALUE(ref line, out int fileNum, out string key, out string value, out string filename);
        return err == "" ? Process_SET_CFG_VALUE(fileNum, key, value, filename) : err;
      }
      if (upper.StartsWith("RUN_PROGRAM"))
      {
        string err = Validate_RUN_PROGRAM(ref line, out int fileNum, out string clp, out bool wait, out string filename);
        return err == "" ? Process_RUN_PROGRAM(fileNum, clp, wait, filename) : err;
      }
      if (upper.StartsWith("RUN_EMULATOR_SEND_KEYS"))
      {
        string err = Validate_RUN_EMULATOR_SEND_KEYS(ref line, out string sendKeys, out string keyDownMs, out bool dosClose);
        return err == "" ? Process_RUN_EMULATOR_SEND_KEYS(sendKeys, keyDownMs, dosClose) : err;
      }
      if (upper.StartsWith("RUN_EMULATOR"))
      {
        string err = Validate_RUN_EMULATOR(ref line, out bool dosClose);
        return err == "" ? Process_RUN_EMULATOR(dosClose) : err;
      }
      if (upper.StartsWith("RUN_GAMEFILE"))
      {
        string err = Validate_RUN_GAMEFILE(ref line, out bool dosClose);
        return err == "" ? Process_RUN_GAMEFILE(dosClose) : err;
      }
      if (upper.StartsWith("SHOW_MESSAGE"))
      {
        string err = Validate_SHOW_MESSAGE(ref line, out string message);
        return err == "" ? Process_SHOW_MESSAGE(message) : err;
      }

      return _langUnknownCommand;
    }

    private string Process_ADD_CLP(string clp)
    {
      string value = clp;
      ReplaceEnvironmentVars(ref value, true, false);
      _commandLine += value;
      _logger.PushLogFileTrace("GemusEngine.Process_ADD_CLP");
      _logger.LogStatus($"Adding to command line: [{value}]");
      _logger.LogStatus($"Command line is now: [{_commandLine}]");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_ADD_CLP2(string clp)
    {
      string value = clp;
      ReplaceEnvironmentVars(ref value, false, false);
      _commandLine += value;
      _logger.PushLogFileTrace("GemusEngine.Process_ADD_CLP2");
      _logger.LogStatus($"Adding to command line: [{value}]");
      _logger.LogStatus($"Command line is now: [{_commandLine}]");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_EDIT_CLP(string prompt)
    {
      string message = prompt;
      ReplaceEnvironmentVars(ref message, false, false);
      var config = new InputBoxConfig
      {
        AllowEmptyReturn = true,
        OkClicked = false,
        MaxLength = 8092,
        CancelButtonText = GetIniMessage("General Buttons", "Cancel", "&Cancel"),
        Caption = "GEMUS",
        OkButtonText = GetIniMessage("General Buttons", "OK", "&OK"),
        Prompt = message + ":",
        ReturnedText = "",
        StartingText = _commandLine
      };

      string? result = Ui?.PromptInput(config);
      if (result != null)
      {
        _commandLine = result;
      }

      _logger.PushLogFileTrace("GemusEngine.Process_EDIT_CLP");
      _logger.LogStatus($"Command line (after edit) is now: [{_commandLine}]");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_CLEAR_CLP()
    {
      _commandLine = string.Empty;
      _logger.PushLogFileTrace("GemusEngine.Process_CLEAR_CLP");
      _logger.LogStatus("Command line cleared");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_SHOW_KV()
    {
      var sb = new StringBuilder();
      foreach (var kv in _notesKeys)
      {
        sb.AppendLine(kv.Key + "=" + kv.Value);
      }
      string kvText = sb.ToString().TrimEnd();
      Ui?.ShowKv(kvText);
      return string.Empty;
    }

    private string Process_SHOW_MESSAGE(string message)
    {
      string text = message;
      ReplaceEnvironmentVars(ref text, false, false);
      if (Ui != null)
      {
        Ui.ShowInfo(text, _appName);
      }
      else
      {
        MessageBox.Show(text, _appName, MessageBoxButtons.OK, MessageBoxIcon.Information);
      }
      return string.Empty;
    }

    private string Process_WAIT(string msText)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_WAIT");
      string value = msText;
      ReplaceEnvironmentVars(ref value, false, false);
      if (!int.TryParse(value, out int ms))
        ms = 0;
      _logger.LogStatus($"Waiting for [{ms}] milliseconds...");
      Thread.Sleep(Math.Max(0, ms));
      _logger.LogStatus("Finished waiting!");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_KILL_PROCESS(string filename)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_KILL_PROCESS");
      string name = filename;
      ReplaceEnvironmentVars(ref name, false, false);
      _logger.LogStatus($"Attempting to kill processes matching [{name}]...");
      int killed = ProcessHelper.KillProcessesByFilename(name);
      _logger.LogStatus($"Killed {killed} matching process(es)");
      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_RUN_PROGRAM(int fileNum, string clp, bool wait, string filename)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_RUN_PROGRAM");
      string file = filename;
      bool found = true;

      if (string.IsNullOrEmpty(file))
      {
        int idx = fileNum - 1;
        if (idx < 0 || idx >= _emulators[_emuIdx].Files.Count)
        {
          found = false;
          if (!ConfirmContinue("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function."))
          {
            _logger.PopLogFileTrace();
            return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
          }
        }
        else
        {
          file = _emulators[_emuIdx].Files[idx];
        }
      }
      else
      {
        ReplaceEnvironmentVars(ref file, true, false);
      }

      _logger.LogStatus($"Determined program to run to be [{file}]");

      if (!File.Exists(file))
      {
        found = false;
        if (!ConfirmContinue("Running Games", "FileNotFound", "The following file was not found", file))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
      }

      if (found)
      {
        string args = clp;
        ReplaceEnvironmentVars(ref args, true, false);
        if (_useShort)
        {
          file = ProcessHelper.GetShortPath(file);
        }

        _funcs.GetFileAndPath(file, out string path, out string filenameOnly);
        _logger.LogStatus($"Changed directory to [{path}]");
        _logger.LogStatus($"Running program [{file}] with command line arguments [{args}]");

        var result = ProcessHelper.LaunchProcess(file, args, path, wait);
        if (!result.success)
        {
          if (!ConfirmContinue("Running Games", "ErrorRunningProgram", "There was an error running the following program", file))
          {
            _logger.PopLogFileTrace();
            return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
          }
        }
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_SET_INI_VALUE(int fileNum, string section, string key, string value, string filename)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_SET_INI_VALUE");

      string file = ResolveAssociatedFile(fileNum, filename, "SET_INI_VALUE");
      if (string.IsNullOrEmpty(file))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      string sectionText = section;
      string keyText = key;
      string valueText = value;

      if (sectionText.StartsWith("[") && sectionText.EndsWith("]") && sectionText.Length > 2)
      {
        sectionText = sectionText[1..^1];
      }

      ReplaceEnvironmentVars(ref sectionText, false, false);
      ReplaceEnvironmentVars(ref keyText, false, false);
      ReplaceEnvironmentVars(ref valueText, false, false);

      if (string.IsNullOrWhiteSpace(sectionText) || string.IsNullOrWhiteSpace(keyText))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      _logger.LogStatus($"Setting ini value [{keyText}={valueText}] in section [{sectionText}] in file [{file}]");
      if (!GemusIniHelper.WriteIniValue(sectionText, keyText, valueText, file))
      {
        if (!ConfirmContinue("Running Games", "ErrorSetINIValue", "Could not set the following INI File value",
          file + "\r\n\r\n[" + sectionText + "]\r\n" + keyText + "=" + valueText))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_SET_CFG_VALUE(int fileNum, string key, string value, string filename)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_SET_CFG_VALUE");

      string file = ResolveAssociatedFile(fileNum, filename, "SET_CFG_VALUE");
      if (string.IsNullOrEmpty(file))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      if (ProcessHelper.IsCdRomDrive(file.Length >= 3 ? file[..3] : file))
      {
        if (!ConfirmContinue("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value",
          file + "\r\n\r\n" + key + "=" + value))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      string keyText = key;
      string valueText = value;
      ReplaceEnvironmentVars(ref keyText, false, false);
      ReplaceEnvironmentVars(ref valueText, false, false);

      if (string.IsNullOrWhiteSpace(keyText))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      _logger.LogStatus($"Setting cfg value [{keyText}={valueText}] in file [{file}]");
      if (!WriteCfgValue(file, keyText, valueText))
      {
        if (!ConfirmContinue("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value",
          file + "\r\n\r\n" + keyText + "=" + valueText))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_SET_CFG_ITEM(int fileNum, string key, string delim, string value, string filename)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_SET_CFG_ITEM");

      string file = ResolveAssociatedFile(fileNum, filename, "SET_CFG_ITEM");
      if (string.IsNullOrEmpty(file))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      if (ProcessHelper.IsCdRomDrive(file.Length >= 3 ? file[..3] : file))
      {
        if (!ConfirmContinue("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value",
          file + "\r\n\r\n" + key + "=" + value))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      string keyText = key;
      string delimText = delim;
      string valueText = value;
      ReplaceEnvironmentVars(ref keyText, false, false);
      if (!string.IsNullOrEmpty(delimText)) ReplaceEnvironmentVars(ref delimText, false, false);
      if (!string.IsNullOrEmpty(valueText)) ReplaceEnvironmentVars(ref valueText, false, false);

      if (string.IsNullOrWhiteSpace(keyText))
      {
        _logger.PopLogFileTrace();
        return string.Empty;
      }

      _logger.LogStatus($"Setting cfg item [{keyText}{delimText}{valueText}] in file [{file}]");
      if (!WriteCfgValue2(file, keyText, delimText, valueText))
      {
        if (!ConfirmContinue("Running Games", "ErrorSetCFGValue", "Could not set the following Config File value",
          file + "\r\n\r\n" + keyText + delimText + valueText))
        {
          _logger.PopLogFileTrace();
          return GetIniMessage("Running Games", "ScriptAborted", "Script aborted!");
        }
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_RUN_EMULATOR(bool dosClose)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_RUN_EMULATOR");
      string exeLine = _emulators[_emuIdx].ShortPathAndFilename;

      if (File.Exists(exeLine))
      {
        _funcs.GetFileAndPath(exeLine, out string path, out string file);
        string commandLine = dosClose ? _commandCom + file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine) :
          file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine);

        _logger.LogStatus($"Changed directory to [{path}]");
        _logger.LogStatus($"Running emulator with command line [{commandLine}]");

        var result = ProcessHelper.LaunchProcess(exeLine, commandLine.Replace(file, ""), path, false);
        if (result.success && result.process != null)
        {
          if (_attemptRepack && !_silentMode)
          {
            _logger.LogStatus("Started to wait for emulator to finish (for repacking)");
            if (_emulators[_emuIdx].Wait == 1)
            {
              ProcessHelper.WaitForProcessExit(result.process);
            }
            _logger.LogStatus("Finished waiting for emulator to finish");
          }
          else if (_attemptRepack || _waitIfNotRezipping)
          {
            _logger.LogStatus("Started to wait for emulator to finish");
            ProcessHelper.WaitForProcessExit(result.process);
            _logger.LogStatus("Finished waiting for emulator to finish");
          }
        }
      }
      else
      {
        _logger.PopLogFileTrace();
        return GetIniMessage("Running Games", "FileNotFound", "The following file was not found") + ":\r\n\r\n" + exeLine;
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_RUN_EMULATOR_SEND_KEYS(string sendKeys, string keyDownDurationText, bool dosClose)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_RUN_EMULATOR_SEND_KEYS");
      string exeLine = _emulators[_emuIdx].ShortPathAndFilename;

      if (!File.Exists(exeLine))
      {
        _logger.PopLogFileTrace();
        return GetIniMessage("Running Games", "FileNotFound", "The following file was not found") + ":\r\n\r\n" + exeLine;
      }

      _funcs.GetFileAndPath(exeLine, out string path, out string file);
      string commandLine = dosClose ? _commandCom + file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine) :
        file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine);

      _logger.LogStatus($"Changed directory to [{path}]");
      _logger.LogStatus($"Running emulator with command line [{commandLine}]");

      var result = ProcessHelper.LaunchProcess(exeLine, commandLine.Replace(file, ""), path, false);
      if (!result.success || result.process == null)
      {
        _logger.PopLogFileTrace();
        return GetIniMessage("Running Games", "ErrorRunningProgram", "There was an error running the following program") + ":\r\n\r\n" + exeLine;
      }

      _logger.LogStatus("Started to send keystrokes");
      int keyDownDuration = 0;
      ReplaceEnvironmentVars(ref keyDownDurationText, false, false);
      if (!int.TryParse(keyDownDurationText, out keyDownDuration)) keyDownDuration = 0;
      var blocks = ParseSendKeys(sendKeys);

      foreach (var block in blocks)
      {
        if (block.SecondsToWait == 0)
        {
          if (result.process.HasExited)
          {
            _logger.LogStatus("Aborted sending any remaining keystrokes as emulator was closed");
            break;
          }
          try
          {
            SendKeys.SendWait(block.KeysToSend);
          }
          catch
          {
            _logger.PopLogFileTrace();
            return "Failed Sending Keys \"" + block.KeysToSend + "\" to " + exeLine;
          }
        }
        else if (block.SecondsToWait < 0)
        {
          keyDownDuration = Math.Abs(block.SecondsToWait);
        }
        else
        {
          ProcessHelper.WaitWhileProcessRunning(block.SecondsToWait * 1000, result.process);
        }
      }

      _logger.LogStatus("Finished sending keystrokes");
      if (_attemptRepack || _waitIfNotRezipping)
      {
        _logger.LogStatus("Started to wait for emulator to finish");
        ProcessHelper.WaitForProcessExit(result.process);
        _logger.LogStatus("Finished waiting for emulator to finish");
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string Process_RUN_GAMEFILE(bool dosClose)
    {
      _logger.PushLogFileTrace("GemusEngine.Process_RUN_GAMEFILE");

      string exeLine = _envGamePathFile;
      if (File.Exists(exeLine))
      {
        _funcs.GetFileAndPath(exeLine, out string path, out string file);
        string ext = Path.GetExtension(file).ToUpperInvariant();

        if (ext is ".EXE" or ".COM" or ".BAT")
        {
          string commandLine = dosClose ? _commandCom + file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine) :
            file + (string.IsNullOrEmpty(_commandLine) ? "" : " " + _commandLine);

          _logger.LogStatus($"Running gamefile with command line [{commandLine}]");
          ProcessHelper.LaunchProcess(exeLine, commandLine.Replace(file, ""), path, false);
        }
        else
        {
          _logger.LogStatus($"Running gamefile [{exeLine}]");
          var result = ProcessHelper.ShellExecuteFile(exeLine, "", false);
          if (result.process != null)
          {
            _logger.LogStatus("Started to wait for gamefile to finish");
            ProcessHelper.WaitForProcessExit(result.process);
            _logger.LogStatus("Finished waiting for gamefile to finish");
          }
        }
      }
      else
      {
        _logger.PopLogFileTrace();
        return GetIniMessage("Running Games", "FileNotFound", "The following file was not found") + ":\r\n\r\n" + exeLine;
      }

      _logger.PopLogFileTrace();
      return string.Empty;
    }

    private string ResolveAssociatedFile(int fileNum, string filename, string context)
    {
      if (!string.IsNullOrEmpty(filename))
      {
        string path = filename;
        ReplaceEnvironmentVars(ref path, false, false);
        return path;
      }

      int idx = fileNum - 1;
      if (idx < 0 || idx >= _emulators[_emuIdx].Files.Count)
      {
        if (!_scriptFileErrorDone)
        {
          _scriptFileErrorDone = true;
          _logger.LogStatus($"Associated script file [{fileNum}] does not point to valid file");
          if (!ConfirmContinue("Running Games", "ErrorInvalidFile", "Invalid file referenced in script function."))
          {
            return string.Empty;
          }
          _logger.LogStatus("Unable to set cfg/ini value but user chose to continue anyway.");
        }
        return string.Empty;
      }

      string file = _emulators[_emuIdx].Files[idx];
      _logger.LogStatus($"Associated script file [{fileNum}] determined to be [{file}]");
      return file;
    }

    private bool ConfirmContinue(string section, string key, string defaultMessage, string? detail = null)
    {
      string msg = GetIniMessage(section, key, defaultMessage);
      if (!string.IsNullOrEmpty(detail))
      {
        msg += ":\r\n\r\n" + detail;
      }
      msg += "\r\n\r\n" + GetIniMessage("Running Games", "Continue", "Do you want to continue running the game anyway?");

      if (Ui != null)
      {
        return Ui.Confirm(msg, _appName + " - " + _langScriptError);
      }

      return MessageBox.Show(msg, _appName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
    }

    private List<SendKeysBlock> ParseSendKeys(string sendKeys)
    {
      var blocks = new List<SendKeysBlock>();
      if (string.IsNullOrEmpty(sendKeys))
        return blocks;

      string keys = sendKeys;
      ReplaceEnvironmentVars(ref keys, false, true);
      keys = keys.Replace("{DOUBLEPIPE}", "||");

      int start = 0;
      while (true)
      {
        int openPos = keys.IndexOf('[', start);
        if (openPos >= 0)
        {
          if (openPos > start)
          {
            blocks.Add(new SendKeysBlock { SecondsToWait = 0, KeysToSend = keys[start..openPos] });
          }

          int closePos = keys.IndexOf(']', openPos + 1);
          if (closePos >= 0)
          {
            if (closePos > openPos + 1)
            {
              if (int.TryParse(keys[(openPos + 1)..closePos], out int secs) && secs != 0)
              {
                blocks.Add(new SendKeysBlock { SecondsToWait = secs, KeysToSend = string.Empty });
              }
            }
            start = closePos + 1;
          }
          else
          {
            break;
          }
        }
        else
        {
          if (start < keys.Length)
          {
            blocks.Add(new SendKeysBlock { SecondsToWait = 0, KeysToSend = keys[start..] });
          }
          break;
        }
      }

      foreach (var block in blocks)
      {
        if (block.SecondsToWait == 0)
        {
          block.KeysToSend = block.KeysToSend.Replace("{OPENSQB}", "[").Replace("{CLOSESQB}", "]");
        }
      }

      return blocks;
    }

    private string VbSendKeysEscape(string text, bool doEscape)
    {
      if (!doEscape)
        return text;

      string result = text.Replace("{", "{{}");
      result = result.Replace("{{}", "~~~~~TEMP_REPLACEMENT~~~~~");
      result = result.Replace("}", "{}}");
      result = result.Replace("~~~~~TEMP_REPLACEMENT~~~~~", "{{}");
      result = result.Replace("+", "{+}");
      result = result.Replace("^", "{^}");
      result = result.Replace("%", "{%}");
      result = result.Replace("~", "{~}");
      return result;
    }

    private void ReplaceEnvironmentVars(ref string text, bool addQuotes, bool doSendKeys)
    {
      text = text.Replace("%dbdriveletter%", _appPathWithBackSlash.Substring(0, 1));

      if (!_useShort)
      {
        if (_sidPath.PathType != PathLocationType.Invalid)
          text = text.Replace("%storedmusic%", VbSendKeysEscape(_sidPath.RealPath, doSendKeys));
        if (_photosPath.PathType != PathLocationType.Invalid)
          text = text.Replace("%storedphotos%", VbSendKeysEscape(_photosPath.RealPath, doSendKeys));

        for (int i = 0; i < _gamePaths.Count; i++)
        {
          if (_gamePaths[i].PathType != PathLocationType.Invalid)
            text = text.Replace($"%storedgames({i})%", VbSendKeysEscape(_gamePaths[i].RealPath, doSendKeys));
        }
        for (int i = 0; i < _screenshotPaths.Count; i++)
        {
          if (_screenshotPaths[i].PathType != PathLocationType.Invalid)
            text = text.Replace($"%storedpictures({i})%", VbSendKeysEscape(_screenshotPaths[i].RealPath, doSendKeys));
        }
        for (int i = 0; i < _extrasPaths.Count; i++)
        {
          if (_extrasPaths[i].PathType != PathLocationType.Invalid)
            text = text.Replace($"%storedextras({i})%", VbSendKeysEscape(_extrasPaths[i].RealPath, doSendKeys));
        }

        text = text.Replace("%gbgamepath%", VbSendKeysEscape(_gbGamePath, doSendKeys));
        text = text.Replace("%dbpath%", VbSendKeysEscape(Path.Combine(_appPathWithBackSlash, _gbName), doSendKeys));
        text = text.Replace("%emupath%", VbSendKeysEscape(_emulators[_emuIdx].FullPathOnly, doSendKeys));

        ReplaceGamePathTokens(ref text, addQuotes, doSendKeys, _envGamePathFile, _envGameFile, _envGamePath, _envGameFileNoExt);

        if (_fldNumGameFiles > 1)
        {
          for (int i = 1; i < _fldNumGameFiles; i++)
          {
            ReplaceGamePathTokens(ref text, addQuotes, doSendKeys, _extraGamePathFile[i], _extraGameFile[i], _extraGamePath[i], _extraGameFileNoExt[i], i);
          }
        }
      }
      else
      {
        string shortGamePathFile = _envGamePathFile;
        _funcs.GetFileAndPath(shortGamePathFile, out string shortPath, out string shortFile);
        string shortFileNoExt = Path.GetFileNameWithoutExtension(shortFile);

        text = text.Replace("%storedmusic%", VbSendKeysEscape(_sidPath.RealPath, doSendKeys));
        text = text.Replace("%storedphotos%", VbSendKeysEscape(_photosPath.RealPath, doSendKeys));
        text = text.Replace("%gbgamepath%", VbSendKeysEscape(_gbGamePath, doSendKeys));
        text = text.Replace("%dbpath%", VbSendKeysEscape(Path.Combine(_appPathWithBackSlash, _gbName), doSendKeys));
        text = text.Replace("%emupath%", VbSendKeysEscape(_emulators[_emuIdx].FullPathOnly, doSendKeys));

        text = text.Replace("%gamepathfile%", VbSendKeysEscape(shortGamePathFile, doSendKeys));
        text = text.Replace("%gamefile%", VbSendKeysEscape(shortFile, doSendKeys));
        text = text.Replace("%gamepath%", VbSendKeysEscape(shortPath, doSendKeys));
        text = text.Replace("%gamefilenoext%", VbSendKeysEscape(shortFileNoExt, doSendKeys));
      }

      text = text.Replace("%imageindex%", _imageIndex.ToString());
      text = text.Replace("%imagename%", VbSendKeysEscape(_imageName, doSendKeys));
      text = text.Replace("%c64imagename%", VbSendKeysEscape(_c64ImageName, doSendKeys));
      text = text.Replace("%commandline%", VbSendKeysEscape(_commandLine, doSendKeys));
      text = text.Replace("%numfiles%", _fldNumGameFiles.ToString());

      text = text.Replace("%crlf%", "\r\n");
      text = text.Replace("%crlfx2%", "\r\n\r\n");
      text = text.Replace("%tab%", "\t");

      foreach (var kv in _notesKeys)
      {
        string find = "%" + kv.Key.ToLowerInvariant() + "_value%";
        text = text.Replace(find, VbSendKeysEscape(kv.Value, false));
      }
    }

    private void ReplaceGamePathTokens(ref string text, bool addQuotes, bool doSendKeys, string fullPathFile,
      string file, string path, string fileNoExt, int? index = null)
    {
      string suffix = index.HasValue ? "(" + index.Value + ")" : string.Empty;
      ReplaceToken(ref text, "%gamepathfile" + suffix + "%", fullPathFile, addQuotes, doSendKeys);
      ReplaceToken(ref text, "%gamefile" + suffix + "%", file, addQuotes, doSendKeys);
      ReplaceToken(ref text, "%gamepath" + suffix + "%", path, addQuotes, doSendKeys);
      ReplaceToken(ref text, "%gamefilenoext" + suffix + "%", fileNoExt, addQuotes, doSendKeys);
    }

    private void ReplaceToken(ref string text, string token, string value, bool addQuotes, bool doSendKeys)
    {
      if (text.Contains(token, StringComparison.OrdinalIgnoreCase))
      {
        string replacement = VbSendKeysEscape(value, doSendKeys);
        if (addQuotes && replacement.Contains(' '))
        {
          replacement = "\"" + replacement + "\"";
        }
        text = text.Replace(token, replacement);
      }
    }

    private string ValidateAndGetScriptLines(string script)
    {
      _scriptLines = script.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).ToList();
      _scriptLinesPre = new List<string>(new string[_scriptLines.Count]);

      for (int i = 0; i < _scriptLines.Count; i++)
      {
        string line = _scriptLines[i];
        string error = ValidateScriptLine(ref line);
        if (!string.IsNullOrEmpty(error))
        {
          return error + "\r\n\r\n" + _langLine + " " + (i + 1);
        }
        _scriptLines[i] = line;
      }

      return string.Empty;
    }

    private string ValidateScriptLine(ref string line)
    {
      if (string.IsNullOrEmpty(line))
        return string.Empty;

      GetScriptLineParts(line, out string pre, out string actual);
      if (string.IsNullOrEmpty(actual))
      {
        line = actual;
        return string.Empty;
      }

      if (actual.StartsWith(";"))
      {
        line = actual;
        return string.Empty;
      }

      if (actual.StartsWith("IF ", StringComparison.OrdinalIgnoreCase))
      {
        string err = ValidateComparisonLine(3, ref actual, out _, out _, out _);
        if (string.IsNullOrEmpty(err))
        {
          line = actual;
          return string.Empty;
        }
        return err;
      }
      if (actual.StartsWith("ELSE IF ", StringComparison.OrdinalIgnoreCase))
      {
        string err = ValidateComparisonLine(8, ref actual, out _, out _, out _);
        if (string.IsNullOrEmpty(err))
        {
          line = actual;
          return string.Empty;
        }
        return err;
      }
      if (actual.StartsWith("ELSEIF ", StringComparison.OrdinalIgnoreCase))
      {
        string err = ValidateComparisonLine(7, ref actual, out _, out _, out _);
        if (string.IsNullOrEmpty(err))
        {
          line = actual;
          return string.Empty;
        }
        return err;
      }
      if (actual.StartsWith("ELSE", StringComparison.OrdinalIgnoreCase))
      {
        if (actual.Length > 4)
          return _langEndOfLineExpected;
        line = "Else";
        return string.Empty;
      }
      if (actual.StartsWith("ENDIF", StringComparison.OrdinalIgnoreCase) || actual.StartsWith("END IF", StringComparison.OrdinalIgnoreCase))
      {
        if (actual.Length > 5 && !actual.StartsWith("END IF", StringComparison.OrdinalIgnoreCase))
          return _langEndOfLineExpected;
        line = "End If";
        return string.Empty;
      }

      string functionLine = actual;
      string funcError = ValidateFunctionLine(ref functionLine);
      if (!string.IsNullOrEmpty(funcError))
        return funcError;

      line = functionLine;
      return string.Empty;
    }

    private void GetScriptLineParts(string line, out string pre, out string actual)
    {
      pre = string.Empty;
      actual = string.Empty;
      for (int i = 0; i < line.Length; i++)
      {
        char c = line[i];
        if (c == '\t' || c == ' ')
        {
          pre += c;
        }
        else
        {
          actual = line[i..];
          break;
        }
      }

      if (!string.IsNullOrEmpty(actual))
      {
        actual = actual.TrimEnd(' ', '\t').Replace("\t", " ");
      }
    }

    private string ValidateComparisonLine(int statementLength, ref string line, out string field, out string op, out string data)
    {
      field = string.Empty;
      op = string.Empty;
      data = string.Empty;

      string exp = line[statementLength..].Trim();
      if (string.IsNullOrEmpty(exp))
        return _langExpressionReqd;

      if (line.Length <= statementLength || line[statementLength - 1] != ' ')
        return _langExpected + ": \" \"";

      if (exp.StartsWith("QUESTION", StringComparison.OrdinalIgnoreCase))
      {
        field = "QUESTION";
        exp = exp[8..];
        if (exp.TrimEnd().EndsWith("NO", StringComparison.OrdinalIgnoreCase))
        {
          data = "NO";
          exp = exp[..^2].Trim();
        }
        else if (exp.TrimEnd().EndsWith("YES", StringComparison.OrdinalIgnoreCase))
        {
          data = "YES";
          exp = exp[..^3].Trim();
        }
        else
        {
          return _langExpected + ": \"= YES|NO\"";
        }

        if (!exp.EndsWith("=", StringComparison.Ordinal))
          return _langExpected + ": \"= YES|NO\"";

        exp = exp[..^1].Trim();
        if (!exp.StartsWith("(") || !exp.EndsWith(")"))
          return _langExpected + ": \"(\"";

        if (exp.Length < 3)
          return _langInvalidValue + "\r\n\r\n" + _langExpected + ": \"QUESTION( text ) = YES|NO\"";

        op = exp[1..^1];
        line = NormalizeComparisonPrefix(line, statementLength) + "QUESTION(" + op + ") = " + data;
        return string.Empty;
      }

      for (int i = 0; i < _scriptFields.Count; i++)
      {
        var fieldDef = _scriptFields[i];
        string upper = exp[..Math.Min(exp.Length, fieldDef.Name.Length)].ToUpperInvariant();
        if (upper == fieldDef.Name.ToUpperInvariant())
        {
          field = fieldDef.Name;
          if (fieldDef.Name == "Key_")
          {
            int spacePos = exp.IndexOf(' ', 4);
            if (spacePos > 0)
            {
              field = "Key_" + exp[4..spacePos];
              exp = exp[spacePos..];
            }
          }
          else if (fieldDef.Name == "%")
          {
            int spacePos = exp.IndexOf(' ');
            if (spacePos > 0)
            {
              field = exp[..spacePos];
              exp = exp[spacePos..];
            }
          }
          else
          {
            exp = exp[fieldDef.Name.Length..];
          }

          if (fieldDef.ValueType == ScriptValueType.Text)
          {
            string opText = exp.Length >= 9 ? exp[..9].ToUpperInvariant() : exp.ToUpperInvariant();
            if (opText is not " CONTAINS" and not " EXCLUDES")
              return _langInvalidOp + "\r\n\r\n" + _langExpected + ": \"CONTAINS()\" / \"EXCLUDES()\"";

            op = opText.Trim();
            exp = exp[9..].Trim();

            if (!exp.StartsWith("(") || !exp.EndsWith(")") || exp.Length < 3)
              return _langInvalidValue + "\r\n\r\n" + _langExpected + ": \"CONTAINS( text [ || text || text... ] )\"";

            data = exp[1..^1];
            if (data.Split("||").Any(arg => string.IsNullOrEmpty(arg)))
              return _langInvalidValue + "\r\n\r\n" + _langExpected + ": \"CONTAINS( text [ || text || text... ] )\"";

            line = NormalizeComparisonPrefix(line, statementLength) + field + " " + op + "(" + data + ")";
            return string.Empty;
          }

          exp = exp.TrimStart();
          bool foundOp = false;
          foreach (string allowed in fieldDef.Operators)
          {
            if (exp.StartsWith(allowed))
            {
              op = allowed;
              exp = exp[allowed.Length..].Trim();
              foundOp = true;
              break;
            }
          }

          if (!foundOp)
          {
            string expected = string.Join(", ", fieldDef.Operators.Select(o => "\"" + o + "\"")) + ".";
            return _langInvalidOp + "\r\n\r\n" + _langExpected + ": " + expected;
          }

          if (fieldDef.ValueType == ScriptValueType.Number)
          {
            if (exp.Contains(' '))
              return _langEndOfLineExpected;

            int num = ConvertStringToNumber(exp);
            if (num < 0)
              return _langInvalidValue + "\r\n\r\n" + _langExpected + ": " + _langNumber + " > -1, " + _langNumber + " < 10000.";

            data = num.ToString();
            line = NormalizeComparisonPrefix(line, statementLength) + field + " " + op + " " + data;
            return string.Empty;
          }

          if (exp.Contains(' '))
            return _langEndOfLineExpected;

          string valueUpper = exp.ToUpperInvariant();
          if (!fieldDef.ListValues.Any(v => v.Equals(valueUpper, StringComparison.OrdinalIgnoreCase)))
          {
            string expected = string.Join(", ", fieldDef.ListValues.Select(v => "\"" + v + "\"")) + ".";
            return _langInvalidValue + "\r\n\r\n" + _langExpected + ": " + expected;
          }

          data = fieldDef.ListValues.First(v => v.Equals(valueUpper, StringComparison.OrdinalIgnoreCase));
          line = NormalizeComparisonPrefix(line, statementLength) + field + " " + op + " " + data;
          return string.Empty;
        }
      }

      string expectedFields = string.Join(", ", _scriptFields.Select(f => f.Name == "%" ? "\"%...%\"" : "\"" + f.Name + "\"")) + ".";
      return _langUnknownField + "\r\n\r\n" + _langExpected + ": \"QUESTION( )\", " + expectedFields;
    }

    private string NormalizeComparisonPrefix(string line, int statementLength)
    {
      if (line.StartsWith("IF ", StringComparison.OrdinalIgnoreCase)) return "If ";
      if (line.StartsWith("ELSE IF ", StringComparison.OrdinalIgnoreCase)) return "ElseIf ";
      if (line.StartsWith("ELSEIF ", StringComparison.OrdinalIgnoreCase)) return "ElseIf ";
      return line[..statementLength];
    }

    private string ValidateFunctionLine(ref string line)
    {
      string upper = line.ToUpperInvariant();
      if (upper.StartsWith("ADD_CLP2")) return Validate_ADD_CLP2(ref line, out _);
      if (upper.StartsWith("ADD_CLP")) return Validate_ADD_CLP(ref line, out _);
      if (upper.StartsWith("SET_INI_VALUE")) return Validate_SET_INI_VALUE(ref line, out _, out _, out _, out _, out _);
      if (upper.StartsWith("SET_CFG_ITEM")) return Validate_SET_CFG_ITEM(ref line, out _, out _, out _, out _, out _);
      if (upper.StartsWith("SET_CFG_VALUE")) return Validate_SET_CFG_VALUE(ref line, out _, out _, out _, out _);
      if (upper.StartsWith("RUN_PROGRAM")) return Validate_RUN_PROGRAM(ref line, out _, out _, out _, out _);
      if (upper.StartsWith("RUN_EMULATOR_SEND_KEYS")) return Validate_RUN_EMULATOR_SEND_KEYS(ref line, out _, out _, out _);
      if (upper.StartsWith("RUN_EMULATOR")) return Validate_RUN_EMULATOR(ref line, out _);
      if (upper.StartsWith("SHOW_MESSAGE")) return Validate_SHOW_MESSAGE(ref line, out _);
      if (upper.StartsWith("EDIT_CLP")) return Validate_EDIT_CLP(ref line, out _);
      if (upper.StartsWith("RUN_GAMEFILE")) return Validate_RUN_GAMEFILE(ref line, out _);
      if (upper.StartsWith("CLEAR_CLP")) return Validate_CLEAR_CLP(ref line);
      if (upper.StartsWith("SHOW_KV")) return Validate_SHOW_KV(ref line);
      if (upper.StartsWith("WAIT")) return Validate_WAIT(ref line, out _);
      if (upper.StartsWith("KILL_PROCESS")) return Validate_KILL_PROCESS(ref line, out _);
      return _langUnknownCommand;
    }

    private int ConvertStringToNumber(string number)
    {
      if (number.Length > 4)
        return -1;

      foreach (char c in number)
      {
        if (!NumbersAllowed.Contains(c))
          return -1;
      }

      return int.TryParse(number, out int value) ? value : -1;
    }

    private string Validate_RUN_EMULATOR(ref string line, out bool dosClose)
    {
      dosClose = false;
      string error = GetScriptFunctionArgs(12, line, out List<string> args, 0);
      if (string.IsNullOrEmpty(error))
      {
        dosClose = false;
        line = "Run_Emulator()";
        return string.Empty;
      }

      error = GetScriptFunctionArgs(12, line, out List<string> args2, 1);
      if (!string.IsNullOrEmpty(error))
      {
        dosClose = false;
        return error;
      }

      if (args2[0].Equals("DOSCLOSE", StringComparison.OrdinalIgnoreCase))
      {
        dosClose = true;
        line = "Run_Emulator(DOSCLOSE)";
        return string.Empty;
      }

      dosClose = false;
      return _langInvalidFunctionParam + " 1.  " + _langExpected + ": DOSCLOSE";
    }

    private string Validate_RUN_EMULATOR_SEND_KEYS(ref string line, out string sendKeys, out string keyDownDuration, out bool dosClose)
    {
      sendKeys = string.Empty;
      keyDownDuration = string.Empty;
      dosClose = false;
      string error = GetScriptFunctionArgs(22, line, out List<string> args, 2);
      if (!string.IsNullOrEmpty(error))
      {
        sendKeys = string.Empty;
        keyDownDuration = string.Empty;
        dosClose = false;
        return error;
      }

      sendKeys = args[0];
      keyDownDuration = args[1];
      dosClose = false;

      line = "Run_Emulator_Send_Keys(" + sendKeys + "||" + keyDownDuration + ")";
      return string.Empty;
    }

    private string Validate_RUN_GAMEFILE(ref string line, out bool dosClose)
    {
      dosClose = false;
      string error = GetScriptFunctionArgs(12, line, out List<string> args, 0);
      if (string.IsNullOrEmpty(error))
      {
        dosClose = false;
        line = "Run_GameFile()";
        return string.Empty;
      }

      error = GetScriptFunctionArgs(12, line, out List<string> args2, 1);
      if (!string.IsNullOrEmpty(error))
      {
        dosClose = false;
        return error;
      }

      if (args2[0].Equals("DOSCLOSE", StringComparison.OrdinalIgnoreCase))
      {
        dosClose = true;
        line = "Run_GameFile(DOSCLOSE)";
        return string.Empty;
      }

      dosClose = false;
      return _langInvalidFunctionParam + " 1.  " + _langExpected + ": DOSCLOSE";
    }

    private string Validate_RUN_PROGRAM(ref string line, out int fileNum, out string clp, out bool wait, out string filename)
    {
      fileNum = 0;
      clp = string.Empty;
      wait = false;
      filename = string.Empty;
      string error = GetScriptFunctionArgs(11, line, out List<string> args, 3);
      if (!string.IsNullOrEmpty(error))
      {
        fileNum = 0;
        clp = string.Empty;
        wait = false;
        filename = string.Empty;
        return error;
      }

      filename = string.Empty;
      fileNum = ConvertStringToNumber(args[0]);
      if (fileNum < 1)
      {
        if (string.IsNullOrEmpty(args[0]))
          return _langInvalidFunctionParam + " 1.  " + _langExpected + ": " + _langNumber;
        filename = args[0];
      }

      clp = args[1];
      wait = args[2].Equals("WAIT", StringComparison.OrdinalIgnoreCase);
      line = "Run_Program(" + args[0] + "||" + args[1] + "||" + args[2] + ")";
      return string.Empty;
    }

    private string Validate_SET_INI_VALUE(ref string line, out int fileNum, out string section, out string key, out string value, out string filename)
    {
      fileNum = 0;
      section = string.Empty;
      key = string.Empty;
      value = string.Empty;
      filename = string.Empty;
      string error = GetScriptFunctionArgs(13, line, out List<string> args, 4);
      if (!string.IsNullOrEmpty(error))
      {
        fileNum = 0; section = key = value = filename = string.Empty;
        return error;
      }

      filename = string.Empty;
      fileNum = ConvertStringToNumber(args[0]);
      if (fileNum < 1)
      {
        if (string.IsNullOrEmpty(args[0]))
          return _langInvalidFunctionParam + " 1.  " + _langExpected + ": " + _langNumber;
        filename = args[0];
      }

      if (string.IsNullOrEmpty(args[1]))
        return _langInvalidFunctionParam + " 2.  " + _langExpected + ": SECTION";
      if (string.IsNullOrEmpty(args[2]))
        return _langInvalidFunctionParam + " 3.  " + _langExpected + ": KEY";

      section = args[1];
      key = args[2];
      value = args[3];
      line = "Set_INI_Value(" + args[0] + "||" + args[1] + "||" + args[2] + "||" + args[3] + ")";
      return string.Empty;
    }

    private string Validate_SET_CFG_ITEM(ref string line, out int fileNum, out string key, out string delim, out string value, out string filename)
    {
      fileNum = 0;
      key = string.Empty;
      delim = string.Empty;
      value = string.Empty;
      filename = string.Empty;
      string error = GetScriptFunctionArgs(12, line, out List<string> args, 4);
      if (!string.IsNullOrEmpty(error))
      {
        fileNum = 0; key = delim = value = filename = string.Empty;
        return error;
      }

      filename = string.Empty;
      fileNum = ConvertStringToNumber(args[0]);
      if (fileNum < 1)
      {
        if (string.IsNullOrEmpty(args[0]))
          return _langInvalidFunctionParam + " 1.  " + _langExpected + ": " + _langNumber;
        filename = args[0];
      }

      if (string.IsNullOrEmpty(args[1]))
        return _langInvalidFunctionParam + " 2.  " + _langExpected + ": KEY";

      key = args[1];
      delim = args[2];
      value = args[3];
      line = "Set_CFG_Item(" + args[0] + "||" + args[1] + "||" + args[2] + "||" + args[3] + ")";
      return string.Empty;
    }

    private string Validate_SET_CFG_VALUE(ref string line, out int fileNum, out string key, out string value, out string filename)
    {
      fileNum = 0;
      key = string.Empty;
      value = string.Empty;
      filename = string.Empty;
      string error = GetScriptFunctionArgs(13, line, out List<string> args, 3);
      if (!string.IsNullOrEmpty(error))
      {
        fileNum = 0; key = value = filename = string.Empty;
        return error;
      }

      filename = string.Empty;
      fileNum = ConvertStringToNumber(args[0]);
      if (fileNum < 1)
      {
        if (string.IsNullOrEmpty(args[0]))
          return _langInvalidFunctionParam + " 1.  " + _langExpected + ": " + _langNumber;
        filename = args[0];
      }

      if (string.IsNullOrEmpty(args[1]))
        return _langInvalidFunctionParam + " 2.  " + _langExpected + ": KEY";

      key = args[1];
      value = args[2];
      line = "Set_CFG_Value(" + args[0] + "||" + args[1] + "||" + args[2] + ")";
      return string.Empty;
    }

    private string Validate_SHOW_MESSAGE(ref string line, out string message)
    {
      message = string.Empty;
      string error = GetScriptFunctionArgs(12, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        message = string.Empty;
        return error;
      }

      message = args[0];
      line = "Show_Message(" + args[0] + ")";
      return string.Empty;
    }

    private string Validate_EDIT_CLP(ref string line, out string message)
    {
      message = string.Empty;
      string error = GetScriptFunctionArgs(8, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        message = string.Empty;
        return error;
      }

      message = args[0];
      line = "Edit_CLP(" + args[0] + ")";
      return string.Empty;
    }

    private string Validate_SHOW_KV(ref string line)
    {
      string error = GetScriptFunctionArgs(7, line, out _, 0);
      if (!string.IsNullOrEmpty(error))
        return error;

      line = "Show_KV()";
      return string.Empty;
    }

    private string Validate_WAIT(ref string line, out string ms)
    {
      ms = string.Empty;
      string error = GetScriptFunctionArgs(4, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        ms = string.Empty;
        return error;
      }

      ms = args[0];
      line = "Wait(" + args[0] + ")";
      return string.Empty;
    }

    private string Validate_KILL_PROCESS(ref string line, out string filename)
    {
      filename = string.Empty;
      string error = GetScriptFunctionArgs(12, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        filename = string.Empty;
        return error;
      }

      filename = args[0];
      line = "Kill_Process(" + args[0] + ")";
      return string.Empty;
    }

    private string Validate_ADD_CLP(ref string line, out string clp)
    {
      clp = string.Empty;
      string error = GetScriptFunctionArgs(7, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        clp = string.Empty;
        return error;
      }

      clp = args[0];
      line = "Add_CLP(" + clp + ")";
      return string.Empty;
    }

    private string Validate_ADD_CLP2(ref string line, out string clp)
    {
      clp = string.Empty;
      string error = GetScriptFunctionArgs(8, line, out List<string> args, 1);
      if (!string.IsNullOrEmpty(error))
      {
        clp = string.Empty;
        return error;
      }

      clp = args[0];
      line = "Add_CLP2(" + clp + ")";
      return string.Empty;
    }

    private string Validate_CLEAR_CLP(ref string line)
    {
      string error = GetScriptFunctionArgs(9, line, out _, 0);
      if (!string.IsNullOrEmpty(error))
        return error;

      line = "Clear_CLP()";
      return string.Empty;
    }

    private string GetScriptFunctionArgs(int funcNameLength, string line, out List<string> args, int requiredArgs)
    {
      args = new List<string>();
      string argText = line[funcNameLength..].Trim();
      if (!argText.StartsWith("("))
        return _langExpected + ": \"(\"";
      if (!argText.EndsWith(")"))
        return _langExpected + ": \")\"";

      if (argText.Length == 2)
      {
        if (requiredArgs > 0)
          return _langMissingParams + "\r\n" + _langExpected + ": " + requiredArgs;
        return string.Empty;
      }

      string inner = argText[1..^1];
      args.AddRange(inner.Split("||"));

      if (args.Count > requiredArgs)
        return _langTooManyParams + "\r\n" + _langExpected + ": " + requiredArgs;
      if (args.Count < requiredArgs)
        return _langMissingParams + "\r\n" + _langExpected + ": " + requiredArgs;

      return string.Empty;
    }

    private string ValidateScriptStructure(string script, out int errorLine)
    {
      errorLine = -1;
      if (string.IsNullOrEmpty(script))
        return string.Empty;

      int numIfs = GetNumSpecificScriptLines("if");
      int numEndifs = GetNumSpecificScriptLines("end if") + GetNumSpecificScriptLines("endif");
      if (numIfs < numEndifs)
        return _langTooMany + " 'END IF' " + _langStatements + "!";
      if (numIfs > numEndifs)
        return _langTooMany + " 'IF' " + _langStatements + "!";

      int numElseIfs = GetNumSpecificScriptLines("else if") + GetNumSpecificScriptLines("elseif");
      if (numElseIfs > numIfs && numIfs == 0)
        return _langTooMany + " 'ELSE IF' " + _langStatements + "!";

      int numElses = GetNumSpecificScriptLines("else");
      if (numElses > numIfs)
        return _langTooMany + " 'ELSE' " + _langStatements + "!";

      string err = ValidateScriptBlock(0, out int nextLine, out int errLine, 0);
      if (!string.IsNullOrEmpty(err))
      {
        errorLine = errLine + 1;
      }
      return err;
    }

    private int GetNumSpecificScriptLines(string text)
    {
      string find = text.ToUpperInvariant();
      int count = 0;
      for (int i = 0; i < _scriptLines.Count; i++)
      {
        string line = _scriptLines[i].ToUpperInvariant();
        if (find == "ELSE")
        {
          if (line == "ELSE") count++;
        }
        else
        {
          if (line.StartsWith(find)) count++;
        }
      }
      return count;
    }

    private string GetScriptLine(int index)
    {
      if (index >= _scriptLines.Count) return "__GEMUS_EOF__";
      return _scriptLines[index];
    }

    private string ValidateScriptBlock(int startLine, out int nextLine, out int errLine, int level)
    {
      bool doneIf = false;
      bool doneElse = false;
      int lineIdx = startLine;
      string line = GetScriptLine(lineIdx);

      while (line != "__GEMUS_EOF__")
      {
        string upper = line.ToUpperInvariant();
        if (upper.StartsWith("IF "))
        {
          doneIf = true;
          _scriptLinesPre[lineIdx] = new string('\t', level);
          string err = ValidateScriptBlock(lineIdx + 1, out nextLine, out errLine, level + 1);
          if (!string.IsNullOrEmpty(err)) return err;
          lineIdx = nextLine - 1;
        }
        else if (upper.StartsWith("ELSEIF ") || upper.StartsWith("ELSE IF "))
        {
          if (doneIf)
          {
            if (doneElse)
            {
              errLine = lineIdx;
              nextLine = lineIdx;
              return _langExpected + " 'END IF'";
            }
            _scriptLinesPre[lineIdx] = new string('\t', level);
            string err = ValidateScriptBlock(lineIdx + 1, out nextLine, out errLine, level + 1);
            if (!string.IsNullOrEmpty(err)) return err;
            lineIdx = nextLine - 1;
          }
          else
          {
            if (level == 0)
            {
              errLine = lineIdx;
              nextLine = lineIdx;
              return "'ELSE IF' " + _langWithout + " 'IF'";
            }
            nextLine = lineIdx;
            errLine = lineIdx;
            return string.Empty;
          }
        }
        else if (upper.StartsWith("ELSE"))
        {
          if (doneIf)
          {
            if (doneElse)
            {
              errLine = lineIdx;
              nextLine = lineIdx;
              return _langExpected + " 'END IF'";
            }
            doneElse = true;
            _scriptLinesPre[lineIdx] = new string('\t', level);
            string err = ValidateScriptBlock(lineIdx + 1, out nextLine, out errLine, level + 1);
            if (!string.IsNullOrEmpty(err)) return err;
            lineIdx = nextLine - 1;
          }
          else
          {
            if (level == 0)
            {
              errLine = lineIdx;
              nextLine = lineIdx;
              return "'ELSE' " + _langWithout + " 'IF'";
            }
            nextLine = lineIdx;
            errLine = lineIdx;
            return string.Empty;
          }
        }
        else if (upper.StartsWith("ENDIF") || upper.StartsWith("END IF"))
        {
          if (doneIf)
          {
            doneIf = false;
            doneElse = false;
            _scriptLinesPre[lineIdx] = new string('\t', level);
          }
          else
          {
            if (level == 0)
            {
              errLine = lineIdx;
              nextLine = lineIdx;
              return "'END IF' " + _langWithout + " 'IF'";
            }
            nextLine = lineIdx;
            errLine = lineIdx;
            return string.Empty;
          }
        }
        else
        {
          _scriptLinesPre[lineIdx] = string.IsNullOrEmpty(line) ? string.Empty : new string('\t', level);
        }

        lineIdx++;
        line = GetScriptLine(lineIdx);
      }

      nextLine = lineIdx;
      errLine = lineIdx;
      if (level == 0)
      {
        return doneIf ? _langExpected + " 'END IF'" : string.Empty;
      }
      return _langExpected + " 'END IF'";
    }

    private void GetScriptFunctionsToProcess(int startLine, out int nextLine, int level, bool process)
    {
      bool doneIf = false;
      bool blockDone = false;

      int lineIdx = startLine;
      string line = GetScriptLine(lineIdx);

      while (line != "__GEMUS_EOF__")
      {
        string upper = line.ToUpperInvariant();
        if (upper.StartsWith("IF "))
        {
          doneIf = true;
          bool evaluated = process && EvaluateLine(line);
          blockDone = evaluated;
          GetScriptFunctionsToProcess(lineIdx + 1, out nextLine, level + 1, evaluated);
          lineIdx = nextLine - 1;
        }
        else if (upper.StartsWith("ELSEIF ") || upper.StartsWith("ELSE IF "))
        {
          if (doneIf)
          {
            bool evaluated = process && !blockDone && EvaluateLine(line);
            if (evaluated) blockDone = true;
            GetScriptFunctionsToProcess(lineIdx + 1, out nextLine, level + 1, evaluated);
            lineIdx = nextLine - 1;
          }
          else
          {
            nextLine = lineIdx;
            return;
          }
        }
        else if (upper.StartsWith("ELSE"))
        {
          if (doneIf)
          {
            bool evaluated = process && !blockDone;
            if (evaluated) blockDone = true;
            GetScriptFunctionsToProcess(lineIdx + 1, out nextLine, level + 1, evaluated);
            lineIdx = nextLine - 1;
          }
          else
          {
            nextLine = lineIdx;
            return;
          }
        }
        else if (upper.StartsWith("ENDIF") || upper.StartsWith("END IF"))
        {
          if (doneIf)
          {
            doneIf = false;
            blockDone = false;
          }
          else
          {
            nextLine = lineIdx;
            return;
          }
        }
        else
        {
          if (!string.IsNullOrWhiteSpace(line) && !line.TrimStart().StartsWith(";"))
          {
            if (process)
              _scriptLinesToProcess.Add(line);
          }
        }

        lineIdx++;
        line = GetScriptLine(lineIdx);
      }

      nextLine = lineIdx;
    }

    private bool EvaluateLine(string line)
    {
      string field;
      string op;
      string data;

      int statementLen = line.StartsWith("IF ", StringComparison.OrdinalIgnoreCase) ? 3
        : line.StartsWith("ELSEIF ", StringComparison.OrdinalIgnoreCase) ? 7
        : 8;

      string lineTest = line;
      string err = ValidateComparisonLine(statementLen, ref lineTest, out field, out op, out data);
      if (!string.IsNullOrEmpty(err))
        return false;

      if (field.StartsWith("KEY_", StringComparison.OrdinalIgnoreCase))
      {
        string key = field[4..].Trim();
        var kv = _notesKeys.FirstOrDefault(k => k.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (kv == null) return false;
        bool result = EvaluateContainsText(kv.Value, data);
        return op.Equals("EXCLUDES", StringComparison.OrdinalIgnoreCase) ? !result : result;
      }

      if (field.StartsWith("%"))
      {
        string value = field;
        ReplaceEnvironmentVars(ref value, false, false);
        bool result = EvaluateContainsText(value, data);
        return op.Equals("EXCLUDES", StringComparison.OrdinalIgnoreCase) ? !result : result;
      }

      switch (field)
      {
        case "QUESTION":
          string question = op;
          ReplaceEnvironmentVars(ref question, false, false);
          bool yes = AskQuestion(question);
          return yes ? data == "YES" : data == "NO";
        case "NumPlayers":
          return CompareNumber(_fldNumPlayersTo, _fldNumPlayersFrom, data, op);
        case "ImageIndex":
          return CompareSingleNumber(_imageIndex, data, op);
        case "ImageName":
          if (_imageName == "") return false;
          return op == "EXCLUDES" ? !EvaluateContainsText(_imageName, data) : EvaluateContainsText(_imageName, data);
        case "Control":
          return EvaluateControl(data, op);
        case "GameType":
          return op == "EXCLUDES" ? !EvaluateContainsText(_fldGameType, data) : EvaluateContainsText(_fldGameType, data);
        case "GameComment":
          if (_fldGameComment == "") return false;
          return op == "EXCLUDES" ? !EvaluateContainsText(_fldGameComment, data) : EvaluateContainsText(_fldGameComment, data);
        case "VersionComment":
          if (_fldVersionComment == "") return false;
          return op == "EXCLUDES" ? !EvaluateContainsText(_fldVersionComment, data) : EvaluateContainsText(_fldVersionComment, data);
        case "PalNTSC":
          return EvaluatePalNtsc(data, op);
        case "TrueDriveEmu":
          return (data.Equals("YES", StringComparison.OrdinalIgnoreCase) && _fldTrueDriveEmu) ||
            (data.Equals("NO", StringComparison.OrdinalIgnoreCase) && !_fldTrueDriveEmu);
        case "NumGameFiles":
          return CompareSingleNumber(_fldNumGameFiles, data, op);
        default:
          return false;
      }
    }

    private bool CompareNumber(int toValue, int fromValue, string data, string op)
    {
      int compare = int.TryParse(data, out int val) ? val : 0;
      if (op == "=")
        return toValue == compare;
      if (op == "<")
        return toValue != PlayersUnknown && toValue < compare;
      if (op == ">")
      {
        if (fromValue == PlayersUnknown)
        {
          return toValue != PlayersUnknown && toValue > compare;
        }
        if (toValue == PlayersUnknown)
          return fromValue > compare;
        return toValue > compare;
      }
      return toValue != compare;
    }

    private bool CompareSingleNumber(int value, string data, string op)
    {
      int compare = int.TryParse(data, out int val) ? val : 0;
      return op switch
      {
        "=" => value == compare,
        "<" => value < compare,
        ">" => value > compare,
        _ => value != compare
      };
    }

    private bool EvaluateControl(string data, string op)
    {
      int? control = data switch
      {
        "JoyPort2" => 0,
        "JoyPort1" => 1,
        "Keyboard" => 2,
        "PaddlePort2" => 3,
        "PaddlePort1" => 4,
        "Mouse" => 5,
        "LightPen" => 6,
        "KoalaPad" => 7,
        "LightGun" => 8,
        _ => null
      };

      if (control == null)
        return op != "=";

      return op == "=" ? _fldControl == control : _fldControl != control;
    }

    private bool EvaluatePalNtsc(string data, string op)
    {
      string upper = data.ToUpperInvariant();
      if (_fldPalNtsc == 0)
        return op == "=" ? upper == "PAL" : upper != "PAL";
      if (_fldPalNtsc == 1)
        return op == "=" ? upper == "BOTH" : upper != "BOTH";
      if (_fldPalNtsc == 2)
        return op == "=" ? upper == "NTSC" : upper != "NTSC";
      return op != "=";
    }

    private bool EvaluateContainsText(string haystack, string needle)
    {
      string hayUpper = haystack.ToUpperInvariant();
      string[] args = needle.Split("||");

      foreach (string arg in args)
      {
        string test = arg.ToUpperInvariant();
        if (string.IsNullOrEmpty(test)) continue;
        if (test == "*" || test == "**")
          return haystack != "";

        if (test.Contains('*'))
        {
          if (test.StartsWith("*") && test.EndsWith("*"))
          {
            if (hayUpper.Contains(test[1..^1])) return true;
          }
          else if (test.StartsWith("*"))
          {
            if (hayUpper.EndsWith(test[1..])) return true;
          }
          else if (test.EndsWith("*"))
          {
            if (hayUpper.StartsWith(test[..^1])) return true;
          }
          else
          {
            if (hayUpper == test) return true;
          }
        }
        else
        {
          if (hayUpper == test) return true;
        }
      }

      return false;
    }

    private bool AskQuestion(string question)
    {
      if (Ui != null)
        return Ui.Confirm(question, _appName);
      return MessageBox.Show(question, _appName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
    }

    private bool WriteCfgValue(string filePath, string key, string value)
    {
      return GemusIniHelper.WriteCfgValue(filePath, key, value);
    }

    private bool WriteCfgValue2(string filePath, string key, string delim, string value)
    {
      try
      {
        if (!File.Exists(filePath))
          File.WriteAllText(filePath, "");

        string delimiter = string.IsNullOrEmpty(delim) ? " " : delim;
        var lines = new List<string>(File.ReadAllLines(filePath));
        bool found = false;
        string keyUpper = key.ToUpperInvariant();

        for (int i = 0; i < lines.Count; i++)
        {
          string line = lines[i].Trim();
          if (line.StartsWith(keyUpper, StringComparison.OrdinalIgnoreCase))
          {
            if (line.Length == keyUpper.Length ||
              line.Substring(keyUpper.Length).StartsWith(delimiter) ||
              NextNonWhiteSpaceCharIs(line, keyUpper.Length + 1, delimiter))
            {
              lines[i] = string.IsNullOrEmpty(value) ? key : key + delimiter + value;
              found = true;
            }
          }
        }

        if (!found)
        {
          lines.Add(key + delimiter + value);
        }

        File.WriteAllLines(filePath, lines);
        return true;
      }
      catch
      {
        return false;
      }
    }

    private bool NextNonWhiteSpaceCharIs(string text, int start, string token)
    {
      for (int i = start; i < text.Length; i++)
      {
        string c = text[i].ToString();
        if (c != " " && c != "\t")
        {
          return text.Substring(i).StartsWith(token);
        }
      }
      return false;
    }

    private void LoadScriptData()
    {
      _scriptFields.Clear();
      _scriptFunctions.Clear();
      _scriptEnvVars.Clear();

      _scriptFields.Add(new ScriptField
      {
        Name = "GameType",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { "=", "!=" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "NumPlayers",
        ValueType = ScriptValueType.Number,
        Operators = new List<string> { "=", "!=", "<", ">" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "Control",
        ValueType = ScriptValueType.List,
        Operators = new List<string> { "=", "!=" },
        ListValues = new List<string>
        {
          "JoyPort2", "JoyPort1", "Keyboard", "PaddlePort2", "PaddlePort1",
          "Mouse", "LightPen", "KoalaPad", "LightGun"
        }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "GameComment",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { string.Empty }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "PalNTSC",
        ValueType = ScriptValueType.List,
        Operators = new List<string> { "=", "!=" },
        ListValues = new List<string> { "PAL", "NTSC", "BOTH" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "TrueDriveEmu",
        ValueType = ScriptValueType.List,
        Operators = new List<string> { "=" },
        ListValues = new List<string> { "YES", "NO" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "VersionComment",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { string.Empty }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "NumGameFiles",
        ValueType = ScriptValueType.Number,
        Operators = new List<string> { "=", "!=", "<", ">" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "ImageIndex",
        ValueType = ScriptValueType.Number,
        Operators = new List<string> { "=", "!=", "<", ">" }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "ImageName",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { string.Empty }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "Key_",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { string.Empty }
      });
      _scriptFields.Add(new ScriptField
      {
        Name = "%",
        ValueType = ScriptValueType.Text,
        Operators = new List<string> { string.Empty }
      });

      _scriptFunctions.Add(new ScriptFunction { Name = "Add_CLP2", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Add_CLP", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Set_INI_Value", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.FileNumber, ScriptFunctionParamType.IniSection, ScriptFunctionParamType.Text, ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Set_CFG_Item", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.FileNumber, ScriptFunctionParamType.Text, ScriptFunctionParamType.Text, ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Set_CFG_Value", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.FileNumber, ScriptFunctionParamType.Text, ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Run_Program", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.FileNumber, ScriptFunctionParamType.Text, ScriptFunctionParamType.List }, ListValues = new List<string> { "WAIT", "NOWAIT" } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Run_Emulator", ParamTypes = new List<ScriptFunctionParamType>() });
      _scriptFunctions.Add(new ScriptFunction { Name = "Show_Message", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Edit_CLP", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Run_GameFile", ParamTypes = new List<ScriptFunctionParamType>() });
      _scriptFunctions.Add(new ScriptFunction { Name = "Clear_CLP", ParamTypes = new List<ScriptFunctionParamType>() });
      _scriptFunctions.Add(new ScriptFunction { Name = "Run_Emulator_Send_Keys", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text, ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Show_KV", ParamTypes = new List<ScriptFunctionParamType>() });
      _scriptFunctions.Add(new ScriptFunction { Name = "Wait", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });
      _scriptFunctions.Add(new ScriptFunction { Name = "Kill_Process", ParamTypes = new List<ScriptFunctionParamType> { ScriptFunctionParamType.Text } });

      _scriptEnvVars.Add("%gamepathfile%");
      _scriptEnvVars.Add("%gamefile%");
      _scriptEnvVars.Add("%gamepath%");
      _scriptEnvVars.Add("%gamefilenoext%");
      _scriptEnvVars.Add("%imageindex%");
      _scriptEnvVars.Add("%imagename%");

      for (int i = 1; i <= 20; i++)
      {
        _scriptEnvVars.Add($"%gamepathfile({i})%");
        _scriptEnvVars.Add($"%gamefile({i})%");
        _scriptEnvVars.Add($"%gamepath({i})%");
        _scriptEnvVars.Add($"%gamefilenoext({i})%");
      }
    }

    private sealed class SendKeysBlock
    {
      public int SecondsToWait { get; set; }
      public string KeysToSend { get; set; } = string.Empty;
    }
  }
}
