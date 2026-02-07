// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// Gemus: Logging service for script execution tracing
// Converted from VB6 modFuncs.bas (LogStatus, PushLogFileTrace, PopLogFileTrace)

namespace Gemus
{
  /// <summary>
  /// Provides structured log file output for GEMUS script execution.
  /// Replaces the VB6 global log file number and trace stack.
  /// </summary>
  public class GemusLogger : IDisposable
  {
    private StreamWriter? _writer;
    private readonly Stack<string> _traceStack = new();
    private string _tracePrefix = string.Empty;
    private const string TraceSeparator = "    ";

    /// <summary>
    /// Creates or opens the log file.
    /// </summary>
    public void CreateLogFile(string logFilePath)
    {
      try
      {
        _traceStack.Clear();
        _tracePrefix = string.Empty;

        if (File.Exists(logFilePath))
          File.Delete(logFilePath);

        _writer = new StreamWriter(logFilePath, false);
        _writer.WriteLine($"***** GEMUS LOG FILE *****");
        _writer.WriteLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        _writer.Flush();
      }
      catch
      {
        _writer = null;
      }
    }

    /// <summary>
    /// Closes the log file.
    /// </summary>
    public void CloseLogFile()
    {
      try
      {
        if (_writer != null)
        {
          _writer.WriteLine("Bye!");
          _writer.Flush();
          _writer.Close();
          _writer = null;
        }
      }
      catch
      {
        _writer = null;
      }
    }

    /// <summary>
    /// Writes a status line to the log file.
    /// </summary>
    public void LogStatus(string status)
    {
      try
      {
        _writer?.WriteLine($"{_tracePrefix}{status}");
        _writer?.Flush();
      }
      catch
      {
        // Silently fail if log file can't be written
      }
    }

    /// <summary>
    /// Pushes a function name onto the trace stack and logs a separator line.
    /// </summary>
    public void PushLogFileTrace(string funcName)
    {
      _traceStack.Push(funcName);
      UpdatePrefix();
      LogStatus($"-----------------{funcName}-----------------");
    }

    /// <summary>
    /// Pops from the trace stack and logs a closing separator line.
    /// </summary>
    public void PopLogFileTrace()
    {
      if (_traceStack.Count > 0)
      {
        string funcName = _traceStack.Pop();
        LogStatus($"-----------------{new string('-', funcName.Length)}-----------------");
        UpdatePrefix();
      }
    }

    private void UpdatePrefix()
    {
      _tracePrefix = string.Concat(Enumerable.Repeat(TraceSeparator, _traceStack.Count));
    }

    public void Dispose()
    {
      CloseLogFile();
      GC.SuppressFinalize(this);
    }
  }
}
