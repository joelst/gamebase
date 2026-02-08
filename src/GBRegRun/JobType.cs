// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: Job type enumeration

namespace GBRegRun;

/// <summary>
/// Types of jobs that GBRegRun can execute.
/// </summary>
internal enum JobType
{
  /// <summary>
  /// Unregister a DLL using regsvr32 /u.
  /// </summary>
  Unregister,

  /// <summary>
  /// Register a DLL using regsvr32.
  /// </summary>
  Register,

  /// <summary>
  /// Run an executable without waiting for it to complete.
  /// </summary>
  RunNoWait,

  /// <summary>
  /// Run an executable and wait for it to complete.
  /// </summary>
  RunAndWait
}
