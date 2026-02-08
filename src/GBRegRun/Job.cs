// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1999-2010 James Burrows - Migrated to .NET 2026
// GBRegRun: Job model

namespace GBRegRun;

/// <summary>
/// Represents a single job to be executed by GBRegRun.
/// </summary>
internal class Job
{
  public JobType Type { get; set; }
  public string RelativePathFile { get; set; } = string.Empty;
}
