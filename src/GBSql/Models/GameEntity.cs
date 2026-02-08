//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  Data Access Models
//==============================================

namespace GBSql.Models;

/// <summary>
/// Represents a Game entity in the GameBase database.
/// </summary>
public class GameEntity
{
  public int GA_Id { get; set; }
  public string Name { get; set; } = string.Empty;
  public string? CCode { get; set; }
  public int? PU_Id { get; set; }
  public int? YE_Id { get; set; }
  public int? PR_Id { get; set; }
  public int? MU_Id { get; set; }
  public int? LA_Id { get; set; }
  public int? GE_Id { get; set; }
  public int? PG_Id { get; set; }
  public int? DI_Id { get; set; }
  public int? CR_Id { get; set; }
  public int? AR_Id { get; set; }
  public int? DE_Id { get; set; }
  public int? LI_Id { get; set; }
  public int? RA_Id { get; set; }
  public int Fav { get; set; }
  public int Classic { get; set; }
  public int Adult { get; set; }
  public int PlayersFrom { get; set; }
  public int PlayersTo { get; set; }
  public int? Prequel { get; set; }
  public int? Sequel { get; set; }
  public int? Related { get; set; }
  public int? CloneOf { get; set; }
  public int? V_Trainers { get; set; }
  public int? V_Length { get; set; }
  public int? V_LengthType { get; set; }
  public int? ReviewRating { get; set; }
  public string? Comment { get; set; }
  public string? Filename { get; set; }
  public string? ScreenshotFilename { get; set; }
}

/// <summary>
/// Represents a lookup table entity (Publishers, Years, etc.).
/// </summary>
public class LookupEntity
{
  public int Id { get; set; }
  public string Name { get; set; } = string.Empty;
}

/// <summary>
/// Represents an Extra (additional content) entity.
/// </summary>
public class ExtraEntity
{
  public int EX_Id { get; set; }
  public int GA_Id { get; set; }
  public string Name { get; set; } = string.Empty;
  public string? Filename { get; set; }
  public int? EX_Type { get; set; }
}
