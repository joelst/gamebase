//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  Data Access Interfaces
//==============================================

using GBSql.Models;

namespace GBSql.Data;

/// <summary>
/// Repository interface for GameBase database operations.
/// </summary>
public interface IGameBaseRepository : IDisposable
{
  /// <summary>
  /// Executes a custom SQL query and returns results.
  /// </summary>
  Task<IEnumerable<T>> QueryAsync<T>(string sql, object? parameters = null) where T : new();

  /// <summary>
  /// Executes a raw SQL query built by QueryBuilder.
  /// </summary>
  Task<IEnumerable<dynamic>> ExecuteQueryAsync(string sql);

  /// <summary>
  /// Gets a game by ID.
  /// </summary>
  Task<GameEntity?> GetGameByIdAsync(int gaId);

  /// <summary>
  /// Gets all games (with optional limit).
  /// </summary>
  Task<IEnumerable<GameEntity>> GetGamesAsync(int? limit = null);

  /// <summary>
  /// Searches games by name.
  /// </summary>
  Task<IEnumerable<GameEntity>> SearchGamesByNameAsync(string searchTerm);

  /// <summary>
  /// Gets extras for a specific game.
  /// </summary>
  Task<IEnumerable<ExtraEntity>> GetExtrasForGameAsync(int gaId);

  /// <summary>
  /// Inserts or updates an extra. Returns the resulting EX_Id.
  /// </summary>
  Task<int> UpsertExtraAsync(ExtraEntity extra);

  /// <summary>
  /// Deletes an extra by EX_Id.
  /// </summary>
  Task DeleteExtraAsync(int exId);

  /// <summary>
  /// Replaces all extras for a game with the provided list.
  /// </summary>
  Task SaveExtrasForGameAsync(int gaId, IEnumerable<ExtraEntity> extras);

  /// <summary>
  /// Gets a lookup table by name (Publishers, Years, etc.).
  /// </summary>
  Task<IEnumerable<LookupEntity>> GetLookupTableAsync(string tableName);
}
