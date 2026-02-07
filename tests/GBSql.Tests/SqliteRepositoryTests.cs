//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql.Tests:  SQLite Repository Tests
//==============================================

using GBSql.Data;
using GBSql.Models;

namespace GBSql.Tests;

/// <summary>
/// Integration tests for SQLite repository.
/// </summary>
public class SqliteRepositoryTests : IDisposable
{
  private readonly string _testDbPath;
  private readonly GameBaseSqliteRepository _repository;

  public SqliteRepositoryTests()
  {
    // Create a temporary database for testing
    _testDbPath = Path.Combine(Path.GetTempPath(), $"gamebase_test_{Guid.NewGuid()}.db");
    _repository = DatabaseMigration.CreateNewDatabaseAsync(_testDbPath, overwriteExisting: true).GetAwaiter().GetResult();

    // Seed with test data - create new context and dispose properly
    using (var context = new GameBaseSqliteContext(_testDbPath))
    {
      DatabaseMigration.SeedSampleDataAsync(context).GetAwaiter().GetResult();
    }
  }

  [Fact]
  public async Task GetGamesAsync_ReturnsGames()
  {
    // Act
    var games = await _repository.GetGamesAsync();

    // Assert
    Assert.NotNull(games);
    Assert.NotEmpty(games);
    Assert.Equal(2, games.Count());
  }

  [Fact]
  public async Task GetGameByIdAsync_ReturnsCorrectGame()
  {
    // Act
    var game = await _repository.GetGameByIdAsync(1);

    // Assert
    Assert.NotNull(game);
    Assert.Equal("Test Game 1", game.Name);
    Assert.Equal("TG1", game.CCode);
    Assert.Equal(1, game.Fav);
  }

  [Fact]
  public async Task SearchGamesByNameAsync_FindsMatchingGames()
  {
    // Act
    var games = await _repository.SearchGamesByNameAsync("Test Game 1");

    // Assert
    Assert.NotNull(games);
    Assert.Single(games);
    Assert.Equal("Test Game 1", games.First().Name);
  }

  [Fact]
  public async Task GetExtrasForGameAsync_ReturnsExtras()
  {
    // Act
    var extras = await _repository.GetExtrasForGameAsync(1);

    // Assert
    Assert.NotNull(extras);
    Assert.Single(extras);
    Assert.Equal("Manual", extras.First().Name);
  }

  [Fact]
  public async Task GetLookupTableAsync_ReturnsPublishers()
  {
    // Act
    var publishers = await _repository.GetLookupTableAsync("Publishers");

    // Assert
    Assert.NotNull(publishers);
    Assert.Equal(2, publishers.Count());
    Assert.Contains(publishers, p => p.Name == "Electronic Arts");
    Assert.Contains(publishers, p => p.Name == "Activision");
  }

  [Fact]
  public async Task ExecuteQueryAsync_ReturnsResults()
  {
    // Act
    var results = await _repository.ExecuteQueryAsync("SELECT Name, CCode FROM Games WHERE Fav = 1");

    // Assert
    Assert.NotNull(results);
    var resultsList = results.ToList();
    Assert.Single(resultsList);

    dynamic firstResult = resultsList[0];
    Assert.Equal("Test Game 1", firstResult.Name);
  }

  [Fact]
  public async Task GetLookupTableAsync_InvalidTableName_ThrowsException()
  {
    // Act & Assert
    await Assert.ThrowsAsync<ArgumentException>(() =>
      _repository.GetLookupTableAsync("InvalidTable"));
  }

  public void Dispose()
  {
    _repository?.Dispose();

    // Allow time for SQLite to release file locks
    GC.Collect();
    GC.WaitForPendingFinalizers();
    
    // Try to delete the test database, but don't fail if it's locked
    if (File.Exists(_testDbPath))
    {
      try
      {
        File.Delete(_testDbPath);
      }
      catch (IOException)
      {
        // File is still locked, skip deletion (OS will clean up temp files)
      }
    }
  }
}
