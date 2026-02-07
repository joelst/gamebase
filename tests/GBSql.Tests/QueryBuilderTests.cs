//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql.Tests:  QueryBuilder Tests
//==============================================

using GBSql;

namespace GBSql.Tests;

/// <summary>
/// Unit tests for QueryBuilder with SQLite syntax.
/// </summary>
public class QueryBuilderTests
{
  [Fact]
  public void GetViewQueryString_SqliteSyntax_UsesPercentWildcards()
  {
    // Arrange
    var qb = new QueryBuilder();
    qb.UseSqliteSyntax = true;

    // Create a simple filter buffer array (1-based indexing from VB6)
    var filters = new string[2];
    filters[0] = string.Empty; // Index 0 unused in VB6 style
    
    // Build a filter for "Name contains 'Test'"
    var filter = BuildFilterBuffer("Games", "Name", FilterOperator.Equal, FilterClauseType.FieldContainsText, "Test");
    filters[1] = filter;

    // Act
    string sql = qb.GetViewQueryString(
      isGamesTable: true,
      secColFields: "",
      filterCount: 1,
      filterBuffers: filters,
      andMode: true,
      noAdultItems: false);

    // Assert
    Assert.Contains("LIKE \"%Test%\"", sql);
    Assert.DoesNotContain("LIKE \"*Test*\"", sql);
  }

  [Fact]
  public void GetViewQueryString_SqliteSyntax_UsesNumericBooleans()
  {
    // Arrange
    var qb = new QueryBuilder();
    qb.UseSqliteSyntax = true;

    var filters = new string[2];
    filters[0] = string.Empty;
    
    // Build a filter for file exists check
    var filter = BuildFilterBuffer("Games", "Filename", FilterOperator.Equal, FilterClauseType.FileExists, "");
    filters[1] = filter;

    // Act
    string sql = qb.GetViewQueryString(
      isGamesTable: true,
      secColFields: "",
      filterCount: 1,
      filterBuffers: filters,
      andMode: true,
      noAdultItems: false);

    // Assert
    Assert.Contains("= 1", sql); // SQLite boolean true
    Assert.DoesNotContain("= True", sql);
  }

  [Fact]
  public void GetViewQueryString_NoAdultItems_UsesSqliteBoolSyntax()
  {
    // Arrange
    var qb = new QueryBuilder();
    qb.UseSqliteSyntax = true;

    var filters = new string[1];
    filters[0] = string.Empty;

    // Act
    string sql = qb.GetViewQueryString(
      isGamesTable: true,
      secColFields: "",
      filterCount: 0,
      filterBuffers: filters,
      andMode: true,
      noAdultItems: true);

    // Assert
    Assert.Contains("Adult = 0", sql); // SQLite boolean false
    Assert.DoesNotContain("Adult = False", sql);
  }

  [Fact]
  public void GetViewQueryString_AccessSyntax_UsesAsterisks()
  {
    // Arrange
    var qb = new QueryBuilder();
    qb.UseSqliteSyntax = false; // Use Access syntax

    var filters = new string[2];
    filters[0] = string.Empty;
    var filter = BuildFilterBuffer("Games", "Name", FilterOperator.Equal, FilterClauseType.FieldContainsText, "Test");
    filters[1] = filter;

    // Act
    string sql = qb.GetViewQueryString(
      isGamesTable: true,
      secColFields: "",
      filterCount: 1,
      filterBuffers: filters,
      andMode: true,
      noAdultItems: false);

    // Assert
    Assert.Contains("LIKE \"*Test*\"", sql);
    Assert.DoesNotContain("LIKE \"%Test%\"", sql);
  }

  [Fact]
  public void GetViewQueryString_IncludesOrderBy()
  {
    // Arrange
    var qb = new QueryBuilder();
    var filters = new string[1];
    filters[0] = string.Empty;

    // Act
    string sql = qb.GetViewQueryString(
      isGamesTable: true,
      secColFields: "",
      filterCount: 0,
      filterBuffers: filters,
      andMode: true,
      noAdultItems: false);

    // Assert
    Assert.Contains("ORDER BY Games.Name", sql);
  }

  /// <summary>
  /// Helper to build a filter buffer matching VB6 CVFilterFixedType structure.
  /// </summary>
  private static string BuildFilterBuffer(
    string fieldTable, 
    string fieldName, 
    FilterOperator op, 
    FilterClauseType clauseType, 
    string clauseData)
  {
    // Format: FieldTable(128) + FieldName(128) + MusicFieldTable(128) + MusicFieldName(128) +
    //         Operator(4) + ClauseType(4) + ClauseData(64)
    var buffer = new char[580];
    Array.Fill(buffer, ' ');

    // FieldTable at position 0
    fieldTable.AsSpan().CopyTo(buffer.AsSpan(0, Math.Min(fieldTable.Length, 128)));
    
    // FieldName at position 128
    fieldName.AsSpan().CopyTo(buffer.AsSpan(128, Math.Min(fieldName.Length, 128)));

    // MusicFieldTable at position 256 (empty for this test)
    // MusicFieldName at position 384 (empty for this test)

    // Operator at position 512 (4 chars)
    string opStr = ((int)op).ToString().PadLeft(4);
    opStr.AsSpan().CopyTo(buffer.AsSpan(512, 4));

    // ClauseType at position 516 (4 chars)
    string ctStr = ((int)clauseType).ToString().PadLeft(4);
    ctStr.AsSpan().CopyTo(buffer.AsSpan(516, 4));

    // ClauseData at position 520 (60 chars)
    clauseData.AsSpan().CopyTo(buffer.AsSpan(520, Math.Min(clauseData.Length, 60)));

    return new string(buffer);
  }
}
