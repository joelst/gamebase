//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   GBSql:  Database Functions used in GB
//==============================================

using System.Text;

namespace GBSql
{
  /// <summary>
  /// Filter clause types for Custom View filters.
  /// Converted from VB6 FilterClauseTypeEnum.
  /// </summary>
  public enum FilterClauseType
  {
    FieldContainsText = 0,
    FileExists = 1,
    FieldIsFilled = 2,
    IsDatabaseField = 3,
    NoClause = 4
  }

  /// <summary>
  /// Filter operators for Custom View filters.
  /// Converted from VB6 FilterOperatorEnum.
  /// </summary>
  public enum FilterOperator
  {
    Equal = 0,
    DoesNotEqual = 1,
    IsBefore = 2,
    IsAfter = 3,
    IsLessThan = 4,
    IsMoreThan = 5,
    StartsWith = 6,
    EndsWith = 7
  }

  /// <summary>
  /// Represents a single Custom View filter definition.
  /// Converted from VB6 CVFilterFixedType.
  /// </summary>
  public class CVFilter
  {
    public string FieldTable { get; set; } = string.Empty;
    public string FieldName { get; set; } = string.Empty;
    public string MusicFieldTable { get; set; } = string.Empty;
    public string MusicFieldName { get; set; } = string.Empty;
    public FilterOperator Operator { get; set; } = FilterOperator.Equal;
    public FilterClauseType ClauseType { get; set; } = FilterClauseType.NoClause;
    public string ClauseData { get; set; } = string.Empty;

    /// <summary>
    /// Size of the fixed-length VB6 buffer structure for interop.
    /// FieldTable(128) + FieldName(128) + MusicFieldTable(128) + MusicFieldName(128) +
    /// Operator(4) + ClauseType(4) + ClauseData(64) = 580 bytes (approx).
    /// </summary>
    public const int FixedBufferSize = 580;

    /// <summary>
    /// Deserialise a filter from a VB6-format fixed-length string buffer.
    /// </summary>
    public static CVFilter FromBuffer(string buffer)
    {
      var filter = new CVFilter();
      if (string.IsNullOrEmpty(buffer) || buffer.Length < FixedBufferSize)
      {
        return filter;
      }

      filter.FieldTable = buffer.Substring(0, 128).Trim();
      filter.FieldName = buffer.Substring(128, 128).Trim();
      filter.MusicFieldTable = buffer.Substring(256, 128).Trim();
      filter.MusicFieldName = buffer.Substring(384, 128).Trim();

      // Operator and ClauseType are stored as 4-byte longs in VB6
      if (int.TryParse(buffer.Substring(512, 4).Trim(), out int op))
        filter.Operator = (FilterOperator)op;
      if (int.TryParse(buffer.Substring(516, 4).Trim(), out int ct))
        filter.ClauseType = (FilterClauseType)ct;

      filter.ClauseData = buffer.Substring(520, 60).Trim();

      return filter;
    }
  }

  /// <summary>
  /// Builds SQL query strings for GameBase Custom Views.
  /// Converted from VB6 clsQuery class.
  /// </summary>
  public class QueryBuilder
  {
    private CVFilter[] _filters = Array.Empty<CVFilter>();
    private int _filterCount;
    private bool _isGamesTable;
    private string _secColFields = string.Empty;
    private string _defaultFields = string.Empty;
    private string _defaultTable = string.Empty;
    private bool _andMode;
    private bool _noAdultItems;

    private const string DefGamesFields = "Games.Name, Games.CCode, Games.GA_Id, Games.Fav AS Bold1, Games.Classic AS Bold2";
    private const string DefGamesTable = "Games";
    private const string DefMusicFields = "Music.Name, Music.GA_Id, Music.SFav AS Bold1";
    private const string DefMusicTable = "Music";

    /// <summary>
    /// Gets the SQL Query string for a View in GameBase.
    /// </summary>
    /// <param name="isGamesTable">Whether we are operating on the Games or Music table.</param>
    /// <param name="secColFields">The fields shown in the second column.</param>
    /// <param name="filterCount">The number of filters in the view.</param>
    /// <param name="filterBuffers">The string buffer array containing the filters (1-based indexing from VB6).</param>
    /// <param name="andMode">Filter mode: true=AND, false=OR.</param>
    /// <param name="noAdultItems">Whether to filter out Adult items.</param>
    /// <returns>The SQL string for the view.</returns>
    public string GetViewQueryString(
        bool isGamesTable,
        string secColFields,
        int filterCount,
        string[] filterBuffers,
        bool andMode,
        bool noAdultItems)
    {
      // Set up the filters array
      if (filterCount == 0)
      {
        _filterCount = 0;
        _filters = Array.Empty<CVFilter>();
      }
      else
      {
        _filterCount = filterCount;
        _filters = new CVFilter[_filterCount + 1]; // 1-based indexing
        _filters[0] = new CVFilter(); // placeholder for index 0
        for (int i = 1; i <= _filterCount; i++)
        {
          _filters[i] = CVFilter.FromBuffer(filterBuffers[i]);
        }
      }

      // Store the rest of the data
      _andMode = andMode;
      _secColFields = secColFields;
      _isGamesTable = isGamesTable;
      _noAdultItems = noAdultItems;

      if (_isGamesTable)
      {
        _defaultFields = DefGamesFields;
        _defaultTable = DefGamesTable;
      }
      else
      {
        _defaultFields = DefMusicFields;
        _defaultTable = DefMusicTable;
      }

      // Build SELECT / FROM / INNER JOIN
      string query = GetTablesAndFields();

      // Build WHERE clause
      if (_filterCount > 0)
      {
        query += " " + GetWhereClause();
      }
      else if (_noAdultItems)
      {
        query += $" WHERE ({_defaultTable}.Adult = False)";
      }

      // ORDER BY
      query += $" ORDER BY {_defaultTable}.Name";

      // DISTINCT for Extras
      if (query.Contains("Extras.GA_Id"))
      {
        query = query.Replace("SELECT ", "SELECT DISTINCT ", StringComparison.Ordinal);
        // Only replace the first occurrence
        int extrasIdx = query.IndexOf("FROM Extras INNER ", StringComparison.Ordinal);
        if (extrasIdx >= 0)
        {
          query = query.Remove(extrasIdx, "FROM Extras INNER ".Length)
                       .Insert(extrasIdx, "FROM Extras RIGHT ");
        }
      }

      return query;
    }

    /// <summary>
    /// Gets the filter text for the SQL query from a filter definition.
    /// </summary>
    private string GetFilter(CVFilter filter)
    {
      string fieldName, tableName;

      if (_isGamesTable)
      {
        fieldName = filter.FieldName.Trim();
        tableName = filter.FieldTable.Trim();
      }
      else
      {
        fieldName = filter.MusicFieldName.Trim();
        tableName = filter.MusicFieldTable.Trim();
      }

      string tableDotField = $"{tableName}.{fieldName}";
      string clauseData = filter.ClauseData.Trim();

      switch (filter.ClauseType)
      {
        case FilterClauseType.FieldContainsText:
          return GetFieldContainsTextFilter(filter.Operator, tableName, tableDotField, clauseData);

        case FilterClauseType.FieldIsFilled:
          return filter.Operator == FilterOperator.Equal
              ? $"{tableDotField} <> \"\""
              : $"{tableDotField} = \"\"";

        case FilterClauseType.FileExists:
          return filter.Operator == FilterOperator.Equal
              ? $"{tableDotField} = True"
              : $"{tableDotField} = False";

        case FilterClauseType.IsDatabaseField:
          return GetDatabaseFieldFilter(filter.Operator, fieldName, tableDotField, clauseData);

        default: // NoClause
          return $"{_defaultTable}.Name <> \"\"";
      }
    }

    private string GetFieldContainsTextFilter(FilterOperator op, string tableName, string tableDotField, string clauseData)
    {
      string escaped = WildCardReplace(clauseData);
      bool isGenres = tableName == "Genres";

      return op switch
      {
        FilterOperator.Equal => isGenres
            ? $"{tableDotField} LIKE \"*{escaped}*\" OR PGenres.ParentGenre LIKE \"*{escaped}*\""
            : $"{tableDotField} LIKE \"*{escaped}*\"",

        FilterOperator.DoesNotEqual => isGenres
            ? $"{tableDotField} NOT LIKE \"*{escaped}*\" OR PGenres.ParentGenre LIKE \"*{escaped}*\""
            : $"{tableDotField} NOT LIKE \"*{escaped}*\"",

        FilterOperator.StartsWith => isGenres
            ? $"{tableDotField} LIKE \"{escaped}*\" OR PGenres.ParentGenre LIKE \"{escaped}*\""
            : $"{tableDotField} LIKE \"{escaped}*\"",

        FilterOperator.EndsWith => isGenres
            ? $"{tableDotField} LIKE \"*{escaped}\" OR PGenres.ParentGenre LIKE \"*{escaped}\""
            : $"{tableDotField} LIKE \"*{escaped}\"",

        _ => $"{tableDotField} LIKE \"*{escaped}*\""
      };
    }

    private string GetDatabaseFieldFilter(FilterOperator op, string fieldName, string tableDotField, string clauseData)
    {
      switch (op)
      {
        case FilterOperator.Equal:
          if (fieldName == "PlayersFrom")
            return $"({tableDotField} = {clauseData}) AND (Games.PlayersTo = {clauseData})";
          if (fieldName == "PlayersTo")
            return $"({tableDotField} = {clauseData})";
          if (fieldName is "Prequel" or "Sequel" or "Related" or "CloneOf")
            return clauseData == "-1"
                ? $"{tableDotField} > 0"
                : $"{tableDotField} = 0";
          return $"{tableDotField} = {clauseData}";

        case FilterOperator.DoesNotEqual:
          if (fieldName == "PlayersFrom")
            return $"NOT ({tableDotField} = {clauseData}) AND (Games.PlayersTo = {clauseData})";
          return $"NOT ({tableDotField} = {clauseData})";

        case FilterOperator.IsAfter:
          return $"({tableDotField} > {clauseData}) AND ({tableDotField} < 9991)";

        case FilterOperator.IsBefore:
          return $"{tableDotField} < {clauseData}";

        case FilterOperator.IsLessThan:
          if (fieldName is "PlayersTo" or "V_Trainers" or "ReviewRating")
            return $"({tableDotField} < {clauseData}) AND ({tableDotField} > -1)";
          if (fieldName is "V_Length" or "V_LengthType")
            return $"({tableDotField} < {clauseData})";
          return $"({tableDotField} < {clauseData}) AND ({tableDotField} > 0)";

        case FilterOperator.IsMoreThan:
          if (fieldName == "PlayersFrom")
            return $"({tableDotField} > {clauseData}) OR (PlayersTo > {clauseData})";
          if (fieldName is "V_Trainers" or "Rating" or "ReviewRating" or "V_Length" or "V_LengthType")
            return $"{tableDotField} > {clauseData}";
          return string.Empty;

        default:
          return string.Empty;
      }
    }

    /// <summary>
    /// Gets the link field to use in the INNER JOIN part of a query.
    /// </summary>
    private static string GetLinkField(string tableName)
    {
      return tableName switch
      {
        "Publishers" => "PU_Id",
        "Years" => "YE_Id",
        "Programmers" => "PR_Id",
        "Musicians" => "MU_Id",
        "Languages" => "LA_Id",
        "Genres" => "GE_Id",
        "PGenres" => "PG_Id",
        "Difficulty" => "DI_Id",
        "Crackers" => "CR_Id",
        "Artists" => "AR_Id",
        "Developers" => "DE_Id",
        "Licenses" => "LI_Id",
        "Rarities" => "RA_Id",
        "Extras" => "GA_Id",
        _ => string.Empty
      };
    }

    /// <summary>
    /// Traverses all the filters and formulates the WHERE clause.
    /// </summary>
    private string GetWhereClause()
    {
      string andOr = _andMode ? " AND " : " OR ";
      var sb = new StringBuilder("WHERE (");

      for (int i = 1; i <= _filterCount; i++)
      {
        sb.Append('(');
        sb.Append(GetFilter(_filters[i]));
        sb.Append(')');

        if (i < _filterCount)
          sb.Append(andOr);
      }

      if (_noAdultItems)
      {
        sb.Append($" AND ({_defaultTable}.Adult = False)");
      }

      sb.Append(')');
      return sb.ToString();
    }

    /// <summary>
    /// Constructs the SELECT fields and FROM / INNER JOIN part of the SQL.
    /// </summary>
    private string GetTablesAndFields()
    {
      // Parse all fields into a list
      var allFields = new List<string>();
      ParseFieldsInto(allFields, _defaultFields);
      if (!string.IsNullOrEmpty(_secColFields))
      {
        ParseFieldsInto(allFields, _secColFields);
      }

      // Remove duplicates (preserve aliases)
      var correctFields = new List<string>();
      foreach (var rawField in allFields)
      {
        string field = rawField.Trim();
        string fieldWithoutAlias = field;
        int aliasStart = field.IndexOf(" AS ", StringComparison.OrdinalIgnoreCase);
        if (aliasStart > 0)
          fieldWithoutAlias = field.Substring(0, aliasStart).Trim();

        bool found = false;
        for (int j = 0; j < correctFields.Count; j++)
        {
          if (correctFields[j].Contains(fieldWithoutAlias, StringComparison.OrdinalIgnoreCase))
          {
            // Preserve alias if existing entry has none
            if (!correctFields[j].Contains(" AS ", StringComparison.OrdinalIgnoreCase) &&
                field.Contains(" AS ", StringComparison.OrdinalIgnoreCase))
            {
              correctFields[j] = field;
            }
            found = true;
            break;
          }
        }

        if (!found)
          correctFields.Add(field);
      }

      // Build SELECT part
      var sb = new StringBuilder("SELECT ");
      sb.Append(string.Join(", ", correctFields));
      sb.Append(' ');

      // Extract unique table names from fields
      var tables = new List<string>();
      foreach (var field in correctFields)
      {
        int dotIdx = field.IndexOf('.');
        if (dotIdx > 0)
        {
          string table = field.Substring(0, dotIdx);
          if (!tables.Contains(table))
            tables.Add(table);
        }
      }

      // Add unique tables from filters
      bool includePGenre = false;
      bool includeGenre = false;

      if (_filterCount > 0)
      {
        for (int i = 1; i <= _filterCount; i++)
        {
          string table = _isGamesTable
              ? _filters[i].FieldTable.Trim()
              : _filters[i].MusicFieldTable.Trim();

          if (table == "Genres") includePGenre = true;
          if (table == "PGenres") includeGenre = true;

          if (!tables.Contains(table))
            tables.Add(table);
        }

        if (includePGenre && !tables.Contains("PGenres"))
          tables.Add("PGenres");
        if (includeGenre && !tables.Contains("Genres"))
          tables.Add("Genres");
      }

      // Re-order: PGenres first, Games/Music last
      if (tables.Count > 1)
      {
        int pgenresIdx = tables.IndexOf("PGenres");
        if (pgenresIdx > 0)
        {
          (tables[0], tables[pgenresIdx]) = (tables[pgenresIdx], tables[0]);
        }

        int defaultIdx = tables.IndexOf(_defaultTable);
        if (defaultIdx >= 0 && defaultIdx < tables.Count - 1)
        {
          (tables[^1], tables[defaultIdx]) = (tables[defaultIdx], tables[^1]);
        }
      }

      // Build FROM / INNER JOIN
      sb.Append("FROM ");
      if (tables.Count > 1)
      {
        // Inbound recursive INNER JOINs
        for (int i = 0; i < tables.Count; i++)
        {
          sb.Append(tables[i]);
          if (i < tables.Count - 1)
            sb.Append(" INNER JOIN ");
          if (i < tables.Count - 2)
            sb.Append('(');
        }

        // Outbound recursive ON clauses
        for (int i = tables.Count - 2; i >= 0; i--)
        {
          string linkField = GetLinkField(tables[i]);
          string linkTable = tables[i] == "PGenres" ? "Genres" : _defaultTable;
          sb.Append($" ON {tables[i]}.{linkField} = {linkTable}.{linkField}");
          if (i > 0)
            sb.Append(')');
        }
      }
      else
      {
        sb.Append(tables[0]);
      }

      return sb.ToString();
    }

    /// <summary>
    /// Parses a comma-separated field string into a list.
    /// </summary>
    private static void ParseFieldsInto(List<string> target, string fieldString)
    {
      foreach (string field in fieldString.Split(','))
      {
        string trimmed = field.Trim();
        if (!string.IsNullOrEmpty(trimmed))
          target.Add(trimmed);
      }
    }

    /// <summary>
    /// Replaces wildcard characters for use in LIKE clauses.
    /// </summary>
    private static string WildCardReplace(string searchString)
    {
      string result = searchString.Replace("[", "[[]");
      result = result.Replace("?", "[?]");
      result = result.Replace("*", "[*]");
      result = result.Replace("!", "[!]");
      result = result.Replace("#", "[#]");
      return result;
    }
  }
}
