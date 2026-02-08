using Microsoft.Data.Sqlite;

namespace BAC;

/// <summary>
/// Engine for merging diff databases into master databases.
/// NOTE: This functionality was not completed in the original VB6 version.
/// This is a stub for future implementation.
/// </summary>
internal class MergeEngine
{
    private readonly SqliteConnection _masterConnection;
    private readonly SqliteConnection _diffConnection;

    public MergeEngine(SqliteConnection masterConnection, SqliteConnection diffConnection)
    {
        _masterConnection = masterConnection ?? throw new ArgumentNullException(nameof(masterConnection));
        _diffConnection = diffConnection ?? throw new ArgumentNullException(nameof(diffConnection));
    }

    /// <summary>
    /// Analyzes differences between master and diff databases.
    /// </summary>
    /// <returns>A merge analysis result.</returns>
    public MergeAnalysis Analyze()
    {
        // TODO: Implement merge analysis
        // - Find new games in diff that don't exist in master
        // - Find games that exist in both but have differences
        // - Find games deleted from diff (marked for removal?)
        throw new NotImplementedException("Merge analysis not yet implemented");
    }

    /// <summary>
    /// Applies changes from diff database to master database.
    /// </summary>
    /// <param name="analysis">The merge analysis to apply.</param>
    /// <param name="options">Options controlling the merge behavior.</param>
    public void Apply(MergeAnalysis analysis, MergeOptions options)
    {
        // TODO: Implement merge application
        // - Use transactions for safety
        // - Insert new games
        // - Update modified games (handle conflicts)
        // - Delete games if requested
        // - Update related tables (musicians, publishers, etc.)
        throw new NotImplementedException("Merge application not yet implemented");
    }
}

/// <summary>
/// Results of analyzing differences between master and diff databases.
/// </summary>
internal class MergeAnalysis
{
    public List<int> NewGames { get; set; } = new();
    public List<int> ModifiedGames { get; set; } = new();
    public List<int> DeletedGames { get; set; } = new();
    public Dictionary<int, List<string>> Conflicts { get; set; } = new();
}

/// <summary>
/// Options controlling merge behavior.
/// </summary>
internal class MergeOptions
{
    public bool OverwriteExisting { get; set; } = false;
    public bool DeleteGames { get; set; } = false;
    public bool PromptOnConflicts { get; set; } = true;
}
