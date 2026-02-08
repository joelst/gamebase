namespace BAC;

/// <summary>
/// Handles resolution of merge conflicts between master and diff databases.
/// NOTE: This functionality was not completed in the original VB6 version.
/// This is a stub for future implementation.
/// </summary>
internal class MergeConflictResolver
{
    /// <summary>
    /// Resolves a conflict for a specific game field.
    /// </summary>
    /// <param name="gameId">The ID of the game with a conflict.</param>
    /// <param name="fieldName">The name of the conflicting field.</param>
    /// <param name="masterValue">The value in the master database.</param>
    /// <param name="diffValue">The value in the diff database.</param>
    /// <returns>The resolved value to use.</returns>
    public object? ResolveConflict(int gameId, string fieldName, object? masterValue, object? diffValue)
    {
        // TODO: Implement conflict resolution
        // - Show dialog to user with both values
        // - Allow user to choose master, diff, or manual edit
        // - Return the chosen value
        throw new NotImplementedException("Conflict resolution not yet implemented");
    }

    /// <summary>
    /// Prompts the user to resolve multiple conflicts for a game.
    /// </summary>
    /// <param name="gameId">The ID of the game with conflicts.</param>
    /// <param name="conflicts">Dictionary of field names to (master, diff) value pairs.</param>
    /// <returns>Dictionary of field names to resolved values.</returns>
    public Dictionary<string, object?> ResolveGameConflicts(int gameId, Dictionary<string, (object? master, object? diff)> conflicts)
    {
        // TODO: Implement batch conflict resolution
        // - Show form with all conflicts for the game
        // - Allow user to resolve all at once
        // - Return all resolved values
        throw new NotImplementedException("Batch conflict resolution not yet implemented");
    }
}
