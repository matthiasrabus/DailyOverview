namespace TeamsMessageFetcher;

// ─── Azure DevOps Commit Model ────────────────────────────────────────────────

public class DevOpsCommit
{
    public string CommitId    { get; set; } = string.Empty;
    public string Message     { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; }
    public string Project     { get; set; } = string.Empty;
    public string Repository  { get; set; } = string.Empty;
    public string Url         { get; set; } = string.Empty;
}
