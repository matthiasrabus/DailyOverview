namespace TeamsMessageFetcher;

// ─── Azure DevOps Work Item Model ────────────────────────────────────────────

public class DevOpsWorkItem
{
    public int      Id          { get; set; }
    public string   Title       { get; set; } = string.Empty;
    public string   Type        { get; set; } = string.Empty;
    public string   State       { get; set; } = string.Empty;
    public string   Project     { get; set; } = string.Empty;
    public string   AssignedTo  { get; set; } = string.Empty;
    public DateTime ChangedDate { get; set; }
    public string   Url         { get; set; } = string.Empty;
}
