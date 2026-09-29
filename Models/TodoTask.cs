namespace TeamsMessageFetcher;

// ─── Microsoft To Do Task Model ───────────────────────────────────────────────

public class TodoTask
{
    public string    Id                { get; set; } = string.Empty;
    public string    Title             { get; set; } = string.Empty;
    public string    Status            { get; set; } = string.Empty;
    public string    Importance        { get; set; } = string.Empty;
    public DateTime  CreatedDateTime   { get; set; }
    public DateTime? CompletedDateTime { get; set; }
    public DateTime? DueDateTime       { get; set; }

    public bool IsCompleted => Status.Equals("completed", StringComparison.OrdinalIgnoreCase);
}
