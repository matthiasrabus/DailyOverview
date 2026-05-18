namespace TeamsMessageFetcher;

// ─── Mail Message Model ───────────────────────────────────────────────────────

public class MailMessage
{
    public string Id           { get; set; } = string.Empty;
    public string Subject      { get; set; } = string.Empty;
    public DateTime Timestamp  { get; set; }
    public string From         { get; set; } = string.Empty;
    public List<string> To     { get; set; } = [];
    public bool IsSent         { get; set; }
    public string BodyPreview  { get; set; } = string.Empty;
}
