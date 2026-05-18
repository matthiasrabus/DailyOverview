namespace TeamsMessageFetcher;

// ─── Message Model ────────────────────────────────────────────────────────────

public class TeamMessage
{
    public string Id { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;       // Chat name or "Team › #Channel"
    public string SenderName { get; set; } = string.Empty;
    public bool IsSentByMe { get; set; }
    public DateTime Timestamp { get; set; }
    public string BodyPreview { get; set; } = string.Empty;  // Truncated, HTML stripped
    public string FullBody { get; set; } = string.Empty;     // Full text
    public string MessageType { get; set; } = "message";
}
