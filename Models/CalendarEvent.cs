namespace TeamsMessageFetcher;

// ─── Calendar Event Model ─────────────────────────────────────────────────────

public class CalendarEvent
{
    public string Id          { get; set; } = string.Empty;
    public string Subject     { get; set; } = string.Empty;
    public DateTime Start     { get; set; }
    public DateTime End       { get; set; }
    public string Organizer   { get; set; } = string.Empty;
    public bool IsOrganizer   { get; set; }
    public bool IsOnline      { get; set; }
    public List<string> Attendees { get; set; } = [];

    public TimeSpan Duration => End - Start;
}
