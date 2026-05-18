using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsMessageFetcher;

public class CalendarFetcher(GraphServiceClient graph)
{
    private readonly GraphServiceClient _graph = graph;

    public async Task<List<CalendarEvent>> GetEventsForDateAsync(DateOnly date)
    {
        var result = new List<CalendarEvent>();

        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        try
        {
            var response = await _graph.Me.CalendarView.GetAsync(config =>
            {
                config.QueryParameters.StartDateTime = startUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");
                config.QueryParameters.EndDateTime   = endUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");
                config.QueryParameters.Select        = ["subject", "start", "end", "organizer",
                                                        "isOrganizer", "isOnlineMeeting", "attendees"];
                config.QueryParameters.Top           = 50;
            });

            var events = await GraphPaginator.PaginateAsync<Event, EventCollectionResponse>(_graph, response);

            result.AddRange(events.Select(Map));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not fetch calendar events: {ex.Message}");
        }

        return result;
    }

    // ─── Mapping ──────────────────────────────────────────────────────────────

    private static CalendarEvent Map(Event ev) => new()
    {
        Id          = ev.Id ?? string.Empty,
        Subject     = ev.Subject ?? "(No subject)",
        Start       = ParseEventTime(ev.Start),
        End         = ParseEventTime(ev.End),
        Organizer   = ev.Organizer?.EmailAddress?.Name ?? "Unknown",
        IsOrganizer = ev.IsOrganizer ?? false,
        IsOnline    = ev.IsOnlineMeeting ?? false,
        Attendees   = ev.Attendees?
                        .Select(a => a.EmailAddress?.Name ?? a.EmailAddress?.Address ?? "?")
                        .ToList() ?? [],
    };

    private static DateTime ParseEventTime(DateTimeTimeZone? dt)
    {
        if (dt?.DateTime == null) return DateTime.MinValue;

        var parsed = DateTime.Parse(dt.DateTime,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None);

        if (string.IsNullOrEmpty(dt.TimeZone) || dt.TimeZone == "UTC")
            return DateTime.SpecifyKind(parsed, DateTimeKind.Utc).ToLocalTime();

        try
        {
            var tz = TimeZoneInfo.FindSystemTimeZoneById(dt.TimeZone);
            return TimeZoneInfo.ConvertTimeToUtc(parsed, tz).ToLocalTime();
        }
        catch
        {
            return parsed; // fall back to whatever was parsed
        }
    }
}
